//! Cross-process serialization and optimistic commit checks for direct writers.

use std::fs::{self, File, Metadata, OpenOptions};
use std::io::Read;
use std::os::windows::fs::OpenOptionsExt;
use std::os::windows::io::AsRawHandle;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Mutex, MutexGuard};
use std::time::{Duration, SystemTime};

use windows::core::{HRESULT, PCWSTR};
use windows::Win32::Foundation::{
    CloseHandle, FILETIME, HANDLE, WAIT_ABANDONED, WAIT_FAILED, WAIT_OBJECT_0,
};
use windows::Win32::Storage::FileSystem::{
    GetFileInformationByHandle, BY_HANDLE_FILE_INFORMATION, FILE_SHARE_DELETE, FILE_SHARE_READ,
};
use windows::Win32::System::Threading::{
    CreateMutexW, GetCurrentProcess, GetProcessTimes, OpenProcess, ReleaseMutex,
    WaitForSingleObject, INFINITE, PROCESS_QUERY_LIMITED_INFORMATION,
};

use crate::error::WincentError;
use crate::WincentResult;

static DIRECT_DESTLIST_LOCK: Mutex<()> = Mutex::new(());
static TEMP_FILE_COUNTER: AtomicU64 = AtomicU64::new(0);

const STALE_TEMP_MIN_AGE: Duration = Duration::from_secs(60 * 60);
const ERROR_INVALID_PARAMETER_HRESULT: HRESULT = HRESULT::from_win32(87);

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct ProcessIdentity {
    pid: u32,
    creation_time: u64,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum TempOwner {
    Version1(ProcessIdentity),
    Legacy(u32),
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ProcessState {
    Missing,
    Present(u64),
    Unknown,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct FileIdentity {
    volume_serial_number: u32,
    file_index: u64,
}

pub(super) struct BackingFileSnapshot {
    identity: FileIdentity,
    hash: [u8; 16],
    bytes: Vec<u8>,
}

impl BackingFileSnapshot {
    pub(super) fn bytes(&self) -> &[u8] {
        &self.bytes
    }
}

/// Keeps the verified file identity write-locked until the caller commits.
pub(super) struct VerifiedBackingFile {
    _file: File,
}

pub(super) struct DirectWriteGuard {
    mutex_handle: HANDLE,
    _process_guard: MutexGuard<'static, ()>,
}

impl DirectWriteGuard {
    pub(super) fn acquire(path: &Path) -> WincentResult<Self> {
        let process_guard = DIRECT_DESTLIST_LOCK.lock().map_err(|_| {
            WincentError::SystemError("direct DestList write lock was poisoned".to_string())
        })?;
        let canonical_path = fs::canonicalize(path).map_err(WincentError::Io)?;
        let mutex_name = mutex_name(&canonical_path);
        let wide_name: Vec<u16> = mutex_name.encode_utf16().chain(Some(0)).collect();

        // SAFETY: The name is a valid null-terminated UTF-16 buffer and remains alive for the
        // call. The returned handle is owned by DirectWriteGuard and closed exactly once.
        let mutex_handle =
            unsafe { CreateMutexW(None, false, PCWSTR::from_raw(wide_name.as_ptr())) }.map_err(
                |error| {
                    WincentError::SystemError(format!(
                        "failed to create direct DestList mutex for {}: {error}",
                        path.display()
                    ))
                },
            )?;

        // SAFETY: mutex_handle is a live waitable mutex handle owned by this function.
        let wait_result = unsafe { WaitForSingleObject(mutex_handle, INFINITE) };
        if wait_result != WAIT_OBJECT_0 && wait_result != WAIT_ABANDONED {
            // SAFETY: mutex_handle is live and has not been transferred or closed.
            unsafe {
                let _ = CloseHandle(mutex_handle);
            }
            let detail = if wait_result == WAIT_FAILED {
                windows::core::Error::from_thread().to_string()
            } else {
                format!("unexpected wait result 0x{:08x}", wait_result.0)
            };
            return Err(WincentError::SystemError(format!(
                "failed to acquire direct DestList mutex for {}: {detail}",
                path.display()
            )));
        }

        let guard = Self {
            mutex_handle,
            _process_guard: process_guard,
        };
        cleanup_stale_replacement_files(&canonical_path)?;
        Ok(guard)
    }

    /// Reads a stable snapshot after both the process and named mutexes are held.
    pub(super) fn read_snapshot(&self, path: &Path) -> WincentResult<BackingFileSnapshot> {
        read_snapshot(path).map(|(snapshot, _file)| snapshot)
    }

    /// Reopens and verifies the path immediately before commit.
    ///
    /// The returned handle denies writers but allows `ReplaceFileW`; keep it alive until an
    /// atomic replacement finishes.
    pub(super) fn verify_unchanged(
        &self,
        path: &Path,
        expected: &BackingFileSnapshot,
    ) -> WincentResult<VerifiedBackingFile> {
        let (current, file) = read_snapshot(path).map_err(|error| match error {
            WincentError::Io(ref io_error) if matches!(io_error.raw_os_error(), Some(32 | 33)) => {
                WincentError::DestListConcurrentModification(path.to_path_buf())
            }
            other => other,
        })?;
        if current.identity != expected.identity
            || current.hash != expected.hash
            || current.bytes != expected.bytes
        {
            return Err(WincentError::DestListConcurrentModification(
                path.to_path_buf(),
            ));
        }
        Ok(VerifiedBackingFile { _file: file })
    }
}

pub(super) fn temporary_replacement_path(path: &Path) -> WincentResult<PathBuf> {
    let parent = path.parent().ok_or_else(|| {
        WincentError::InvalidArgument(format!(
            "DestList path has no parent directory: {}",
            path.display()
        ))
    })?;
    let name = path.file_name().ok_or_else(|| {
        WincentError::InvalidArgument(format!(
            "DestList path has no file name: {}",
            path.display()
        ))
    })?;
    let identity = current_process_identity()?;
    let counter = TEMP_FILE_COUNTER.fetch_add(1, Ordering::Relaxed);
    Ok(parent.join(format!(
        ".{}.wincent-v1-{}-{:016x}-{counter:016x}.tmp",
        name.to_string_lossy(),
        identity.pid,
        identity.creation_time
    )))
}

fn cleanup_stale_replacement_files(path: &Path) -> WincentResult<()> {
    cleanup_stale_replacement_files_at(path, SystemTime::now())
}

fn cleanup_stale_replacement_files_at(path: &Path, now: SystemTime) -> WincentResult<()> {
    let parent = path.parent().ok_or_else(|| {
        WincentError::InvalidArgument(format!(
            "DestList path has no parent directory: {}",
            path.display()
        ))
    })?;
    let backing_name = path
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| {
            WincentError::InvalidArgument(format!(
                "DestList path has no Unicode file name: {}",
                path.display()
            ))
        })?;
    let entries = fs::read_dir(parent)
        .map_err(|error| cleanup_io_error("enumerate temporary replacements in", parent, error))?;

    for entry in entries {
        let entry = entry.map_err(|error| {
            cleanup_io_error("read a temporary replacement entry in", parent, error)
        })?;
        let Some(candidate_name) = entry.file_name().to_str().map(str::to_owned) else {
            continue;
        };
        let Some(owner) = parse_temp_candidate(backing_name, &candidate_name) else {
            continue;
        };
        let candidate_path = entry.path();
        let file_type = entry.file_type().map_err(|error| {
            cleanup_io_error(
                "read temporary replacement type for",
                &candidate_path,
                error,
            )
        })?;
        if !file_type.is_file() {
            continue;
        }
        let metadata = entry.metadata().map_err(|error| {
            cleanup_io_error(
                "read temporary replacement metadata for",
                &candidate_path,
                error,
            )
        })?;
        if !candidate_is_old(&candidate_path, &metadata, now, STALE_TEMP_MIN_AGE)?
            || !owner_is_definitely_inactive(owner)
        {
            continue;
        }
        match fs::remove_file(&candidate_path) {
            Ok(()) => {}
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {}
            Err(error) => {
                return Err(cleanup_io_error(
                    "remove stale temporary replacement",
                    &candidate_path,
                    error,
                ));
            }
        }
    }
    Ok(())
}

fn parse_temp_candidate(backing_name: &str, candidate_name: &str) -> Option<TempOwner> {
    let prefix = format!(".{backing_name}.wincent-");
    let body = candidate_name.strip_prefix(&prefix)?.strip_suffix(".tmp")?;
    if let Some(versioned) = body.strip_prefix("v1-") {
        let mut fields = versioned.split('-');
        let pid = parse_pid(fields.next()?)?;
        let creation_time = parse_fixed_hex_u64(fields.next()?)?;
        parse_fixed_hex_u64(fields.next()?)?;
        if fields.next().is_some() {
            return None;
        }
        return Some(TempOwner::Version1(ProcessIdentity { pid, creation_time }));
    }

    let mut fields = body.split('-');
    let pid = parse_pid(fields.next()?)?;
    parse_decimal_u64(fields.next()?)?;
    if fields.next().is_some() {
        return None;
    }
    Some(TempOwner::Legacy(pid))
}

fn parse_pid(value: &str) -> Option<u32> {
    let pid = parse_decimal_u64(value).and_then(|pid| u32::try_from(pid).ok())?;
    (pid != 0).then_some(pid)
}

fn parse_decimal_u64(value: &str) -> Option<u64> {
    if value.is_empty()
        || (value.len() > 1 && value.starts_with('0'))
        || !value.bytes().all(|byte| byte.is_ascii_digit())
    {
        return None;
    }
    value.parse().ok()
}

fn parse_fixed_hex_u64(value: &str) -> Option<u64> {
    if value.len() != 16
        || !value
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
    {
        return None;
    }
    u64::from_str_radix(value, 16).ok()
}

fn candidate_is_old(
    path: &Path,
    metadata: &Metadata,
    now: SystemTime,
    minimum_age: Duration,
) -> WincentResult<bool> {
    let modified = metadata.modified().map_err(|error| {
        cleanup_io_error(
            "read temporary replacement modification time for",
            path,
            error,
        )
    })?;
    Ok(now
        .duration_since(modified)
        .is_ok_and(|age| age >= minimum_age))
}

fn owner_is_definitely_inactive(owner: TempOwner) -> bool {
    match (owner, query_process_state(owner_pid(owner))) {
        (_, ProcessState::Missing) => true,
        (TempOwner::Version1(identity), ProcessState::Present(creation_time)) => {
            identity.creation_time != creation_time
        }
        (TempOwner::Legacy(_), ProcessState::Present(_)) | (_, ProcessState::Unknown) => false,
    }
}

const fn owner_pid(owner: TempOwner) -> u32 {
    match owner {
        TempOwner::Version1(identity) => identity.pid,
        TempOwner::Legacy(pid) => pid,
    }
}

fn current_process_identity() -> WincentResult<ProcessIdentity> {
    // SAFETY: GetCurrentProcess returns a pseudo-handle valid in this process. It must not be
    // closed and is used only for the duration of GetProcessTimes.
    let handle = unsafe { GetCurrentProcess() };
    let creation_time = process_creation_time(handle).map_err(|error| {
        WincentError::SystemError(format!("failed to query current process identity: {error}"))
    })?;
    Ok(ProcessIdentity {
        pid: std::process::id(),
        creation_time,
    })
}

fn query_process_state(pid: u32) -> ProcessState {
    // SAFETY: The requested access is query-only, handle inheritance is disabled, and any
    // returned owned handle is closed below.
    let handle = match unsafe { OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) } {
        Ok(handle) => handle,
        Err(error) if error.code() == ERROR_INVALID_PARAMETER_HRESULT => {
            return ProcessState::Missing;
        }
        Err(_) => return ProcessState::Unknown,
    };
    let result = process_creation_time(handle);
    // SAFETY: handle was returned by OpenProcess and is closed exactly once here.
    unsafe {
        let _ = CloseHandle(handle);
    }
    match result {
        Ok(creation_time) => ProcessState::Present(creation_time),
        Err(_) => ProcessState::Unknown,
    }
}

fn process_creation_time(handle: HANDLE) -> windows::core::Result<u64> {
    let mut creation_time = FILETIME::default();
    let mut exit_time = FILETIME::default();
    let mut kernel_time = FILETIME::default();
    let mut user_time = FILETIME::default();
    // SAFETY: handle has process query rights. All four pointers refer to initialized writable
    // FILETIME values that remain valid for the duration of the call.
    unsafe {
        GetProcessTimes(
            handle,
            &mut creation_time,
            &mut exit_time,
            &mut kernel_time,
            &mut user_time,
        )
    }?;
    Ok((u64::from(creation_time.dwHighDateTime) << 32) | u64::from(creation_time.dwLowDateTime))
}

fn cleanup_io_error(action: &str, path: &Path, error: std::io::Error) -> WincentError {
    WincentError::SystemError(format!("failed to {action} {}: {error}", path.display()))
}

impl Drop for DirectWriteGuard {
    fn drop(&mut self) {
        // SAFETY: This guard owns the acquired mutex and its handle. ReleaseMutex is called only
        // after a successful wait, then the handle is closed exactly once.
        unsafe {
            let _ = ReleaseMutex(self.mutex_handle);
            let _ = CloseHandle(self.mutex_handle);
        }
    }
}

fn mutex_name(path: &Path) -> String {
    let normalized = path.as_os_str().to_string_lossy().to_lowercase();
    format!(
        "Local\\Wincent.DestList.Write.{:x}",
        md5::compute(normalized.as_bytes())
    )
}

fn read_snapshot(path: &Path) -> WincentResult<(BackingFileSnapshot, File)> {
    let mut file = OpenOptions::new()
        .read(true)
        .share_mode(FILE_SHARE_READ.0 | FILE_SHARE_DELETE.0)
        .open(path)
        .map_err(WincentError::Io)?;
    let identity = file_identity(&file)?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes).map_err(WincentError::Io)?;
    let hash = md5::compute(&bytes).into();
    Ok((
        BackingFileSnapshot {
            identity,
            hash,
            bytes,
        },
        file,
    ))
}

fn file_identity(file: &File) -> WincentResult<FileIdentity> {
    let mut info = BY_HANDLE_FILE_INFORMATION::default();
    // SAFETY: as_raw_handle returns the live handle borrowed from file. info points to writable,
    // properly aligned storage and remains valid for the duration of the call.
    unsafe { GetFileInformationByHandle(HANDLE(file.as_raw_handle()), &mut info) }.map_err(
        |error| {
            WincentError::SystemError(format!("failed to read DestList file identity: {error}"))
        },
    )?;
    Ok(FileIdentity {
        volume_serial_number: info.dwVolumeSerialNumber,
        file_index: (u64::from(info.nFileIndexHigh) << 32) | u64::from(info.nFileIndexLow),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use filetime::{set_file_mtime, FileTime};
    use std::path::PathBuf;
    use std::process::{Command, Stdio};
    use std::time::{Duration, Instant, SystemTime};

    const CHILD_PATH_ENV: &str = "WINCENT_DESTLIST_MUTEX_TEST_PATH";
    const CHILD_READY_ENV: &str = "WINCENT_DESTLIST_MUTEX_TEST_READY";
    const CHILD_ACQUIRED_ENV: &str = "WINCENT_DESTLIST_MUTEX_TEST_ACQUIRED";

    #[test]
    fn named_mutex_serializes_another_process() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let path = temp_dir.path().join("backing.bin");
        let ready = temp_dir.path().join("child-ready");
        let acquired = temp_dir.path().join("child-acquired");
        fs::write(&path, b"stable").map_err(WincentError::Io)?;
        let guard = DirectWriteGuard::acquire(&path)?;

        let mut child = Command::new(std::env::current_exe().map_err(WincentError::Io)?)
            .args([
                "--exact",
                "destlist::transaction::tests::named_mutex_child",
                "--nocapture",
            ])
            .env(CHILD_PATH_ENV, &path)
            .env(CHILD_READY_ENV, &ready)
            .env(CHILD_ACQUIRED_ENV, &acquired)
            .stdout(Stdio::null())
            .spawn()
            .map_err(WincentError::Io)?;

        wait_for_path(&ready, Duration::from_secs(5))?;
        std::thread::sleep(Duration::from_millis(100));
        assert!(!acquired.exists());
        drop(guard);

        let deadline = Instant::now() + Duration::from_secs(5);
        let status = loop {
            if let Some(status) = child.try_wait().map_err(WincentError::Io)? {
                break status;
            }
            if Instant::now() >= deadline {
                child.kill().map_err(WincentError::Io)?;
                return Err(WincentError::Timeout(
                    "child did not acquire the named DestList mutex".to_string(),
                ));
            }
            std::thread::sleep(Duration::from_millis(10));
        };
        assert!(status.success());
        assert!(acquired.exists());
        Ok(())
    }

    #[test]
    fn named_mutex_child() -> WincentResult<()> {
        let Some(path) = std::env::var_os(CHILD_PATH_ENV).map(PathBuf::from) else {
            return Ok(());
        };
        let ready = std::env::var_os(CHILD_READY_ENV)
            .map(PathBuf::from)
            .ok_or_else(|| WincentError::InvalidArgument("missing child ready path".to_string()))?;
        let acquired = std::env::var_os(CHILD_ACQUIRED_ENV)
            .map(PathBuf::from)
            .ok_or_else(|| {
                WincentError::InvalidArgument("missing child acquired path".to_string())
            })?;
        fs::write(ready, b"ready").map_err(WincentError::Io)?;
        let _guard = DirectWriteGuard::acquire(&path)?;
        fs::write(acquired, b"acquired").map_err(WincentError::Io)?;
        Ok(())
    }

    #[test]
    fn snapshot_rejects_same_file_with_changed_content() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let path = temp_dir.path().join("backing.bin");
        fs::write(&path, b"before").map_err(WincentError::Io)?;
        let guard = DirectWriteGuard::acquire(&path)?;
        let snapshot = guard.read_snapshot(&path)?;

        fs::write(&path, b"after").map_err(WincentError::Io)?;

        assert!(matches!(
            guard.verify_unchanged(&path, &snapshot),
            Err(WincentError::DestListConcurrentModification(changed)) if changed == path
        ));
        Ok(())
    }

    #[test]
    fn snapshot_rejects_replacement_with_identical_content() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let path = temp_dir.path().join("backing.bin");
        let replacement = temp_dir.path().join("replacement.bin");
        fs::write(&path, b"same bytes").map_err(WincentError::Io)?;
        let guard = DirectWriteGuard::acquire(&path)?;
        let snapshot = guard.read_snapshot(&path)?;

        fs::write(&replacement, b"same bytes").map_err(WincentError::Io)?;
        fs::remove_file(&path).map_err(WincentError::Io)?;
        fs::rename(&replacement, &path).map_err(WincentError::Io)?;

        assert!(matches!(
            guard.verify_unchanged(&path, &snapshot),
            Err(WincentError::DestListConcurrentModification(changed)) if changed == path
        ));
        Ok(())
    }

    #[test]
    fn atomic_replace_succeeds_after_snapshot_verification() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let path = temp_dir.path().join("backing.bin");
        fs::write(&path, b"before").map_err(WincentError::Io)?;
        let guard = DirectWriteGuard::acquire(&path)?;
        let snapshot = guard.read_snapshot(&path)?;

        super::super::remove::replace_file_atomically(&path, b"after", &guard, &snapshot)?;

        assert_eq!(fs::read(path).map_err(WincentError::Io)?, b"after");
        Ok(())
    }

    #[test]
    fn atomic_replace_preserves_concurrent_content() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let path = temp_dir.path().join("backing.bin");
        fs::write(&path, b"before").map_err(WincentError::Io)?;
        let guard = DirectWriteGuard::acquire(&path)?;
        let snapshot = guard.read_snapshot(&path)?;
        fs::write(&path, b"external change").map_err(WincentError::Io)?;

        assert!(matches!(
            super::super::remove::replace_file_atomically(
                &path,
                b"stale replacement",
                &guard,
                &snapshot,
            ),
            Err(WincentError::DestListConcurrentModification(changed)) if changed == path
        ));
        assert_eq!(
            fs::read(path).map_err(WincentError::Io)?,
            b"external change"
        );
        Ok(())
    }

    #[test]
    fn replacement_path_uses_versioned_process_identity() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let path = temp_dir.path().join("backing.bin");
        let identity = current_process_identity()?;

        let first = temporary_replacement_path(&path)?;
        let second = temporary_replacement_path(&path)?;

        assert_eq!(first.parent(), Some(temp_dir.path()));
        assert_ne!(first, second);
        let first_name = first
            .file_name()
            .and_then(|name| name.to_str())
            .expect("generated name is Unicode");
        assert_eq!(
            parse_temp_candidate("backing.bin", first_name),
            Some(TempOwner::Version1(identity))
        );
        Ok(())
    }

    #[test]
    fn temporary_candidate_parser_is_strict() {
        let valid = ".backing.bin.wincent-v1-123-0123456789abcdef-fedcba9876543210.tmp";
        assert_eq!(
            parse_temp_candidate("backing.bin", valid),
            Some(TempOwner::Version1(ProcessIdentity {
                pid: 123,
                creation_time: 0x0123_4567_89ab_cdef,
            }))
        );
        assert_eq!(
            parse_temp_candidate("backing.bin", ".backing.bin.wincent-123-45.tmp"),
            Some(TempOwner::Legacy(123))
        );

        for invalid in [
            "backing.bin.wincent-v1-123-0123456789abcdef-fedcba9876543210.tmp",
            ".other.bin.wincent-v1-123-0123456789abcdef-fedcba9876543210.tmp",
            ".backing.bin.wincent-v2-123-0123456789abcdef-fedcba9876543210.tmp",
            ".backing.bin.wincent-v1-x-0123456789abcdef-fedcba9876543210.tmp",
            ".backing.bin.wincent-v1-0-0123456789abcdef-fedcba9876543210.tmp",
            ".backing.bin.wincent-v1-0123-0123456789abcdef-fedcba9876543210.tmp",
            ".backing.bin.wincent-v1-123-1234-fedcba9876543210.tmp",
            ".backing.bin.wincent-v1-123-0123456789ABCDEF-fedcba9876543210.tmp",
            ".backing.bin.wincent-v1-123-0123456789abcdef-fedcba9876543210-extra.tmp",
            ".backing.bin.wincent-v1-123-0123456789abcdef-fedcba9876543210",
            ".backing.bin.wincent-123.tmp",
            ".backing.bin.wincent-123-045.tmp",
            ".backing.bin.wincent-123-4x.tmp",
            ".backing.bin.wincent-123-45-extra.tmp",
        ] {
            assert_eq!(
                parse_temp_candidate("backing.bin", invalid),
                None,
                "{invalid}"
            );
        }
    }

    #[test]
    fn cleanup_removes_only_old_candidates_with_inactive_owners() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let path = temp_dir.path().join("backing.bin");
        fs::write(&path, b"backing").map_err(WincentError::Io)?;
        let identity = current_process_identity()?;
        let now = SystemTime::now();
        let old = now - STALE_TEMP_MIN_AGE - Duration::from_secs(1);

        let reused_pid = temp_dir.path().join(format!(
            ".backing.bin.wincent-v1-{}-{:016x}-0000000000000001.tmp",
            identity.pid,
            identity.creation_time ^ 1
        ));
        let live_owner = temp_dir.path().join(format!(
            ".backing.bin.wincent-v1-{}-{:016x}-0000000000000002.tmp",
            identity.pid, identity.creation_time
        ));
        let live_legacy = temp_dir
            .path()
            .join(format!(".backing.bin.wincent-{}-3.tmp", identity.pid));
        let fresh_inactive = temp_dir.path().join(format!(
            ".backing.bin.wincent-v1-{}-{:016x}-0000000000000004.tmp",
            identity.pid,
            identity.creation_time ^ 1
        ));
        let other_backing = temp_dir.path().join(format!(
            ".other.bin.wincent-v1-{}-{:016x}-0000000000000005.tmp",
            identity.pid,
            identity.creation_time ^ 1
        ));
        let similar_name = temp_dir.path().join(format!(
            ".backing.bin.wincent-v1-{}-{:016x}-5.tmp",
            identity.pid,
            identity.creation_time ^ 1
        ));

        for candidate in [
            &reused_pid,
            &live_owner,
            &live_legacy,
            &fresh_inactive,
            &other_backing,
            &similar_name,
        ] {
            fs::write(candidate, b"temporary").map_err(WincentError::Io)?;
        }
        for candidate in [
            &reused_pid,
            &live_owner,
            &live_legacy,
            &other_backing,
            &similar_name,
        ] {
            set_file_mtime(candidate, FileTime::from_system_time(old)).map_err(WincentError::Io)?;
        }

        cleanup_stale_replacement_files_at(&path, now)?;

        assert!(!reused_pid.exists());
        assert!(live_owner.exists());
        assert!(live_legacy.exists());
        assert!(fresh_inactive.exists());
        assert!(other_backing.exists());
        assert!(similar_name.exists());
        Ok(())
    }

    #[test]
    fn cleanup_removes_old_legacy_candidate_for_missing_pid() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let path = temp_dir.path().join("backing.bin");
        fs::write(&path, b"backing").map_err(WincentError::Io)?;
        let candidate = temp_dir
            .path()
            .join(format!(".backing.bin.wincent-{}-0.tmp", u32::MAX));
        fs::write(&candidate, b"temporary").map_err(WincentError::Io)?;
        let now = SystemTime::now();
        set_file_mtime(
            &candidate,
            FileTime::from_system_time(now - STALE_TEMP_MIN_AGE - Duration::from_secs(1)),
        )
        .map_err(WincentError::Io)?;

        cleanup_stale_replacement_files_at(&path, now)?;

        assert!(!candidate.exists());
        Ok(())
    }

    #[test]
    fn cleanup_ignores_matching_directory() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let path = temp_dir.path().join("backing.bin");
        fs::write(&path, b"backing").map_err(WincentError::Io)?;
        let candidate = temp_dir
            .path()
            .join(".backing.bin.wincent-4294967295-0.tmp");
        fs::create_dir(&candidate).map_err(WincentError::Io)?;

        cleanup_stale_replacement_files_at(
            &path,
            SystemTime::now() + STALE_TEMP_MIN_AGE + Duration::from_secs(1),
        )?;

        assert!(candidate.is_dir());
        Ok(())
    }

    #[test]
    fn verified_snapshot_keeps_other_writers_out() -> WincentResult<()> {
        let temp_dir = tempfile::tempdir().map_err(WincentError::Io)?;
        let path = temp_dir.path().join("backing.bin");
        fs::write(&path, b"stable").map_err(WincentError::Io)?;
        let guard = DirectWriteGuard::acquire(&path)?;
        let snapshot = guard.read_snapshot(&path)?;
        let verified = guard.verify_unchanged(&path, &snapshot)?;

        let write_error = OpenOptions::new().write(true).open(&path).unwrap_err();
        assert_eq!(write_error.raw_os_error(), Some(32));
        drop(verified);
        OpenOptions::new()
            .write(true)
            .open(&path)
            .map_err(WincentError::Io)?;
        Ok(())
    }

    fn wait_for_path(path: &Path, timeout: Duration) -> WincentResult<()> {
        let deadline = Instant::now() + timeout;
        while !path.exists() {
            if Instant::now() >= deadline {
                return Err(WincentError::Timeout(format!(
                    "timed out waiting for {}",
                    path.display()
                )));
            }
            std::thread::sleep(Duration::from_millis(10));
        }
        Ok(())
    }
}
