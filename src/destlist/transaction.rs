//! Cross-process serialization and optimistic commit checks for direct writers.

use std::fs::{self, File, OpenOptions};
use std::io::Read;
use std::os::windows::fs::OpenOptionsExt;
use std::os::windows::io::AsRawHandle;
use std::path::Path;
use std::sync::{Mutex, MutexGuard};

use windows::core::PCWSTR;
use windows::Win32::Foundation::{CloseHandle, HANDLE, WAIT_ABANDONED, WAIT_FAILED, WAIT_OBJECT_0};
use windows::Win32::Storage::FileSystem::{
    GetFileInformationByHandle, BY_HANDLE_FILE_INFORMATION, FILE_SHARE_DELETE, FILE_SHARE_READ,
};
use windows::Win32::System::Threading::{
    CreateMutexW, ReleaseMutex, WaitForSingleObject, INFINITE,
};

use crate::error::WincentError;
use crate::WincentResult;

static DIRECT_DESTLIST_LOCK: Mutex<()> = Mutex::new(());

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

        Ok(Self {
            mutex_handle,
            _process_guard: process_guard,
        })
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
    use std::path::PathBuf;
    use std::process::{Command, Stdio};
    use std::time::{Duration, Instant};

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
