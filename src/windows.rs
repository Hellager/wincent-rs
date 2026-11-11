//! Windows platform information and known-folder helpers.

use crate::{error::WincentError, WincentResult};
use std::ffi::{OsStr, OsString};
use std::mem;
use std::os::windows::ffi::{OsStrExt, OsStringExt};
use windows::Wdk::System::SystemServices::RtlGetVersion;
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::System::SystemInformation::OSVERSIONINFOW;
use windows::Win32::UI::Shell::{FOLDERID_Recent, SHGetKnownFolderPath, KNOWN_FOLDER_FLAG};
use winreg::enums::HKEY_LOCAL_MACHINE;
use winreg::RegKey;

pub(crate) fn os_str_to_wide_null(value: &OsStr) -> Vec<u16> {
    value.encode_wide().chain(std::iter::once(0)).collect()
}

/// Windows operating-system version information.
///
/// The major, minor, and build fields come from `RtlGetVersion`, which reports
/// the actual kernel version without the application-manifest compatibility
/// behavior of `GetVersionExW`. The remaining fields are optional values read
/// from the Windows `CurrentVersion` registry key.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WindowsVersion {
    major: u32,
    minor: u32,
    build: u32,
    ubr: Option<u32>,
    display_version: Option<String>,
    product_name: Option<String>,
}

impl WindowsVersion {
    /// Major version reported by Windows.
    #[must_use]
    pub fn major(&self) -> u32 {
        self.major
    }

    /// Minor version reported by Windows.
    #[must_use]
    pub fn minor(&self) -> u32 {
        self.minor
    }

    /// Build number reported by Windows.
    #[must_use]
    pub fn build(&self) -> u32 {
        self.build
    }

    /// Update Build Revision from the Windows registry, when available.
    #[must_use]
    pub fn ubr(&self) -> Option<u32> {
        self.ubr
    }

    /// Marketing release label such as `22H2`, when available.
    #[must_use]
    pub fn display_version(&self) -> Option<&str> {
        self.display_version.as_deref()
    }

    /// Product name from the Windows registry, when available.
    #[must_use]
    pub fn product_name(&self) -> Option<&str> {
        self.product_name.as_deref()
    }

    /// Returns the build and UBR in the conventional `build.ubr` form.
    #[must_use]
    pub fn full_build(&self) -> String {
        self.ubr.map_or_else(
            || self.build.to_string(),
            |ubr| format!("{}.{ubr}", self.build),
        )
    }
}

/// Returns the current Windows version.
///
/// Kernel version fields are read with `RtlGetVersion`. Optional release
/// metadata is read from:
/// `HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion`.
///
/// # Errors
///
/// Returns [`WincentError::SystemError`] if Windows cannot provide the kernel
/// version through `RtlGetVersion`. Registry metadata is optional and does not
/// make an otherwise valid version query fail.
pub fn get_windows_version() -> WincentResult<WindowsVersion> {
    let mut os_version = OSVERSIONINFOW {
        dwOSVersionInfoSize: mem::size_of::<OSVERSIONINFOW>() as u32,
        ..Default::default()
    };

    // SAFETY: `os_version` is a writable, correctly sized OSVERSIONINFOW buffer.
    let status = unsafe { RtlGetVersion(&mut os_version) };
    if status.0 < 0 {
        return Err(WincentError::SystemError(format!(
            "RtlGetVersion failed with NTSTATUS 0x{:08x}",
            status.0 as u32
        )));
    }

    let current_version = RegKey::predef(HKEY_LOCAL_MACHINE)
        .open_subkey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion")
        .ok();
    let ubr = current_version
        .as_ref()
        .and_then(|key| key.get_value::<u32, _>("UBR").ok());
    let display_version = current_version
        .as_ref()
        .and_then(|key| key.get_value::<String, _>("DisplayVersion").ok());
    let product_name = current_version
        .as_ref()
        .and_then(|key| key.get_value::<String, _>("ProductName").ok());

    Ok(WindowsVersion {
        major: os_version.dwMajorVersion,
        minor: os_version.dwMinorVersion,
        build: os_version.dwBuildNumber,
        ubr,
        display_version,
        product_name,
    })
}

/// Returns the path to the current user's Windows Recent folder.
pub(crate) fn get_windows_recent_folder() -> WincentResult<String> {
    // SAFETY: SHGetKnownFolderPath allocates a null-terminated path for FOLDERID_Recent.
    let result = unsafe { SHGetKnownFolderPath(&FOLDERID_Recent, KNOWN_FOLDER_FLAG(0), None) }?;
    // SAFETY: The Shell allocation is copied before being released with CoTaskMemFree.
    unsafe {
        let path = OsString::from_wide(result.as_wide());
        CoTaskMemFree(Some(result.as_ptr().cast()));
        path.into_string()
            .map_err(|_| WincentError::SystemError("Invalid UTF-16".to_string()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn recent_folder_exists() -> WincentResult<()> {
        let recent_folder = get_windows_recent_folder()?;
        assert!(!recent_folder.is_empty());
        assert!(std::path::Path::new(&recent_folder).exists());
        Ok(())
    }

    #[test]
    fn windows_version_has_a_valid_kernel_build() -> WincentResult<()> {
        let version = get_windows_version()?;
        assert!(version.major() >= 6);
        assert!(version.build() > 0);
        assert_eq!(
            version.full_build(),
            version.ubr().map_or_else(
                || version.build().to_string(),
                |ubr| format!("{}.{ubr}", version.build())
            )
        );
        Ok(())
    }
}
