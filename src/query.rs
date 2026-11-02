//! Native Shell COM helpers shared by Quick Access mutation operations.

use crate::{error::WincentError, WincentResult};
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER};
use windows::Win32::System::Variant::VARIANT;
use windows::Win32::UI::Shell::{Folder, FolderItem, FolderItems, IShellDispatch, Shell};

/// Shell namespace GUID for frequent folders in Windows Quick Access.
pub(crate) const FREQUENT_FOLDERS_NAMESPACE: &str =
    "shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}";

/// Extracts a non-empty file-system path from a Shell folder item.
pub(crate) fn item_path(item: &FolderItem) -> Option<String> {
    // SAFETY: Callers pass a live FolderItem interface obtained from Shell COM;
    // Path returns a COM BSTR copied into an owned Rust String before returning.
    unsafe {
        item.Path().ok().and_then(|path| {
            let value = path.to_string();
            if value.trim().is_empty() {
                None
            } else {
                Some(value)
            }
        })
    }
}

/// Retrieves the item collection from a Shell folder.
pub(crate) fn folder_items(folder: &Folder) -> WincentResult<FolderItems> {
    // SAFETY: `folder` is a live Shell Folder interface and COM is initialized
    // by the caller's STA worker.
    unsafe {
        folder.Items().map_err(|error| {
            WincentError::SystemError(format!("Failed to enumerate Quick Access items: {error}"))
        })
    }
}

fn shell_dispatch() -> WincentResult<IShellDispatch> {
    // SAFETY: Called only after COM initialization. CoCreateInstance returns an
    // interface pointer managed by the windows crate wrapper.
    unsafe {
        CoCreateInstance(&Shell, None, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER).map_err(
            |error| {
                WincentError::SystemError(format!("Failed to create Shell COM object: {error}"))
            },
        )
    }
}

/// Opens a Shell namespace by path or GUID.
pub(crate) fn shell_folder(namespace: &str) -> WincentResult<Folder> {
    let shell = shell_dispatch()?;
    let variant = VARIANT::from(namespace);

    // SAFETY: `shell` is a live IShellDispatch interface and `variant`
    // contains a namespace string that lives until NameSpace returns.
    unsafe {
        shell.NameSpace(&variant).map_err(|error| {
            WincentError::SystemError(format!(
                "Failed to open shell namespace {namespace}: {error}"
            ))
        })
    }
}
