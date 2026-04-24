use crate::{error::WincentError, WincentResult};
use crate::com::{classify_coinit_result, ComInitStatus};
use std::panic::AssertUnwindSafe;
use std::sync::mpsc;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};

/// Runs a closure on a dedicated STA thread, allowing COM cross-process calls to succeed.
///
/// On Windows 11, Shell COM operations like `pintohome` require cross-process calls to
/// explorer.exe. STA COM threads must pump a Windows message loop to receive marshaled
/// replies. COM's internal machinery handles this automatically when the closure runs on
/// a fresh STA thread with no other blocking waits. Without this wrapper, calling
/// `shell_folder(path)` from an arbitrary thread hangs indefinitely on Windows 11.
pub(crate) fn run_on_sta_thread<F, T>(f: F, timeout: std::time::Duration) -> WincentResult<T>
where
    F: FnOnce() -> WincentResult<T> + Send + 'static,
    T: Send + 'static,
{
    if timeout.is_zero() {
        return Err(WincentError::InvalidArgument(
            "timeout must be greater than zero".to_string(),
        ));
    }

    let (tx, rx) = mpsc::channel::<WincentResult<T>>();

    std::thread::Builder::new()
        .name("wincent-com-sta".into())
        .spawn(move || {
            let hr = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };

            match classify_coinit_result(hr) {
                ComInitStatus::Success => {
                    // Normal case: new thread first-time initialization
                }
                ComInitStatus::AlreadyInitialized => {
                    // Unexpected: new thread should not return S_FALSE
                    // But continue for robustness
                    eprintln!("Warning: CoInitializeEx returned S_FALSE on new thread");
                }
                ComInitStatus::ApartmentMismatch => {
                    let _ = tx.send(Err(WincentError::ComApartmentMismatch(
                        "Thread already initialized with incompatible COM apartment model".to_string()
                    )));
                    return;
                }
                ComInitStatus::OtherError(hr_code) => {
                    let _ = tx.send(Err(WincentError::SystemError(format!(
                        "COM init on STA thread failed: 0x{:08X}",
                        hr_code
                    ))));
                    return;
                }
            }

            let result = std::panic::catch_unwind(AssertUnwindSafe(f))
                .unwrap_or_else(|payload| {
                    let msg = payload
                        .downcast_ref::<&str>()
                        .copied()
                        .or_else(|| payload.downcast_ref::<String>().map(String::as_str))
                        .unwrap_or("unknown panic");
                    Err(WincentError::SystemError(format!(
                        "COM STA closure panicked: {}",
                        msg
                    )))
                });
            unsafe { CoUninitialize() };
            let _ = tx.send(result);
        })
        .map_err(|e| WincentError::SystemError(format!("Failed to spawn COM thread: {}", e)))?;

    rx.recv_timeout(timeout)
        .map_err(|_| {
            WincentError::SystemError(format!(
                "COM STA thread timed out or disconnected after {}s",
                timeout.as_secs()
            ))
        })?
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_run_on_sta_thread_zero_timeout_rejected() {
        // Duration::ZERO must be rejected immediately with InvalidArgument,
        // not silently converted to a near-instant recv_timeout that races.
        let result: WincentResult<()> =
            run_on_sta_thread(|| Ok(()), std::time::Duration::ZERO);
        assert!(
            matches!(result, Err(WincentError::InvalidArgument(_))),
            "Expected InvalidArgument for zero timeout, got: {:?}",
            result
        );
    }

    #[test]
    fn test_run_on_sta_thread_success() {
        // Tests the normal path: new thread, S_OK from CoInitializeEx
        let result = run_on_sta_thread(|| Ok(42), std::time::Duration::from_secs(10));
        assert_eq!(result.unwrap(), 42);
    }

    #[test]
    fn test_run_on_sta_thread_error_propagation() {
        // Tests that errors from the closure are properly propagated
        let result: WincentResult<()> = run_on_sta_thread(|| {
            Err(WincentError::InvalidPath("test".to_string()))
        }, std::time::Duration::from_secs(10));
        assert!(matches!(result, Err(WincentError::InvalidPath(_))));
    }

    #[test]
    fn test_run_on_sta_thread_multiple_calls() {
        // Verify multiple calls work correctly (no COM reference leaks)
        // Each call spawns a new thread, so each should get S_OK
        for i in 0..5 {
            let result = run_on_sta_thread(move || Ok(i), std::time::Duration::from_secs(10));
            assert_eq!(result.unwrap(), i);
        }
    }

    #[test]
    fn test_run_on_sta_thread_panic_becomes_system_error() {
        // A panicking closure must not leak COM state or produce a misleading
        // "timed out / disconnected" error — it should surface as SystemError.
        let result: WincentResult<()> =
            run_on_sta_thread(|| panic!("deliberate test panic"), std::time::Duration::from_secs(10));
        assert!(
            matches!(result, Err(WincentError::SystemError(_))),
            "Expected SystemError for panicking closure, got: {:?}",
            result
        );
    }

    // Note: The S_FALSE and RPC_E_CHANGED_MODE branches in run_on_sta_thread()
    // are defensive code paths that are difficult to trigger in tests:
    // - S_FALSE: Would require COM to already be initialized on a new thread (unlikely)
    // - RPC_E_CHANGED_MODE: Would require the new thread to have incompatible apartment (unlikely)
    // These branches are tested indirectly through the main implementation tests
    // (handle.rs, empty.rs, query.rs) which verify the same logic.
}
