//! COM initialization infrastructure
//!
//! Provides unified COM initialization classification logic and RAII lifetime management.

use windows::core::HRESULT;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};

/// Raw result classification for COM initialization
#[derive(Debug, PartialEq)]
pub(crate) enum ComInitStatus {
    /// S_OK (0): COM initialized successfully for the first time on this thread
    Success,
    /// S_FALSE (1): COM was already initialized on the current thread
    AlreadyInitialized,
    /// RPC_E_CHANGED_MODE: thread apartment model conflicts with existing initialization
    ApartmentMismatch,
    /// Any other unexpected HRESULT
    OtherError(i32),
}

/// Classifies a CoInitializeEx HRESULT into a ComInitStatus (single source of truth)
pub(crate) fn classify_coinit_result(hr: HRESULT) -> ComInitStatus {
    const S_OK: i32 = 0;
    const S_FALSE: i32 = 1;
    const RPC_E_CHANGED_MODE: i32 = -2147417850_i32;

    match hr.0 {
        S_OK => ComInitStatus::Success,
        S_FALSE => ComInitStatus::AlreadyInitialized,
        RPC_E_CHANGED_MODE => ComInitStatus::ApartmentMismatch,
        other => ComInitStatus::OtherError(other),
    }
}

/// RAII guard that uninitializes COM when dropped
#[derive(Debug)]
pub(crate) struct ComGuard {
    should_uninitialize: bool,
}

impl ComGuard {
    /// Attempts to initialize COM in STA mode
    ///
    /// # Returns
    ///
    /// - `Ok(guard)`: initialization succeeded (S_OK or S_FALSE)
    /// - `Err(ComInitStatus)`: initialization failed
    pub(crate) fn try_initialize() -> Result<Self, ComInitStatus> {
        unsafe {
            let hr = CoInitializeEx(Some(std::ptr::null_mut()), COINIT_APARTMENTTHREADED);

            match classify_coinit_result(hr) {
                ComInitStatus::Success => Ok(Self {
                    should_uninitialize: true,
                }),
                // S_FALSE still represents a successful CoInitializeEx call.
                // COM requires every successful call, including S_FALSE, to be
                // balanced by CoUninitialize on the same thread.
                ComInitStatus::AlreadyInitialized => Ok(Self {
                    should_uninitialize: true,
                }),
                other => Err(other),
            }
        }
    }
}

impl Drop for ComGuard {
    fn drop(&mut self) {
        if self.should_uninitialize {
            unsafe {
                CoUninitialize();
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows::Win32::System::Com::COINIT_MULTITHREADED;

    struct TestComGuard;

    impl Drop for TestComGuard {
        fn drop(&mut self) {
            unsafe {
                CoUninitialize();
            }
        }
    }

    #[test]
    fn test_classify_s_ok() {
        assert_eq!(classify_coinit_result(HRESULT(0)), ComInitStatus::Success);
    }

    #[test]
    fn test_classify_s_false() {
        assert_eq!(
            classify_coinit_result(HRESULT(1)),
            ComInitStatus::AlreadyInitialized
        );
    }

    #[test]
    fn test_classify_apartment_mismatch() {
        assert_eq!(
            classify_coinit_result(HRESULT(-2147417850_i32)),
            ComInitStatus::ApartmentMismatch
        );
    }

    #[test]
    fn test_classify_other_error() {
        assert_eq!(
            classify_coinit_result(HRESULT(-2147024809_i32)),
            ComInitStatus::OtherError(-2147024809_i32)
        );
    }

    #[test]
    fn test_guard_initialization() {
        let result = ComGuard::try_initialize();
        assert!(result.is_ok());
    }

    #[test]
    fn test_guard_balances_already_initialized_sta() {
        unsafe {
            {
                let hr = CoInitializeEx(Some(std::ptr::null_mut()), COINIT_APARTMENTTHREADED);
                if hr.0 != 0 {
                    if hr.is_ok() {
                        CoUninitialize();
                    }
                    eprintln!(
                        "skipping test_guard_balances_already_initialized_sta: initial STA setup returned HRESULT {}",
                        hr.0
                    );
                    return;
                }
                let _initial_guard = TestComGuard;

                {
                    let guard = ComGuard::try_initialize()
                        .expect("ComGuard should accept an already initialized STA thread");
                    assert!(guard.should_uninitialize);
                }

                let hr = CoInitializeEx(Some(std::ptr::null_mut()), COINIT_MULTITHREADED);
                assert!(
                    matches!(classify_coinit_result(hr), ComInitStatus::ApartmentMismatch),
                    "dropping the S_FALSE guard must not uninitialize the original STA reference"
                );
            }

            // If the S_FALSE guard failed to call CoUninitialize, the leaked
            // STA reference would keep this thread in STA and MTA init would fail.
            let hr = CoInitializeEx(Some(std::ptr::null_mut()), COINIT_MULTITHREADED);
            assert!(
                hr.is_ok(),
                "S_FALSE guard should balance its own COM reference; got HRESULT {}",
                hr.0
            );
            let _mta_guard = TestComGuard;
        }
    }

    #[test]
    fn test_guard_apartment_mismatch() {
        unsafe {
            let hr = CoInitializeEx(Some(std::ptr::null_mut()), COINIT_MULTITHREADED);
            if hr.is_ok() {
                let _mta_guard = TestComGuard;
                let result = ComGuard::try_initialize();
                assert!(matches!(result, Err(ComInitStatus::ApartmentMismatch)));
            } else {
                eprintln!(
                    "skipping test_guard_apartment_mismatch: MTA initialization failed with HRESULT {}",
                    hr.0
                );
            }
        }
    }
}
