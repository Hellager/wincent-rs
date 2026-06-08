use crate::com::{classify_coinit_result, ComInitStatus};
use crate::{error::WincentError, WincentResult};
use std::panic::AssertUnwindSafe;
use std::sync::{
    atomic::{AtomicUsize, Ordering},
    mpsc::{self, Receiver, RecvTimeoutError},
};
use std::time::Duration;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};

// Normal Quick Access use usually has at most one or two concurrent Shell COM
// operations. This cap is a guardrail against repeated timeouts accumulating
// background STA workers; it is not a throughput tuning parameter.
const MAX_ACTIVE_STA_WORKERS: usize = 4;
static ACTIVE_STA_WORKERS: AtomicUsize = AtomicUsize::new(0);

struct StaComGuard;

impl Drop for StaComGuard {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

struct ActiveStaWorkerGuard;

impl Drop for ActiveStaWorkerGuard {
    fn drop(&mut self) {
        ACTIVE_STA_WORKERS.fetch_sub(1, Ordering::AcqRel);
    }
}

fn acquire_active_sta_worker() -> WincentResult<ActiveStaWorkerGuard> {
    loop {
        let current = ACTIVE_STA_WORKERS.load(Ordering::Acquire);
        if current >= MAX_ACTIVE_STA_WORKERS {
            return Err(WincentError::Timeout(
                "Too many Shell COM operations are still running; previous timed-out operations may not have finished yet.".to_string(),
            ));
        }

        if ACTIVE_STA_WORKERS
            .compare_exchange(current, current + 1, Ordering::AcqRel, Ordering::Acquire)
            .is_ok()
        {
            return Ok(ActiveStaWorkerGuard);
        }
    }
}

fn recv_sta_result<T>(rx: Receiver<WincentResult<T>>, timeout: Duration) -> WincentResult<T> {
    match rx.recv_timeout(timeout) {
        Ok(result) => result,
        Err(RecvTimeoutError::Timeout) => Err(WincentError::Timeout(format!(
            "COM STA thread timed out after {}s; the caller stopped waiting, but the underlying Shell COM operation may still complete later and affect Explorer state.",
            timeout.as_secs_f64()
        ))),
        Err(RecvTimeoutError::Disconnected) => Err(WincentError::SystemError(
            "COM STA thread disconnected before sending a result".to_string(),
        )),
    }
}

/// Runs a closure on a dedicated STA thread, allowing COM cross-process calls to succeed.
///
/// On Windows 11, Shell COM operations like `pintohome` require cross-process calls to
/// explorer.exe. STA COM threads must pump a Windows message loop to receive marshaled
/// replies. COM's internal machinery handles this automatically when the closure runs on
/// a fresh STA thread with no other blocking waits. Without this wrapper, calling
/// `shell_folder(path)` from an arbitrary thread hangs indefinitely on Windows 11.
///
/// If the timeout elapses, this function returns while the worker thread keeps
/// running. COM is uninitialized by that worker when it exits naturally; if the
/// receiver has already gone away, sending the late result is intentionally
/// ignored.
///
/// Active workers are capped to prevent repeated timeouts from accumulating an
/// unbounded number of background STA threads. The cap is checked before a worker
/// is spawned; when it is reached this function returns immediately instead of
/// waiting for a slot. A slot is released only when the worker thread exits
/// naturally, not when the caller stops waiting.
pub(crate) fn run_on_sta_thread<F, T>(f: F, timeout: Duration) -> WincentResult<T>
where
    F: FnOnce() -> WincentResult<T> + Send + 'static,
    T: Send + 'static,
{
    if timeout.is_zero() {
        return Err(WincentError::InvalidArgument(
            "timeout must be greater than zero".to_string(),
        ));
    }

    let active_worker = acquire_active_sta_worker()?;
    let (tx, rx) = mpsc::channel::<WincentResult<T>>();

    std::thread::Builder::new()
        .name("wincent-com-sta".into())
        .spawn(move || {
            let _active_worker = active_worker;
            let hr = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };

            match classify_coinit_result(hr) {
                ComInitStatus::Success => {
                    // Normal case: new thread first-time initialization
                }
                ComInitStatus::AlreadyInitialized => {
                    // Unexpected: new thread should not return S_FALSE.
                    // Continue for robustness; S_FALSE is still a successful
                    // CoInitializeEx call and is balanced by StaComGuard below.
                }
                ComInitStatus::ApartmentMismatch => {
                    // Failed CoInitializeEx calls do not require CoUninitialize.
                    let _ = tx.send(Err(WincentError::ComApartmentMismatch(
                        "Thread already initialized with incompatible COM apartment model"
                            .to_string(),
                    )));
                    return;
                }
                ComInitStatus::OtherError(hr_code) => {
                    // Failed CoInitializeEx calls do not require CoUninitialize.
                    let _ = tx.send(Err(WincentError::SystemError(format!(
                        "COM init on STA thread failed: 0x{:08X}",
                        hr_code
                    ))));
                    return;
                }
            }

            let _com = StaComGuard;
            let result = std::panic::catch_unwind(AssertUnwindSafe(f)).unwrap_or_else(|payload| {
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
            // The receiver may have timed out and returned already. In that
            // case the worker still finishes normally and StaComGuard releases COM.
            let _ = tx.send(result);
        })
        .map_err(|e| WincentError::SystemError(format!("Failed to spawn COM thread: {}", e)))?;

    recv_sta_result(rx, timeout)
}

#[cfg(test)]
mod tests {
    use super::*;
    use serial_test::serial;
    use std::sync::{mpsc, Arc, Condvar, Mutex};
    use std::time::{Duration, Instant};

    fn wait_for_active_worker_count(expected: usize, timeout: Duration) -> bool {
        let started = Instant::now();
        while started.elapsed() < timeout {
            if ACTIVE_STA_WORKERS.load(Ordering::Acquire) == expected {
                return true;
            }
            std::thread::sleep(Duration::from_millis(5));
        }
        ACTIVE_STA_WORKERS.load(Ordering::Acquire) == expected
    }

    #[test]
    #[serial(com_thread_active_workers)]
    fn test_run_on_sta_thread_zero_timeout_rejected() {
        // Duration::ZERO must be rejected immediately with InvalidArgument,
        // not silently converted to a near-instant recv_timeout that races.
        let result: WincentResult<()> = run_on_sta_thread(|| Ok(()), std::time::Duration::ZERO);
        assert!(
            matches!(result, Err(WincentError::InvalidArgument(_))),
            "Expected InvalidArgument for zero timeout, got: {:?}",
            result
        );
    }

    #[test]
    fn test_recv_sta_result_timeout_is_distinct() {
        let (_tx, rx) = mpsc::channel::<WincentResult<()>>();
        let result = recv_sta_result(rx, std::time::Duration::from_millis(1));

        match result {
            Err(WincentError::Timeout(message)) => {
                assert!(
                    message.contains("may still complete later"),
                    "Timeout message should mention late completion risk, got: {message}"
                );
            }
            other => panic!("Expected Timeout for recv timeout, got: {:?}", other),
        }
    }

    #[test]
    fn test_recv_sta_result_disconnected_is_distinct() {
        let (tx, rx) = mpsc::channel::<WincentResult<()>>();
        drop(tx);

        let result = recv_sta_result(rx, std::time::Duration::from_secs(10));

        assert!(
            matches!(result, Err(WincentError::SystemError(_))),
            "Expected SystemError for disconnected worker, got: {:?}",
            result
        );
        let rendered = result.unwrap_err().to_string();
        assert!(
            rendered.contains("disconnected"),
            "Disconnected error should be explicit, got: {}",
            rendered
        );
    }

    #[test]
    #[serial(com_thread_active_workers)]
    fn test_run_on_sta_thread_success() {
        // Tests the normal path: new thread, S_OK from CoInitializeEx
        let result = run_on_sta_thread(|| Ok(42), std::time::Duration::from_secs(10));
        assert_eq!(result.unwrap(), 42);
    }

    #[test]
    #[serial(com_thread_active_workers)]
    fn test_run_on_sta_thread_error_propagation() {
        // Tests that errors from the closure are properly propagated
        let result: WincentResult<()> = run_on_sta_thread(
            || Err(WincentError::invalid_path_reason("test")),
            std::time::Duration::from_secs(10),
        );
        assert!(matches!(result, Err(WincentError::InvalidPath(_))));
    }

    #[test]
    #[serial(com_thread_active_workers)]
    fn test_run_on_sta_thread_multiple_calls() {
        // Verify multiple calls work correctly (no COM reference leaks)
        // Each call spawns a new thread, so each should get S_OK
        for i in 0..5 {
            let result = run_on_sta_thread(move || Ok(i), std::time::Duration::from_secs(10));
            assert_eq!(result.unwrap(), i);
        }
    }

    #[test]
    #[serial(com_thread_active_workers)]
    fn test_run_on_sta_thread_panic_becomes_system_error() {
        // A panicking closure must not leak COM state or produce a misleading
        // "timed out / disconnected" error — it should surface as SystemError.
        let result: WincentResult<()> = run_on_sta_thread(
            || panic!("deliberate test panic"),
            std::time::Duration::from_secs(10),
        );
        assert!(
            matches!(result, Err(WincentError::SystemError(_))),
            "Expected SystemError for panicking closure, got: {:?}",
            result
        );
    }

    #[test]
    #[serial(com_thread_active_workers)]
    fn test_active_worker_limit_rejects_extra_worker_and_recovers() {
        let release_pair = Arc::new((Mutex::new(false), Condvar::new()));
        let (ready_tx, ready_rx) = mpsc::channel();
        let (result_tx, result_rx) = mpsc::channel();
        let mut callers = Vec::new();

        for _ in 0..MAX_ACTIVE_STA_WORKERS {
            let release_pair = Arc::clone(&release_pair);
            let ready_tx = ready_tx.clone();
            let result_tx = result_tx.clone();
            callers.push(std::thread::spawn(move || {
                let result = run_on_sta_thread(
                    move || {
                        ready_tx.send(()).expect("test receiver should be alive");
                        let (lock, condvar) = &*release_pair;
                        let mut released = lock.lock().expect("release lock poisoned");
                        while !*released {
                            released = condvar
                                .wait(released)
                                .expect("release condvar wait poisoned");
                        }
                        Ok(())
                    },
                    std::time::Duration::from_millis(20),
                );
                result_tx
                    .send(result)
                    .expect("test result receiver should be alive");
            }));
        }
        drop(ready_tx);
        drop(result_tx);

        // Wait until every worker has entered its blocking closure. At that
        // point each active-worker slot is definitely held. The short caller
        // timeout above is intentional: callers stop waiting quickly, while the
        // worker closures remain alive until this test releases the condvar.
        for _ in 0..MAX_ACTIVE_STA_WORKERS {
            ready_rx
                .recv_timeout(std::time::Duration::from_secs(5))
                .expect("worker should enter blocking closure");
        }

        // Collect caller-side timeouts before trying the extra worker so the
        // test exercises the exact production hazard: timed-out callers with
        // still-running Shell COM worker threads.
        for _ in 0..MAX_ACTIVE_STA_WORKERS {
            let result = result_rx
                .recv_timeout(std::time::Duration::from_secs(5))
                .expect("caller should time out while worker remains blocked");
            assert!(
                matches!(result, Err(WincentError::Timeout(_))),
                "callers should have timed out while workers continued, got: {:?}",
                result
            );
        }

        let extra_ran = Arc::new(std::sync::atomic::AtomicBool::new(false));
        let extra_ran_for_worker = Arc::clone(&extra_ran);
        let extra = run_on_sta_thread(
            move || {
                extra_ran_for_worker.store(true, Ordering::SeqCst);
                Ok(())
            },
            std::time::Duration::from_secs(5),
        );

        match extra {
            Err(WincentError::Timeout(message)) => {
                assert!(
                    message.contains("Too many Shell COM operations"),
                    "limit timeout should explain active worker cap, got: {message}"
                );
            }
            other => panic!("Expected active worker limit Timeout, got: {:?}", other),
        }
        assert!(
            !extra_ran.load(Ordering::SeqCst),
            "extra closure must not run when active worker limit is reached"
        );

        {
            let (lock, condvar) = &*release_pair;
            let mut released = lock.lock().expect("release lock poisoned");
            *released = true;
            condvar.notify_all();
        }

        for caller in callers {
            caller.join().expect("caller thread should not panic");
        }
        assert!(
            wait_for_active_worker_count(0, std::time::Duration::from_secs(5)),
            "active worker count should return to zero after workers exit"
        );

        let result = run_on_sta_thread(|| Ok(()), std::time::Duration::from_secs(5));
        assert!(
            result.is_ok(),
            "active worker count should recover after workers exit: {:?}",
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
