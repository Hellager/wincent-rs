//! Polling monitor for Quick Access item changes.

use crate::manager::QuickAccessManager;
use crate::utils::paths_equal;
use crate::{QuickAccess, WincentError, WincentResult};
use std::sync::mpsc;
use std::thread::{self, JoinHandle};
use std::time::Duration;

const DEFAULT_POLL_INTERVAL: Duration = Duration::from_secs(1);

/// Options for monitoring Quick Access changes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct QuickAccessMonitorOptions {
    qa_type: QuickAccess,
    poll_interval: Duration,
}

impl Default for QuickAccessMonitorOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl QuickAccessMonitorOptions {
    /// Creates default monitor options.
    #[must_use]
    pub fn new() -> Self {
        Self {
            qa_type: QuickAccess::All,
            poll_interval: DEFAULT_POLL_INTERVAL,
        }
    }

    /// Quick Access category to monitor.
    #[must_use]
    pub fn qa_type(&self) -> QuickAccess {
        self.qa_type
    }

    /// Sets the Quick Access category to monitor.
    #[must_use]
    pub fn with_qa_type(mut self, qa_type: QuickAccess) -> Self {
        self.qa_type = qa_type;
        self
    }

    /// Poll interval used by the background monitor thread.
    #[must_use]
    pub fn poll_interval(&self) -> Duration {
        self.poll_interval
    }

    /// Sets the poll interval.
    ///
    /// # Panics
    ///
    /// Panics when `poll_interval` is zero.
    #[must_use]
    pub fn with_poll_interval(mut self, poll_interval: Duration) -> Self {
        assert!(
            !poll_interval.is_zero(),
            "poll interval must be greater than zero"
        );
        self.poll_interval = poll_interval;
        self
    }

    /// Tries to set the poll interval.
    ///
    /// # Errors
    ///
    /// Returns [`WincentError::InvalidArgument`] when `poll_interval` is zero.
    pub fn try_poll_interval(mut self, poll_interval: Duration) -> WincentResult<Self> {
        if poll_interval.is_zero() {
            return Err(WincentError::InvalidArgument(
                "poll interval must be greater than zero".to_string(),
            ));
        }
        self.poll_interval = poll_interval;
        Ok(self)
    }
}

/// Quick Access snapshot change event.
///
/// Events may represent added items, removed items, or order-only changes.
/// [`QuickAccessChangeEvent::added_items`] and
/// [`QuickAccessChangeEvent::removed_items`] report only membership changes;
/// they are both empty for a pure reorder. Use
/// [`QuickAccessChangeEvent::is_reorder`] to detect that case.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QuickAccessChangeEvent {
    qa_type: QuickAccess,
    previous_items: Vec<String>,
    current_items: Vec<String>,
    added_items: Vec<String>,
    removed_items: Vec<String>,
}

impl QuickAccessChangeEvent {
    fn new(qa_type: QuickAccess, previous_items: Vec<String>, current_items: Vec<String>) -> Self {
        let added_items = diff_new_items(&previous_items, &current_items);
        let removed_items = diff_new_items(&current_items, &previous_items);

        Self {
            qa_type,
            previous_items,
            current_items,
            added_items,
            removed_items,
        }
    }

    /// Quick Access category that was monitored.
    #[must_use]
    pub fn qa_type(&self) -> QuickAccess {
        self.qa_type
    }

    /// Previous snapshot.
    #[must_use]
    pub fn previous_items(&self) -> &[String] {
        &self.previous_items
    }

    /// Current snapshot.
    #[must_use]
    pub fn current_items(&self) -> &[String] {
        &self.current_items
    }

    /// Items present in the current snapshot but not in the previous snapshot.
    ///
    /// This does not include order changes. For a pure reorder, this slice is
    /// empty and [`QuickAccessChangeEvent::is_reorder`] returns `true`.
    #[must_use]
    pub fn added_items(&self) -> &[String] {
        &self.added_items
    }

    /// Items present in the previous snapshot but not in the current snapshot.
    ///
    /// This does not include order changes. For a pure reorder, this slice is
    /// empty and [`QuickAccessChangeEvent::is_reorder`] returns `true`.
    #[must_use]
    pub fn removed_items(&self) -> &[String] {
        &self.removed_items
    }

    /// Returns whether this event is only a change in item order.
    #[must_use]
    pub fn is_reorder(&self) -> bool {
        self.added_items.is_empty()
            && self.removed_items.is_empty()
            && !items_equal(&self.previous_items, &self.current_items)
    }
}

/// Guard for a running Quick Access monitor.
pub struct QuickAccessMonitor {
    stop_tx: Option<mpsc::Sender<()>>,
    worker: Option<JoinHandle<()>>,
}

impl QuickAccessMonitor {
    pub(crate) fn start<F>(
        manager: QuickAccessManager,
        options: QuickAccessMonitorOptions,
        mut callback: F,
    ) -> WincentResult<Self>
    where
        F: FnMut(WincentResult<QuickAccessChangeEvent>) + Send + 'static,
    {
        let mut previous = manager.get_items(options.qa_type())?;
        let (stop_tx, stop_rx) = mpsc::channel();
        let worker = thread::spawn(move || {
            while stop_rx.recv_timeout(options.poll_interval()).is_err() {
                match manager.get_items(options.qa_type()) {
                    Ok(current) => {
                        if items_equal(&previous, &current) {
                            continue;
                        }
                        let event = QuickAccessChangeEvent::new(
                            options.qa_type(),
                            previous.clone(),
                            current.clone(),
                        );
                        previous = current;
                        callback(Ok(event));
                    }
                    Err(error) => callback(Err(error)),
                }
            }
        });

        Ok(Self {
            stop_tx: Some(stop_tx),
            worker: Some(worker),
        })
    }
}

impl Drop for QuickAccessMonitor {
    fn drop(&mut self) {
        if let Some(stop_tx) = self.stop_tx.take() {
            let _ = stop_tx.send(());
        }
        if let Some(worker) = self.worker.take() {
            let _ = worker.join();
        }
    }
}

fn items_equal(left: &[String], right: &[String]) -> bool {
    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(left, right)| paths_equal(left, right))
}

fn diff_new_items(previous: &[String], current: &[String]) -> Vec<String> {
    current
        .iter()
        .filter(|item| !previous.iter().any(|previous| paths_equal(previous, item)))
        .cloned()
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::backend::QuickAccessBackend;
    use crate::utils::PathType;
    use std::collections::VecDeque;
    use std::path::{Path, PathBuf};
    use std::sync::{Arc, Mutex};

    enum FakeResponse {
        Items(Vec<String>),
        Error(&'static str),
    }

    struct FakeBackend {
        responses: Mutex<VecDeque<FakeResponse>>,
    }

    impl FakeBackend {
        fn new(responses: Vec<FakeResponse>) -> Self {
            Self {
                responses: Mutex::new(responses.into()),
            }
        }
    }

    impl QuickAccessBackend for FakeBackend {
        fn validate_path(&self, _path: &str, _expected: PathType) -> WincentResult<()> {
            Ok(())
        }

        fn get_items(
            &self,
            _qa_type: QuickAccess,
            _timeout: Duration,
        ) -> WincentResult<Vec<String>> {
            match self.responses.lock().unwrap().pop_front() {
                Some(FakeResponse::Items(items)) => Ok(items),
                Some(FakeResponse::Error(message)) => {
                    Err(WincentError::SystemError(message.to_string()))
                }
                None => Ok(Vec::new()),
            }
        }

        fn add_recent_file(&self, _path: &str, _timeout: Duration) -> WincentResult<()> {
            Ok(())
        }

        fn add_frequent_folder(&self, _path: &str, _timeout: Duration) -> WincentResult<()> {
            Ok(())
        }

        fn remove_recent_file(&self, _path: &str, _timeout: Duration) -> WincentResult<()> {
            Ok(())
        }

        fn remove_frequent_folder(&self, _path: &str, _timeout: Duration) -> WincentResult<()> {
            Ok(())
        }

        fn delete_recent_links_for_target(
            &self,
            _path: &str,
            _timeout: Duration,
        ) -> WincentResult<()> {
            Ok(())
        }

        fn list_recent_lnk_files(&self) -> WincentResult<Vec<PathBuf>> {
            Ok(Vec::new())
        }

        fn delete_lnk_file(&self, _path: &Path) -> WincentResult<()> {
            Ok(())
        }

        fn resolve_lnk_with_type(
            &self,
            _path: &Path,
            _timeout: Duration,
        ) -> WincentResult<Option<crate::recent_links::LnkResolution>> {
            Ok(None)
        }

        fn delete_recent_files_backing_data(&self) -> WincentResult<()> {
            Ok(())
        }

        fn delete_frequent_folders_backing_file(&self) -> WincentResult<()> {
            Ok(())
        }

        fn wait_for_frequent_folders_rebuild(
            &self,
            _timeout: Duration,
        ) -> WincentResult<Vec<crate::destlist::DestListEntry>> {
            Ok(Vec::new())
        }

        fn clear_recent_files(&self, _timeout: Duration) -> WincentResult<()> {
            Ok(())
        }

        fn clear_frequent_folders_jumplist(&self) -> WincentResult<()> {
            Ok(())
        }

        fn refresh_explorer(&self, _timeout: Duration) -> WincentResult<()> {
            Ok(())
        }
    }

    fn manager_with(responses: Vec<FakeResponse>) -> QuickAccessManager {
        QuickAccessManager::with_backend_for_tests(
            Duration::from_secs(1),
            Arc::new(FakeBackend::new(responses)),
        )
    }

    #[test]
    fn monitor_options_default_values() {
        let options = QuickAccessMonitorOptions::new();

        assert_eq!(options.qa_type(), QuickAccess::All);
        assert_eq!(options.poll_interval(), Duration::from_secs(1));
    }

    #[test]
    fn monitor_options_reject_zero_poll_interval() {
        let result = QuickAccessMonitorOptions::new().try_poll_interval(Duration::ZERO);

        assert!(matches!(result, Err(WincentError::InvalidArgument(_))));
    }

    #[test]
    fn diff_uses_windows_path_semantics() {
        let previous = vec!["C:\\Projects".to_string()];
        let current = vec!["c:/projects".to_string(), "C:\\New".to_string()];

        assert!(items_equal(&previous, &["c:/projects".to_string()]));
        assert_eq!(diff_new_items(&previous, &current), vec!["C:\\New"]);
    }

    #[test]
    fn change_event_reports_added_and_removed_items() {
        let event = QuickAccessChangeEvent::new(
            QuickAccess::All,
            vec!["C:\\Old".to_string(), "C:\\Same".to_string()],
            vec!["C:\\Same".to_string(), "C:\\New".to_string()],
        );

        assert_eq!(event.added_items(), &["C:\\New".to_string()]);
        assert_eq!(event.removed_items(), &["C:\\Old".to_string()]);
        assert!(!event.is_reorder());
    }

    #[test]
    fn added_item_event_is_not_reorder() {
        let event = QuickAccessChangeEvent::new(
            QuickAccess::All,
            vec!["C:\\Same".to_string()],
            vec!["C:\\Same".to_string(), "C:\\New".to_string()],
        );

        assert_eq!(event.added_items(), &["C:\\New".to_string()]);
        assert!(event.removed_items().is_empty());
        assert!(!event.is_reorder());
    }

    #[test]
    fn removed_item_event_is_not_reorder() {
        let event = QuickAccessChangeEvent::new(
            QuickAccess::All,
            vec!["C:\\Old".to_string(), "C:\\Same".to_string()],
            vec!["C:\\Same".to_string()],
        );

        assert!(event.added_items().is_empty());
        assert_eq!(event.removed_items(), &["C:\\Old".to_string()]);
        assert!(!event.is_reorder());
    }

    #[test]
    fn order_change_is_a_change_without_added_or_removed_items() {
        let previous = vec!["C:\\A".to_string(), "C:\\B".to_string()];
        let current = vec!["C:\\B".to_string(), "C:\\A".to_string()];
        let event = QuickAccessChangeEvent::new(QuickAccess::All, previous.clone(), current);

        assert!(!items_equal(&previous, event.current_items()));
        assert!(event.added_items().is_empty());
        assert!(event.removed_items().is_empty());
        assert!(event.is_reorder());
    }

    #[test]
    fn watch_returns_error_when_initial_snapshot_fails() {
        let manager = manager_with(vec![FakeResponse::Error("initial failed")]);

        let result = manager.watch_quick_access(QuickAccessMonitorOptions::new(), |_| {});

        assert!(
            matches!(result, Err(WincentError::SystemError(message)) if message == "initial failed")
        );
    }

    #[test]
    fn watch_reports_changes() -> WincentResult<()> {
        let manager = manager_with(vec![
            FakeResponse::Items(vec!["C:\\Old".to_string()]),
            FakeResponse::Items(vec!["C:\\Old".to_string(), "C:\\New".to_string()]),
        ]);
        let options =
            QuickAccessMonitorOptions::new().try_poll_interval(Duration::from_millis(10))?;
        let (tx, rx) = mpsc::channel();
        let _monitor = manager.watch_quick_access(options, move |result| {
            tx.send(result.map(|event| event.added_items().to_vec()))
                .unwrap();
        })?;

        assert_eq!(
            rx.recv_timeout(Duration::from_secs(1)).unwrap().unwrap(),
            vec!["C:\\New".to_string()]
        );
        Ok(())
    }

    #[test]
    fn watch_reports_errors_and_continues() -> WincentResult<()> {
        let manager = manager_with(vec![
            FakeResponse::Items(vec!["C:\\Old".to_string()]),
            FakeResponse::Error("poll failed"),
            FakeResponse::Items(vec!["C:\\New".to_string()]),
        ]);
        let options =
            QuickAccessMonitorOptions::new().try_poll_interval(Duration::from_millis(10))?;
        let (tx, rx) = mpsc::channel();
        let _monitor = manager.watch_quick_access(options, move |result| {
            tx.send(match result {
                Ok(event) => format!("ok:{}", event.current_items().join(",")),
                Err(error) => format!("err:{error}"),
            })
            .unwrap();
        })?;

        assert!(rx
            .recv_timeout(Duration::from_secs(1))
            .unwrap()
            .starts_with("err:"));
        assert_eq!(
            rx.recv_timeout(Duration::from_secs(1)).unwrap(),
            "ok:C:\\New"
        );
        Ok(())
    }

    #[test]
    fn dropping_monitor_stops_worker() -> WincentResult<()> {
        let manager = manager_with(vec![FakeResponse::Items(Vec::new())]);
        let options =
            QuickAccessMonitorOptions::new().try_poll_interval(Duration::from_secs(60))?;
        let monitor = manager.watch_quick_access(options, |_| {})?;

        drop(monitor);
        Ok(())
    }
}
