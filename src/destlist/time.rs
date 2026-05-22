use std::time::{Duration, SystemTime, UNIX_EPOCH};

/// Converts a Windows FILETIME value to a [`SystemTime`].
///
/// Returns `None` if the FILETIME predates the Unix epoch (January 1, 1970).
pub fn filetime_to_system_time(filetime: u64) -> Option<SystemTime> {
    // Number of 100-nanosecond intervals between 1601-01-01 and 1970-01-01.
    const FILETIME_UNIX_EPOCH: u64 = 116_444_736_000_000_000;
    if filetime < FILETIME_UNIX_EPOCH {
        return None;
    }

    let intervals = filetime - FILETIME_UNIX_EPOCH;
    let secs = intervals / 10_000_000;
    let nanos = (intervals % 10_000_000) * 100;
    Some(UNIX_EPOCH + Duration::new(secs, nanos as u32))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pre_epoch_returns_none() {
        // FILETIME 0 is 1601-01-01, well before Unix epoch
        assert!(filetime_to_system_time(0).is_none());
        // Just below the epoch offset
        assert!(filetime_to_system_time(116_444_736_000_000_000 - 1).is_none());
    }

    #[test]
    fn unix_epoch_itself() {
        // FILETIME equal to the epoch offset should give UNIX_EPOCH
        let result = filetime_to_system_time(116_444_736_000_000_000);
        assert!(result.is_some());
        assert_eq!(result.unwrap(), UNIX_EPOCH);
    }

    #[test]
    fn known_filetime_conversion() {
        // 2009-07-25 23:59:59 UTC
        // FILETIME = 128930687990000000
        // = 116444736000000000 + 12485951990000000
        // secs = 12485951990000000 / 10_000_000 = 1248595199
        // nanos = (12485951990000000 % 10_000_000) * 100 = 0
        let filetime: u64 = 128_930_687_990_000_000;
        let result = filetime_to_system_time(filetime);
        assert!(result.is_some());
        let st = result.unwrap();
        let since_epoch = st.duration_since(UNIX_EPOCH).unwrap();
        assert_eq!(since_epoch.as_secs(), 1_248_595_199);
        assert_eq!(since_epoch.subsec_nanos(), 0);
    }

    #[test]
    fn sub_second_precision() {
        // Add 5_000_000 intervals (0.5 seconds = 500_000_000 ns) to epoch
        let filetime = 116_444_736_000_000_000 + 5_000_000;
        let result = filetime_to_system_time(filetime).unwrap();
        let since_epoch = result.duration_since(UNIX_EPOCH).unwrap();
        assert_eq!(since_epoch.as_secs(), 0);
        assert_eq!(since_epoch.subsec_nanos(), 500_000_000);
    }
}
