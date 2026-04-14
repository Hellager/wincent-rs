//! Retry policy and strategy for handling transient errors
//!
//! Provides configurable retry mechanisms with exponential backoff and jitter
//! to improve reliability when dealing with temporary failures.
//!
//! # Key Features
//! - Exponential backoff with configurable parameters
//! - Jitter to prevent thundering herd
//! - Predefined policies for common scenarios
//! - Smart retry only for transient errors
//!
//! # Example
//! ```rust
//! use wincent::retry::RetryPolicy;
//! use std::time::Duration;
//!
//! // Use default policy (3 retries, exponential backoff)
//! let policy = RetryPolicy::default();
//!
//! // Use aggressive policy for critical operations
//! let policy = RetryPolicy::aggressive();
//!
//! // Custom policy
//! let policy = RetryPolicy {
//!     max_attempts: 5,
//!     initial_delay: Duration::from_millis(200),
//!     max_delay: Duration::from_secs(10),
//!     backoff_factor: 2.0,
//!     jitter: true,
//! };
//! ```

use std::time::Duration;

/// Retry policy configuration
///
/// Defines how retries should be performed when transient errors occur.
/// Only errors identified as transient (via `PowerShellError::is_transient()`)
/// will be retried.
#[derive(Debug, Clone)]
pub struct RetryPolicy {
    /// Maximum number of retry attempts (not including the initial attempt)
    ///
    /// For example, `max_attempts = 3` means:
    /// - 1 initial attempt
    /// - Up to 3 retry attempts
    /// - Total of 4 attempts maximum
    pub max_attempts: u32,

    /// Initial delay before the first retry
    pub initial_delay: Duration,

    /// Maximum delay between retries (caps exponential growth)
    pub max_delay: Duration,

    /// Backoff multiplier for exponential backoff
    ///
    /// Each retry delay is multiplied by this factor:
    /// - delay(n) = initial_delay * backoff_factor^n
    pub backoff_factor: f64,

    /// Enable jitter to prevent thundering herd effect
    ///
    /// When enabled, adds ±25% random variation to delays
    pub jitter: bool,
}

impl Default for RetryPolicy {
    /// Creates a standard retry policy
    ///
    /// - 3 retry attempts
    /// - 100ms initial delay
    /// - 5s maximum delay
    /// - 2x exponential backoff
    /// - Jitter enabled
    fn default() -> Self {
        Self {
            max_attempts: 3,
            initial_delay: Duration::from_millis(100),
            max_delay: Duration::from_secs(5),
            backoff_factor: 2.0,
            jitter: true,
        }
    }
}

impl RetryPolicy {
    /// Creates a policy with no retries
    ///
    /// Use this when you want to disable retry behavior entirely.
    ///
    /// # Example
    /// ```rust
    /// use wincent::retry::RetryPolicy;
    ///
    /// let policy = RetryPolicy::no_retry();
    /// assert_eq!(policy.max_attempts, 0);
    /// ```
    pub fn no_retry() -> Self {
        Self {
            max_attempts: 0,
            initial_delay: Duration::from_millis(100),
            max_delay: Duration::from_secs(5),
            backoff_factor: 2.0,
            jitter: true,
        }
    }

    /// Creates a fast retry policy for lightweight operations
    ///
    /// - 2 retry attempts
    /// - 50ms initial delay
    /// - 1s maximum delay
    /// - 1.5x backoff
    /// - Jitter enabled
    ///
    /// # Example
    /// ```rust
    /// use wincent::retry::RetryPolicy;
    ///
    /// let policy = RetryPolicy::fast();
    /// assert_eq!(policy.max_attempts, 2);
    /// ```
    pub fn fast() -> Self {
        Self {
            max_attempts: 2,
            initial_delay: Duration::from_millis(50),
            max_delay: Duration::from_secs(1),
            backoff_factor: 1.5,
            jitter: true,
        }
    }

    /// Creates a standard retry policy (same as default)
    ///
    /// # Example
    /// ```rust
    /// use wincent::retry::RetryPolicy;
    ///
    /// let policy = RetryPolicy::standard();
    /// assert_eq!(policy.max_attempts, 3);
    /// ```
    pub fn standard() -> Self {
        Self::default()
    }

    /// Creates an aggressive retry policy for critical operations
    ///
    /// - 5 retry attempts
    /// - 200ms initial delay
    /// - 10s maximum delay
    /// - 2x exponential backoff
    /// - Jitter enabled
    ///
    /// # Example
    /// ```rust
    /// use wincent::retry::RetryPolicy;
    ///
    /// let policy = RetryPolicy::aggressive();
    /// assert_eq!(policy.max_attempts, 5);
    /// ```
    pub fn aggressive() -> Self {
        Self {
            max_attempts: 5,
            initial_delay: Duration::from_millis(200),
            max_delay: Duration::from_secs(10),
            backoff_factor: 2.0,
            jitter: true,
        }
    }

    /// Calculates the delay for the nth retry attempt
    ///
    /// Uses exponential backoff: delay(n) = min(initial_delay * backoff_factor^n, max_delay)
    /// If jitter is enabled, adds ±25% random variation.
    ///
    /// # Arguments
    /// * `attempt` - The retry attempt number (0-indexed)
    ///
    /// # Example
    /// ```rust
    /// use wincent::retry::RetryPolicy;
    ///
    /// let policy = RetryPolicy::default();
    /// let delay1 = policy.calculate_delay(0);
    /// let delay2 = policy.calculate_delay(1);
    /// // delay2 should be roughly 2x delay1 (with jitter variation)
    /// ```
    pub fn calculate_delay(&self, attempt: u32) -> Duration {
        // Calculate base delay with exponential backoff
        let base_delay = self.initial_delay.as_secs_f64()
            * self.backoff_factor.powi(attempt as i32);

        // Cap at max_delay
        let delay = base_delay.min(self.max_delay.as_secs_f64());

        // Apply jitter if enabled
        let final_delay = if self.jitter {
            use rand::Rng;
            let mut rng = rand::thread_rng();
            // Add ±25% random variation
            let jitter_factor = rng.gen_range(0.75..=1.25);
            delay * jitter_factor
        } else {
            delay
        };

        Duration::from_secs_f64(final_delay)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_policy() {
        let policy = RetryPolicy::default();
        assert_eq!(policy.max_attempts, 3);
        assert_eq!(policy.initial_delay, Duration::from_millis(100));
        assert_eq!(policy.max_delay, Duration::from_secs(5));
        assert_eq!(policy.backoff_factor, 2.0);
        assert!(policy.jitter);
    }

    #[test]
    fn test_no_retry_policy() {
        let policy = RetryPolicy::no_retry();
        assert_eq!(policy.max_attempts, 0);
    }

    #[test]
    fn test_fast_policy() {
        let policy = RetryPolicy::fast();
        assert_eq!(policy.max_attempts, 2);
        assert_eq!(policy.initial_delay, Duration::from_millis(50));
        assert_eq!(policy.max_delay, Duration::from_secs(1));
        assert_eq!(policy.backoff_factor, 1.5);
    }

    #[test]
    fn test_aggressive_policy() {
        let policy = RetryPolicy::aggressive();
        assert_eq!(policy.max_attempts, 5);
        assert_eq!(policy.initial_delay, Duration::from_millis(200));
        assert_eq!(policy.max_delay, Duration::from_secs(10));
    }

    #[test]
    fn test_exponential_backoff_without_jitter() {
        let policy = RetryPolicy {
            max_attempts: 3,
            initial_delay: Duration::from_millis(100),
            max_delay: Duration::from_secs(10),
            backoff_factor: 2.0,
            jitter: false,
        };

        let delay0 = policy.calculate_delay(0);
        let delay1 = policy.calculate_delay(1);
        let delay2 = policy.calculate_delay(2);

        assert_eq!(delay0, Duration::from_millis(100));
        assert_eq!(delay1, Duration::from_millis(200));
        assert_eq!(delay2, Duration::from_millis(400));
    }

    #[test]
    fn test_max_delay_cap() {
        let policy = RetryPolicy {
            max_attempts: 10,
            initial_delay: Duration::from_millis(100),
            max_delay: Duration::from_secs(1),
            backoff_factor: 2.0,
            jitter: false,
        };

        // After several attempts, delay should be capped at max_delay
        let delay10 = policy.calculate_delay(10);
        assert_eq!(delay10, Duration::from_secs(1));
    }

    #[test]
    fn test_jitter_variation() {
        let policy = RetryPolicy {
            max_attempts: 3,
            initial_delay: Duration::from_millis(100),
            max_delay: Duration::from_secs(10),
            backoff_factor: 2.0,
            jitter: true,
        };

        // Generate multiple delays and verify they vary
        let delays: Vec<Duration> = (0..10)
            .map(|_| policy.calculate_delay(1))
            .collect();

        // With jitter, not all delays should be identical
        let all_same = delays.windows(2).all(|w| w[0] == w[1]);
        assert!(!all_same, "Jitter should cause variation in delays");

        // All delays should be within ±25% of base delay (200ms)
        for delay in delays {
            let millis = delay.as_millis();
            assert!(millis >= 150 && millis <= 250, "Delay {} out of expected range", millis);
        }
    }

    #[test]
    fn test_standard_policy() {
        let standard = RetryPolicy::standard();
        let default = RetryPolicy::default();

        assert_eq!(standard.max_attempts, default.max_attempts);
        assert_eq!(standard.initial_delay, default.initial_delay);
        assert_eq!(standard.max_delay, default.max_delay);
        assert_eq!(standard.backoff_factor, default.backoff_factor);
        assert_eq!(standard.jitter, default.jitter);
    }
}
