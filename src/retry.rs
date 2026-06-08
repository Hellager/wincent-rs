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
//! use wincent::RetryPolicy;
//! use std::time::Duration;
//!
//! // Use default policy (3 retries, exponential backoff)
//! let policy = RetryPolicy::default();
//!
//! // Use aggressive policy for critical operations
//! let policy = RetryPolicy::aggressive();
//!
//! // Custom policy
//! let policy = RetryPolicy::new()
//!     .with_max_attempts(5)
//!     .with_initial_delay(Duration::from_millis(200))
//!     .with_max_delay(Duration::from_secs(10))
//!     .with_backoff_factor(2.0)
//!     .with_jitter(true);
//! ```

use crate::{WincentError, WincentResult};
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
    max_attempts: u32,

    /// Initial delay before the first retry
    initial_delay: Duration,

    /// Maximum delay between retries (caps exponential growth)
    max_delay: Duration,

    /// Backoff multiplier for exponential backoff
    ///
    /// Each retry delay is multiplied by this factor:
    /// - delay(n) = initial_delay * backoff_factor^n
    backoff_factor: f64,

    /// Enable jitter to prevent thundering herd effect
    ///
    /// When enabled, adds ±25% random variation to delays
    jitter: bool,
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
    /// Creates a standard retry policy.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Maximum number of retry attempts, not including the initial attempt.
    #[must_use]
    pub fn max_attempts(&self) -> u32 {
        self.max_attempts
    }

    /// Sets the maximum number of retry attempts.
    #[must_use]
    pub fn with_max_attempts(mut self, max_attempts: u32) -> Self {
        self.max_attempts = max_attempts;
        self
    }

    /// Initial delay before the first retry.
    #[must_use]
    pub fn initial_delay(&self) -> Duration {
        self.initial_delay
    }

    /// Sets the initial delay before the first retry.
    #[must_use]
    pub fn with_initial_delay(mut self, initial_delay: Duration) -> Self {
        self.initial_delay = initial_delay;
        self
    }

    /// Maximum delay between retries.
    #[must_use]
    pub fn max_delay(&self) -> Duration {
        self.max_delay
    }

    /// Sets the maximum delay between retries.
    #[must_use]
    pub fn with_max_delay(mut self, max_delay: Duration) -> Self {
        self.max_delay = max_delay;
        self
    }

    /// Backoff multiplier for exponential backoff.
    #[must_use]
    pub fn backoff_factor(&self) -> f64 {
        self.backoff_factor
    }

    /// Sets the backoff multiplier for exponential backoff.
    #[must_use]
    pub fn with_backoff_factor(mut self, backoff_factor: f64) -> Self {
        self.backoff_factor = backoff_factor;
        self
    }

    /// Whether jitter is enabled.
    #[must_use]
    pub fn jitter(&self) -> bool {
        self.jitter
    }

    /// Sets whether jitter is enabled.
    #[must_use]
    pub fn with_jitter(mut self, jitter: bool) -> Self {
        self.jitter = jitter;
        self
    }

    /// Validates that this retry policy has coherent final configuration.
    ///
    /// Custom policies should be validated before direct use with
    /// [`RetryPolicy::calculate_delay`]. Policies passed through
    /// [`crate::QuickAccessManagerBuilder`] are validated when the manager is
    /// built.
    ///
    /// # Errors
    ///
    /// Returns [`WincentError::InvalidArgument`] when:
    /// - `backoff_factor` is not finite or is less than `1.0`
    /// - retries are enabled but `initial_delay` is zero
    /// - `max_delay` is less than `initial_delay`
    pub fn validate(&self) -> WincentResult<()> {
        if !self.backoff_factor.is_finite() || self.backoff_factor < 1.0 {
            return Err(WincentError::InvalidArgument(
                "retry backoff_factor must be finite and greater than or equal to 1.0".to_string(),
            ));
        }

        if self.max_attempts > 0 && self.initial_delay.is_zero() {
            return Err(WincentError::InvalidArgument(
                "retry initial_delay must be greater than zero when retries are enabled"
                    .to_string(),
            ));
        }

        if self.max_delay < self.initial_delay {
            return Err(WincentError::InvalidArgument(
                "retry max_delay must be greater than or equal to initial_delay".to_string(),
            ));
        }

        Ok(())
    }

    /// Validates this retry policy and returns it unchanged on success.
    ///
    /// This is useful at the end of a `RetryPolicy::new().with_*()` chain when
    /// the caller wants a fallible finalization step.
    ///
    /// # Errors
    ///
    /// Returns [`WincentError::InvalidArgument`] for the same conditions as
    /// [`RetryPolicy::validate`].
    pub fn validated(self) -> WincentResult<Self> {
        self.validate()?;
        Ok(self)
    }

    /// Creates a policy with no retries
    ///
    /// Use this when you want to disable retry behavior entirely.
    ///
    /// # Example
    /// ```rust
    /// use wincent::RetryPolicy;
    ///
    /// let policy = RetryPolicy::no_retry();
    /// assert_eq!(policy.max_attempts(), 0);
    /// ```
    #[must_use]
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
    /// use wincent::RetryPolicy;
    ///
    /// let policy = RetryPolicy::fast();
    /// assert_eq!(policy.max_attempts(), 2);
    /// ```
    #[must_use]
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
    /// use wincent::RetryPolicy;
    ///
    /// let policy = RetryPolicy::standard();
    /// assert_eq!(policy.max_attempts(), 3);
    /// ```
    #[must_use]
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
    /// use wincent::RetryPolicy;
    ///
    /// let policy = RetryPolicy::aggressive();
    /// assert_eq!(policy.max_attempts(), 5);
    /// ```
    #[must_use]
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
    /// This method assumes the policy is valid. Use one of the predefined
    /// policies, pass the policy through [`crate::QuickAccessManagerBuilder`],
    /// or call [`RetryPolicy::validate`] / [`RetryPolicy::validated`] before
    /// directly calculating delays for a custom policy.
    ///
    /// # Arguments
    /// * `attempt` - The retry attempt number (0-indexed)
    ///
    /// # Example
    /// ```rust
    /// use wincent::RetryPolicy;
    ///
    /// let policy = RetryPolicy::default();
    /// let delay1 = policy.calculate_delay(0);
    /// let delay2 = policy.calculate_delay(1);
    /// // delay2 should be roughly 2x delay1 (with jitter variation)
    /// ```
    #[must_use]
    pub fn calculate_delay(&self, attempt: u32) -> Duration {
        // Calculate base delay with exponential backoff
        let base_delay =
            self.initial_delay.as_secs_f64() * self.backoff_factor.powi(attempt as i32);

        // Cap at max_delay
        let delay = base_delay.min(self.max_delay.as_secs_f64());

        // Apply jitter if enabled
        let final_delay = if self.jitter {
            // Add ±25% random variation
            let jitter_factor = rand::random_range(0.75..=1.25);
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
        assert_eq!(policy.max_attempts(), 3);
        assert_eq!(policy.initial_delay(), Duration::from_millis(100));
        assert_eq!(policy.max_delay(), Duration::from_secs(5));
        assert_eq!(policy.backoff_factor(), 2.0);
        assert!(policy.jitter());
    }

    #[test]
    fn test_no_retry_policy() {
        let policy = RetryPolicy::no_retry();
        assert_eq!(policy.max_attempts(), 0);
    }

    #[test]
    fn test_fast_policy() {
        let policy = RetryPolicy::fast();
        assert_eq!(policy.max_attempts(), 2);
        assert_eq!(policy.initial_delay(), Duration::from_millis(50));
        assert_eq!(policy.max_delay(), Duration::from_secs(1));
        assert_eq!(policy.backoff_factor(), 1.5);
    }

    #[test]
    fn test_aggressive_policy() {
        let policy = RetryPolicy::aggressive();
        assert_eq!(policy.max_attempts(), 5);
        assert_eq!(policy.initial_delay(), Duration::from_millis(200));
        assert_eq!(policy.max_delay(), Duration::from_secs(10));
    }

    #[test]
    fn test_exponential_backoff_without_jitter() {
        let policy = RetryPolicy::new()
            .with_max_attempts(3)
            .with_initial_delay(Duration::from_millis(100))
            .with_max_delay(Duration::from_secs(10))
            .with_backoff_factor(2.0)
            .with_jitter(false);

        let delay0 = policy.calculate_delay(0);
        let delay1 = policy.calculate_delay(1);
        let delay2 = policy.calculate_delay(2);

        assert_eq!(delay0, Duration::from_millis(100));
        assert_eq!(delay1, Duration::from_millis(200));
        assert_eq!(delay2, Duration::from_millis(400));
    }

    #[test]
    fn test_max_delay_cap() {
        let policy = RetryPolicy::new()
            .with_max_attempts(10)
            .with_initial_delay(Duration::from_millis(100))
            .with_max_delay(Duration::from_secs(1))
            .with_backoff_factor(2.0)
            .with_jitter(false);

        // After several attempts, delay should be capped at max_delay
        let delay10 = policy.calculate_delay(10);
        assert_eq!(delay10, Duration::from_secs(1));
    }

    #[test]
    fn test_jitter_variation() {
        let policy = RetryPolicy::new()
            .with_max_attempts(3)
            .with_initial_delay(Duration::from_millis(100))
            .with_max_delay(Duration::from_secs(10))
            .with_backoff_factor(2.0)
            .with_jitter(true);

        // Generate multiple delays and verify they vary
        let delays: Vec<Duration> = (0..10).map(|_| policy.calculate_delay(1)).collect();

        // With jitter, not all delays should be identical
        let all_same = delays.windows(2).all(|w| w[0] == w[1]);
        assert!(!all_same, "Jitter should cause variation in delays");

        // All delays should be within ±25% of base delay (200ms)
        for delay in delays {
            let millis = delay.as_millis();
            assert!(
                (150..=250).contains(&millis),
                "Delay {} out of expected range",
                millis
            );
        }
    }

    #[test]
    fn test_standard_policy() {
        let standard = RetryPolicy::standard();
        let default = RetryPolicy::default();

        assert_eq!(standard.max_attempts(), default.max_attempts());
        assert_eq!(standard.initial_delay(), default.initial_delay());
        assert_eq!(standard.max_delay(), default.max_delay());
        assert_eq!(standard.backoff_factor(), default.backoff_factor());
        assert_eq!(standard.jitter(), default.jitter());
    }

    #[test]
    fn predefined_policies_are_valid() {
        let policies = [
            RetryPolicy::default(),
            RetryPolicy::fast(),
            RetryPolicy::standard(),
            RetryPolicy::aggressive(),
            RetryPolicy::no_retry(),
        ];

        for policy in policies {
            policy
                .validate()
                .expect("predefined policy should be valid");
        }
    }

    #[test]
    fn validate_rejects_invalid_backoff_factor() {
        for factor in [0.0, 0.99, f64::NAN, f64::INFINITY] {
            let result = RetryPolicy::new().with_backoff_factor(factor).validate();
            assert!(
                matches!(result, Err(WincentError::InvalidArgument(_))),
                "factor {factor:?} should be rejected, got: {result:?}"
            );
        }
    }

    #[test]
    fn validated_returns_self_or_error() {
        let policy = RetryPolicy::new().validated().expect("default is valid");
        assert_eq!(policy.max_attempts(), RetryPolicy::default().max_attempts());

        let result = RetryPolicy::new().with_backoff_factor(f64::NAN).validated();
        assert!(matches!(result, Err(WincentError::InvalidArgument(_))));
    }

    #[test]
    fn validate_rejects_zero_initial_delay_when_retries_enabled() {
        let result = RetryPolicy::new()
            .with_max_attempts(1)
            .with_initial_delay(Duration::ZERO)
            .validate();

        assert!(matches!(result, Err(WincentError::InvalidArgument(_))));
    }

    #[test]
    fn validate_rejects_max_delay_less_than_initial_delay() {
        let result = RetryPolicy::new()
            .with_initial_delay(Duration::from_millis(100))
            .with_max_delay(Duration::from_millis(50))
            .validate();

        assert!(matches!(result, Err(WincentError::InvalidArgument(_))));
    }

    #[test]
    fn validate_is_order_insensitive_for_final_state() {
        let policy = RetryPolicy::new()
            .with_max_delay(Duration::from_millis(50))
            .with_initial_delay(Duration::from_millis(10));

        assert!(policy.validate().is_ok());
    }
}
