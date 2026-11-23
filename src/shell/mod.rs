//! Windows Shell and COM integration used by Quick Access operations.

/// `SHAddToRecentDocs` operation for adding or clearing Recent Items.
pub(crate) const SHARD_PATHW: u32 = 0x0000_0003;

pub(crate) mod clear;
mod com;
pub(crate) mod explorer;
pub(crate) mod mutation;
mod query;
pub(crate) mod thread;
