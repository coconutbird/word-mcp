//! Live Microsoft Word automation on Windows.
//!
//! Word's objects live in a single-threaded COM apartment, so one dedicated thread owns
//! them. [`LiveBackend`] validates each call on the caller's thread, then hands the typed
//! operation to that thread and waits for its reply.
use anyhow::Result;
#[cfg(not(windows))]
use anyhow::bail;
use rmcp::model::Tool;
use serde_json::Value;

#[cfg_attr(
    not(windows),
    expect(
        dead_code,
        reason = "Only the Windows backend reads the validated arguments"
    )
)]
mod areas;
mod common;
#[cfg(windows)]
mod word;

use crate::tool::Output;

/// Handle to the Word automation thread, started on first use.
#[derive(Default)]
pub struct LiveBackend {
    #[cfg(windows)]
    worker: Option<word::Worker>,
}

impl LiveBackend {
    /// Construct the backend without starting a thread or Word.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Validate and execute one live Word tool.
    ///
    /// # Errors
    /// Returns an error for invalid arguments, when Word is unavailable or does not
    /// answer in time, or when automation fails. A failed edit may be partially applied;
    /// the message says so, and `word_live_edit` undo can revert it.
    #[cfg_attr(
        not(windows),
        expect(
            clippy::unused_self,
            reason = "Only the Windows backend keeps a worker"
        )
    )]
    pub(crate) fn call(&mut self, name: &str, arguments: Value) -> Result<Output> {
        let operation = areas::parse(name, arguments)?;
        #[cfg(windows)]
        {
            let worker = match &mut self.worker {
                Some(worker) => worker,
                slot => slot.insert(word::Worker::start()?),
            };
            let result = worker.execute(operation);
            if result.is_err() && !worker.is_alive() {
                // A dead worker is restarted by the next call.
                self.worker = None;
            }
            result
        }
        #[cfg(not(windows))]
        {
            drop(operation);
            bail!(
                "live Word tools require Windows with desktop Microsoft Word; use the docx_* tools on this platform"
            )
        }
    }
}

/// The live Word tool definitions.
#[must_use]
pub fn tools() -> Vec<Tool> {
    areas::tools()
}
