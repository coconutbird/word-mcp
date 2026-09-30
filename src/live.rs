//! Live Microsoft Word automation on Windows.
//!
//! Word's objects live in a single-threaded COM apartment, so one dedicated thread owns
//! them. [`LiveBackend`] validates each call on the caller's thread, then hands the typed
//! request to that thread and waits for its reply.
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
mod args;
#[cfg(windows)]
mod word;

use crate::tool::{Effect, tool};
use args::{
    CloseArgs, DeleteArgs, DocumentArgs, ExportArgs, FormatArgs, InsertArgs, OpenArgs,
    ParagraphsArgs, ReadArgs, ReplaceArgs, Request, StatusArgs, UndoArgs,
};

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
    /// the message says so, and `word_live_undo` can revert it.
    pub fn call(&mut self, name: &str, arguments: Value) -> Result<Value> {
        let request = Request::parse(name, arguments)?;
        #[cfg(windows)]
        {
            let worker = match &mut self.worker {
                Some(worker) => worker,
                slot => slot.insert(word::Worker::start()?),
            };
            let result = worker.execute(request);
            if result.is_err() && !worker.is_alive() {
                // A dead worker is restarted by the next call.
                self.worker = None;
            }
            result
        }
        #[cfg(not(windows))]
        {
            drop(request);
            bail!(
                "live Word tools require Windows with desktop Microsoft Word; use the saved DOCX tools on this platform"
            )
        }
    }
}

/// The live Word tool definitions.
#[must_use]
pub fn tools() -> Vec<Tool> {
    vec![
        tool::<StatusArgs>(
            "word_live_status",
            "Report whether desktop Word is installed and running, and list its open documents. Never launches Word.",
            Effect::ReadOnly,
        ),
        tool::<OpenArgs>(
            "word_live_open",
            "Open a document in Microsoft Word with macros disabled, launching Word if needed. Reuses the document if it is already open.",
            Effect::Additive,
        ),
        tool::<CloseArgs>(
            "word_live_close",
            "Close an open document. A document with unsaved changes needs save=true (save) or save=false (discard). Word itself keeps running.",
            Effect::Destructive,
        ),
        tool::<ReadArgs>(
            "word_live_read",
            "Read main-story text, including unsaved edits. Positions are Word UTF-16 offsets with an exclusive end; paragraphs end in \\r.",
            Effect::ReadOnly,
        ),
        tool::<ParagraphsArgs>(
            "word_live_paragraphs",
            "List main-story paragraphs with their Word start/end positions, style name, and text. Use the positions for insert, delete, and format.",
            Effect::ReadOnly,
        ),
        tool::<ReplaceArgs>(
            "word_live_replace_text",
            "Replace literal, case-sensitive main-story text as one undo record. Replaces the first match unless all=true. Unsaved until word_live_save.",
            Effect::Destructive,
        ),
        tool::<InsertArgs>(
            "word_live_insert_text",
            "Insert text at a Word position as one undo record. Unsaved until word_live_save.",
            Effect::Destructive,
        ),
        tool::<DeleteArgs>(
            "word_live_delete_range",
            "Delete the text between two Word positions as one undo record. Unsaved until word_live_save.",
            Effect::Destructive,
        ),
        tool::<FormatArgs>(
            "word_live_format",
            "Apply character formatting, a style, or paragraph alignment to a Word range as one undo record. Unsaved until word_live_save.",
            Effect::Destructive,
        ),
        tool::<DocumentArgs>(
            "word_live_save",
            "Save an open document to its current file.",
            Effect::Destructive,
        ),
        tool::<ExportArgs>(
            "word_live_export_pdf",
            "Export the document through Word's layout engine to PDF, including unsaved changes. Does not save the document.",
            Effect::Additive,
        ),
        tool::<UndoArgs>(
            "word_live_undo",
            "Undo recent actions in an open document, including the user's own actions. Each edit tool call is one action.",
            Effect::Destructive,
        ),
        tool::<DocumentArgs>(
            "word_live_view",
            "Show Word and bring an open document to the front for visual inspection.",
            Effect::Additive,
        ),
    ]
}
