//! Native OOXML operations on saved DOCX files, without Microsoft Word.
//!
//! Edits patch only the targeted spans of the affected parts, keep every other package
//! part byte-for-byte, refuse documents Word holds open, detect concurrent changes, and
//! leave a backup of the original beside the document.
//!
//! Each area (document, read, edit, format, table, review, layout, references, media,
//! controls) is one grouped MCP tool, `docx_<area>`, whose `operation.action` selects
//! the operation.

mod controls;
mod document;
mod edit;
mod format;
mod layout;
mod media;
mod ooxml;
mod package;
mod props;
mod read;
mod references;
mod review;
mod table;
#[cfg(test)]
mod testing;
mod xml;

use anyhow::{Result, bail};
use rmcp::model::Tool;
use serde_json::Value;

use crate::tool::Output;

/// One grouped saved-file tool.
struct Definition {
    name: &'static str,
    tool: fn() -> Tool,
    call: fn(Value) -> Result<Output>,
}

const TOOLS: &[Definition] = &[
    document::TOOL,
    read::TOOL,
    edit::TOOL,
    format::TOOL,
    table::TOOL,
    review::TOOL,
    layout::TOOL,
    references::TOOL,
    media::TOOL,
    controls::TOOL,
];

/// The saved-file tool definitions.
#[must_use]
pub fn tools() -> Vec<Tool> {
    TOOLS.iter().map(|definition| (definition.tool)()).collect()
}

/// Run one saved-file tool.
///
/// # Errors
/// Returns an error for an unknown tool, invalid arguments, or an invalid or
/// unsupported document.
pub(crate) fn call(name: &str, arguments: Value) -> Result<Output> {
    let Some(definition) = TOOLS.iter().find(|definition| definition.name == name) else {
        bail!("unknown saved-file tool: {name}");
    };
    (definition.call)(arguments)
}
