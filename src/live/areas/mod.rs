//! The live Word tools, one grouped tool per area.
//!
//! Each area module defines its argument types (which double as the MCP input schema),
//! validates them on the caller's thread, and implements [`Operation`] to run on the
//! Word automation thread.

mod controls;
mod document;
mod edit;
mod format;
mod layout;
mod media;
mod read;
mod references;
mod review;
mod table;

use anyhow::{Result, bail};
use rmcp::model::Tool;
use serde_json::Value;

#[cfg(windows)]
use super::word::Session;
#[cfg(windows)]
use crate::tool::Output;

/// A validated live operation, ready for the Word automation thread.
pub(in crate::live) trait Operation: Send {
    /// Execute against Word on the automation thread.
    #[cfg(windows)]
    fn run(self: Box<Self>, session: &mut Session) -> Result<Output>;
}

/// One grouped live tool.
pub(in crate::live) struct Definition {
    name: &'static str,
    tool: fn() -> Tool,
    parse: fn(Value) -> Result<Box<dyn Operation>>,
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

/// The live tool definitions.
pub(in crate::live) fn tools() -> Vec<Tool> {
    TOOLS.iter().map(|definition| (definition.tool)()).collect()
}

/// Parse and validate one live tool call.
pub(in crate::live) fn parse(name: &str, arguments: Value) -> Result<Box<dyn Operation>> {
    let Some(definition) = TOOLS.iter().find(|definition| definition.name == name) else {
        bail!("unknown live Word tool: {name}");
    };
    (definition.parse)(arguments)
}
