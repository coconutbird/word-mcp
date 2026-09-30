//! `docx_references`: hyperlinks, bookmarks, fields, table of contents, footnotes, and endnotes of a saved .docx.
//!
//! SCAFFOLD: the area owner replaces this file; `Operation` has no actions yet.

use std::path::PathBuf;

use anyhow::Result;
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;

use super::Definition;
use crate::tool::{Effect, Output, parse, tool};

const NAME: &str = "docx_references";

pub(super) const TOOL: Definition = Definition {
    name: NAME,
    tool: || {
        tool::<Arguments>(
            NAME,
            "Hyperlinks, bookmarks, fields, table of contents, footnotes, and endnotes of a saved .docx.",
            Effect::Destructive,
        )
    },
    call,
};

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct Arguments {
    /// Local .docx path.
    #[schemars(length(min = 1))]
    #[expect(dead_code, reason = "SCAFFOLD until the area has actions")]
    path: PathBuf,
    operation: Operation,
}

#[derive(Deserialize, JsonSchema)]
#[serde(tag = "action", rename_all = "snake_case", deny_unknown_fields)]
enum Operation {}

fn call(arguments: Value) -> Result<Output> {
    let arguments: Arguments = parse(NAME, arguments)?;
    match arguments.operation {}
}
