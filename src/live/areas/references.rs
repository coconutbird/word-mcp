//! `word_live_references`: hyperlinks, bookmarks, fields, tables of contents, captions, cross-references, footnotes, and endnotes of an open Word document.
//!
//! SCAFFOLD: the area owner replaces this file; `Action` has no actions yet.

use anyhow::Result;
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;

use super::{Definition, Operation};
use crate::tool::{Effect, parse, tool};

const NAME: &str = "word_live_references";

pub(super) const TOOL: Definition = Definition {
    name: NAME,
    tool: || {
        tool::<Arguments>(
            NAME,
            "Hyperlinks, bookmarks, fields, tables of contents, captions, cross-references, footnotes, and endnotes of an open Word document.",
            Effect::Destructive,
        )
    },
    parse: parse_operation,
};

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct Arguments {
    operation: Action,
}

#[derive(Deserialize, JsonSchema)]
#[serde(tag = "action", rename_all = "snake_case", deny_unknown_fields)]
enum Action {}

impl Operation for Action {
    #[cfg(windows)]
    fn run(
        self: Box<Self>,
        _session: &mut crate::live::word::Session,
    ) -> Result<crate::tool::Output> {
        match *self {}
    }
}

fn parse_operation(arguments: Value) -> Result<Box<dyn Operation>> {
    let arguments: Arguments = parse(NAME, arguments)?;
    Ok(Box::new(arguments.operation))
}
