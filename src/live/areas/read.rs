//! `word_live_read`: read an open document, including unsaved edits.

use std::path::PathBuf;

use anyhow::{Result, ensure};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;

use super::{Definition, Operation};
use crate::{
    live::common::document_path,
    tool::{Effect, parse, tool},
};

const NAME: &str = "word_live_read";

pub(super) const TOOL: Definition = Definition {
    name: NAME,
    tool: || {
        tool::<Arguments>(
            NAME,
            "Read an open Word document's main story, including unsaved edits: text between Word positions (UTF-16 offsets, exclusive end; paragraphs end in \\r), or paragraphs with their positions, style names, and text.",
            Effect::ReadOnly,
        )
    },
    parse: parse_operation,
};

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct Arguments {
    /// Absolute path of a document already open in Word.
    path: PathBuf,
    operation: Action,
}

#[derive(Deserialize, JsonSchema)]
#[serde(tag = "action", rename_all = "snake_case", deny_unknown_fields)]
enum Action {
    /// Text between two Word positions.
    Text {
        /// First Word position to read. Defaults to the story start.
        #[schemars(range(min = 0))]
        start: Option<i32>,
        /// Exclusive end position. Defaults to the story end.
        #[schemars(range(min = 0))]
        end: Option<i32>,
    },
    /// Paragraphs with start/end positions, style name, and text. Use the positions for
    /// edits and formatting.
    Paragraphs {
        /// Zero-based index of the first paragraph.
        #[serde(default)]
        start: u32,
        /// Maximum number of paragraphs.
        #[serde(default = "default_limit")]
        #[schemars(range(min = 1, max = 500))]
        limit: u32,
    },
}

fn default_limit() -> u32 {
    100
}

struct Request {
    path: PathBuf,
    action: Action,
}

fn parse_operation(arguments: Value) -> Result<Box<dyn Operation>> {
    let Arguments { path, operation } = parse(NAME, arguments)?;
    match &operation {
        Action::Text { start, end } => ensure!(
            start.is_none_or(|start| start >= 0)
                && end.is_none_or(|end| end >= 0)
                && !matches!((start, end), (Some(start), Some(end)) if start > end),
            "range must satisfy 0 <= start <= end"
        ),
        Action::Paragraphs { limit, .. } => {
            ensure!((1..=500).contains(limit), "limit must be between 1 and 500");
        }
    }
    Ok(Box::new(Request {
        path: document_path(&path)?,
        action: operation,
    }))
}

impl Operation for Request {
    #[cfg(windows)]
    fn run(
        self: Box<Self>,
        session: &mut crate::live::word::Session,
    ) -> Result<crate::tool::Output> {
        let (_, document) = session.document(&self.path)?;
        Ok(match self.action {
            Action::Text { start, end } => com::text(&document, start, end)?,
            Action::Paragraphs { start, limit } => com::paragraphs(&document, start, limit)?,
        }
        .into())
    }
}

#[cfg(windows)]
mod com {
    use anyhow::{Context, Result};
    use serde_json::{Value, json};

    use crate::live::word::{IDispatch, check_range, item, range};

    pub(super) fn text(
        document: &IDispatch,
        start: Option<i32>,
        end: Option<i32>,
    ) -> Result<Value> {
        let content = document.object("Content")?;
        let start = start.map_or_else(|| content.int("Start"), Ok)?;
        let end = end.map_or_else(|| content.int("End"), Ok)?;
        check_range(document, start, end)?;
        Ok(json!({
            "path": document.string("FullName")?,
            "text": range(document, start, end)?.string("Text")?,
            "start": start,
            "end": end,
            "saved": document.flag("Saved")?,
        }))
    }

    pub(super) fn paragraphs(document: &IDispatch, start: u32, limit: u32) -> Result<Value> {
        let paragraphs = document.object("Paragraphs")?;
        let total = paragraphs.int("Count")?;
        let first = i32::try_from(start).context("start is too large")?;
        let mut listed = Vec::new();
        let mut next = if first < total {
            Some(item(&paragraphs, first + 1)?)
        } else {
            None
        };
        let mut index = first;
        while let Some(paragraph) = next {
            if listed.len() == limit as usize {
                break;
            }
            let range = paragraph.object("Range")?;
            let text = range.string("Text")?;
            listed.push(json!({
                "index": index,
                "start": range.int("Start")?,
                "end": range.int("End")?,
                "style": style_name(&paragraph)?,
                "text": text.strip_suffix('\r').unwrap_or(&text),
            }));
            next = paragraph.call("Next", vec![])?.into_object()?;
            index += 1;
        }
        Ok(json!({
            "path": document.string("FullName")?,
            "total_paragraphs": total,
            "start": start,
            "paragraphs": listed,
        }))
    }

    /// The display name of a paragraph's (or range's) style.
    pub(in crate::live) fn style_name(target: &IDispatch) -> Result<Option<String>> {
        let style = target.get("Style")?;
        if let Ok(name) = style.string() {
            return Ok(Some(name));
        }
        style
            .into_object()?
            .map(|style| style.string("NameLocal"))
            .transpose()
    }
}

#[cfg(all(test, windows))]
mod tests {
    use serde_json::json;

    use crate::live::word::testing::{text, with_word};

    #[test]
    #[ignore = "Requires desktop Microsoft Word; starts a private hidden instance"]
    fn reads_text_and_paragraph_positions() {
        with_word(|harness| {
            let (path, _document) =
                harness.document("read.docx", "alpha 🦀 beta\rsecond paragraph")?;
            let session = &mut harness.session;
            let read = session.json(
                "word_live_read",
                json!({"path": path, "operation": {"action": "text"}}),
            )?;
            assert!(text(&read).contains("alpha 🦀 beta"));
            let listed = session.json(
                "word_live_read",
                json!({"path": path, "operation": {"action": "paragraphs"}}),
            )?;
            assert_eq!(listed["total_paragraphs"], 2);
            assert_eq!(listed["paragraphs"][1]["text"], "second paragraph");
            let second = session.json(
                "word_live_read",
                json!({"path": path, "operation": {"action": "paragraphs", "start": 1, "limit": 1}}),
            )?;
            assert_eq!(second["paragraphs"][0]["index"], 1);
            assert!(
                session
                    .json(
                        "word_live_read",
                        json!({"path": path, "operation": {"action": "text", "start": 5, "end": 2}}),
                    )
                    .is_err()
            );
            Ok(())
        });
    }
}
