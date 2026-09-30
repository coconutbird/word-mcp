//! `word_live_format`: character, paragraph, list, and style formatting of an open
//! document. Each call is one undo step.

use std::path::PathBuf;

use anyhow::{Result, ensure};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;

use super::{Definition, Operation};
use crate::{
    live::common::{check_positions, document_path},
    tool::{Effect, parse, tool},
};

const NAME: &str = "word_live_format";

pub(super) const TOOL: Definition = Definition {
    name: NAME,
    tool: || {
        tool::<Arguments>(
            NAME,
            "Format an open Word document: apply character formatting, a style, or paragraph alignment to a range between Word positions. Each call is one undo step and stays unsaved until word_live_document save.",
            Effect::Destructive,
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

/// Paragraph alignment.
#[derive(Clone, Copy, Deserialize, JsonSchema)]
#[serde(rename_all = "lowercase")]
enum Alignment {
    Left,
    Center,
    Right,
    Justify,
}

impl Alignment {
    /// `WdParagraphAlignment`.
    #[cfg(windows)]
    fn word_value(self) -> i32 {
        match self {
            Self::Left => 0,
            Self::Center => 1,
            Self::Right => 2,
            Self::Justify => 3,
        }
    }
}

#[derive(Deserialize, JsonSchema)]
#[serde(tag = "action", rename_all = "snake_case", deny_unknown_fields)]
enum Action {
    /// Apply formatting to the range `start..end`.
    Text {
        /// First Word position of the range.
        #[schemars(range(min = 0))]
        start: i32,
        /// Exclusive end position; must be greater than `start`.
        #[schemars(range(min = 1))]
        end: i32,
        /// Set or clear bold.
        bold: Option<bool>,
        /// Set or clear italic.
        italic: Option<bool>,
        /// Set single underline or remove underline.
        underline: Option<bool>,
        /// Font family name.
        #[schemars(length(min = 1))]
        font_name: Option<String>,
        /// Font size in points (half-point steps).
        #[schemars(range(min = 1.0, max = 1638.0))]
        font_size_pt: Option<f64>,
        /// Style name (for example "Heading 1"), applied to the range.
        #[schemars(length(min = 1))]
        style: Option<String>,
        /// Alignment of the paragraphs the range touches.
        alignment: Option<Alignment>,
        /// Temporarily turn Track Changes on or off for this edit.
        tracked_changes: Option<bool>,
    },
}

struct Request {
    path: PathBuf,
    action: Action,
}

fn parse_operation(arguments: Value) -> Result<Box<dyn Operation>> {
    let Arguments { path, operation } = parse(NAME, arguments)?;
    match &operation {
        Action::Text {
            start,
            end,
            bold,
            italic,
            underline,
            font_name,
            font_size_pt,
            style,
            alignment,
            ..
        } => {
            check_positions(*start, *end, false)?;
            ensure!(
                bold.is_some()
                    || italic.is_some()
                    || underline.is_some()
                    || font_name.is_some()
                    || font_size_pt.is_some()
                    || style.is_some()
                    || alignment.is_some(),
                "provide at least one formatting property"
            );
            if let Some(size) = font_size_pt {
                ensure!(
                    size.is_finite()
                        && (1.0..=1638.0).contains(size)
                        && (size * 2.0).fract() == 0.0,
                    "font_size_pt must be 1-1638 in half-point steps"
                );
            }
            ensure!(
                font_name
                    .as_deref()
                    .is_none_or(|name| !name.trim().is_empty())
                    && style
                        .as_deref()
                        .is_none_or(|style| !style.trim().is_empty()),
                "font_name and style must not be blank"
            );
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
        use serde_json::json;

        use crate::live::word::{check_range, mutation, range};

        /// `wdUnderlineSingle` and `wdUnderlineNone`.
        const UNDERLINE_SINGLE: i32 = 1;
        const UNDERLINE_NONE: i32 = 0;

        let (word, document) = session.document(&self.path)?;
        match self.action {
            Action::Text {
                start,
                end,
                bold,
                italic,
                underline,
                font_name,
                font_size_pt,
                style,
                alignment,
                tracked_changes,
            } => {
                check_range(&document, start, end)?;
                mutation(&word, &document, tracked_changes, "Word MCP format", || {
                    let target = range(&document, start, end)?;
                    // A style first, so explicit character formatting applies on top.
                    if let Some(style) = &style {
                        target.put("Style", style.as_str().into())?;
                    }
                    let font = target.object("Font")?;
                    if let Some(bold) = bold {
                        font.put("Bold", bold.into())?;
                    }
                    if let Some(italic) = italic {
                        font.put("Italic", italic.into())?;
                    }
                    if let Some(underline) = underline {
                        let value = if underline {
                            UNDERLINE_SINGLE
                        } else {
                            UNDERLINE_NONE
                        };
                        font.put("Underline", value.into())?;
                    }
                    if let Some(name) = &font_name {
                        font.put("Name", name.as_str().into())?;
                    }
                    if let Some(size) = font_size_pt {
                        font.put("Size", size.into())?;
                    }
                    if let Some(alignment) = alignment {
                        target
                            .object("ParagraphFormat")?
                            .put("Alignment", alignment.word_value().into())?;
                    }
                    Ok(())
                })?;
                Ok(json!({"start": start, "end": end, "saved": document.flag("Saved")?}).into())
            }
        }
    }
}

#[cfg(all(test, windows))]
mod tests {
    use serde_json::json;

    use crate::live::word::{range, testing::with_word};

    #[test]
    #[ignore = "Requires desktop Microsoft Word; starts a private hidden instance"]
    fn formats_a_range() {
        with_word(|harness| {
            let (path, document) = harness.document("format.docx", "alpha\rsecond paragraph")?;
            let session = &mut harness.session;
            // "second" starts after "alpha\r".
            session.json(
                "word_live_format",
                json!({"path": path, "operation": {"action": "text", "start": 6, "end": 12, "bold": true, "font_size_pt": 15.5, "alignment": "center"}}),
            )?;
            let formatted = range(&document, 6, 12)?;
            assert_eq!(formatted.object("Font")?.int("Bold")?, -1);
            assert!((formatted.object("Font")?.get("Size")?.number()? - 15.5).abs() < f64::EPSILON);
            assert_eq!(formatted.object("ParagraphFormat")?.int("Alignment")?, 1);
            assert!(
                session
                    .json(
                        "word_live_format",
                        json!({"path": path, "operation": {"action": "text", "start": 0, "end": 4}}),
                    )
                    .is_err()
            );
            Ok(())
        });
    }
}
