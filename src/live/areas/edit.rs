//! `word_live_edit`: change the text of an open document. Each call is one undo step.

use std::path::PathBuf;

use anyhow::{Result, ensure};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;

use super::{Definition, Operation};
use crate::{
    live::common::{check_find, check_positions, document_path},
    tool::{Effect, parse, tool},
};

const NAME: &str = "word_live_edit";

pub(super) const TOOL: Definition = Definition {
    name: NAME,
    tool: || {
        tool::<Arguments>(
            NAME,
            "Edit an open Word document's main story: replace literal text, insert text at a Word position, delete a range, or undo recent actions. Each edit is one Word undo step and stays unsaved until word_live_document save. `tracked_changes` temporarily turns Track Changes on or off for the edit.",
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

#[derive(Deserialize, JsonSchema)]
#[serde(tag = "action", rename_all = "snake_case", deny_unknown_fields)]
enum Action {
    /// Replace literal, case-sensitive text. Replaces the first match unless `all`.
    Replace {
        /// Literal text to find (at most 255 UTF-16 units after escaping).
        #[schemars(length(min = 1))]
        find: String,
        /// Replacement text.
        replacement: String,
        /// Replace every match.
        #[serde(default)]
        all: bool,
        /// Temporarily turn Track Changes on or off for this edit.
        tracked_changes: Option<bool>,
    },
    /// Insert text at a Word position. `\r` starts a new paragraph; `\u000b` is a
    /// line break.
    InsertText {
        /// Word position (UTF-16 units) to insert at.
        #[schemars(range(min = 0))]
        position: i32,
        /// Text to insert.
        text: String,
        /// Temporarily turn Track Changes on or off for this edit.
        tracked_changes: Option<bool>,
    },
    /// Delete the text between two Word positions.
    DeleteRange {
        /// First Word position to delete.
        #[schemars(range(min = 0))]
        start: i32,
        /// Exclusive end position; must be greater than `start`.
        #[schemars(range(min = 1))]
        end: i32,
        /// Temporarily turn Track Changes on or off for this edit.
        tracked_changes: Option<bool>,
    },
    /// Undo recent actions, including the user's own actions.
    Undo {
        /// Number of actions to undo.
        #[serde(default = "one")]
        #[schemars(range(min = 1, max = 100))]
        count: i32,
    },
}

fn one() -> i32 {
    1
}

struct Request {
    path: PathBuf,
    action: Action,
}

fn parse_operation(arguments: Value) -> Result<Box<dyn Operation>> {
    let Arguments { path, operation } = parse(NAME, arguments)?;
    match &operation {
        Action::Replace { find, .. } => check_find(find)?,
        Action::InsertText { position, .. } => {
            ensure!(*position >= 0, "position must be nonnegative");
        }
        Action::DeleteRange { start, end, .. } => check_positions(*start, *end, false)?,
        Action::Undo { count } => {
            ensure!((1..=100).contains(count), "count must be between 1 and 100");
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
        let (word, document) = session.document(&self.path)?;
        Ok(match self.action {
            Action::Replace {
                find,
                replacement,
                all,
                tracked_changes,
            } => com::replace(&word, &document, &find, &replacement, all, tracked_changes)?,
            Action::InsertText {
                position,
                text,
                tracked_changes,
            } => com::insert(&word, &document, position, &text, tracked_changes)?,
            Action::DeleteRange {
                start,
                end,
                tracked_changes,
            } => com::delete(&word, &document, start, end, tracked_changes)?,
            Action::Undo { count } => com::undo(&document, count)?,
        }
        .into())
    }
}

#[cfg(windows)]
mod com {
    use anyhow::Result;
    use serde_json::{Value, json};

    use crate::live::{
        common::word_find_text,
        word::{IDispatch, check_range, mutation, range, writable},
    };

    pub(super) fn replace(
        word: &IDispatch,
        document: &IDispatch,
        find: &str,
        replacement: &str,
        all: bool,
        tracking: Option<bool>,
    ) -> Result<Value> {
        let matches = literal_matches(document, find, all)?;
        if !matches.is_empty() {
            mutation(word, document, tracking, "Word MCP replace text", || {
                // Last to first, so earlier positions stay valid.
                for &(start, end) in matches.iter().rev() {
                    range(document, start, end)?.put("Text", replacement.into())?;
                }
                Ok(())
            })?;
        }
        Ok(json!({"replacements": matches.len(), "saved": document.flag("Saved")?}))
    }

    pub(super) fn insert(
        word: &IDispatch,
        document: &IDispatch,
        position: i32,
        text: &str,
        tracking: Option<bool>,
    ) -> Result<Value> {
        check_range(document, position, position)?;
        if !text.is_empty() {
            mutation(word, document, tracking, "Word MCP insert text", || {
                range(document, position, position)?.put("Text", text.into())
            })?;
        }
        Ok(json!({
            "inserted_utf16": text.encode_utf16().count(),
            "saved": document.flag("Saved")?,
        }))
    }

    pub(super) fn delete(
        word: &IDispatch,
        document: &IDispatch,
        start: i32,
        end: i32,
        tracking: Option<bool>,
    ) -> Result<Value> {
        check_range(document, start, end)?;
        let target = range(document, start, end)?;
        let deleted = target.string("Text")?;
        mutation(word, document, tracking, "Word MCP delete range", || {
            target.call("Delete", vec![]).map(drop)
        })?;
        Ok(json!({"deleted_text": deleted, "saved": document.flag("Saved")?}))
    }

    pub(super) fn undo(document: &IDispatch, count: i32) -> Result<Value> {
        writable(document)?;
        let undone = document.call("Undo", vec![count.into()])?.flag()?;
        Ok(json!({
            "undone": undone,
            "requested_count": count,
            "saved": document.flag("Saved")?,
        }))
    }

    /// Word-reported ranges of literal matches in the main story.
    ///
    /// Word's own Find locates each match because fields and hidden text make
    /// plain-text offsets diverge from Word positions.
    pub(in crate::live) fn literal_matches(
        document: &IDispatch,
        text: &str,
        all: bool,
    ) -> Result<Vec<(i32, i32)>> {
        // Word reports manual line breaks as VT; accept LF for them.
        let literal = text.replace('\n', "\u{b}");
        let pattern = word_find_text(text);
        let content = document.object("Content")?;
        let mut position = content.int("Start")?;
        let limit = content.int("End")?;
        let mut matches = Vec::new();
        while position < limit {
            let scope = range(document, position, limit)?;
            let find = scope.object("Find")?;
            find.call("ClearFormatting", vec![])?;
            for option in ["IgnorePunct", "IgnoreSpace", "MatchPrefix", "MatchSuffix"] {
                find.put(option, false.into())?;
            }
            let found = find
                .call_named(
                    "Execute",
                    &[
                        "FindText",
                        "MatchCase",
                        "MatchWholeWord",
                        "MatchWildcards",
                        "MatchSoundsLike",
                        "MatchAllWordForms",
                        "Forward",
                        "Wrap",
                        "Format",
                    ],
                    vec![
                        pattern.as_str().into(),
                        true.into(),
                        false.into(),
                        false.into(),
                        false.into(),
                        false.into(),
                        true.into(),
                        // wdFindStop
                        0.into(),
                        false.into(),
                    ],
                )?
                .flag()?;
            if !found {
                break;
            }
            // A successful Find moves `scope` onto the match.
            let start = scope.int("Start")?;
            let end = scope.int("End")?;
            anyhow::ensure!(end > position, "Word search made no forward progress");
            if scope.string("Text")? == literal {
                matches.push((start, end));
                if !all {
                    break;
                }
            }
            position = end;
        }
        Ok(matches)
    }
}

#[cfg(all(test, windows))]
mod tests {
    use serde_json::{Value, json};

    use crate::live::word::{
        Variant, range,
        testing::{text, with_word},
    };

    #[test]
    #[ignore = "Requires desktop Microsoft Word; starts a private hidden instance"]
    fn replaces_inserts_deletes_and_undoes() {
        with_word(|harness| {
            let (path, document) = harness.document("edit.docx", "alpha 🦀 beta\rsecond")?;
            let session = &mut harness.session;
            let mut edit = |operation: Value| {
                session.json(
                    "word_live_edit",
                    json!({"path": path, "operation": operation}),
                )
            };
            edit(
                json!({"action": "insert_text", "position": 0, "text": "PREFIX ", "tracked_changes": false}),
            )?;
            assert_eq!(
                edit(
                    json!({"action": "replace", "find": "beta", "replacement": "delta", "all": true, "tracked_changes": false})
                )?["replacements"],
                1
            );
            assert_eq!(edit(json!({"action": "undo"}))?["undone"], true);
            assert_eq!(
                edit(json!({"action": "delete_range", "start": 0, "end": 7}))?["deleted_text"],
                "PREFIX "
            );

            // Word positions stay authoritative when a field precedes the match.
            let field_range = range(&document, 0, 0)?;
            document.object("Fields")?.call(
                "Add",
                vec![
                    Variant::object(&field_range),
                    (-1).into(),
                    "DATE".into(),
                    true.into(),
                ],
            )?;
            assert_eq!(
                edit(json!({"action": "replace", "find": "beta", "replacement": "gamma"}))?["replacements"],
                1
            );
            // Find special codes are literal text, and LF matches a manual line break.
            edit(json!({"action": "insert_text", "position": 0, "text": "^p left\u{b}right "}))?;
            assert_eq!(
                edit(json!({"action": "replace", "find": "^p", "replacement": "[literal]"}))?["replacements"],
                1
            );
            assert_eq!(
                edit(json!({"action": "replace", "find": "left\nright", "replacement": "LF"}))?["replacements"],
                1
            );

            let original = document.flag("TrackRevisions")?;
            edit(
                json!({"action": "insert_text", "position": 0, "text": "TRACKED ", "tracked_changes": true}),
            )?;
            assert_eq!(document.flag("TrackRevisions")?, original);
            let read = session.json(
                "word_live_read",
                json!({"path": path, "operation": {"action": "text"}}),
            )?;
            assert!(text(&read).contains("[literal] LF") && text(&read).contains("alpha 🦀 gamma"));
            Ok(())
        });
    }
}
