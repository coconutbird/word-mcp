//! `word_live_read`: read an open document, including unsaved edits.

#[cfg(windows)]
mod com;

use std::path::PathBuf;

use anyhow::{Result, bail, ensure};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;

use super::{Definition, Operation};
use crate::{
    live::common::{FIND_LIMIT, check_find, check_positions, document_path},
    tool::{Effect, parse, tool},
};

const NAME: &str = "word_live_read";
/// Most pages `page_text` returns in one call.
const MAX_PAGES: u32 = 50;

pub(super) const TOOL: Definition = Definition {
    name: NAME,
    tool: || {
        tool::<Arguments>(
            NAME,
            "Read an open Word document's main story, including unsaved edits. Positions are Word character positions (UTF-16 offsets, exclusive end; paragraphs end in \\r); paragraph indices are zero-based and pages one-based. Actions: text between positions; paragraphs with positions and styles; find with context; heading outline; text of pages; paragraph and font formatting of a range; snapshot and diff of paragraph texts; the user's current selection.",
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
    /// Find text in the main story with Word's Find, returning each match's positions,
    /// surrounding context, and paragraph index.
    Find {
        /// Text to find; literal unless `wildcards`.
        #[schemars(length(min = 1, max = 255))]
        text: String,
        /// Match letter case exactly.
        #[serde(default = "yes")]
        match_case: bool,
        /// Only match whole words.
        #[serde(default)]
        whole_word: bool,
        /// Interpret `text` with Word's wildcard syntax (`?`, `*`, `[a-z]`, `{n}`, `<`,
        /// `>`...).
        #[serde(default)]
        wildcards: bool,
        /// Maximum number of matches to return.
        #[serde(default = "default_max_results")]
        #[schemars(range(min = 1, max = 500))]
        max_results: u32,
        /// Characters of context to return before and after each match.
        #[serde(default = "default_context_chars")]
        #[schemars(range(min = 0, max = 1000))]
        context_chars: i32,
    },
    /// Paragraphs with an outline level 1-9 (headings), with positions and text.
    Outline {},
    /// Text and positions of one page or a range of pages (at most 50), plus the total
    /// page count.
    PageText {
        /// One-based first page.
        #[schemars(range(min = 1))]
        page: u32,
        /// One-based last page, inclusive. Defaults to `page`.
        #[schemars(range(min = 1))]
        end_page: Option<u32>,
    },
    /// Paragraph formatting and font of a range (`start` and `end`) or of one
    /// paragraph (`paragraph`). Values that differ across the range are null.
    Formatting {
        /// First Word position of the range.
        #[schemars(range(min = 0))]
        start: Option<i32>,
        /// Exclusive end position of the range.
        #[schemars(range(min = 0))]
        end: Option<i32>,
        /// Zero-based paragraph index, instead of `start`/`end`.
        paragraph: Option<u32>,
    },
    /// Remember the current paragraph texts for a later `diff`, replacing any earlier
    /// snapshot of this document.
    Snapshot {},
    /// Compare the current paragraphs with the last `snapshot`: changed, inserted, and
    /// deleted paragraphs with their indices. The snapshot is kept.
    Diff {},
    /// The user's current selection when this document is Word's active document.
    Selection {},
}

fn yes() -> bool {
    true
}

fn default_limit() -> u32 {
    100
}

fn default_max_results() -> u32 {
    50
}

fn default_context_chars() -> i32 {
    40
}

/// Where `formatting` reads.
enum Target {
    Range { start: i32, end: i32 },
    Paragraph(u32),
}

/// A validated live read.
enum Read {
    Text {
        start: Option<i32>,
        end: Option<i32>,
    },
    Paragraphs {
        start: u32,
        limit: u32,
    },
    Find {
        /// The Find text, escaped unless wildcards are on.
        text: String,
        match_case: bool,
        whole_word: bool,
        wildcards: bool,
        max_results: u32,
        context_chars: i32,
    },
    Outline,
    PageText {
        first: u32,
        last: u32,
    },
    Formatting(Target),
    Snapshot,
    Diff,
    Selection,
}

struct Request {
    path: PathBuf,
    read: Read,
}

fn parse_operation(arguments: Value) -> Result<Box<dyn Operation>> {
    let Arguments { path, operation } = parse(NAME, arguments)?;
    let read = validate(operation)?;
    Ok(Box::new(Request {
        path: document_path(&path)?,
        read,
    }))
}

fn validate(action: Action) -> Result<Read> {
    Ok(match action {
        Action::Text { start, end } => {
            ensure!(
                start.is_none_or(|start| start >= 0)
                    && end.is_none_or(|end| end >= 0)
                    && !matches!((start, end), (Some(start), Some(end)) if start > end),
                "range must satisfy 0 <= start <= end"
            );
            Read::Text { start, end }
        }
        Action::Paragraphs { start, limit } => {
            ensure!(
                (1..=500).contains(&limit),
                "limit must be between 1 and 500"
            );
            Read::Paragraphs { start, limit }
        }
        Action::Find {
            text,
            match_case,
            whole_word,
            wildcards,
            max_results,
            context_chars,
        } => {
            ensure!(
                (1..=500).contains(&max_results),
                "max_results must be between 1 and 500"
            );
            ensure!(
                (0..=1000).contains(&context_chars),
                "context_chars must be between 0 and 1000"
            );
            let text = if wildcards {
                ensure!(!text.is_empty(), "text must not be empty");
                ensure!(
                    text.encode_utf16().count() <= FIND_LIMIT,
                    "text exceeds Word's {FIND_LIMIT} UTF-16 unit search limit"
                );
                text
            } else {
                check_find(&text)?;
                crate::live::common::word_find_text(&text)
            };
            Read::Find {
                text,
                match_case,
                whole_word,
                wildcards,
                max_results,
                context_chars,
            }
        }
        Action::Outline {} => Read::Outline,
        Action::PageText { page, end_page } => {
            let last = end_page.unwrap_or(page);
            ensure!(page >= 1, "page is one-based and must be at least 1");
            ensure!(last >= page, "end_page must not be before page");
            ensure!(
                last - page < MAX_PAGES,
                "at most {MAX_PAGES} pages can be read at once"
            );
            Read::PageText { first: page, last }
        }
        Action::Formatting {
            start,
            end,
            paragraph,
        } => Read::Formatting(match (start, end, paragraph) {
            (Some(start), Some(end), None) => {
                check_positions(start, end, true)?;
                Target::Range { start, end }
            }
            (None, None, Some(paragraph)) => Target::Paragraph(paragraph),
            _ => bail!("give either start and end, or paragraph"),
        }),
        Action::Snapshot {} => Read::Snapshot,
        Action::Diff {} => Read::Diff,
        Action::Selection {} => Read::Selection,
    })
}

impl Operation for Request {
    #[cfg(windows)]
    fn run(
        self: Box<Self>,
        session: &mut crate::live::word::Session,
    ) -> Result<crate::tool::Output> {
        let (word, document) = session.document(&self.path)?;
        Ok(match self.read {
            Read::Text { start, end } => com::text(&document, start, end)?,
            Read::Paragraphs { start, limit } => com::paragraphs(&document, start, limit)?,
            Read::Find {
                text,
                match_case,
                whole_word,
                wildcards,
                max_results,
                context_chars,
            } => com::find(
                &document,
                &com::Search {
                    text,
                    match_case,
                    whole_word,
                    wildcards,
                    max_results,
                    context_chars,
                },
            )?,
            Read::Outline => com::outline(&document)?,
            Read::PageText { first, last } => com::page_text(&document, first, last)?,
            Read::Formatting(target) => com::formatting(&document, &target)?,
            Read::Snapshot => {
                let texts = com::paragraph_texts(&document)?;
                let count = texts.len();
                session.snapshots.insert(self.path.clone(), texts);
                serde_json::json!({
                    "path": document.string("FullName")?,
                    "paragraphs": count,
                })
            }
            Read::Diff => {
                let snapshot = session.snapshots.get(&self.path).ok_or_else(|| {
                    anyhow::anyhow!(
                        "no snapshot of this document; call word_live_read snapshot first"
                    )
                })?;
                let current = com::paragraph_texts(&document)?;
                let changes = diff(snapshot, &current);
                serde_json::json!({
                    "path": document.string("FullName")?,
                    "unchanged": changes.is_empty(),
                    "snapshot_paragraphs": snapshot.len(),
                    "current_paragraphs": current.len(),
                    "changes": changes,
                })
            }
            Read::Selection => com::selection(&word, &document)?,
        }
        .into())
    }
}

/// Paragraph-level differences from `old` to `new`, in document order. Runs of
/// deletions and insertions between common paragraphs pair up as changes.
#[cfg(any(windows, test))]
fn diff(old: &[String], new: &[String]) -> Vec<Value> {
    use serde_json::json;

    /// Largest LCS table (old × new middle paragraphs) computed exactly.
    const MAX_CELLS: usize = 4_000_000;

    let prefix = old
        .iter()
        .zip(new)
        .take_while(|(old, new)| old == new)
        .count();
    let suffix = old[prefix..]
        .iter()
        .rev()
        .zip(new[prefix..].iter().rev())
        .take_while(|(old, new)| old == new)
        .count();
    let old_middle = &old[prefix..old.len() - suffix];
    let new_middle = &new[prefix..new.len() - suffix];
    // Common paragraph pairs (old index, new index) within the middles.
    let mut common = Vec::new();
    let (rows, columns) = (old_middle.len(), new_middle.len());
    if rows > 0 && columns > 0 && rows.saturating_mul(columns) <= MAX_CELLS {
        let mut table = vec![0_u32; (rows + 1) * (columns + 1)];
        let at = |row: usize, column: usize| row * (columns + 1) + column;
        for row in (0..rows).rev() {
            for column in (0..columns).rev() {
                table[at(row, column)] = if old_middle[row] == new_middle[column] {
                    table[at(row + 1, column + 1)] + 1
                } else {
                    table[at(row + 1, column)].max(table[at(row, column + 1)])
                };
            }
        }
        let (mut row, mut column) = (0, 0);
        while row < rows && column < columns {
            if old_middle[row] == new_middle[column] {
                common.push((row, column));
                row += 1;
                column += 1;
            } else if table[at(row + 1, column)] >= table[at(row, column + 1)] {
                row += 1;
            } else {
                column += 1;
            }
        }
    }
    common.push((rows, columns));
    let mut changes = Vec::new();
    let (mut row, mut column) = (0, 0);
    for (next_row, next_column) in common {
        let deleted = next_row - row;
        let inserted = next_column - column;
        for pair in 0..deleted.max(inserted) {
            let old_index = prefix + row + pair;
            let new_index = prefix + column + pair;
            changes.push(if pair < deleted && pair < inserted {
                json!({"type": "changed", "index": new_index, "old_index": old_index, "old_text": old[old_index], "text": new[new_index]})
            } else if pair < deleted {
                json!({"type": "deleted", "old_index": old_index, "old_text": old[old_index]})
            } else {
                json!({"type": "inserted", "index": new_index, "text": new[new_index]})
            });
        }
        row = next_row + 1;
        column = next_column + 1;
    }
    changes
}

#[cfg(test)]
mod tests {
    use serde_json::{Value, json};

    use super::diff;

    fn texts(values: &[&str]) -> Vec<String> {
        values.iter().map(|value| (*value).to_owned()).collect()
    }

    #[test]
    fn diff_pairs_changes_and_reports_insertions_and_deletions() {
        let old = texts(&["title", "alpha", "beta", "gamma", "end"]);
        let new = texts(&["title", "alpha!", "gamma", "new", "end", "tail"]);
        assert_eq!(
            Value::from(diff(&old, &new)),
            json!([
                {"type": "changed", "index": 1, "old_index": 1, "old_text": "alpha", "text": "alpha!"},
                {"type": "deleted", "old_index": 2, "old_text": "beta"},
                {"type": "inserted", "index": 3, "text": "new"},
                {"type": "inserted", "index": 5, "text": "tail"},
            ])
        );
        assert!(diff(&old, &old).is_empty());
        assert_eq!(
            Value::from(diff(&old, &[])).as_array().map(Vec::len),
            Some(5)
        );
    }

    #[test]
    fn validation_rejects_bad_requests_before_word() {
        let parse = |operation: Value| {
            super::parse_operation(json!({"path": "C:/missing/doc.docx", "operation": operation}))
                .err()
                .map(|error| format!("{error:#}"))
                .unwrap_or_default()
        };
        assert!(parse(json!({"action": "find", "text": ""})).contains("empty"));
        assert!(
            parse(json!({"action": "find", "text": "x", "max_results": 0})).contains("max_results")
        );
        assert!(parse(json!({"action": "page_text", "page": 0})).contains("page"));
        assert!(
            parse(json!({"action": "page_text", "page": 3, "end_page": 2})).contains("end_page")
        );
        assert!(
            parse(json!({"action": "page_text", "page": 1, "end_page": 60})).contains("at most")
        );
        assert!(parse(json!({"action": "formatting", "start": 1})).contains("either"));
        assert!(
            parse(json!({"action": "formatting", "start": 1, "end": 2, "paragraph": 0}))
                .contains("either")
        );
        assert!(parse(json!({"action": "formatting", "start": 5, "end": 2})).contains("range"));
        assert!(parse(json!({"action": "snapshot", "extra": 1})).contains("invalid arguments"));
    }
}

#[cfg(all(test, windows))]
mod live_tests;
