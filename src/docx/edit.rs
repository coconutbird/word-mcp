//! `docx_edit`: change the text and paragraph structure of saved documents.

use std::path::PathBuf;

use anyhow::{Context, Result, ensure};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::{Value, json};

use super::{
    Definition,
    package::edit,
    xml::{Document, expand_empty, paragraph_xml, patch, text_xml, validate_text},
};
use crate::tool::{Effect, Output, parse, tool};

const NAME: &str = "docx_edit";

pub(super) const TOOL: Definition = Definition {
    name: NAME,
    tool: || {
        tool::<Arguments>(
            NAME,
            "Edit a saved .docx: replace literal text across runs, insert or delete paragraphs. Paragraph indices are zero-based and include table paragraphs. Every edit keeps a backup.",
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
    path: PathBuf,
    operation: Operation,
}

#[derive(Deserialize, JsonSchema)]
#[serde(tag = "action", rename_all = "snake_case", deny_unknown_fields)]
enum Operation {
    /// Replace literal text, including matches split across runs; the replacement takes
    /// the formatting of the run where the match starts.
    Replace {
        /// Literal text to find.
        #[schemars(length(min = 1))]
        find: String,
        /// Replacement text.
        replacement: String,
        /// Restrict matching to this zero-based paragraph.
        paragraph: Option<usize>,
        /// Replace every match instead of only the first.
        #[serde(default)]
        all: bool,
    },
    /// Insert a paragraph before a zero-based index, or append when omitted.
    InsertParagraph {
        /// Paragraph text; `\n` becomes a line break and `\t` a tab.
        text: String,
        /// Zero-based paragraph index to insert before.
        index: Option<usize>,
        /// Paragraph style id, for example `Heading1`.
        style: Option<String>,
    },
    /// Delete a paragraph. Refuses section breaks, anchored ranges, and the last
    /// paragraph of the body or a table cell.
    DeleteParagraph {
        /// Zero-based paragraph index.
        index: usize,
    },
}

fn call(arguments: Value) -> Result<Output> {
    let Arguments { path, operation } = parse(NAME, arguments)?;
    Ok(edit(&path, |package| {
        let (xml, document) = package.parse()?;
        let (xml, result) = match operation {
            Operation::Replace {
                find,
                replacement,
                paragraph,
                all,
            } => {
                let Some((xml, count)) =
                    replace(&document, xml, &find, &replacement, paragraph, all)?
                else {
                    return Ok(json!({"replacements": 0}));
                };
                (xml, json!({"replacements": count}))
            }
            Operation::InsertParagraph { text, index, style } => (
                insert(&document, xml, &text, index, style.as_deref())?,
                json!({}),
            ),
            Operation::DeleteParagraph { index } => {
                let (xml, deleted) = delete(&document, xml, index)?;
                (xml, json!({"deleted_text": deleted}))
            }
        };
        package.set_document(xml)?;
        Ok(result)
    })?
    .into())
}

/// The element index of the zero-based paragraph `index`.
pub(super) fn paragraph_at(document: &Document, index: usize) -> Result<usize> {
    document
        .paragraphs
        .get(index)
        .copied()
        .context("paragraph index is out of bounds")
}

/// Replace literal text across a paragraph's runs; `None` when nothing matched.
fn replace(
    document: &Document,
    xml: &str,
    find: &str,
    replacement: &str,
    only: Option<usize>,
    all: bool,
) -> Result<Option<(String, usize)>> {
    ensure!(!find.is_empty(), "find must not be empty");
    validate_text(replacement)?;
    if let Some(index) = only {
        paragraph_at(document, index)?;
    }
    let mut changes = Vec::new();
    let mut count = 0;
    for (index, paragraph) in document.paragraphs.iter().copied().enumerate() {
        if only.is_some_and(|selected| index != selected) {
            continue;
        }
        let text = document.text(paragraph);
        let mut matches: Vec<usize> = text.match_indices(find).map(|(offset, _)| offset).collect();
        if matches.is_empty() {
            continue;
        }
        document.editable(paragraph)?;
        document.unanchored(paragraph)?;
        if !all {
            matches.truncate(1);
        }
        let texts = document.texts(paragraph);
        let mut values: Vec<String> = texts
            .iter()
            .map(|index| document.nodes[*index].text.clone())
            .collect();
        let mut offsets = Vec::with_capacity(values.len());
        let mut offset = 0;
        for value in &values {
            offsets.push((offset, offset + value.len()));
            offset += value.len();
        }
        for start in matches.iter().rev().copied() {
            let end = start + find.len();
            let first = offsets
                .iter()
                .position(|(low, high)| *low <= start && start < *high)
                .context("match starts outside editable text")?;
            for (node, (low, high)) in offsets.iter().copied().enumerate() {
                if low >= end || high <= start {
                    continue;
                }
                let local_start = start.saturating_sub(low);
                let local_end = (end - low).min(high - low);
                values[node].replace_range(
                    local_start..local_end,
                    if node == first { replacement } else { "" },
                );
            }
        }
        for (node, value) in texts.into_iter().zip(values) {
            let node = &document.nodes[node];
            if value != node.text {
                changes.push((node.start, node.end, text_xml(&value)));
            }
        }
        count += matches.len();
        if !all {
            break;
        }
    }
    if count == 0 {
        return Ok(None);
    }
    Ok(Some((patch(xml, changes)?, count)))
}

/// Insert a paragraph before the zero-based `index`, or append before the body's
/// section properties.
fn insert(
    document: &Document,
    xml: &str,
    text: &str,
    index: Option<usize>,
    style: Option<&str>,
) -> Result<String> {
    let index = index.unwrap_or(document.paragraphs.len());
    ensure!(
        index <= document.paragraphs.len(),
        "paragraph index is out of bounds"
    );
    let paragraph = paragraph_xml(text, style)?;
    let body = &document.nodes[document.body];
    if body.self_closing() {
        return patch(
            xml,
            vec![(body.start, body.end, expand_empty(xml, body, &paragraph)?)],
        );
    }
    let offset = if let Some(target) = document.paragraphs.get(index) {
        document.editable(*target)?;
        document.nodes[*target].start
    } else {
        document
            .child(document.body, "sectPr")
            .map_or(body.close_start, |node| document.nodes[node].start)
    };
    patch(xml, vec![(offset, offset, paragraph)])
}

/// Remove the zero-based paragraph, returning the new XML and the deleted text.
fn delete(document: &Document, xml: &str, index: usize) -> Result<(String, String)> {
    let paragraph = paragraph_at(document, index)?;
    document.editable(paragraph)?;
    document.unanchored(paragraph)?;
    ensure!(
        document
            .child(paragraph, "pPr")
            .is_none_or(|properties| document.child(properties, "sectPr").is_none()),
        "paragraph holds a section break (w:sectPr); deleting it would merge sections"
    );
    let node = &document.nodes[paragraph];
    let container = node.parent.context("paragraph has no container")?;
    ensure!(
        document
            .children(container)
            .filter(|(_, child)| child.is("p"))
            .nth(1)
            .is_some(),
        "cannot delete the only paragraph in its {} element",
        document.nodes[container].name
    );
    if document.nodes[container].is("tc") {
        // A table cell must end with a paragraph, for example after a nested table.
        let last_block = document
            .children(container)
            .filter(|&(child_index, child)| child_index != paragraph && !child.is("tcPr"))
            .last();
        ensure!(
            last_block.is_some_and(|(_, child)| child.is("p")),
            "a table cell must end with a paragraph"
        );
    }
    Ok((
        patch(xml, vec![(node.start, node.end, String::new())])?,
        document.text(paragraph),
    ))
}

#[cfg(test)]
mod tests {
    use std::fs;

    use anyhow::{Context, Result};
    use serde_json::{Value, json};

    use super::super::{
        testing::{PRESERVED_BYTES, document_xml, fixture, preserved_part, run},
        xml::{Document, WORD_NS},
    };
    use super::*;

    fn edit_op(path: &std::path::Path, operation: &Value) -> Result<Value> {
        run("docx_edit", json!({"path": path, "operation": operation}))
    }

    fn text(path: &std::path::Path) -> Result<Value> {
        Ok(run(
            "docx_read",
            json!({"path": path, "operation": {"action": "paragraphs"}}),
        )?["text"]
            .clone())
    }

    #[test]
    fn split_runs_preserve_format_parts_and_original_backup() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("split.docx");
        fixture(
            &path,
            "<q:p><q:r><q:rPr><q:b/></q:rPr><q:t>café </q:t></q:r><q:r><q:rPr><q:i/></q:rPr><q:t>🍎 today café 🍎</q:t></q:r></q:p>",
        )?;
        let original = fs::read(&path)?;
        let result = edit_op(
            &path,
            &json!({"action": "replace", "find": "café 🍎", "replacement": "A & B", "all": true}),
        )?;
        assert_eq!(result["replacements"], 2);
        let backup = result["backup_path"].as_str().context("missing backup")?;
        assert_eq!(fs::read(backup)?, original);
        let xml = document_xml(&path)?;
        assert!(xml.contains("<q:b/>") && xml.contains("<q:i/>"));
        assert_eq!(preserved_part(&path)?, PRESERVED_BYTES);
        assert_eq!(text(&path)?, "A & B today A & B");
        let unchanged = edit_op(
            &path,
            &json!({"action": "replace", "find": "absent", "replacement": "x"}),
        )?;
        assert_eq!(unchanged["modified"], false);
        Ok(())
    }

    #[test]
    fn insert_empty_body_and_table_indices() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("table.docx");
        fixture(
            &path,
            "<q:p><q:r><q:t>before</q:t></q:r></q:p><q:tbl><q:tr><q:tc><q:p><q:r><q:t>cell</q:t></q:r></q:p></q:tc></q:tr></q:tbl><q:sectPr/>",
        )?;
        edit_op(
            &path,
            &json!({"action": "insert_paragraph", "index": 1, "text": "inserted"}),
        )?;
        assert_eq!(text(&path)?, "before\ninserted\ncell");
        let empty = format!("<q:document xmlns:q=\"{WORD_NS}\"><q:body/></q:document>");
        let document = Document::parse(&empty)?;
        let edited = insert(&document, &empty, "new", None, None)?;
        let parsed = Document::parse(&edited)?;
        assert_eq!(parsed.nodes[parsed.paragraphs[0]].parent, Some(parsed.body));
        Ok(())
    }

    #[test]
    fn delete_paragraph_removes_only_target() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("delete.docx");
        fixture(
            &path,
            "<q:p><q:r><q:t>keep one</q:t></q:r></q:p><q:p><q:r><q:rPr><q:b/></q:rPr><q:t>drop me</q:t></q:r></q:p><q:tbl><q:tr><q:tc><q:p><q:r><q:t>cell a</q:t></q:r></q:p><q:p><q:r><q:t>cell b</q:t></q:r></q:p></q:tc></q:tr></q:tbl><q:p><q:r><q:t>keep two</q:t></q:r></q:p><q:sectPr/>",
        )?;
        let original = fs::read(&path)?;
        let result = edit_op(&path, &json!({"action": "delete_paragraph", "index": 1}))?;
        assert_eq!(result["modified"], true);
        assert_eq!(result["deleted_text"], "drop me");
        let backup = result["backup_path"].as_str().context("missing backup")?;
        assert_eq!(fs::read(backup)?, original);
        assert_eq!(preserved_part(&path)?, PRESERVED_BYTES);
        edit_op(&path, &json!({"action": "delete_paragraph", "index": 2}))?;
        assert_eq!(text(&path)?, "keep one\ncell a\nkeep two");
        assert!(!document_xml(&path)?.contains("<q:b/>"));
        Ok(())
    }

    #[test]
    fn delete_paragraph_refusals_leave_document_unchanged() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("refuse.docx");
        for (body, index) in [
            ("<q:p><q:r><q:t>only</q:t></q:r></q:p><q:sectPr/>", 0),
            (
                "<q:p><q:r><q:t>body</q:t></q:r></q:p><q:tbl><q:tr><q:tc><q:p><q:r><q:t>cell</q:t></q:r></q:p></q:tc></q:tr></q:tbl><q:p/>",
                1,
            ),
            (
                "<q:p><q:pPr><q:sectPr/></q:pPr><q:r><q:t>section one</q:t></q:r></q:p><q:p><q:r><q:t>after</q:t></q:r></q:p><q:sectPr/>",
                0,
            ),
            (
                "<q:p><q:bookmarkStart q:id=\"1\"/><q:r><q:t>marked</q:t></q:r><q:bookmarkEnd q:id=\"1\"/></q:p><q:p/>",
                0,
            ),
            (
                "<q:p/><q:tbl><q:tr><q:tc><q:p/><q:tbl><q:tr><q:tc><q:p/></q:tc></q:tr></q:tbl><q:p/></q:tc></q:tr></q:tbl><q:p/>",
                3,
            ),
            ("<q:p/><q:p/>", 2),
        ] {
            fixture(&path, body)?;
            let before = fs::read(&path)?;
            assert!(
                edit_op(
                    &path,
                    &json!({"action": "delete_paragraph", "index": index})
                )
                .is_err(),
                "deleted paragraph {index} of {body}"
            );
            assert_eq!(fs::read(&path)?, before);
        }
        Ok(())
    }

    #[test]
    fn locks_and_unsupported_edits_leave_original_unchanged() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("locked.docx");
        fixture(&path, "<q:p><q:r><q:t>text</q:t></q:r></q:p>")?;
        let original = fs::read(&path)?;
        let replace = json!({"action": "replace", "find": "text", "replacement": "changed"});
        for name in ["~$locked.docx", "~$cked.docx"] {
            let lock = directory.path().join(name);
            fs::write(&lock, [])?;
            assert!(edit_op(&path, &replace).is_err());
            assert_eq!(fs::read(&path)?, original);
            fs::remove_file(lock)?;
        }
        for body in [
            "<q:sdt><q:sdtContent><q:p><q:r><q:t>text</q:t></q:r></q:p></q:sdtContent></q:sdt>",
            "<q:p><q:fldSimple q:instr=\"DATE\"><q:r><q:t>text</q:t></q:r></q:fldSimple></q:p>",
            "<q:p><q:bookmarkStart q:id=\"1\"/><q:r><q:t>text</q:t></q:r><q:bookmarkEnd q:id=\"1\"/></q:p>",
        ] {
            fixture(&path, body)?;
            let before = fs::read(&path)?;
            assert!(edit_op(&path, &replace).is_err());
            assert_eq!(fs::read(&path)?, before);
        }
        Ok(())
    }
}
