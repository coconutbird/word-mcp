//! `docx_read`: read saved documents without changing them.

use std::{
    fs,
    path::{Path, PathBuf},
};

use anyhow::{Result, ensure};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::{Value, json};

use super::{
    Definition,
    package::{read, write_atomic},
    xml::{Document, escaped},
};
use crate::tool::{Effect, Output, parse, tool};

const NAME: &str = "docx_read";

pub(super) const TOOL: Definition = Definition {
    name: NAME,
    tool: || {
        tool::<Arguments>(
            NAME,
            "Read a saved .docx: paragraphs (with style and table membership, zero-based and paginated) or a logical HTML preview.",
            Effect::ReadOnly,
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
    /// Main-document paragraphs, including table paragraphs, in document order.
    Paragraphs {
        /// Zero-based index of the first paragraph to return.
        start: Option<usize>,
        /// Maximum number of paragraphs to return.
        limit: Option<usize>,
    },
    /// An escaped logical HTML preview (not a page-accurate rendering).
    Preview {
        /// Also write the preview to this `.html` or `.htm` file.
        output_path: Option<PathBuf>,
        /// Replace an existing file at `output_path`.
        #[serde(default)]
        overwrite: bool,
    },
}

fn call(arguments: Value) -> Result<Output> {
    let Arguments { path, operation } = parse(NAME, arguments)?;
    let package = read(&path)?;
    let (_, document) = package.parse()?;
    Ok(match operation {
        Operation::Paragraphs { start, limit } => paragraphs(&path, &document, start, limit),
        Operation::Preview {
            output_path,
            overwrite,
        } => preview(&path, &document, output_path.as_deref(), overwrite)?,
    }
    .into())
}

fn paragraphs(
    path: &Path,
    document: &Document,
    start: Option<usize>,
    limit: Option<usize>,
) -> Value {
    let start = start.unwrap_or(0);
    let mut texts = Vec::new();
    let paragraphs: Vec<Value> = document
        .paragraphs
        .iter()
        .enumerate()
        .skip(start)
        .take(limit.unwrap_or(usize::MAX))
        .map(|(index, paragraph)| {
            let text = document.text(*paragraph);
            let value = json!({
                "index": index,
                "text": text,
                "style": document.style(*paragraph),
                "in_table": document.in_table(*paragraph),
            });
            texts.push(text);
            value
        })
        .collect();
    json!({
        "path": path,
        "paragraphs": paragraphs,
        "text": texts.join("\n"),
        "total_paragraphs": document.paragraphs.len(),
        "start": start,
    })
}

/// Render escaped logical HTML, optionally saving it beside the document.
fn preview(
    path: &Path,
    document: &Document,
    output: Option<&Path>,
    overwrite: bool,
) -> Result<Value> {
    let body = document
        .paragraphs
        .iter()
        .map(|paragraph| {
            format!(
                "<p>{}</p>",
                escaped(&document.text(*paragraph)).replace('\n', "<br>")
            )
        })
        .collect::<Vec<_>>()
        .join("\n");
    let html = format!(
        "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><title>Document preview</title><style>body{{max-width:52rem;margin:2rem auto;font-family:system-ui}}p{{white-space:pre-wrap}}</style></head><body>{body}</body></html>"
    );
    if let Some(output) = output {
        ensure!(
            output
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("html")
                    || extension.eq_ignore_ascii_case("htm")),
            "preview output must use .html or .htm extension"
        );
        ensure!(output != path, "preview cannot overwrite source document");
        if output.exists() {
            ensure!(
                fs::canonicalize(output)? != fs::canonicalize(path)?,
                "preview cannot overwrite source document"
            );
        }
        write_atomic(output, html.as_bytes(), None, overwrite)?;
    }
    Ok(json!({"path": path, "html": html, "output_path": output, "preview_type": "logical_html"}))
}

#[cfg(test)]
mod tests {
    use anyhow::{Context, Result};
    use serde_json::json;

    use super::super::testing::{fixture, run};

    #[test]
    fn unicode_entities_controls_and_preview_roundtrip() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("unicode.docx");
        let text = "é 🍎 <script>& \"quotes\"\nline\ttab";
        run(
            "docx_document",
            json!({"path": path, "operation": {"action": "create", "paragraphs": [text]}}),
        )?;
        let read = run(
            "docx_read",
            json!({"path": path, "operation": {"action": "paragraphs"}}),
        )?;
        assert_eq!(read["text"], text);
        let html = run(
            "docx_read",
            json!({"path": path, "operation": {"action": "preview"}}),
        )?;
        assert!(
            html["html"]
                .as_str()
                .context("missing HTML")?
                .contains("&lt;script&gt;&amp;")
        );
        Ok(())
    }

    #[test]
    fn paragraphs_report_style_and_table_membership() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("styles.docx");
        fixture(
            &path,
            "<q:p><q:pPr><q:pStyle q:val=\"Heading1\"/></q:pPr><q:r><q:t>title</q:t></q:r></q:p><q:tbl><q:tr><q:tc><q:p><q:pPr><q:pStyle q:val=\"TableText\"/></q:pPr><q:r><q:t>cell</q:t></q:r></q:p></q:tc></q:tr></q:tbl><q:p/>",
        )?;
        let read = run(
            "docx_read",
            json!({"path": path, "operation": {"action": "paragraphs", "start": 0}}),
        )?;
        assert_eq!(
            read["paragraphs"],
            json!([
                {"index": 0, "text": "title", "style": "Heading1", "in_table": false},
                {"index": 1, "text": "cell", "style": "TableText", "in_table": true},
                {"index": 2, "text": "", "style": null, "in_table": false},
            ])
        );
        Ok(())
    }
}
