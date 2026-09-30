//! Opt-in interoperability test between the DOCX writer and real Microsoft Word.
use anyhow::{Result, bail};
use serde_json::{Value, json};

use super::{automation_path, testing::with_word};
use crate::tool::Output;

fn docx(tool: &str, path: &str, operation: &Value) -> Result<Value> {
    match crate::docx::call(tool, json!({"path": path, "operation": operation}))? {
        Output::Json(value) => Ok(value),
        Output::Image { .. } => bail!("unexpected image"),
    }
}

#[test]
#[ignore = "Requires desktop Microsoft Word; starts a private hidden instance"]
fn rust_docx_word_interoperability() {
    with_word(|harness| {
        let path = automation_path(&harness.path("rust-created.docx")?);
        let pdf = harness.path("rust-created.pdf")?;
        docx(
            "docx_document",
            &path,
            &json!({"action": "create", "paragraphs": ["Draft 🦀 report", "Revenue <forecast> & actual", "Line one\nLine two\tstop"]}),
        )?;
        docx(
            "docx_edit",
            &path,
            &json!({"action": "replace", "find": "Draft", "replacement": "Final"}),
        )?;
        docx(
            "docx_format",
            &path,
            &json!({"action": "paragraph", "index": 0, "bold": true, "font_size_pt": 18, "alignment": "center", "style": "Normal"}),
        )?;

        let session = &mut harness.session;
        let live = |session: &mut super::Session, tool: &str, operation: Value| {
            session.json(tool, json!({"path": path, "operation": operation}))
        };
        live(
            session,
            "word_live_document",
            json!({"action": "open", "visible": false}),
        )?;
        let read = live(session, "word_live_read", json!({"action": "text"}))?;
        let text = read["text"].as_str().unwrap_or_default();
        assert!(text.contains("Final 🦀 report"));
        assert!(text.contains("Line one\u{b}Line two\tstop"));
        assert!(text.contains("Revenue <forecast> & actual"));

        assert!(
            docx(
                "docx_edit",
                &path,
                &json!({"action": "replace", "find": "Final", "replacement": "Must not overwrite an open document"}),
            )
            .is_err(),
            "saved-file edits must refuse a document open in Word"
        );

        live(
            session,
            "word_live_edit",
            json!({"action": "replace", "find": "Final", "replacement": "Published", "tracked_changes": false}),
        )?;
        live(session, "word_live_document", json!({"action": "save"}))?;
        live(
            session,
            "word_live_document",
            json!({"action": "export_pdf", "output_path": pdf}),
        )?;
        assert!(std::fs::read(&pdf)?.starts_with(b"%PDF-"));
        live(session, "word_live_document", json!({"action": "close"}))?;

        let saved = docx("docx_read", &path, &json!({"action": "paragraphs"}))?;
        assert!(saved.to_string().contains("Published 🦀 report"));
        Ok(())
    });
}
