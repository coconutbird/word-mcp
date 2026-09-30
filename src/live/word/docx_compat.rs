//! Opt-in interoperability test between the DOCX writer and real Microsoft Word.
use anyhow::{Context, Result};
use serde_json::json;

use super::{Apartment, Session, TestWord, automation_path, launch_word, word_class};

fn roundtrip() -> Result<()> {
    let _apartment = Apartment::enter()?;
    let directory = tempfile::tempdir()?;
    let path = automation_path(&directory.path().canonicalize()?.join("rust-created.docx"));
    let pdf = automation_path(&directory.path().join("rust-created.pdf"));
    crate::docx::call(
        "create_document",
        json!({
            "path": path,
            "paragraphs": ["Draft 🦀 report", "Revenue <forecast> & actual", "Line one\nLine two\tstop"],
        }),
    )?;
    crate::docx::call(
        "replace_text",
        json!({"path": path, "find": "Draft", "replacement": "Final"}),
    )?;
    crate::docx::call(
        "format_paragraph",
        json!({"path": path, "index": 0, "bold": true, "font_size_pt": 18, "alignment": "center", "style": "Normal"}),
    )?;

    // A fresh instance owned by this test, never the user's Word.
    let word = launch_word(&word_class().context("Word is not installed")?)?;
    let _guard = TestWord(word.clone());
    let mut session = Session { word: Some(word) };
    session.tool("word_live_open", json!({"path": path, "visible": false}))?;
    let read = session.tool("word_live_read", json!({"path": path}))?;
    let text = read["text"].as_str().context("no text")?;
    assert!(text.contains("Final 🦀 report"));
    assert!(text.contains("Line one\u{b}Line two\tstop"));
    assert!(text.contains("Revenue <forecast> & actual"));

    assert!(
        crate::docx::call(
            "replace_text",
            json!({"path": path, "find": "Final", "replacement": "Must not overwrite an open document"}),
        )
        .is_err(),
        "saved-file edits must refuse a document open in Word"
    );

    session.tool(
        "word_live_replace_text",
        json!({"path": path, "find": "Final", "replacement": "Published", "tracked_changes": false}),
    )?;
    session.tool("word_live_save", json!({"path": path}))?;
    session.tool(
        "word_live_export_pdf",
        json!({"path": path, "output_path": pdf}),
    )?;
    assert!(std::fs::read(&pdf)?.starts_with(b"%PDF-"));
    session.tool("word_live_close", json!({"path": path}))?;

    let saved = crate::docx::call("read_document", json!({"path": path}))?;
    assert!(saved.to_string().contains("Published 🦀 report"));
    Ok(())
}

#[test]
#[ignore = "Requires desktop Microsoft Word; starts a private hidden instance"]
fn rust_docx_word_interoperability() {
    std::thread::spawn(roundtrip)
        .join()
        .expect("Word interoperability test thread panicked")
        .expect("a Rust-written DOCX must round-trip through Word");
}
