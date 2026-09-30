//! Opt-in interoperability test between our DOCX writer and actual Microsoft Word.
use super::{
    Apartment, CLSCTX_LOCAL_SERVER, CLSIDFromProgID, COINIT_APARTMENTTHREADED, CoCreateInstance,
    CoInitializeEx, Dispatch, IDispatch, Result, State, Variant, automation_path, json, w,
};

struct TestWord(Dispatch);

impl Drop for TestWord {
    fn drop(&mut self) {
        // Only the instance created by this test is owned by this guard.
        let _ = self.0.method("Quit", vec![Variant::int(0)]);
    }
}

fn roundtrip() -> Result<()> {
    // SAFETY: This fresh test thread owns its COM apartment until the final guard drops.
    unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) }.ok()?;
    let _apartment = Apartment;
    let directory = tempfile::tempdir()?;
    let path = automation_path(&directory.path().join("rust-created.docx"));
    let pdf = automation_path(&directory.path().join("rust-created.pdf"));
    crate::docx::call(
        "create_document",
        json!({
            "path":path,"paragraphs":["Draft 🦀 report","Revenue <forecast> & actual","Line one\nLine two\tstop"]
        }),
    )?;
    crate::docx::call(
        "replace_text",
        json!({
            "path":path,"find":"Draft","replacement":"Final"
        }),
    )?;
    crate::docx::call(
        "format_paragraph",
        json!({
            "path":path,"index":0,"bold":true,"font_size_pt":18,"alignment":"center","style":"Normal"
        }),
    )?;
    // SAFETY: COM is initialized on this thread; Word's registered CLSID is queried.
    let clsid = unsafe { CLSIDFromProgID(w!("Word.Application")) }?;
    // SAFETY: Requesting the standard IDispatch interface of a fresh local Word instance.
    let app = Dispatch::from_windows(unsafe {
        CoCreateInstance::<_, IDispatch>(&raw const clsid, None, CLSCTX_LOCAL_SERVER)
    }?);
    let _word = TestWord(app.clone());
    let mut state = State { app: Some(app) };
    state.call("word_live_open", &json!({"path":path,"visible":false}))?;
    let read = state.call("word_live_read", &json!({"path":path}))?;
    assert!(read["text"].as_str().unwrap().contains("Final 🦀 report"));
    assert!(
        read["text"]
            .as_str()
            .unwrap()
            .contains("Line one\u{000b}Line two\tstop")
    );
    assert!(
        read["text"]
            .as_str()
            .unwrap()
            .contains("Revenue <forecast> & actual")
    );
    let refusal = crate::docx::call(
        "replace_text",
        json!({
            "path":path,"find":"Final","replacement":"Must not overwrite open document"
        }),
    );
    assert!(
        refusal.is_err(),
        "offline edits must reject a document open in Word"
    );
    state.call(
        "word_live_replace_text",
        &json!({
            "path":path,"find":"Final","replacement":"Published","tracked_changes":false
        }),
    )?;
    state.call("word_live_save", &json!({"path":path}))?;
    state.call(
        "word_live_export_pdf",
        &json!({"path":path,"output_path":pdf}),
    )?;
    assert!(std::fs::read(&pdf)?.starts_with(b"%PDF-"));
    state
        .document(&path)?
        .1
        .method("Close", vec![Variant::int(0)])?;
    let saved = crate::docx::call("read_document", json!({"path":path}))?;
    assert!(saved.to_string().contains("Published 🦀 report"));
    Ok(())
}

#[test]
#[ignore = "Requires installed desktop Microsoft Word; creates its own instance and temporary files"]
fn rust_docx_word_interoperability() {
    std::thread::spawn(roundtrip)
        .join()
        .expect("Word interoperability test thread panicked")
        .expect("Rust DOCX must round-trip through Word");
}
