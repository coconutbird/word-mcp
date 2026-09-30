//! Live Word automation. All COM objects stay on one dedicated STA thread.
use anyhow::{Context, Result, bail};
use serde::Deserialize;
use serde_json::{Value, json};
use std::path::{Path, PathBuf};
#[cfg(windows)]
use std::sync::{Mutex, mpsc};

#[cfg(windows)]
mod windows;

#[cfg(windows)]
type Reply = std::result::Result<Value, String>;
#[cfg(windows)]
struct Job {
    name: String,
    args: Value,
    reply: mpsc::Sender<Reply>,
}

/// Thread-safe facade for Word objects owned by a dedicated COM apartment.
#[derive(Default)]
pub struct LiveBackend {
    #[cfg(windows)]
    sender: Mutex<Option<mpsc::Sender<Job>>>,
}

impl LiveBackend {
    /// Construct the facade without starting a thread or launching Word.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Execute one validated live Word tool on its owning STA thread.
    ///
    /// # Errors
    /// Returns an error for invalid arguments, unavailable Word, or failed automation.
    pub fn call(&self, name: &str, args: Value) -> Result<Value> {
        // Validate before starting Word or handing work to the COM apartment.
        validate(name, &args)?;
        #[cfg(not(windows))]
        {
            bail!(
                "Live Microsoft Word automation requires Windows with desktop Microsoft Word installed; use the DOCX tools on this platform"
            )
        }
        #[cfg(windows)]
        {
            let sender = {
                let mut slot = self
                    .sender
                    .lock()
                    .map_err(|_| anyhow::anyhow!("Word worker lock poisoned"))?;
                if slot.is_none() {
                    let (tx, rx) = mpsc::channel();
                    std::thread::Builder::new()
                        .name("word-com-sta".into())
                        .spawn(move || windows::run(rx))?;
                    *slot = Some(tx);
                }
                slot.as_ref()
                    .context("Word worker channel missing")?
                    .clone()
            };
            let (tx, rx) = mpsc::channel();
            sender
                .send(Job {
                    name: name.into(),
                    args,
                    reply: tx,
                })
                .context("Word worker stopped")?;
            rx.recv()
                .context("Word worker stopped before replying")?
                .map_err(anyhow::Error::msg)
        }
    }
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct EmptyArgs {}
#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct PathArgs {
    path: String,
}
#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct OpenArgs {
    path: String,
    #[cfg_attr(
        not(windows),
        expect(
            dead_code,
            reason = "Automation fields are consumed by the Windows backend"
        )
    )]
    #[serde(default)]
    read_only: bool,
    #[cfg_attr(
        not(windows),
        expect(
            dead_code,
            reason = "Automation fields are consumed by the Windows backend"
        )
    )]
    #[serde(default = "yes")]
    visible: bool,
}
#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct ReadArgs {
    path: String,
    start: Option<i32>,
    end: Option<i32>,
}
#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct ReplaceArgs {
    path: String,
    find: String,
    #[cfg_attr(
        not(windows),
        expect(
            dead_code,
            reason = "Automation fields are consumed by the Windows backend"
        )
    )]
    replacement: String,
    #[cfg_attr(
        not(windows),
        expect(
            dead_code,
            reason = "Automation fields are consumed by the Windows backend"
        )
    )]
    #[serde(default)]
    all: bool,
    #[cfg_attr(
        not(windows),
        expect(
            dead_code,
            reason = "Automation fields are consumed by the Windows backend"
        )
    )]
    tracked_changes: Option<bool>,
}
#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct InsertArgs {
    path: String,
    position: i32,
    #[cfg_attr(
        not(windows),
        expect(
            dead_code,
            reason = "Automation fields are consumed by the Windows backend"
        )
    )]
    text: String,
    #[cfg_attr(
        not(windows),
        expect(
            dead_code,
            reason = "Automation fields are consumed by the Windows backend"
        )
    )]
    tracked_changes: Option<bool>,
}
#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct ExportArgs {
    path: String,
    output_path: String,
    #[serde(default)]
    overwrite: bool,
}
#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct UndoArgs {
    path: String,
    #[serde(default = "one")]
    count: i32,
}
fn yes() -> bool {
    true
}
fn one() -> i32 {
    1
}
fn parse<T: serde::de::DeserializeOwned>(args: &Value) -> Result<T> {
    serde_json::from_value(args.clone()).context("Invalid live Word tool arguments")
}

fn document_path(path: &str) -> Result<PathBuf> {
    let p = Path::new(path);
    if !p.is_absolute() {
        bail!("path must be an absolute document path");
    }
    if !p.is_file() {
        bail!("document does not exist: {}", p.display());
    }
    let extension = p
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or("")
        .to_ascii_lowercase();
    if !matches!(
        extension.as_str(),
        "docx" | "doc" | "docm" | "dotx" | "dotm" | "rtf"
    ) {
        bail!("Unsupported Word document extension");
    }
    p.canonicalize().context("Cannot resolve document path")
}

fn output_path(path: &str, overwrite: bool) -> Result<PathBuf> {
    let p = Path::new(path);
    if !p.is_absolute() || !p.extension().is_some_and(|s| s.eq_ignore_ascii_case("pdf")) {
        bail!("output_path must be an absolute .pdf path");
    }
    if p.exists() && (!overwrite || !p.is_file()) {
        bail!("PDF output already exists; use overwrite=true to replace an existing file");
    }
    let parent = p
        .parent()
        .context("PDF output requires a parent directory")?;
    if !parent.is_dir() {
        bail!("PDF output directory does not exist");
    }
    Ok(parent
        .canonicalize()?
        .join(p.file_name().context("PDF output requires a filename")?))
}

fn validate(name: &str, args: &Value) -> Result<()> {
    match name {
        "word_live_status" => {
            parse::<EmptyArgs>(args)?;
        }
        "word_live_open" => {
            let a: OpenArgs = parse(args)?;
            document_path(&a.path)?;
        }
        "word_live_read" => {
            let a: ReadArgs = parse(args)?;
            document_path(&a.path)?;
            if a.start.is_some_and(|n| n < 0)
                || a.end.is_some_and(|n| n < 0)
                || matches!((a.start,a.end),(Some(s),Some(e)) if s > e)
            {
                bail!("range must satisfy 0 <= start <= end");
            }
        }
        "word_live_replace_text" => {
            let a: ReplaceArgs = parse(args)?;
            document_path(&a.path)?;
            if a.find.is_empty() {
                bail!("find cannot be empty");
            }
            if word_find_text(&a.find).encode_utf16().count() > 255 {
                bail!("find exceeds Word's 255 UTF-16 character search limit after escaping");
            }
        }
        "word_live_insert_text" => {
            let a: InsertArgs = parse(args)?;
            document_path(&a.path)?;
            if a.position < 0 {
                bail!("position must be nonnegative");
            }
        }
        "word_live_save" | "word_live_view" => {
            let a: PathArgs = parse(args)?;
            document_path(&a.path)?;
        }
        "word_live_export_pdf" => {
            let a: ExportArgs = parse(args)?;
            document_path(&a.path)?;
            output_path(&a.output_path, a.overwrite)?;
        }
        "word_live_undo" => {
            let a: UndoArgs = parse(args)?;
            document_path(&a.path)?;
            if !(1..=100).contains(&a.count) {
                bail!("count must be between 1 and 100");
            }
        }
        _ => bail!("Unknown live Word tool: {name}"),
    }
    Ok(())
}

fn word_find_text(text: &str) -> String {
    text.replace('^', "^^")
        .replace('\r', "^p")
        .replace(['\n', '\u{b}'], "^l")
        .replace('\t', "^t")
}

/// Return the live Word tools' strict MCP JSON schemas.
#[must_use]
pub fn tools() -> Vec<Value> {
    let path = json!({"type":"string","description":"Absolute path of the exact document. Except open, the document must already be open in Word."});
    let tracking = json!({"type":"boolean","description":"Temporarily enable/disable Track Changes for this edit; original setting is restored."});
    let mut out = Vec::new();
    let mut add = |name: &str,
                   description: &str,
                   properties: Value,
                   required: &[&str],
                   read_only: bool| {
        out.push(json!({"name":name,"description":description,"inputSchema":{"type":"object","properties":properties,"required":required,"additionalProperties":false},"annotations":{"readOnlyHint":read_only,"destructiveHint":!read_only,"openWorldHint":false}}));
    };
    add(
        "word_live_status",
        "Check whether desktop Word is installed/running and list open documents. Does not launch Word.",
        json!({}),
        &[],
        true,
    );
    add(
        "word_live_open",
        "Open an existing document in Microsoft Word with automation macros disabled. May launch Word. Reuses an already open exact document.",
        json!({"path":path,"read_only":{"type":"boolean","default":false},"visible":{"type":"boolean","default":true}}),
        &["path"],
        false,
    );
    add(
        "word_live_read",
        "Read text from the document's main story, including unsaved live edits. Offsets are Word UTF-16 positions; end is exclusive.",
        json!({"path":path,"start":{"type":"integer","minimum":0},"end":{"type":"integer","minimum":0}}),
        &["path"],
        true,
    );
    add(
        "word_live_replace_text",
        "Replace literal, case-sensitive main-story text in a single Word undo record. Edits remain unsaved until save. all=false replaces the first match.",
        json!({"path":path,"find":{"type":"string","minLength":1},"replacement":{"type":"string"},"all":{"type":"boolean","default":false},"tracked_changes":tracking}),
        &["path", "find", "replacement"],
        false,
    );
    add(
        "word_live_insert_text",
        "Insert text at a Word UTF-16 position in the main story as a single undo record. Edits remain unsaved until save.",
        json!({"path":path,"position":{"type":"integer","minimum":0},"text":{"type":"string"},"tracked_changes":tracking}),
        &["path", "position", "text"],
        false,
    );
    add(
        "word_live_save",
        "Save the exact open document to its current file.",
        json!({"path":path}),
        &["path"],
        false,
    );
    add(
        "word_live_export_pdf",
        "Export the live document through Word's layout engine to PDF, including unsaved changes. Does not save the Word document.",
        json!({"path":path,"output_path":{"type":"string"},"overwrite":{"type":"boolean","default":false}}),
        &["path", "output_path"],
        false,
    );
    add(
        "word_live_undo",
        "Undo recent actions in the exact open Word document. count includes user actions and tool actions.",
        json!({"path":path,"count":{"type":"integer","minimum":1,"maximum":100,"default":1}}),
        &["path"],
        false,
    );
    add(
        "word_live_view",
        "Show Word and activate the exact open document for visual inspection.",
        json!({"path":path}),
        &["path"],
        false,
    );
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn escapes_word_find_special_codes() {
        assert_eq!(word_find_text("literal ^p\r\n\t"), "literal ^^p^p^l^t");
        assert_eq!(word_find_text("left\u{b}right"), "left^lright");
        assert_eq!(word_find_text("left\nright"), "left^lright");
    }
    #[test]
    fn strict_status_arguments() {
        assert!(validate("word_live_status", &json!({"extra":true})).is_err());
        assert!(validate("word_live_status", &json!({})).is_ok());
    }
    #[test]
    fn rejects_relative_paths() {
        assert!(document_path("test.docx").is_err());
    }
    #[test]
    fn backend_is_send_sync() {
        fn check<T: Send + Sync>() {}
        check::<LiveBackend>();
    }
    #[test]
    fn schemas_are_strict_and_unique() {
        let t = tools();
        let mut names = std::collections::HashSet::new();
        for v in t {
            assert!(names.insert(v["name"].as_str().unwrap().to_owned()));
            assert_eq!(v["inputSchema"]["additionalProperties"], false);
        }
    }
}
