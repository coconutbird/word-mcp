//! The Word automation thread and the live tool operations it runs.
//!
//! Every COM object stays on this thread. Requests arrive as validated Rust values and
//! results leave as JSON, so nothing apartment-bound crosses the boundary.
use std::{
    path::{Path, PathBuf},
    sync::mpsc::{self, RecvTimeoutError},
    thread,
    time::Duration,
};

use anyhow::{Context, Result, anyhow, bail};
use serde_json::{Value, json};

mod automation;
#[cfg(test)]
mod docx_compat;

use super::args::{
    Alignment, CloseArgs, DeleteArgs, DocumentArgs, ExportArgs, FormatArgs, InsertArgs, OpenArgs,
    ParagraphsArgs, ReadArgs, ReplaceArgs, Request, UndoArgs, word_find_text,
};
#[cfg(test)]
use automation::Variant;
use automation::{Apartment, IDispatch, Object, launch_word, pump, running_word, word_class};

/// How long a caller waits for Word before giving up on a reply.
const REPLY_TIMEOUT: Duration = Duration::from_secs(180);
/// How often the idle thread pumps messages for Word's callbacks.
const IDLE_PUMP: Duration = Duration::from_millis(25);
/// `msoAutomationSecurityForceDisable`: open documents with macros disabled.
const FORCE_DISABLE_MACROS: i32 = 3;
/// `wdFormatPDF` for `ExportAsFixedFormat`.
const FORMAT_PDF: i32 = 17;
/// `wdSaveChanges` and `wdDoNotSaveChanges` for `Document.Close`.
const SAVE_CHANGES: i32 = -1;
const DISCARD_CHANGES: i32 = 0;
/// `wdUnderlineSingle` and `wdUnderlineNone`.
const UNDERLINE_SINGLE: i32 = 1;
const UNDERLINE_NONE: i32 = 0;

struct Job {
    request: Request,
    reply: mpsc::Sender<Result<Value>>,
}

/// The Word automation thread.
pub(super) struct Worker {
    jobs: mpsc::Sender<Job>,
    thread: thread::JoinHandle<()>,
}

impl Worker {
    pub(super) fn start() -> Result<Self> {
        let (jobs, receiver) = mpsc::channel();
        let thread = thread::Builder::new()
            .name("word-sta".into())
            .spawn(move || run(&receiver))
            .context("cannot start the Word automation thread")?;
        Ok(Self { jobs, thread })
    }

    pub(super) fn execute(&self, request: Request) -> Result<Value> {
        let (reply, response) = mpsc::channel();
        self.jobs
            .send(Job { request, reply })
            .map_err(|_| anyhow!("the Word automation thread stopped; retry to restart it"))?;
        match response.recv_timeout(REPLY_TIMEOUT) {
            Ok(result) => result,
            Err(RecvTimeoutError::Timeout) => bail!(
                "Word did not answer within {} seconds; it may be showing a dialog. The operation may still complete, so inspect the document before retrying",
                REPLY_TIMEOUT.as_secs()
            ),
            Err(RecvTimeoutError::Disconnected) => {
                bail!("the Word automation thread stopped before replying; retry to restart it")
            }
        }
    }

    pub(super) fn is_alive(&self) -> bool {
        !self.thread.is_finished()
    }
}

fn run(jobs: &mpsc::Receiver<Job>) {
    let apartment = match Apartment::enter() {
        Ok(apartment) => apartment,
        Err(error) => {
            let message = format!("{error:#}");
            for job in jobs {
                let _ = job.reply.send(Err(anyhow!(message.clone())));
            }
            return;
        }
    };
    let mut session = Session::default();
    loop {
        pump();
        match jobs.recv_timeout(IDLE_PUMP) {
            Ok(job) => {
                let _ = job.reply.send(session.handle(job.request));
            }
            Err(RecvTimeoutError::Timeout) => {}
            Err(RecvTimeoutError::Disconnected) => break,
        }
    }
    // Release every Word object before leaving the apartment. Word keeps running: the
    // user may be working in it.
    drop(session);
    drop(apartment);
}

/// The state of the automation thread: the attached Word instance.
#[derive(Default)]
struct Session {
    word: Option<Object>,
}

impl Session {
    fn handle(&mut self, request: Request) -> Result<Value> {
        match request {
            Request::Status => self.status(),
            Request::Open(args) => self.open(&args),
            Request::Close(args) => self.close(&args),
            Request::Read(args) => self.read(&args),
            Request::Paragraphs(args) => self.paragraphs(&args),
            Request::Replace(args) => self.replace(&args),
            Request::Insert(args) => self.insert(&args),
            Request::Delete(args) => self.delete(&args),
            Request::Format(args) => self.format(&args),
            Request::Save(args) => self.save(&args),
            Request::ExportPdf(args) => self.export_pdf(&args),
            Request::Undo(args) => self.undo(&args),
            Request::View(args) => self.view(&args),
        }
    }

    /// The Word instance: the cached one if still alive, else the running one, else
    /// (only when `launch`) a new one.
    fn word(&mut self, launch: bool) -> Result<Option<Object>> {
        if let Some(word) = &self.word {
            if word.get("Version").is_ok() {
                return Ok(Some(word.clone()));
            }
            self.word = None;
        }
        let class = word_class().context("desktop Microsoft Word is not installed")?;
        let word = match running_word(&class)? {
            Some(word) => word,
            None if launch => launch_word(&class)?,
            None => return Ok(None),
        };
        self.word = Some(word.clone());
        Ok(Some(word))
    }

    /// The Word instance and the open document at `path`.
    fn document(&mut self, path: &Path) -> Result<(Object, Object)> {
        let word = self
            .word(false)?
            .context("Microsoft Word is not running; call word_live_open first")?;
        let documents = word.object("Documents")?;
        let document = find_document(&documents, path)?.with_context(|| {
            format!(
                "{} is not open in Word; call word_live_open first",
                path.display()
            )
        })?;
        Ok((word, document))
    }

    fn status(&mut self) -> Result<Value> {
        if word_class().is_none() {
            return Ok(json!({"installed": false, "running": false, "documents": []}));
        }
        let Some(word) = self.word(false)? else {
            return Ok(json!({"installed": true, "running": false, "documents": []}));
        };
        let documents = word.object("Documents")?;
        let documents = (1..=documents.int("Count")?)
            .map(|index| {
                let document = item(&documents, index)?;
                Ok(json!({
                    "name": document.string("Name")?,
                    "path": document.string("FullName")?,
                    "saved": document.flag("Saved")?,
                    "read_only": document.flag("ReadOnly")?,
                }))
            })
            .collect::<Result<Vec<_>>>()?;
        Ok(json!({
            "installed": true,
            "running": true,
            "version": word.string("Version")?,
            "visible": word.flag("Visible")?,
            "documents": documents,
        }))
    }

    fn open(&mut self, args: &OpenArgs) -> Result<Value> {
        let word = self.word(true)?.context("cannot start Microsoft Word")?;
        let documents = word.object("Documents")?;
        let existing = find_document(&documents, &args.path)?;
        let reused = existing.is_some();
        let document = match existing {
            Some(document) => document,
            None => open_without_macros(&word, &documents, args)?,
        };
        if args.visible {
            word.put("Visible", true.into())?;
            document.call("Activate", vec![])?;
        }
        Ok(json!({
            "path": document.string("FullName")?,
            "reused": reused,
            "read_only": document.flag("ReadOnly")?,
            "visible": word.flag("Visible")?,
        }))
    }

    fn close(&mut self, args: &CloseArgs) -> Result<Value> {
        let (_, document) = self.document(&args.path)?;
        let unsaved = !document.flag("Saved")?;
        let save = match args.save {
            Some(save) => save,
            None if unsaved => bail!(
                "{} has unsaved changes; pass save=true to save them or save=false to discard them",
                args.path.display()
            ),
            None => false,
        };
        if save && unsaved {
            writable(&document)?;
        }
        let path = document.string("FullName")?;
        let disposition = if save { SAVE_CHANGES } else { DISCARD_CHANGES };
        document.call("Close", vec![disposition.into()])?;
        Ok(json!({
            "path": path,
            "closed": true,
            "saved_changes": save && unsaved,
            "discarded_changes": !save && unsaved,
        }))
    }

    fn read(&mut self, args: &ReadArgs) -> Result<Value> {
        let (_, document) = self.document(&args.path)?;
        let content = document.object("Content")?;
        let start = args.start.map_or_else(|| content.int("Start"), Ok)?;
        let end = args.end.map_or_else(|| content.int("End"), Ok)?;
        check_range(&document, start, end)?;
        Ok(json!({
            "path": document.string("FullName")?,
            "text": range(&document, start, end)?.string("Text")?,
            "start": start,
            "end": end,
            "saved": document.flag("Saved")?,
        }))
    }

    fn paragraphs(&mut self, args: &ParagraphsArgs) -> Result<Value> {
        let (_, document) = self.document(&args.path)?;
        let paragraphs = document.object("Paragraphs")?;
        let total = paragraphs.int("Count")?;
        let first = i32::try_from(args.start).context("start is too large")?;
        let mut listed = Vec::new();
        let mut next = if first < total {
            Some(item(&paragraphs, first + 1)?)
        } else {
            None
        };
        let mut index = first;
        while let Some(paragraph) = next {
            if listed.len() == args.limit as usize {
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
            "start": args.start,
            "paragraphs": listed,
        }))
    }

    fn replace(&mut self, args: &ReplaceArgs) -> Result<Value> {
        let (word, document) = self.document(&args.path)?;
        let matches = literal_matches(&document, &args.find, args.all)?;
        if !matches.is_empty() {
            mutation(
                &word,
                &document,
                args.tracked_changes,
                "Word MCP replace text",
                || {
                    // Last to first, so earlier positions stay valid.
                    for &(start, end) in matches.iter().rev() {
                        range(&document, start, end)?
                            .put("Text", args.replacement.as_str().into())?;
                    }
                    Ok(())
                },
            )?;
        }
        Ok(json!({"replacements": matches.len(), "saved": document.flag("Saved")?}))
    }

    fn insert(&mut self, args: &InsertArgs) -> Result<Value> {
        let (word, document) = self.document(&args.path)?;
        check_range(&document, args.position, args.position)?;
        if !args.text.is_empty() {
            mutation(
                &word,
                &document,
                args.tracked_changes,
                "Word MCP insert text",
                || {
                    range(&document, args.position, args.position)?
                        .put("Text", args.text.as_str().into())
                },
            )?;
        }
        Ok(json!({
            "inserted_utf16": args.text.encode_utf16().count(),
            "saved": document.flag("Saved")?,
        }))
    }

    fn delete(&mut self, args: &DeleteArgs) -> Result<Value> {
        let (word, document) = self.document(&args.path)?;
        check_range(&document, args.start, args.end)?;
        let target = range(&document, args.start, args.end)?;
        let deleted = target.string("Text")?;
        mutation(
            &word,
            &document,
            args.tracked_changes,
            "Word MCP delete range",
            || target.call("Delete", vec![]).map(drop),
        )?;
        Ok(json!({"deleted_text": deleted, "saved": document.flag("Saved")?}))
    }

    fn format(&mut self, args: &FormatArgs) -> Result<Value> {
        let (word, document) = self.document(&args.path)?;
        check_range(&document, args.start, args.end)?;
        mutation(
            &word,
            &document,
            args.tracked_changes,
            "Word MCP format",
            || {
                let target = range(&document, args.start, args.end)?;
                // A style first, so explicit character formatting applies on top of it.
                if let Some(style) = &args.style {
                    target.put("Style", style.as_str().into())?;
                }
                let font = target.object("Font")?;
                if let Some(bold) = args.bold {
                    font.put("Bold", bold.into())?;
                }
                if let Some(italic) = args.italic {
                    font.put("Italic", italic.into())?;
                }
                if let Some(underline) = args.underline {
                    let value = if underline {
                        UNDERLINE_SINGLE
                    } else {
                        UNDERLINE_NONE
                    };
                    font.put("Underline", value.into())?;
                }
                if let Some(name) = &args.font_name {
                    font.put("Name", name.as_str().into())?;
                }
                if let Some(size) = args.font_size_pt {
                    font.put("Size", size.into())?;
                }
                if let Some(alignment) = args.alignment {
                    target
                        .object("ParagraphFormat")?
                        .put("Alignment", alignment_value(alignment).into())?;
                }
                Ok(())
            },
        )?;
        Ok(json!({"start": args.start, "end": args.end, "saved": document.flag("Saved")?}))
    }

    fn save(&mut self, args: &DocumentArgs) -> Result<Value> {
        let (_, document) = self.document(&args.path)?;
        writable(&document)?;
        document.call("Save", vec![])?;
        Ok(json!({
            "path": document.string("FullName")?,
            "saved": document.flag("Saved")?,
        }))
    }

    fn export_pdf(&mut self, args: &ExportArgs) -> Result<Value> {
        let (_, document) = self.document(&args.path)?;
        // Export beside the target, then publish atomically so a failed export never
        // leaves a truncated PDF behind.
        let temporary = tempfile::Builder::new()
            .prefix(".word-mcp-")
            .suffix(".pdf")
            .tempfile_in(
                args.output_path
                    .parent()
                    .context("PDF path has no parent")?,
            )?
            .into_temp_path();
        document.call(
            "ExportAsFixedFormat",
            vec![
                automation_path(&temporary).as_str().into(),
                FORMAT_PDF.into(),
                false.into(),
            ],
        )?;
        anyhow::ensure!(
            temporary.metadata()?.len() > 0,
            "Word reported success but wrote an empty PDF"
        );
        if args.overwrite {
            temporary.persist(&args.output_path)
        } else {
            temporary.persist_noclobber(&args.output_path)
        }
        .context("cannot publish the exported PDF")?;
        Ok(json!({
            "output_path": automation_path(&args.output_path),
            "bytes": args.output_path.metadata()?.len(),
        }))
    }

    fn undo(&mut self, args: &UndoArgs) -> Result<Value> {
        let (_, document) = self.document(&args.path)?;
        writable(&document)?;
        let undone = document.call("Undo", vec![args.count.into()])?.flag()?;
        Ok(json!({
            "undone": undone,
            "requested_count": args.count,
            "saved": document.flag("Saved")?,
        }))
    }

    fn view(&mut self, args: &DocumentArgs) -> Result<Value> {
        let (word, document) = self.document(&args.path)?;
        word.put("Visible", true.into())?;
        document.call("Activate", vec![])?;
        Ok(json!({"visible": true, "path": document.string("FullName")?}))
    }
}

/// `Documents.Open` with macros force-disabled, restoring Word's security setting.
fn open_without_macros(word: &IDispatch, documents: &IDispatch, args: &OpenArgs) -> Result<Object> {
    let previous = word.int("AutomationSecurity")?;
    word.put("AutomationSecurity", FORCE_DISABLE_MACROS.into())?;
    let opened = documents.call_named(
        "Open",
        &[
            "FileName",
            "ConfirmConversions",
            "ReadOnly",
            "AddToRecentFiles",
            "Visible",
            "NoEncodingDialog",
        ],
        vec![
            automation_path(&args.path).as_str().into(),
            false.into(),
            args.read_only.into(),
            false.into(),
            args.visible.into(),
            true.into(),
        ],
    );
    let restored = word.put("AutomationSecurity", previous.into());
    match (opened, restored) {
        (Ok(document), Ok(())) => document.into_object()?.context("Word opened no document"),
        (Err(error), Ok(())) => Err(error),
        (Ok(_), Err(error)) => Err(error).context(
            "the document opened, but Word's macro security setting could not be restored",
        ),
        (Err(open), Err(restore)) => bail!(
            "opening the document failed: {open:#}; restoring Word's macro security setting also failed: {restore:#}"
        ),
    }
}

/// The open document whose file is `target`, if any.
fn find_document(documents: &IDispatch, target: &Path) -> Result<Option<Object>> {
    let mut found = None;
    for index in 1..=documents.int("Count")? {
        let document = item(documents, index)?;
        if same_file(target, &document.string("FullName")?) {
            anyhow::ensure!(
                found.is_none(),
                "several open Word documents match {}; close the duplicate",
                target.display()
            );
            found = Some(document);
        }
    }
    Ok(found)
}

/// Item `index` (one-based) of a Word collection.
fn item(collection: &IDispatch, index: i32) -> Result<Object> {
    collection
        .call("Item", vec![index.into()])?
        .into_object()?
        .context("Word returned an empty collection item")
}

/// `Document.Range(start, end)`.
fn range(document: &IDispatch, start: i32, end: i32) -> Result<Object> {
    document
        .call("Range", vec![start.into(), end.into()])?
        .into_object()?
        .context("Word returned no range")
}

/// The display name of a paragraph's style.
fn style_name(paragraph: &IDispatch) -> Result<Option<String>> {
    let style = paragraph.get("Style")?;
    if let Ok(name) = style.string() {
        return Ok(Some(name));
    }
    style
        .into_object()?
        .map(|style| style.string("NameLocal"))
        .transpose()
}

fn alignment_value(alignment: Alignment) -> i32 {
    // WdParagraphAlignment.
    match alignment {
        Alignment::Left => 0,
        Alignment::Center => 1,
        Alignment::Right => 2,
        Alignment::Justify => 3,
    }
}

/// Word-reported ranges of literal matches in the main story.
///
/// Word's own Find locates each match because fields and hidden text make plain-text
/// offsets diverge from Word positions.
fn literal_matches(document: &IDispatch, text: &str, all: bool) -> Result<Vec<(i32, i32)>> {
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

/// A path as Word expects it, without the `\\?\` verbatim prefix.
fn automation_path(path: &Path) -> String {
    let text = path.to_string_lossy();
    if let Some(unc) = text.strip_prefix(r"\\?\UNC\") {
        format!(r"\\{unc}")
    } else {
        text.strip_prefix(r"\\?\").unwrap_or(&text).to_owned()
    }
}

/// Whether Word's `FullName` names the canonical file `target`.
fn same_file(target: &Path, full_name: &str) -> bool {
    let path = PathBuf::from(full_name);
    path.is_absolute() && path.canonicalize().is_ok_and(|path| path == target)
}

fn check_range(document: &IDispatch, start: i32, end: i32) -> Result<()> {
    let content = document.object("Content")?;
    let first = content.int("Start")?;
    let last = content.int("End")?;
    anyhow::ensure!(
        first <= start && start <= end && end <= last,
        "range {start}..{end} is outside the main story ({first}..{last})"
    );
    Ok(())
}

fn writable(document: &IDispatch) -> Result<()> {
    anyhow::ensure!(!document.flag("ReadOnly")?, "the document is read-only");
    Ok(())
}

/// Run `edit` as one custom undo record, optionally with Track Changes forced on or
/// off, restoring the document's setting afterwards even if `edit` fails.
fn mutation(
    word: &IDispatch,
    document: &IDispatch,
    tracking: Option<bool>,
    label: &str,
    edit: impl FnOnce() -> Result<()>,
) -> Result<()> {
    writable(document)?;
    // Custom undo records belong to the active document.
    document.call("Activate", vec![])?;
    let previous = document.flag("TrackRevisions")?;
    let undo = word.object("UndoRecord")?;
    undo.call("StartCustomRecord", vec![label.into()])?;
    let result = tracking
        .map_or(Ok(()), |enabled| {
            document.put("TrackRevisions", enabled.into())
        })
        .and_then(|()| edit());
    let restored = match tracking {
        Some(_) => document.put("TrackRevisions", previous.into()),
        None => Ok(()),
    };
    let ended = undo.call("EndCustomRecord", vec![]);
    result.context("the edit failed; any partial change is one undo step (word_live_undo)")?;
    restored.context("the edit succeeded, but Track Changes could not be restored")?;
    ended.context("the edit succeeded, but Word's undo record could not be closed")?;
    Ok(())
}

#[cfg(test)]
impl Session {
    /// Parse and run one tool call, as the worker does.
    fn tool(&mut self, name: &str, arguments: Value) -> Result<Value> {
        self.handle(Request::parse(name, arguments)?)
    }
}

/// Quits a Word instance that a test created; never the user's.
#[cfg(test)]
struct TestWord(Object);

#[cfg(test)]
impl Drop for TestWord {
    fn drop(&mut self) {
        let _ = self.0.call("Quit", vec![DISCARD_CHANGES.into()]);
    }
}

#[cfg(test)]
mod tests {
    use serde_json::json;

    use super::*;

    #[test]
    fn variants_own_and_convert_values() -> Result<()> {
        assert_eq!(Variant::from(42).int()?, 42);
        assert!((Variant::from(10.5).number()? - 10.5).abs() < f64::EPSILON);
        assert!(Variant::from(true).flag()?);
        assert_eq!(Variant::from("hello 🦀").string()?, "hello 🦀");
        assert!(Variant::from("text").int().is_err());
        Ok(())
    }

    #[test]
    fn automation_paths_drop_verbatim_prefixes() {
        assert_eq!(
            automation_path(Path::new(r"\\?\C:\test.docx")),
            r"C:\test.docx"
        );
        assert_eq!(
            automation_path(Path::new(r"\\?\UNC\server\share\test.docx")),
            r"\\server\share\test.docx"
        );
    }

    fn text(value: &Value) -> &str {
        value["text"].as_str().unwrap_or_default()
    }

    #[expect(
        clippy::too_many_lines,
        reason = "One end-to-end session against a real Word instance"
    )]
    fn roundtrip() -> Result<()> {
        let _apartment = Apartment::enter()?;
        let class = word_class().context("Word is not installed")?;
        // A fresh instance owned by this test, never the user's Word.
        let word = launch_word(&class)?;
        let _guard = TestWord(word.clone());
        let directory = tempfile::tempdir()?;
        let path = automation_path(&directory.path().canonicalize()?.join("live-test.docx"));
        let pdf = automation_path(&directory.path().join("live-test.pdf"));
        let document = word
            .object("Documents")?
            .call("Add", vec![])?
            .into_object()?
            .context("no document")?;
        document
            .object("Content")?
            .put("Text", "alpha 🦀 beta\rsecond paragraph".into())?;
        document.call("SaveAs2", vec![path.as_str().into(), 16.into()])?;
        let mut session = Session {
            word: Some(word.clone()),
        };

        assert_eq!(
            session.tool("word_live_status", json!({}))?["running"],
            true
        );
        let opened = session.tool("word_live_open", json!({"path": path, "visible": false}))?;
        assert_eq!(opened["reused"], true);
        assert!(
            text(&session.tool("word_live_read", json!({"path": path}))?).contains("alpha 🦀 beta")
        );

        let listed = session.tool("word_live_paragraphs", json!({"path": path}))?;
        assert_eq!(listed["total_paragraphs"], 2);
        assert_eq!(listed["paragraphs"][1]["text"], "second paragraph");
        let second_start = listed["paragraphs"][1]["start"].as_i64().unwrap();
        assert_eq!(
            session.tool(
                "word_live_paragraphs",
                json!({"path": path, "start": 1, "limit": 1})
            )?["paragraphs"][0]["index"],
            1
        );

        session.tool(
            "word_live_insert_text",
            json!({"path": path, "position": 0, "text": "PREFIX ", "tracked_changes": false}),
        )?;
        let replaced = session.tool(
            "word_live_replace_text",
            json!({"path": path, "find": "beta", "replacement": "delta", "all": true, "tracked_changes": false}),
        )?;
        assert_eq!(replaced["replacements"], 1);
        let edited = session.tool("word_live_read", json!({"path": path}))?;
        assert!(text(&edited).contains("PREFIX alpha 🦀 delta"));
        assert_eq!(edited["saved"], false);
        assert_eq!(
            session.tool("word_live_undo", json!({"path": path}))?["undone"],
            true
        );
        assert!(
            text(&session.tool("word_live_read", json!({"path": path}))?)
                .contains("PREFIX alpha 🦀 beta")
        );

        // "PREFIX " shifted the second paragraph by seven positions.
        let second = i32::try_from(second_start)? + 7;
        session.tool(
            "word_live_format",
            json!({"path": path, "start": second, "end": second + 6, "bold": true, "font_size_pt": 15.5, "alignment": "center"}),
        )?;
        let formatted = range(&document, second, second + 6)?;
        assert_eq!(formatted.object("Font")?.int("Bold")?, -1);
        assert!((formatted.object("Font")?.get("Size")?.number()? - 15.5).abs() < f64::EPSILON);
        assert_eq!(formatted.object("ParagraphFormat")?.int("Alignment")?, 1);
        let deleted = session.tool(
            "word_live_delete_range",
            json!({"path": path, "start": 0, "end": 7}),
        )?;
        assert_eq!(deleted["deleted_text"], "PREFIX ");

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
        let replaced = session.tool(
            "word_live_replace_text",
            json!({"path": path, "find": "beta", "replacement": "gamma"}),
        )?;
        assert_eq!(replaced["replacements"], 1);
        assert!(
            text(&session.tool("word_live_read", json!({"path": path}))?)
                .contains("alpha 🦀 gamma")
        );

        // Find special codes are literal text, and LF matches Word's manual line break.
        session.tool(
            "word_live_insert_text",
            json!({"path": path, "position": 0, "text": "^p "}),
        )?;
        assert_eq!(
            session.tool(
                "word_live_replace_text",
                json!({"path": path, "find": "^p", "replacement": "[literal]"})
            )?["replacements"],
            1
        );
        session.tool(
            "word_live_insert_text",
            json!({"path": path, "position": 0, "text": "left\u{b}right "}),
        )?;
        assert_eq!(
            session.tool(
                "word_live_replace_text",
                json!({"path": path, "find": "left\nright", "replacement": "LF"})
            )?["replacements"],
            1
        );

        let original = document.flag("TrackRevisions")?;
        session.tool(
            "word_live_insert_text",
            json!({"path": path, "position": 0, "text": "TRACKED ", "tracked_changes": true}),
        )?;
        assert_eq!(document.flag("TrackRevisions")?, original);

        assert!(
            session
                .tool("word_live_close", json!({"path": path}))
                .is_err(),
            "unsaved close needs a decision"
        );
        assert_eq!(
            session.tool("word_live_save", json!({"path": path}))?["saved"],
            true
        );
        session.tool(
            "word_live_export_pdf",
            json!({"path": path, "output_path": pdf}),
        )?;
        assert!(std::fs::read(&pdf)?.starts_with(b"%PDF-"));
        assert!(
            session
                .tool(
                    "word_live_export_pdf",
                    json!({"path": path, "output_path": pdf})
                )
                .is_err()
        );
        session.tool(
            "word_live_export_pdf",
            json!({"path": path, "output_path": pdf, "overwrite": true}),
        )?;
        assert_eq!(
            session.tool("word_live_view", json!({"path": path}))?["visible"],
            true
        );
        let closed = session.tool("word_live_close", json!({"path": path}))?;
        assert_eq!(closed["discarded_changes"], false);
        drop((document, formatted, field_range));

        // Opening from disk forces macros off and restores the user's setting.
        let security = word.int("AutomationSecurity")?;
        let reopened = session.tool("word_live_open", json!({"path": path, "visible": false}))?;
        assert_eq!(reopened["reused"], false);
        assert_eq!(word.int("AutomationSecurity")?, security);
        assert!(text(&session.tool("word_live_read", json!({"path": path}))?).contains("TRACKED"));
        session.tool("word_live_close", json!({"path": path, "save": false}))?;
        Ok(())
    }

    #[test]
    #[ignore = "Requires desktop Microsoft Word; starts a private hidden instance"]
    fn live_word_roundtrip() {
        thread::spawn(roundtrip)
            .join()
            .expect("Word test thread panicked")
            .expect("Word automation roundtrip failed");
    }
}
