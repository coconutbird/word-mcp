//! The Word automation thread, its session, and helpers shared by the live areas.
//!
//! Every COM object stays on this thread. Operations arrive as validated Rust values and
//! results leave as JSON or PNG bytes, so nothing apartment-bound crosses the boundary.
use std::{
    path::{Path, PathBuf},
    sync::mpsc::{self, RecvTimeoutError},
    thread,
    time::Duration,
};

use anyhow::{Context, Result, anyhow, bail};

mod automation;
#[cfg(test)]
mod docx_compat;

#[cfg(test)]
pub(in crate::live) use automation::Variant;
use automation::{Apartment, launch_word, pump, running_word, word_class};
pub(in crate::live) use automation::{IDispatch, Object};

use super::areas::Operation;
use crate::tool::Output;

/// How long a caller waits for Word before giving up on a reply.
const REPLY_TIMEOUT: Duration = Duration::from_secs(180);
/// How often the idle thread pumps messages for Word's callbacks.
const IDLE_PUMP: Duration = Duration::from_millis(25);
/// `wdDoNotSaveChanges` for `Document.Close` and `Application.Quit`.
pub(in crate::live) const DISCARD_CHANGES: i32 = 0;

struct Job {
    operation: Box<dyn Operation>,
    reply: mpsc::Sender<Result<Output>>,
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

    pub(super) fn execute(&self, operation: Box<dyn Operation>) -> Result<Output> {
        let (reply, response) = mpsc::channel();
        self.jobs
            .send(Job { operation, reply })
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
                let _ = job.reply.send(job.operation.run(&mut session));
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

/// The state of the automation thread: the attached Word instance, plus per-document
/// state that areas keep between calls (for example read snapshots).
#[derive(Default)]
pub(in crate::live) struct Session {
    word: Option<Object>,
    /// Paragraph texts captured by `word_live_read` snapshots, by canonical path.
    #[expect(dead_code, reason = "SCAFFOLD: used by areas under construction")]
    pub(in crate::live) snapshots: std::collections::HashMap<PathBuf, Vec<String>>,
}

impl Session {
    /// The Word instance: the cached one if still alive, else the running one, else
    /// (only when `launch`) a new one.
    pub(in crate::live) fn word(&mut self, launch: bool) -> Result<Option<Object>> {
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

    /// Whether desktop Word is registered on this machine.
    pub(in crate::live) fn installed() -> bool {
        word_class().is_some()
    }

    /// The Word instance and the open document at `path`.
    pub(in crate::live) fn document(&mut self, path: &Path) -> Result<(Object, Object)> {
        let word = self
            .word(false)?
            .context("Microsoft Word is not running; open the document first")?;
        let documents = word.object("Documents")?;
        let document = find_document(&documents, path)?.with_context(|| {
            format!(
                "{} is not open in Word; open it with word_live_document first",
                path.display()
            )
        })?;
        Ok((word, document))
    }
}

/// The open document whose file is `target`, if any.
pub(in crate::live) fn find_document(
    documents: &IDispatch,
    target: &Path,
) -> Result<Option<Object>> {
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
pub(in crate::live) fn item(collection: &IDispatch, index: i32) -> Result<Object> {
    collection
        .call("Item", vec![index.into()])?
        .into_object()?
        .context("Word returned an empty collection item")
}

/// Every item of a Word collection, in order.
#[expect(dead_code, reason = "SCAFFOLD: used by areas under construction")]
pub(in crate::live) fn items(collection: &IDispatch) -> Result<Vec<Object>> {
    (1..=collection.int("Count")?)
        .map(|index| item(collection, index))
        .collect()
}

/// `Document.Range(start, end)`.
pub(in crate::live) fn range(document: &IDispatch, start: i32, end: i32) -> Result<Object> {
    document
        .call("Range", vec![start.into(), end.into()])?
        .into_object()?
        .context("Word returned no range")
}

/// The range of zero-based paragraph `index` of the main story.
#[expect(dead_code, reason = "SCAFFOLD: used by areas under construction")]
pub(in crate::live) fn paragraph_range(document: &IDispatch, index: usize) -> Result<Object> {
    let paragraphs = document.object("Paragraphs")?;
    let count = paragraphs.int("Count")?;
    let number = i32::try_from(index)?
        .checked_add(1)
        .filter(|number| *number <= count)
        .with_context(|| {
            format!("paragraph {index} is out of bounds (the document has {count})")
        })?;
    item(&paragraphs, number)?.object("Range")
}

/// A path as Word expects it, without the `\\?\` verbatim prefix.
pub(in crate::live) fn automation_path(path: &Path) -> String {
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

/// Refuse a range outside the main story.
pub(in crate::live) fn check_range(document: &IDispatch, start: i32, end: i32) -> Result<()> {
    let content = document.object("Content")?;
    let first = content.int("Start")?;
    let last = content.int("End")?;
    anyhow::ensure!(
        first <= start && start <= end && end <= last,
        "range {start}..{end} is outside the main story ({first}..{last})"
    );
    Ok(())
}

/// Refuse a read-only document.
pub(in crate::live) fn writable(document: &IDispatch) -> Result<()> {
    anyhow::ensure!(!document.flag("ReadOnly")?, "the document is read-only");
    Ok(())
}

/// Run `edit` as one custom undo record, optionally with Track Changes forced on or
/// off, restoring the document's setting afterwards even if `edit` fails.
pub(in crate::live) fn mutation<T>(
    word: &IDispatch,
    document: &IDispatch,
    tracking: Option<bool>,
    label: &str,
    edit: impl FnOnce() -> Result<T>,
) -> Result<T> {
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
    let value = result
        .context("the edit failed; any partial change is one undo step (word_live_edit undo)")?;
    restored.context("the edit succeeded, but Track Changes could not be restored")?;
    ended.context("the edit succeeded, but Word's undo record could not be closed")?;
    Ok(value)
}

/// Test support: a private Word instance and helpers to drive tools against it.
#[cfg(test)]
pub(in crate::live) mod testing {
    use std::path::{Path, PathBuf};

    use anyhow::{Context, Result, bail};
    use serde_json::Value;

    use super::{
        Apartment, DISCARD_CHANGES, Object, Session, automation_path, launch_word, word_class,
    };
    use crate::tool::Output;

    impl Session {
        /// Parse and run one tool call, as the worker does.
        pub(in crate::live) fn tool(&mut self, name: &str, arguments: Value) -> Result<Output> {
            crate::live::areas::parse(name, arguments)?.run(self)
        }

        /// Run one tool call that returns JSON.
        pub(in crate::live) fn json(&mut self, name: &str, arguments: Value) -> Result<Value> {
            match self.tool(name, arguments)? {
                Output::Json(value) => Ok(value),
                Output::Image { .. } => bail!("{name} returned an image"),
            }
        }
    }

    /// Quits a Word instance that a test created; never the user's.
    pub(in crate::live) struct TestWord(pub(in crate::live) Object);

    impl Drop for TestWord {
        fn drop(&mut self) {
            let _ = self.0.call("Quit", vec![DISCARD_CHANGES.into()]);
        }
    }

    /// Everything a live test needs: a session bound to a private hidden Word
    /// instance and a temporary directory.
    pub(in crate::live) struct Harness {
        pub(in crate::live) session: Session,
        pub(in crate::live) word: Object,
        pub(in crate::live) directory: tempfile::TempDir,
    }

    impl Harness {
        /// Create `name` in the temporary directory holding `text` (paragraphs separated
        /// by `\r`), saved as .docx and left open. Returns its path and document.
        pub(in crate::live) fn document(
            &mut self,
            name: &str,
            text: &str,
        ) -> Result<(String, Object)> {
            let path = automation_path(&self.directory.path().canonicalize()?.join(name));
            let document = self
                .word
                .object("Documents")?
                .call("Add", vec![])?
                .into_object()?
                .context("no document")?;
            document.object("Content")?.put("Text", text.into())?;
            // wdFormatXMLDocument
            document.call("SaveAs2", vec![path.as_str().into(), 16.into()])?;
            Ok((path, document))
        }

        /// A path in the temporary directory.
        pub(in crate::live) fn path(&self, name: &str) -> Result<PathBuf> {
            Ok(self.directory.path().canonicalize()?.join(name))
        }
    }

    /// Run `test` on its own STA thread against a fresh hidden Word instance, which is
    /// quit afterwards. Panics when the test fails.
    pub(in crate::live) fn with_word(
        test: impl FnOnce(&mut Harness) -> Result<()> + Send + 'static,
    ) {
        std::thread::spawn(move || -> Result<()> {
            let _apartment = Apartment::enter()?;
            let class = word_class().context("Word is not installed")?;
            let word = launch_word(&class)?;
            let _guard = TestWord(word.clone());
            let mut harness = Harness {
                session: Session {
                    word: Some(word.clone()),
                    ..Session::default()
                },
                word,
                directory: tempfile::tempdir()?,
            };
            let result = test(&mut harness);
            drop(harness);
            result
        })
        .join()
        .expect("Word test thread panicked")
        .expect("Word test failed");
    }

    /// Text of a tool result field.
    pub(in crate::live) fn text(value: &Value) -> &str {
        value["text"].as_str().unwrap_or_default()
    }

    /// Whether `path` names a PDF file.
    pub(in crate::live) fn is_pdf(path: &Path) -> Result<bool> {
        Ok(std::fs::read(path)?.starts_with(b"%PDF-"))
    }
}

#[cfg(test)]
mod tests {
    use anyhow::Result;

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
}
