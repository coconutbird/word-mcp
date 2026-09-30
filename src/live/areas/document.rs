//! `word_live_document`: Word availability, and opening, closing, saving, exporting, and
//! showing documents.

use std::path::PathBuf;

use anyhow::{Result, ensure};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;

use super::{Definition, Operation};
use crate::{
    live::common::{output_path, required_document},
    tool::{Effect, parse, tool},
};

const NAME: &str = "word_live_document";

pub(super) const TOOL: Definition = Definition {
    name: NAME,
    tool: || {
        tool::<Arguments>(
            NAME,
            "Manage documents in desktop Microsoft Word: report Word's status and open documents (never launches Word), open a document with macros disabled (launching Word if needed), close, save, export to PDF through Word's layout engine, or bring a document to the front. Other word_live_* tools need the document open first.",
            Effect::Destructive,
        )
    },
    parse: parse_operation,
};

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct Arguments {
    /// Absolute path of the document. Required by every action except `status`.
    path: Option<PathBuf>,
    operation: Action,
}

#[derive(Deserialize, JsonSchema)]
#[serde(tag = "action", rename_all = "snake_case", deny_unknown_fields)]
enum Action {
    /// Report whether Word is installed and running, and list its open documents.
    Status {},
    /// Open an existing .docx, .docm, .doc, .dotx, .dotm, or .rtf file. Reuses the
    /// document if it is already open.
    Open {
        /// Open the document read-only.
        #[serde(default)]
        read_only: bool,
        /// Show Word and activate the document. An already visible Word window stays
        /// visible either way.
        #[serde(default = "yes")]
        visible: bool,
    },
    /// Close the document. With unsaved changes, `save` is required: `true` saves them,
    /// `false` discards them. Word itself keeps running.
    Close {
        /// Save (`true`) or discard (`false`) unsaved changes.
        save: Option<bool>,
    },
    /// Save the document to its current file.
    Save {},
    /// Export the current layout, including unsaved changes, to PDF. Does not save the
    /// document.
    ExportPdf {
        /// Absolute path of the .pdf file to write.
        output_path: PathBuf,
        /// Replace an existing PDF.
        #[serde(default)]
        overwrite: bool,
    },
    /// Show Word and bring the document to the front.
    View {},
}

fn yes() -> bool {
    true
}

struct Request {
    path: Option<PathBuf>,
    action: Action,
}

fn parse_operation(arguments: Value) -> Result<Box<dyn Operation>> {
    let Arguments { path, operation } = parse(NAME, arguments)?;
    let path = match &operation {
        Action::Status {} => {
            ensure!(path.is_none(), "status takes no path");
            None
        }
        Action::Open { .. } => Some(required_document(path.as_deref(), "open")?),
        Action::Close { .. } => Some(required_document(path.as_deref(), "close")?),
        Action::Save {} => Some(required_document(path.as_deref(), "save")?),
        Action::ExportPdf { .. } => Some(required_document(path.as_deref(), "export_pdf")?),
        Action::View {} => Some(required_document(path.as_deref(), "view")?),
    };
    let action = match operation {
        Action::ExportPdf {
            output_path: output,
            overwrite,
        } => Action::ExportPdf {
            output_path: output_path(&output, &["pdf"], overwrite)?,
            overwrite,
        },
        other => other,
    };
    Ok(Box::new(Request { path, action }))
}

impl Operation for Request {
    #[cfg(windows)]
    fn run(
        self: Box<Self>,
        session: &mut crate::live::word::Session,
    ) -> Result<crate::tool::Output> {
        let path = self.path.as_deref();
        Ok(match self.action {
            Action::Status {} => com::status(session)?,
            Action::Open { read_only, visible } => {
                com::open(session, required(path)?, read_only, visible)?
            }
            Action::Close { save } => com::close(session, required(path)?, save)?,
            Action::Save {} => com::save(session, required(path)?)?,
            Action::ExportPdf {
                output_path,
                overwrite,
            } => com::export_pdf(session, required(path)?, &output_path, overwrite)?,
            Action::View {} => com::view(session, required(path)?)?,
        }
        .into())
    }
}

#[cfg(windows)]
fn required(path: Option<&std::path::Path>) -> Result<&std::path::Path> {
    use anyhow::Context;
    path.context("path is required")
}

#[cfg(windows)]
mod com {
    use std::path::Path;

    use anyhow::{Context, Result, bail};
    use serde_json::{Value, json};

    use crate::live::word::{
        DISCARD_CHANGES, IDispatch, Object, Session, automation_path, find_document, item, writable,
    };

    /// `msoAutomationSecurityForceDisable`: open documents with macros disabled.
    const FORCE_DISABLE_MACROS: i32 = 3;
    /// `wdFormatPDF` for `ExportAsFixedFormat`.
    const FORMAT_PDF: i32 = 17;
    /// `wdSaveChanges` for `Document.Close`.
    const SAVE_CHANGES: i32 = -1;

    pub(super) fn status(session: &mut Session) -> Result<Value> {
        if !Session::installed() {
            return Ok(json!({"installed": false, "running": false, "documents": []}));
        }
        let Some(word) = session.word(false)? else {
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

    pub(super) fn open(
        session: &mut Session,
        path: &Path,
        read_only: bool,
        visible: bool,
    ) -> Result<Value> {
        let word = session.word(true)?.context("cannot start Microsoft Word")?;
        let documents = word.object("Documents")?;
        let existing = find_document(&documents, path)?;
        let reused = existing.is_some();
        let document = match existing {
            Some(document) => document,
            None => open_without_macros(&word, &documents, path, read_only, visible)?,
        };
        if visible {
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

    /// `Documents.Open` with macros force-disabled, restoring Word's security setting.
    fn open_without_macros(
        word: &IDispatch,
        documents: &IDispatch,
        path: &Path,
        read_only: bool,
        visible: bool,
    ) -> Result<Object> {
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
                automation_path(path).as_str().into(),
                false.into(),
                read_only.into(),
                false.into(),
                visible.into(),
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

    pub(super) fn close(session: &mut Session, path: &Path, save: Option<bool>) -> Result<Value> {
        let (_, document) = session.document(path)?;
        let unsaved = !document.flag("Saved")?;
        let save = match save {
            Some(save) => save,
            None if unsaved => bail!(
                "{} has unsaved changes; pass save=true to save them or save=false to discard them",
                path.display()
            ),
            None => false,
        };
        if save && unsaved {
            writable(&document)?;
        }
        let full_name = document.string("FullName")?;
        let disposition = if save { SAVE_CHANGES } else { DISCARD_CHANGES };
        document.call("Close", vec![disposition.into()])?;
        Ok(json!({
            "path": full_name,
            "closed": true,
            "saved_changes": save && unsaved,
            "discarded_changes": !save && unsaved,
        }))
    }

    pub(super) fn save(session: &mut Session, path: &Path) -> Result<Value> {
        let (_, document) = session.document(path)?;
        writable(&document)?;
        document.call("Save", vec![])?;
        Ok(json!({
            "path": document.string("FullName")?,
            "saved": document.flag("Saved")?,
        }))
    }

    pub(super) fn export_pdf(
        session: &mut Session,
        path: &Path,
        output: &Path,
        overwrite: bool,
    ) -> Result<Value> {
        let (_, document) = session.document(path)?;
        // Export beside the target, then publish atomically so a failed export never
        // leaves a truncated PDF behind.
        let temporary = tempfile::Builder::new()
            .prefix(".word-mcp-")
            .suffix(".pdf")
            .tempfile_in(output.parent().context("PDF path has no parent")?)?
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
        if overwrite {
            temporary.persist(output)
        } else {
            temporary.persist_noclobber(output)
        }
        .context("cannot publish the exported PDF")?;
        Ok(json!({
            "output_path": automation_path(output),
            "bytes": output.metadata()?.len(),
        }))
    }

    pub(super) fn view(session: &mut Session, path: &Path) -> Result<Value> {
        let (word, document) = session.document(path)?;
        word.put("Visible", true.into())?;
        document.call("Activate", vec![])?;
        Ok(json!({"visible": true, "path": document.string("FullName")?}))
    }
}

#[cfg(all(test, windows))]
mod tests {
    use serde_json::json;

    use crate::live::word::testing::{is_pdf, text, with_word};

    #[test]
    #[ignore = "Requires desktop Microsoft Word; starts a private hidden instance"]
    fn open_save_export_close_and_reopen() {
        with_word(|harness| {
            let (path, _document) = harness.document("document.docx", "alpha\rbeta")?;
            let pdf = harness.path("document.pdf")?;
            let session = &mut harness.session;
            let tool = |session: &mut crate::live::word::Session, operation: serde_json::Value| {
                session.json(
                    "word_live_document",
                    json!({"path": path, "operation": operation}),
                )
            };
            assert_eq!(
                session.json(
                    "word_live_document",
                    json!({"operation": {"action": "status"}})
                )?["running"],
                true
            );
            assert_eq!(
                tool(session, json!({"action": "open", "visible": false}))?["reused"],
                true
            );
            session.json(
                "word_live_edit",
                json!({"path": path, "operation": {"action": "insert_text", "position": 0, "text": "NEW "}}),
            )?;
            assert!(
                tool(session, json!({"action": "close"})).is_err(),
                "unsaved close needs a decision"
            );
            assert_eq!(tool(session, json!({"action": "save"}))?["saved"], true);
            tool(session, json!({"action": "export_pdf", "output_path": pdf}))?;
            assert!(is_pdf(&pdf)?);
            assert!(tool(session, json!({"action": "export_pdf", "output_path": pdf})).is_err());
            tool(
                session,
                json!({"action": "export_pdf", "output_path": pdf, "overwrite": true}),
            )?;
            assert_eq!(tool(session, json!({"action": "view"}))?["visible"], true);
            assert_eq!(
                tool(session, json!({"action": "close"}))?["discarded_changes"],
                false
            );

            // Opening from disk forces macros off and restores the user's setting.
            let security = harness.word.int("AutomationSecurity")?;
            let session = &mut harness.session;
            assert_eq!(
                tool(session, json!({"action": "open", "visible": false}))?["reused"],
                false
            );
            assert_eq!(harness.word.int("AutomationSecurity")?, security);
            let session = &mut harness.session;
            let read = session.json(
                "word_live_read",
                json!({"path": path, "operation": {"action": "text"}}),
            )?;
            assert!(text(&read).starts_with("NEW alpha"));
            tool(session, json!({"action": "close", "save": false}))?;
            Ok(())
        });
    }
}
