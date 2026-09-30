//! Typed, validated arguments of the live Word tools.
//!
//! Each struct is both the tool's input schema and its parser. [`Request::parse`]
//! validates everything that does not need Word, including resolving the document path,
//! before any work reaches the COM thread.
use std::path::{Path, PathBuf};

use anyhow::{Context, Result, bail, ensure};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;

use crate::tool::parse;

/// Word's limit on the length of a search string, in UTF-16 code units after escaping.
const FIND_LIMIT: usize = 255;

/// No arguments.
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub(super) struct StatusArgs {}

/// Identifies a document that is open in Word.
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub(super) struct DocumentArgs {
    /// Absolute path of a document already open in Word (see `word_live_open`).
    pub path: PathBuf,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub(super) struct OpenArgs {
    /// Absolute path of an existing .docx, .docm, .doc, .dotx, .dotm, or .rtf file.
    pub path: PathBuf,
    /// Open the document read-only.
    #[serde(default)]
    pub read_only: bool,
    /// Show Word and activate the document. A Word window that is already visible stays
    /// visible either way.
    #[serde(default = "yes")]
    pub visible: bool,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub(super) struct CloseArgs {
    /// Absolute path of a document already open in Word.
    pub path: PathBuf,
    /// Save before closing (`true`) or discard unsaved changes (`false`). Required when
    /// the document has unsaved changes.
    pub save: Option<bool>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub(super) struct ReadArgs {
    /// Absolute path of a document already open in Word.
    pub path: PathBuf,
    /// First Word position to read (UTF-16 units). Defaults to the story start.
    #[schemars(range(min = 0))]
    pub start: Option<i32>,
    /// Exclusive end position. Defaults to the story end.
    #[schemars(range(min = 0))]
    pub end: Option<i32>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub(super) struct ParagraphsArgs {
    /// Absolute path of a document already open in Word.
    pub path: PathBuf,
    /// Zero-based index of the first paragraph to list.
    #[serde(default)]
    pub start: u32,
    /// Maximum number of paragraphs to list.
    #[serde(default = "default_paragraph_limit")]
    #[schemars(range(min = 1, max = 500))]
    pub limit: u32,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub(super) struct ReplaceArgs {
    /// Absolute path of a document already open in Word.
    pub path: PathBuf,
    /// Literal, case-sensitive text to find (at most 255 UTF-16 units after escaping).
    #[schemars(length(min = 1))]
    pub find: String,
    /// Replacement text.
    pub replacement: String,
    /// Replace every match instead of only the first.
    #[serde(default)]
    pub all: bool,
    /// Temporarily turn Track Changes on or off for this edit; the document's setting
    /// is restored afterwards. Omit to use the document's setting.
    pub tracked_changes: Option<bool>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub(super) struct InsertArgs {
    /// Absolute path of a document already open in Word.
    pub path: PathBuf,
    /// Word position (UTF-16 units) to insert at.
    #[schemars(range(min = 0))]
    pub position: i32,
    /// Text to insert. `\r` starts a new paragraph; `\u000b` is a line break.
    pub text: String,
    /// Temporarily turn Track Changes on or off for this edit.
    pub tracked_changes: Option<bool>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub(super) struct DeleteArgs {
    /// Absolute path of a document already open in Word.
    pub path: PathBuf,
    /// First Word position to delete.
    #[schemars(range(min = 0))]
    pub start: i32,
    /// Exclusive end position; must be greater than `start`.
    #[schemars(range(min = 1))]
    pub end: i32,
    /// Temporarily turn Track Changes on or off for this edit.
    pub tracked_changes: Option<bool>,
}

/// Paragraph alignment.
#[derive(Clone, Copy, Deserialize, JsonSchema)]
#[serde(rename_all = "lowercase")]
pub(super) enum Alignment {
    Left,
    Center,
    Right,
    Justify,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub(super) struct FormatArgs {
    /// Absolute path of a document already open in Word.
    pub path: PathBuf,
    /// First Word position of the range.
    #[schemars(range(min = 0))]
    pub start: i32,
    /// Exclusive end position; must be greater than `start`.
    #[schemars(range(min = 1))]
    pub end: i32,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    /// Font family name.
    #[schemars(length(min = 1))]
    pub font_name: Option<String>,
    /// Font size in points (half-point steps).
    #[schemars(range(min = 1.0, max = 1638.0))]
    pub font_size_pt: Option<f64>,
    /// Style name (for example "Heading 1"), applied to the range.
    #[schemars(length(min = 1))]
    pub style: Option<String>,
    /// Alignment of the paragraphs the range touches.
    pub alignment: Option<Alignment>,
    /// Temporarily turn Track Changes on or off for this edit.
    pub tracked_changes: Option<bool>,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub(super) struct ExportArgs {
    /// Absolute path of a document already open in Word.
    pub path: PathBuf,
    /// Absolute path of the .pdf file to write.
    pub output_path: PathBuf,
    /// Replace an existing PDF.
    #[serde(default)]
    pub overwrite: bool,
}

#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
pub(super) struct UndoArgs {
    /// Absolute path of a document already open in Word.
    pub path: PathBuf,
    /// Number of actions to undo, including the user's own actions.
    #[serde(default = "one")]
    #[schemars(range(min = 1, max = 100))]
    pub count: i32,
}

fn yes() -> bool {
    true
}

fn one() -> i32 {
    1
}

fn default_paragraph_limit() -> u32 {
    100
}

/// A validated live Word operation, ready for the COM thread.
pub(super) enum Request {
    Status,
    Open(OpenArgs),
    Close(CloseArgs),
    Read(ReadArgs),
    Paragraphs(ParagraphsArgs),
    Replace(ReplaceArgs),
    Insert(InsertArgs),
    Delete(DeleteArgs),
    Format(FormatArgs),
    Save(DocumentArgs),
    ExportPdf(ExportArgs),
    Undo(UndoArgs),
    View(DocumentArgs),
}

impl Request {
    /// Parse and validate one live tool call. Document paths come back canonical.
    pub(super) fn parse(name: &str, arguments: Value) -> Result<Self> {
        Ok(match name {
            "word_live_status" => {
                parse::<StatusArgs>(name, arguments)?;
                Self::Status
            }
            "word_live_open" => {
                let mut args: OpenArgs = parse(name, arguments)?;
                args.path = document_path(&args.path)?;
                Self::Open(args)
            }
            "word_live_close" => {
                let mut args: CloseArgs = parse(name, arguments)?;
                args.path = document_path(&args.path)?;
                Self::Close(args)
            }
            "word_live_read" => {
                let mut args: ReadArgs = parse(name, arguments)?;
                args.path = document_path(&args.path)?;
                ensure!(
                    args.start.is_none_or(|start| start >= 0)
                        && args.end.is_none_or(|end| end >= 0)
                        && !matches!((args.start, args.end), (Some(start), Some(end)) if start > end),
                    "range must satisfy 0 <= start <= end"
                );
                Self::Read(args)
            }
            "word_live_paragraphs" => {
                let mut args: ParagraphsArgs = parse(name, arguments)?;
                args.path = document_path(&args.path)?;
                ensure!(
                    (1..=500).contains(&args.limit),
                    "limit must be between 1 and 500"
                );
                Self::Paragraphs(args)
            }
            "word_live_replace_text" => {
                let mut args: ReplaceArgs = parse(name, arguments)?;
                args.path = document_path(&args.path)?;
                ensure!(!args.find.is_empty(), "find must not be empty");
                ensure!(
                    word_find_text(&args.find).encode_utf16().count() <= FIND_LIMIT,
                    "find exceeds Word's {FIND_LIMIT} UTF-16 unit search limit after escaping"
                );
                Self::Replace(args)
            }
            "word_live_insert_text" => {
                let mut args: InsertArgs = parse(name, arguments)?;
                args.path = document_path(&args.path)?;
                ensure!(args.position >= 0, "position must be nonnegative");
                Self::Insert(args)
            }
            "word_live_delete_range" => {
                let mut args: DeleteArgs = parse(name, arguments)?;
                args.path = document_path(&args.path)?;
                ensure!(
                    0 <= args.start && args.start < args.end,
                    "range must satisfy 0 <= start < end"
                );
                Self::Delete(args)
            }
            "word_live_format" => {
                let mut args: FormatArgs = parse(name, arguments)?;
                args.path = document_path(&args.path)?;
                validate_format(&args)?;
                Self::Format(args)
            }
            "word_live_save" => {
                let mut args: DocumentArgs = parse(name, arguments)?;
                args.path = document_path(&args.path)?;
                Self::Save(args)
            }
            "word_live_view" => {
                let mut args: DocumentArgs = parse(name, arguments)?;
                args.path = document_path(&args.path)?;
                Self::View(args)
            }
            "word_live_export_pdf" => {
                let mut args: ExportArgs = parse(name, arguments)?;
                args.path = document_path(&args.path)?;
                args.output_path = pdf_path(&args.output_path, args.overwrite)?;
                Self::ExportPdf(args)
            }
            "word_live_undo" => {
                let mut args: UndoArgs = parse(name, arguments)?;
                args.path = document_path(&args.path)?;
                ensure!(
                    (1..=100).contains(&args.count),
                    "count must be between 1 and 100"
                );
                Self::Undo(args)
            }
            _ => bail!("unknown live Word tool: {name}"),
        })
    }
}

fn validate_format(args: &FormatArgs) -> Result<()> {
    ensure!(
        0 <= args.start && args.start < args.end,
        "range must satisfy 0 <= start < end"
    );
    ensure!(
        args.bold.is_some()
            || args.italic.is_some()
            || args.underline.is_some()
            || args.font_name.is_some()
            || args.font_size_pt.is_some()
            || args.style.is_some()
            || args.alignment.is_some(),
        "provide at least one formatting property"
    );
    if let Some(size) = args.font_size_pt {
        ensure!(
            size.is_finite() && (1.0..=1638.0).contains(&size) && (size * 2.0).fract() == 0.0,
            "font_size_pt must be 1-1638 in half-point steps"
        );
    }
    ensure!(
        args.font_name
            .as_deref()
            .is_none_or(|name| !name.trim().is_empty())
            && args
                .style
                .as_deref()
                .is_none_or(|style| !style.trim().is_empty()),
        "font_name and style must not be blank"
    );
    Ok(())
}

/// Resolve an existing Word document to its canonical absolute path.
fn document_path(path: &Path) -> Result<PathBuf> {
    ensure!(
        path.is_absolute(),
        "path must be absolute: {}",
        path.display()
    );
    ensure!(
        path.is_file(),
        "document does not exist: {}",
        path.display()
    );
    let extension = path
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    ensure!(
        matches!(
            extension.as_str(),
            "docx" | "docm" | "doc" | "dotx" | "dotm" | "rtf"
        ),
        "not a Word document: {}",
        path.display()
    );
    path.canonicalize()
        .with_context(|| format!("cannot resolve {}", path.display()))
}

/// Validate a PDF destination and return it with a canonical parent directory.
fn pdf_path(path: &Path, overwrite: bool) -> Result<PathBuf> {
    ensure!(
        path.is_absolute()
            && path
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("pdf")),
        "output_path must be an absolute .pdf path"
    );
    if path.exists() {
        ensure!(path.is_file(), "output_path is not a file");
        ensure!(
            overwrite,
            "PDF already exists; pass overwrite=true to replace it"
        );
    }
    let parent = path
        .parent()
        .context("output_path has no parent directory")?;
    ensure!(parent.is_dir(), "output directory does not exist");
    Ok(parent
        .canonicalize()?
        .join(path.file_name().context("output_path has no file name")?))
}

/// Escape literal text for Word's Find, which treats `^` as a special-code prefix.
pub(super) fn word_find_text(text: &str) -> String {
    text.replace('^', "^^")
        .replace('\r', "^p")
        .replace(['\n', '\u{b}'], "^l")
        .replace('\t', "^t")
}

#[cfg(test)]
mod tests {
    use serde_json::json;

    use super::*;

    #[test]
    fn escapes_word_find_special_codes() {
        assert_eq!(word_find_text("literal ^p\r\n\t"), "literal ^^p^p^l^t");
        assert_eq!(word_find_text("left\u{b}right"), "left^lright");
    }

    #[test]
    fn rejects_unknown_arguments_and_relative_paths() {
        assert!(Request::parse("word_live_status", json!({"extra": true})).is_err());
        assert!(Request::parse("word_live_status", json!({})).is_ok());
        assert!(Request::parse("word_live_read", json!({"path": "test.docx"})).is_err());
    }

    #[test]
    fn validates_ranges_and_formatting_before_word() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("doc.docx");
        std::fs::write(&path, b"")?;
        let request = |name: &str, extra: Value| {
            let mut arguments = json!({"path": path});
            arguments
                .as_object_mut()
                .unwrap()
                .extend(extra.as_object().unwrap().clone());
            Request::parse(name, arguments)
        };
        assert!(request("word_live_read", json!({"start": 5, "end": 2})).is_err());
        assert!(request("word_live_delete_range", json!({"start": 3, "end": 3})).is_err());
        assert!(request("word_live_format", json!({"start": 0, "end": 4})).is_err());
        assert!(
            request(
                "word_live_format",
                json!({"start": 0, "end": 4, "font_size_pt": 10.3})
            )
            .is_err()
        );
        assert!(
            request(
                "word_live_format",
                json!({"start": 0, "end": 4, "alignment": "center", "font_size_pt": 10.5})
            )
            .is_ok()
        );
        assert!(
            request(
                "word_live_replace_text",
                json!({"find": "^".repeat(128), "replacement": ""})
            )
            .is_err()
        );
        Ok(())
    }
}
