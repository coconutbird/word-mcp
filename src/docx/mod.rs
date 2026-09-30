//! Native OOXML operations on saved DOCX files.
//!
//! Edits patch only the targeted spans of `word/document.xml`, keep every other
//! package part byte-for-byte, refuse documents Word holds open, detect concurrent
//! changes, and leave a backup of the original beside the document.

mod edit;
mod package;
mod xml;

use std::path::PathBuf;

use anyhow::{Result, bail};
use rmcp::model::Tool;
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::{Value, json};

use crate::tool::{Effect, parse, tool};
use package::{Package, save_edit};

/// Arguments for `create_document`.
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct CreateDocumentArgs {
    /// Local DOCX path to create.
    #[schemars(length(min = 1))]
    path: PathBuf,
    /// Paragraph texts; `\n` becomes a line break and `\t` a tab. Defaults to one empty paragraph.
    #[serde(default)]
    paragraphs: Vec<String>,
    /// Replace an existing file at `path`.
    #[serde(default)]
    overwrite: bool,
}

/// Arguments for `read_document`.
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct ReadDocumentArgs {
    /// Local DOCX path.
    #[schemars(length(min = 1))]
    path: PathBuf,
    /// Zero-based index of the first paragraph to return.
    start: Option<usize>,
    /// Maximum number of paragraphs to return.
    limit: Option<usize>,
}

/// Arguments for `get_document_info`.
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct GetDocumentInfoArgs {
    /// Local DOCX path.
    #[schemars(length(min = 1))]
    path: PathBuf,
}

/// Arguments for `replace_text`.
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct ReplaceTextArgs {
    /// Local DOCX path.
    #[schemars(length(min = 1))]
    path: PathBuf,
    /// Literal text to find; matches may span runs.
    #[schemars(length(min = 1))]
    find: String,
    /// Replacement text; takes the formatting of the run where the match starts.
    replacement: String,
    /// Restrict matching to this zero-based paragraph.
    paragraph_index: Option<usize>,
    /// Replace every match instead of only the first.
    #[serde(default)]
    replace_all: bool,
}

/// Arguments for `insert_paragraph`.
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct InsertParagraphArgs {
    /// Local DOCX path.
    #[schemars(length(min = 1))]
    path: PathBuf,
    /// Paragraph text; `\n` becomes a line break and `\t` a tab.
    text: String,
    /// Zero-based paragraph index to insert before; appends when omitted.
    index: Option<usize>,
    /// Paragraph style id, e.g. `Heading1`.
    style: Option<String>,
}

/// Horizontal paragraph alignment.
#[derive(Clone, Copy, Deserialize, JsonSchema)]
#[serde(rename_all = "lowercase")]
enum Alignment {
    /// Align to the left margin.
    Left,
    /// Center between the margins.
    Center,
    /// Align to the right margin.
    Right,
    /// Justify to both margins.
    Both,
}

impl Alignment {
    /// The `w:jc` value.
    fn val(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Both => "both",
        }
    }
}

/// Arguments for `format_paragraph`.
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct FormatParagraphArgs {
    /// Local DOCX path.
    #[schemars(length(min = 1))]
    path: PathBuf,
    /// Zero-based paragraph index.
    index: usize,
    /// Set or clear bold on every run.
    bold: Option<bool>,
    /// Set or clear italic on every run.
    italic: Option<bool>,
    /// Set single underline or remove underline on every run.
    underline: Option<bool>,
    /// Font size in points, in half-point increments.
    #[schemars(range(min = 1, max = 1638))]
    font_size_pt: Option<f64>,
    /// Font family applied to every script.
    #[schemars(length(min = 1))]
    font_family: Option<String>,
    /// Paragraph alignment.
    alignment: Option<Alignment>,
    /// Paragraph style id, e.g. `Heading1`.
    style: Option<String>,
}

/// Arguments for `delete_paragraph`.
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct DeleteParagraphArgs {
    /// Local DOCX path.
    #[schemars(length(min = 1))]
    path: PathBuf,
    /// Zero-based paragraph index, as returned by `read_document`.
    index: usize,
}

/// Arguments for `preview_document`.
#[derive(Deserialize, JsonSchema)]
#[serde(deny_unknown_fields)]
struct PreviewDocumentArgs {
    /// Local DOCX path.
    #[schemars(length(min = 1))]
    path: PathBuf,
    /// Also write the preview to this `.html` or `.htm` file.
    output_path: Option<PathBuf>,
    /// Replace an existing file at `output_path`.
    #[serde(default)]
    overwrite: bool,
}

/// Return the native DOCX tool definitions.
#[must_use]
pub fn tools() -> Vec<Tool> {
    vec![
        tool::<CreateDocumentArgs>(
            "create_document",
            "Create a DOCX package. Existing output requires overwrite:true.",
            Effect::Additive,
        ),
        tool::<ReadDocumentArgs>(
            "read_document",
            "Read main-document paragraphs, including tables, with zero-based pagination. Each paragraph reports its style id and whether it is in a table.",
            Effect::ReadOnly,
        ),
        tool::<ReplaceTextArgs>(
            "replace_text",
            "Replace literal text across plain runs, preserving surrounding formatting. Edits create a backup.",
            Effect::Destructive,
        ),
        tool::<InsertParagraphArgs>(
            "insert_paragraph",
            "Insert a paragraph before a zero-based index or append. Edits create a backup.",
            Effect::Destructive,
        ),
        tool::<FormatParagraphArgs>(
            "format_paragraph",
            "Set direct paragraph/run formatting. Edits create a backup.",
            Effect::Destructive,
        ),
        tool::<DeleteParagraphArgs>(
            "delete_paragraph",
            "Delete a zero-based paragraph (read_document indexing). Refuses section breaks, anchored ranges, and the last paragraph of the body or a table cell. Edits create a backup.",
            Effect::Destructive,
        ),
        tool::<GetDocumentInfoArgs>(
            "get_document_info",
            "Inspect paragraph count, character count and package parts.",
            Effect::ReadOnly,
        ),
        tool::<PreviewDocumentArgs>(
            "preview_document",
            "Return an escaped logical HTML preview; optional output requires overwrite:true if it exists.",
            Effect::Additive,
        ),
    ]
}

/// Dispatch a native DOCX operation.
///
/// # Errors
/// Returns an error for an unknown operation, invalid arguments, or an invalid or
/// unsupported document.
pub fn call(name: &str, args: Value) -> Result<Value> {
    match name {
        "create_document" => edit::create(&parse(name, args)?),
        "read_document" => read(&parse(name, args)?),
        "get_document_info" => info(&parse(name, args)?),
        "preview_document" => {
            let arguments: PreviewDocumentArgs = parse(name, args)?;
            let (_, document) = Package::open(&arguments.path)?.parse()?;
            edit::preview(&arguments, &document)
        }
        "replace_text" => {
            let arguments: ReplaceTextArgs = parse(name, args)?;
            let package = Package::open(&arguments.path)?;
            let (xml, document) = package.parse()?;
            let Some((xml, count)) = edit::replace(&arguments, &document, xml)? else {
                return Ok(json!({"path":arguments.path,"modified":false,"replacements":0}));
            };
            let mut result = save_edit(&arguments.path, package, xml)?;
            result["replacements"] = json!(count);
            Ok(result)
        }
        "insert_paragraph" => {
            let arguments: InsertParagraphArgs = parse(name, args)?;
            let package = Package::open(&arguments.path)?;
            let (xml, document) = package.parse()?;
            let xml = edit::insert(&arguments, &document, xml)?;
            save_edit(&arguments.path, package, xml)
        }
        "format_paragraph" => {
            let arguments: FormatParagraphArgs = parse(name, args)?;
            let package = Package::open(&arguments.path)?;
            let (xml, document) = package.parse()?;
            let xml = edit::format(&arguments, &document, xml)?;
            save_edit(&arguments.path, package, xml)
        }
        "delete_paragraph" => {
            let arguments: DeleteParagraphArgs = parse(name, args)?;
            let package = Package::open(&arguments.path)?;
            let (xml, document) = package.parse()?;
            let (xml, deleted) = edit::delete(&document, xml, arguments.index)?;
            let mut result = save_edit(&arguments.path, package, xml)?;
            result["deleted_text"] = json!(deleted);
            Ok(result)
        }
        _ => bail!("unknown DOCX tool: {name}"),
    }
}

fn read(arguments: &ReadDocumentArgs) -> Result<Value> {
    let (_, document) = Package::open(&arguments.path)?.parse()?;
    let start = arguments.start.unwrap_or(0);
    let mut texts = Vec::new();
    let paragraphs: Vec<Value> = document
        .paragraphs
        .iter()
        .enumerate()
        .skip(start)
        .take(arguments.limit.unwrap_or(usize::MAX))
        .map(|(index, paragraph)| {
            let text = document.text(*paragraph);
            let value = json!({
                "index":index,
                "text":text,
                "style":document.style(*paragraph),
                "in_table":document.in_table(*paragraph),
            });
            texts.push(text);
            value
        })
        .collect();
    Ok(json!({
        "path":arguments.path,
        "paragraphs":paragraphs,
        "text":texts.join("\n"),
        "total_paragraphs":document.paragraphs.len(),
        "start":start,
    }))
}

fn info(arguments: &GetDocumentInfoArgs) -> Result<Value> {
    let package = Package::open(&arguments.path)?;
    let (_, document) = package.parse()?;
    let characters: usize = document
        .paragraphs
        .iter()
        .map(|paragraph| document.text(*paragraph).chars().count())
        .sum();
    Ok(json!({
        "path":arguments.path,
        "paragraph_count":document.paragraphs.len(),
        "character_count":characters,
        "parts":package.entries.iter().map(|entry| &entry.name).collect::<Vec<_>>(),
    }))
}

#[cfg(test)]
mod tests {
    use std::{fs, path::Path};

    use anyhow::Context;

    use super::{
        package::{Entry, encode_entries},
        xml::{Document, WORD_NS},
        *,
    };

    fn fixture(path: &Path, body: &str) -> Result<()> {
        let xml = format!("<q:document xmlns:q=\"{WORD_NS}\"><q:body>{body}</q:body></q:document>");
        let entries = [
            Entry {
                name: "word/document.xml".into(),
                bytes: xml.into_bytes(),
                directory: false,
            },
            Entry {
                name: "custom/preserved.bin".into(),
                bytes: vec![0, 255, 13, 10],
                directory: false,
            },
        ];
        fs::write(path, encode_entries(&entries)?)?;
        Ok(())
    }

    fn preserved_part(path: &Path) -> Result<Vec<u8>> {
        let package = Package::open(path)?;
        let entry = package
            .entries
            .into_iter()
            .find(|entry| entry.name == "custom/preserved.bin")
            .context("missing part")?;
        Ok(entry.bytes)
    }

    #[test]
    fn schemas_advertise_constraints() -> Result<()> {
        let tools = tools();
        let schema = |name: &str| -> Result<Value> {
            let tool = tools
                .iter()
                .find(|tool| tool.name == name)
                .context("missing tool")?;
            Ok(Value::Object((*tool.input_schema).clone()))
        };
        let format = schema("format_paragraph")?;
        assert_eq!(format["properties"]["path"]["minLength"], 1);
        assert_eq!(format["additionalProperties"], false);
        assert_eq!(format["properties"]["font_size_pt"]["maximum"], 1638);
        assert_eq!(
            schema("replace_text")?["properties"]["find"]["minLength"],
            1
        );
        assert!(call("delete_paragraph", json!({"path":"x.docx"})).is_err());
        assert!(call("read_document", json!({"path":""})).is_err());
        assert!(call("read_document", json!({"path":"x.docx","extra":1})).is_err());
        assert!(call("missing_tool", json!({})).is_err());
        Ok(())
    }

    #[test]
    fn unicode_entities_controls_and_preview_roundtrip() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("unicode.docx");
        let text = "é 🍎 <script>& \"quotes\"\nline\ttab";
        call("create_document", json!({"path":path,"paragraphs":[text]}))?;
        let read = call("read_document", json!({"path":path}))?;
        assert_eq!(read["text"], text);
        let html = call("preview_document", json!({"path":path}))?;
        assert!(
            html["html"]
                .as_str()
                .context("missing HTML")?
                .contains("&lt;script&gt;&amp;")
        );
        assert!(call("create_document", json!({"path":path})).is_err());
        Ok(())
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
        let result = call(
            "replace_text",
            json!({"path":path,"find":"café 🍎","replacement":"A & B","replace_all":true}),
        )?;
        assert_eq!(result["replacements"], 2);
        let backup = result["backup_path"].as_str().context("missing backup")?;
        assert_eq!(fs::read(backup)?, original);
        let package = Package::open(&path)?;
        assert!(package.xml()?.contains("<q:b/>") && package.xml()?.contains("<q:i/>"));
        assert_eq!(preserved_part(&path)?, vec![0, 255, 13, 10]);
        assert_eq!(
            call("read_document", json!({"path":path}))?["text"],
            "A & B today A & B"
        );
        Ok(())
    }

    #[test]
    fn formatting_preserves_properties_and_orders_new_ones() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("format.docx");
        fixture(
            &path,
            "<q:p><q:pPr><q:spacing q:after=\"120\"/><q:jc q:val=\"right\"/></q:pPr><q:r><q:rPr><q:color q:val=\"112233\"/></q:rPr><q:t>hello</q:t></q:r><q:r q:rsidR=\"01234567\"/></q:p><q:p q:rsidR=\"76543210\"/>",
        )?;
        call(
            "format_paragraph",
            json!({"path":path,"index":0,"bold":true,"font_size_pt":12.5,"style":"Heading1","alignment":"center"}),
        )?;
        call(
            "format_paragraph",
            json!({"path":path,"index":1,"italic":true,"alignment":"both","style":"Quote"}),
        )?;
        assert!(
            call(
                "format_paragraph",
                json!({"path":path,"index":0,"alignment":"middle"})
            )
            .is_err()
        );
        assert!(
            call(
                "format_paragraph",
                json!({"path":path,"index":0,"font_size_pt":12.3})
            )
            .is_err()
        );
        let package = Package::open(&path)?;
        let xml = package.xml()?;
        assert!(xml.contains("q:after=\"120\"") && xml.contains("q:val=\"112233\""));
        assert!(xml.contains("q:rsidR=\"01234567\"") && xml.contains("q:rsidR=\"76543210\""));
        assert!(
            xml.find("<w:pStyle").context("missing style")?
                < xml.find("<q:spacing").context("missing spacing")?
        );
        let quote = xml.find("w:val=\"Quote\"").context("missing quote style")?;
        assert!(quote < xml.find("w:val=\"both\"").context("missing alignment")?);
        let document = Document::parse(xml)?;
        for (index, node) in document.nodes.iter().enumerate() {
            if node.is("rPr") {
                assert!(
                    node.parent
                        .is_some_and(|parent| document.nodes[parent].name == "r"),
                    "run properties {index} have wrong parent"
                );
            }
        }
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
        call(
            "insert_paragraph",
            json!({"path":path,"index":1,"text":"inserted"}),
        )?;
        assert_eq!(
            call("read_document", json!({"path":path}))?["text"],
            "before\ninserted\ncell"
        );
        let empty = format!("<q:document xmlns:q=\"{WORD_NS}\"><q:body/></q:document>");
        let document = Document::parse(&empty)?;
        let arguments: InsertParagraphArgs =
            serde_json::from_value(json!({"path":path,"text":"new"}))?;
        let edited = edit::insert(&arguments, &document, &empty)?;
        let parsed = Document::parse(&edited)?;
        assert_eq!(parsed.nodes[parsed.paragraphs[0]].parent, Some(parsed.body));
        Ok(())
    }

    #[test]
    fn read_reports_style_and_table_membership() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("styles.docx");
        fixture(
            &path,
            "<q:p><q:pPr><q:pStyle q:val=\"Heading1\"/></q:pPr><q:r><q:t>title</q:t></q:r></q:p><q:tbl><q:tr><q:tc><q:p><q:pPr><q:pStyle q:val=\"TableText\"/></q:pPr><q:r><q:t>cell</q:t></q:r></q:p></q:tc></q:tr></q:tbl><q:p/>",
        )?;
        let read = call("read_document", json!({"path":path,"start":0}))?;
        assert_eq!(
            read["paragraphs"],
            json!([
                {"index":0,"text":"title","style":"Heading1","in_table":false},
                {"index":1,"text":"cell","style":"TableText","in_table":true},
                {"index":2,"text":"","style":null,"in_table":false},
            ])
        );
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
        let result = call("delete_paragraph", json!({"path":path,"index":1}))?;
        assert_eq!(result["modified"], true);
        assert_eq!(result["deleted_text"], "drop me");
        let backup = result["backup_path"].as_str().context("missing backup")?;
        assert_eq!(fs::read(backup)?, original);
        assert_eq!(preserved_part(&path)?, vec![0, 255, 13, 10]);
        call("delete_paragraph", json!({"path":path,"index":2}))?;
        let read = call("read_document", json!({"path":path}))?;
        assert_eq!(read["text"], "keep one\ncell a\nkeep two");
        assert_eq!(read["paragraphs"][1]["in_table"], true);
        assert!(!Package::open(&path)?.xml()?.contains("<q:b/>"));
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
                call("delete_paragraph", json!({"path":path,"index":index})).is_err(),
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
        for name in ["~$locked.docx", "~$cked.docx"] {
            let lock = directory.path().join(name);
            fs::write(&lock, [])?;
            assert!(
                call(
                    "replace_text",
                    json!({"path":path,"find":"text","replacement":"changed"})
                )
                .is_err()
            );
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
            assert!(
                call(
                    "replace_text",
                    json!({"path":path,"find":"text","replacement":"changed"})
                )
                .is_err()
            );
            assert_eq!(fs::read(&path)?, before);
        }
        Ok(())
    }
}
