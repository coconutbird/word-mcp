//! `docx_document`: create saved documents and inspect packages.

use std::path::PathBuf;

use anyhow::Result;
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::{Value, json};

use super::{
    Definition,
    ooxml::{content_type, ns, rel},
    package::{DOCUMENT_PART, Entry, Package, encode_entries, read, write_atomic},
    xml::{Document, WORD_NS, paragraph_xml},
};
use crate::tool::{Effect, Output, parse, tool};

const NAME: &str = "docx_document";

pub(super) const TOOL: Definition = Definition {
    name: NAME,
    tool: || {
        tool::<Arguments>(
            NAME,
            "Create a saved .docx file or inspect a package (paragraph and character counts, parts).",
            Effect::Additive,
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
    /// Create a new document holding one paragraph per entry.
    Create {
        /// Paragraph texts; `\n` becomes a line break and `\t` a tab.
        #[serde(default)]
        paragraphs: Vec<String>,
        /// Replace an existing file at `path`.
        #[serde(default)]
        overwrite: bool,
    },
    /// Report paragraph count, character count, and package parts.
    Info {},
}

fn call(arguments: Value) -> Result<Output> {
    let Arguments { path, operation } = parse(NAME, arguments)?;
    Ok(match operation {
        Operation::Create {
            paragraphs,
            overwrite,
        } => create(&path, &paragraphs, overwrite)?,
        Operation::Info {} => info(&path)?,
    }
    .into())
}

/// Create a minimal DOCX holding one paragraph per entry (at least one).
fn create(path: &std::path::Path, paragraphs: &[String], overwrite: bool) -> Result<Value> {
    let empty = [String::new()];
    let paragraphs = if paragraphs.is_empty() {
        &empty[..]
    } else {
        paragraphs
    };
    let body = paragraphs
        .iter()
        .map(|text| paragraph_xml(text, None))
        .collect::<Result<String>>()?;
    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"{WORD_NS}\" xmlns:r=\"{}\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>",
        ns::R
    );
    Document::parse(&xml)?;
    let entries = vec![
        Entry {
            name: "[Content_Types].xml".into(),
            bytes: format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"{}\"><Default Extension=\"rels\" ContentType=\"{}\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/{DOCUMENT_PART}\" ContentType=\"{}\"/></Types>",
                ns::CONTENT_TYPES,
                content_type::RELATIONSHIPS,
                content_type::DOCUMENT
            )
            .into_bytes(),
            directory: false,
        },
        Entry {
            name: "_rels/.rels".into(),
            bytes: format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{}\"><Relationship Id=\"rId1\" Type=\"{}\" Target=\"{DOCUMENT_PART}\"/></Relationships>",
                ns::PACKAGE_RELATIONSHIPS,
                rel::OFFICE_DOCUMENT
            )
            .into_bytes(),
            directory: false,
        },
        Entry {
            name: DOCUMENT_PART.into(),
            bytes: xml.into_bytes(),
            directory: false,
        },
    ];
    let package = Package::from_entries(entries);
    write_atomic(path, &encode_entries(&package.entries)?, None, overwrite)?;
    Ok(json!({"path": path, "created": true, "paragraph_count": paragraphs.len()}))
}

fn info(path: &std::path::Path) -> Result<Value> {
    let package = read(path)?;
    let (_, document) = package.parse()?;
    let characters: usize = document
        .paragraphs
        .iter()
        .map(|paragraph| document.text(*paragraph).chars().count())
        .sum();
    Ok(json!({
        "path": path,
        "paragraph_count": document.paragraphs.len(),
        "character_count": characters,
        "parts": package.part_names().collect::<Vec<_>>(),
    }))
}

#[cfg(test)]
mod tests {
    use anyhow::Result;
    use serde_json::json;

    use super::super::testing::run;

    #[test]
    fn create_refuses_to_overwrite_by_default() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("new.docx");
        let created = run(
            "docx_document",
            json!({"path": path, "operation": {"action": "create", "paragraphs": ["a", "b"]}}),
        )?;
        assert_eq!(created["paragraph_count"], 2);
        assert!(
            run(
                "docx_document",
                json!({"path": path, "operation": {"action": "create"}})
            )
            .is_err()
        );
        let info = run(
            "docx_document",
            json!({"path": path, "operation": {"action": "info"}}),
        )?;
        assert_eq!(info["paragraph_count"], 2);
        assert_eq!(info["character_count"], 2);
        Ok(())
    }
}
