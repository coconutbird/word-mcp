//! Fixtures shared by the saved-file tool tests.

use std::{fmt::Write as _, fs, path::Path};

use anyhow::{Context, Result, bail};
use serde_json::Value;

use super::{
    call,
    ooxml::{content_type, ns, rel},
    package::{Entry, Package, encode_entries},
    xml::WORD_NS,
};
use crate::tool::Output;

/// A part that no tool touches, used to prove byte-for-byte preservation.
pub(super) const PRESERVED_PART: &str = "custom/preserved.bin";
pub(super) const PRESERVED_BYTES: [u8; 4] = [0, 255, 13, 10];

/// Write a minimal valid package whose body is `body`, written with the unusual
/// Word prefix `q:` so tests catch prefix assumptions.
pub(super) fn fixture(path: &Path, body: &str) -> Result<()> {
    fixture_with(path, body, &[])
}

/// Like [`fixture`], with extra `(name, xml)` parts related from the main document by
/// `(relationship type, content type)`.
pub(super) fn fixture_with(
    path: &Path,
    body: &str,
    extra: &[(&str, &str, &str, &str)],
) -> Result<()> {
    let document = format!(
        "<q:document xmlns:q=\"{WORD_NS}\" xmlns:r=\"{}\"><q:body>{body}</q:body></q:document>",
        ns::R
    );
    let mut overrides = format!(
        "<Override PartName=\"/word/document.xml\" ContentType=\"{}\"/>",
        content_type::DOCUMENT
    );
    let mut relationships = String::new();
    let mut entries = Vec::new();
    for (index, (name, kind, content, xml)) in extra.iter().enumerate() {
        let _ = write!(
            overrides,
            "<Override PartName=\"/{name}\" ContentType=\"{content}\"/>"
        );
        let target = name
            .strip_prefix("word/")
            .context("extra parts live in word/")?;
        let _ = write!(
            relationships,
            "<Relationship Id=\"rId{}\" Type=\"{kind}\" Target=\"{target}\"/>",
            index + 1
        );
        entries.push(Entry {
            name: (*name).to_owned(),
            bytes: xml.as_bytes().to_vec(),
            directory: false,
        });
    }
    entries.extend([
        Entry {
            name: "[Content_Types].xml".into(),
            bytes: format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"{}\"><Default Extension=\"rels\" ContentType=\"{}\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>{overrides}</Types>",
                ns::CONTENT_TYPES,
                content_type::RELATIONSHIPS
            )
            .into_bytes(),
            directory: false,
        },
        Entry {
            name: "_rels/.rels".into(),
            bytes: format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"{}\"><Relationship Id=\"rId1\" Type=\"{}\" Target=\"word/document.xml\"/></Relationships>",
                ns::PACKAGE_RELATIONSHIPS,
                rel::OFFICE_DOCUMENT
            )
            .into_bytes(),
            directory: false,
        },
        Entry {
            name: "word/document.xml".into(),
            bytes: document.into_bytes(),
            directory: false,
        },
        Entry {
            name: PRESERVED_PART.into(),
            bytes: PRESERVED_BYTES.to_vec(),
            directory: false,
        },
    ]);
    if !relationships.is_empty() {
        entries.push(Entry {
            name: "word/_rels/document.xml.rels".into(),
            bytes: format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"{}\">{relationships}</Relationships>",
                ns::PACKAGE_RELATIONSHIPS
            )
            .into_bytes(),
            directory: false,
        });
    }
    fs::write(path, encode_entries(&entries)?)?;
    Ok(())
}

/// The bytes of the untouched fixture part after edits.
pub(super) fn preserved_part(path: &Path) -> Result<Vec<u8>> {
    Ok(Package::open(path)?
        .part(PRESERVED_PART)
        .context("missing preserved part")?
        .to_vec())
}

/// The main document XML of the package at `path`.
pub(super) fn document_xml(path: &Path) -> Result<String> {
    Ok(Package::open(path)?.xml()?.to_owned())
}

/// Run a grouped tool and return its JSON result.
pub(super) fn run(tool: &str, arguments: Value) -> Result<Value> {
    match call(tool, arguments)? {
        Output::Json(value) => Ok(value),
        Output::Image { .. } => bail!("{tool} returned an image"),
    }
}
