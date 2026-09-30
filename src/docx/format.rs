//! `docx_format`: character, paragraph, list, and style formatting of saved documents.

use std::path::PathBuf;

use anyhow::{Result, ensure};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::{Value, json};

use super::{
    Definition,
    edit::paragraph_at,
    package::edit,
    props::{Property, joined, ordered, property_patches, runs},
    xml::{
        Document, PARAGRAPH_PROPERTY_ORDER, RUN_PROPERTY_ORDER, WORD_NS, escaped, expand_empty,
        patch, property_xml, validate_text,
    },
};
use crate::tool::{Effect, Output, parse, tool};

const NAME: &str = "docx_format";

pub(super) const TOOL: Definition = Definition {
    name: NAME,
    tool: || {
        tool::<Arguments>(
            NAME,
            "Format a saved .docx: direct paragraph and character formatting of a zero-based paragraph. Every edit keeps a backup.",
            Effect::Destructive,
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

/// Horizontal paragraph alignment.
#[derive(Clone, Copy, Deserialize, JsonSchema)]
#[serde(rename_all = "lowercase")]
pub(super) enum Alignment {
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
    pub(super) fn val(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Both => "both",
        }
    }
}

#[derive(Deserialize, JsonSchema)]
#[serde(tag = "action", rename_all = "snake_case", deny_unknown_fields)]
enum Operation {
    /// Set paragraph properties and run formatting on every run of one paragraph.
    Paragraph {
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
        /// Paragraph style id, for example `Heading1`.
        style: Option<String>,
    },
}

fn call(arguments: Value) -> Result<Output> {
    let Arguments { path, operation } = parse(NAME, arguments)?;
    Ok(edit(&path, |package| {
        let (xml, document) = package.parse()?;
        let xml = match operation {
            Operation::Paragraph {
                index,
                bold,
                italic,
                underline,
                font_size_pt,
                font_family,
                alignment,
                style,
            } => {
                let run = RunFormat {
                    bold,
                    italic,
                    underline,
                    font_size_pt,
                    font_family,
                };
                format_paragraph(
                    &document,
                    xml,
                    index,
                    &run.properties()?,
                    &paragraph_properties(alignment, style.as_deref())?,
                )?
            }
        };
        package.set_document(xml)?;
        Ok(json!({}))
    })?
    .into())
}

/// Requested character formatting.
pub(super) struct RunFormat {
    pub(super) bold: Option<bool>,
    pub(super) italic: Option<bool>,
    pub(super) underline: Option<bool>,
    pub(super) font_size_pt: Option<f64>,
    pub(super) font_family: Option<String>,
}

impl RunFormat {
    /// The `w:rPr` children this format sets, in schema order.
    pub(super) fn properties(&self) -> Result<Vec<Property>> {
        let mut properties = Vec::new();
        for (name, value) in [("b", self.bold), ("i", self.italic)] {
            if let Some(value) = value {
                properties.push((name, property_xml(name, if value { "1" } else { "0" })));
            }
        }
        if let Some(value) = self.underline {
            properties.push((
                "u",
                property_xml("u", if value { "single" } else { "none" }),
            ));
        }
        if let Some(size) = self.font_size_pt {
            ensure!(
                size.is_finite() && (1.0..=1638.0).contains(&size) && (size * 2.0).fract() == 0.0,
                "font size must be 1–1638 points in half-point increments"
            );
            properties.push(("sz", property_xml("sz", &format!("{:.0}", size * 2.0))));
        }
        if let Some(font) = &self.font_family {
            validate_text(font)?;
            ensure!(!font.is_empty(), "font family must not be empty");
            let font = escaped(font);
            properties.push((
                "rFonts",
                format!(
                    "<w:rFonts xmlns:w=\"{WORD_NS}\" w:ascii=\"{font}\" w:hAnsi=\"{font}\" w:eastAsia=\"{font}\" w:cs=\"{font}\"/>"
                ),
            ));
        }
        Ok(ordered(properties, RUN_PROPERTY_ORDER))
    }
}

/// Paragraph properties for alignment and style, in schema order.
fn paragraph_properties(
    alignment: Option<Alignment>,
    style: Option<&str>,
) -> Result<Vec<Property>> {
    let mut properties = Vec::new();
    if let Some(alignment) = alignment {
        properties.push(("jc", property_xml("jc", alignment.val())));
    }
    if let Some(style) = style {
        validate_text(style)?;
        properties.push(("pStyle", property_xml("pStyle", style)));
    }
    Ok(ordered(properties, PARAGRAPH_PROPERTY_ORDER))
}

/// Apply paragraph properties and run properties to every run of one paragraph.
fn format_paragraph(
    document: &Document,
    xml: &str,
    index: usize,
    run_properties: &[Property],
    paragraph_properties: &[Property],
) -> Result<String> {
    let paragraph = paragraph_at(document, index)?;
    document.editable(paragraph)?;
    ensure!(
        !run_properties.is_empty() || !paragraph_properties.is_empty(),
        "provide at least one formatting property"
    );
    let node = &document.nodes[paragraph];
    let mut changes = Vec::new();
    if node.self_closing() {
        let content = format!(
            "<w:pPr xmlns:w=\"{WORD_NS}\">{}</w:pPr><w:r xmlns:w=\"{WORD_NS}\"><w:rPr>{}</w:rPr><w:t/></w:r>",
            joined(paragraph_properties),
            joined(run_properties)
        );
        changes.push((node.start, node.end, expand_empty(xml, node, &content)?));
    } else {
        property_patches(
            document,
            xml,
            paragraph,
            "pPr",
            paragraph_properties,
            &mut changes,
        )?;
        let runs = runs(document, paragraph);
        for run in &runs {
            property_patches(document, xml, *run, "rPr", run_properties, &mut changes)?;
        }
        if runs.is_empty() && !run_properties.is_empty() {
            changes.push((
                node.close_start,
                node.close_start,
                format!(
                    "<w:r xmlns:w=\"{WORD_NS}\"><w:rPr>{}</w:rPr><w:t/></w:r>",
                    joined(run_properties)
                ),
            ));
        }
    }
    patch(xml, changes)
}

#[cfg(test)]
mod tests {
    use anyhow::{Context, Result};
    use serde_json::json;

    use super::super::{
        testing::{document_xml, fixture, run},
        xml::Document,
    };

    #[test]
    fn formatting_preserves_properties_and_orders_new_ones() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("format.docx");
        fixture(
            &path,
            "<q:p><q:pPr><q:spacing q:after=\"120\"/><q:jc q:val=\"right\"/></q:pPr><q:r><q:rPr><q:color q:val=\"112233\"/></q:rPr><q:t>hello</q:t></q:r><q:r q:rsidR=\"01234567\"/></q:p><q:p q:rsidR=\"76543210\"/>",
        )?;
        let format = |operation| run("docx_format", json!({"path": path, "operation": operation}));
        format(
            json!({"action": "paragraph", "index": 0, "bold": true, "font_size_pt": 12.5, "style": "Heading1", "alignment": "center"}),
        )?;
        format(
            json!({"action": "paragraph", "index": 1, "italic": true, "alignment": "both", "style": "Quote"}),
        )?;
        assert!(format(json!({"action": "paragraph", "index": 0, "alignment": "middle"})).is_err());
        assert!(format(json!({"action": "paragraph", "index": 0, "font_size_pt": 12.3})).is_err());
        let xml = document_xml(&path)?;
        assert!(xml.contains("q:after=\"120\"") && xml.contains("q:val=\"112233\""));
        assert!(xml.contains("q:rsidR=\"01234567\"") && xml.contains("q:rsidR=\"76543210\""));
        assert!(
            xml.find("<w:pStyle").context("missing style")?
                < xml.find("<q:spacing").context("missing spacing")?
        );
        let quote = xml.find("w:val=\"Quote\"").context("missing quote style")?;
        assert!(quote < xml.find("w:val=\"both\"").context("missing alignment")?);
        let document = Document::parse(&xml)?;
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
}
