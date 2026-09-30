//! Document operations: creation, preview and targeted paragraph edits.
//!
//! Edit operations return the patched main document XML; callers save it.

use std::fs;

use anyhow::{Context, Result, ensure};
use serde_json::{Value, json};

use super::{
    CreateDocumentArgs, FormatParagraphArgs, InsertParagraphArgs, PreviewDocumentArgs,
    ReplaceTextArgs,
    package::{DOCUMENT_PART, Entry, encode_entries, write_atomic},
    xml::{
        Document, PARAGRAPH_PROPERTY_ORDER, Patch, RUN_PROPERTY_ORDER, WORD_NS, escaped,
        expand_empty, paragraph_xml, patch, prepend_child, property_xml, text_xml, validate_text,
    },
};

const CONTENT_TYPES: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
const RELATIONSHIPS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;

/// Create a minimal DOCX holding one paragraph per entry (at least one).
pub(super) fn create(arguments: &CreateDocumentArgs) -> Result<Value> {
    let empty = [String::new()];
    let paragraphs = if arguments.paragraphs.is_empty() {
        &empty[..]
    } else {
        &arguments.paragraphs
    };
    let body = paragraphs
        .iter()
        .map(|text| paragraph_xml(text, None))
        .collect::<Result<String>>()?;
    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"{WORD_NS}\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr></w:body></w:document>"
    );
    Document::parse(&xml)?;
    let entries = [
        Entry {
            name: "[Content_Types].xml".into(),
            bytes: CONTENT_TYPES.to_vec(),
            directory: false,
        },
        Entry {
            name: "_rels/.rels".into(),
            bytes: RELATIONSHIPS.to_vec(),
            directory: false,
        },
        Entry {
            name: DOCUMENT_PART.into(),
            bytes: xml.into_bytes(),
            directory: false,
        },
    ];
    write_atomic(
        &arguments.path,
        &encode_entries(&entries)?,
        None,
        arguments.overwrite,
    )?;
    Ok(json!({"path":arguments.path,"created":true,"paragraph_count":paragraphs.len()}))
}

/// Render escaped logical HTML, optionally saving it beside the document.
pub(super) fn preview(arguments: &PreviewDocumentArgs, document: &Document) -> Result<Value> {
    let body = document
        .paragraphs
        .iter()
        .map(|paragraph| {
            format!(
                "<p>{}</p>",
                escaped(&document.text(*paragraph)).replace('\n', "<br>")
            )
        })
        .collect::<Vec<_>>()
        .join("\n");
    let html = format!(
        "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><title>Document preview</title><style>body{{max-width:52rem;margin:2rem auto;font-family:system-ui}}p{{white-space:pre-wrap}}</style></head><body>{body}</body></html>"
    );
    if let Some(output) = &arguments.output_path {
        ensure!(
            output
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("html")
                    || extension.eq_ignore_ascii_case("htm")),
            "preview output must use .html or .htm extension"
        );
        ensure!(
            output != &arguments.path,
            "preview cannot overwrite source document"
        );
        if output.exists() {
            ensure!(
                fs::canonicalize(output)? != fs::canonicalize(&arguments.path)?,
                "preview cannot overwrite source document"
            );
        }
        write_atomic(output, html.as_bytes(), None, arguments.overwrite)?;
    }
    Ok(
        json!({"path":arguments.path,"html":html,"output_path":arguments.output_path,"preview_type":"logical_html"}),
    )
}

/// The element index of the zero-based paragraph `index`.
fn paragraph_at(document: &Document, index: usize) -> Result<usize> {
    document
        .paragraphs
        .get(index)
        .copied()
        .context("paragraph index is out of bounds")
}

/// Replace literal text across a paragraph's runs; `None` when nothing matched.
pub(super) fn replace(
    arguments: &ReplaceTextArgs,
    document: &Document,
    xml: &str,
) -> Result<Option<(String, usize)>> {
    let find = arguments.find.as_str();
    ensure!(!find.is_empty(), "find must not be empty");
    validate_text(&arguments.replacement)?;
    if let Some(index) = arguments.paragraph_index {
        paragraph_at(document, index)?;
    }
    let mut changes = Vec::new();
    let mut count = 0;
    for (index, paragraph) in document.paragraphs.iter().copied().enumerate() {
        if arguments
            .paragraph_index
            .is_some_and(|selected| index != selected)
        {
            continue;
        }
        let text = document.text(paragraph);
        let mut matches: Vec<usize> = text.match_indices(find).map(|(offset, _)| offset).collect();
        if matches.is_empty() {
            continue;
        }
        document.editable(paragraph)?;
        document.unanchored(paragraph)?;
        if !arguments.replace_all {
            matches.truncate(1);
        }
        let texts = document.texts(paragraph);
        let mut values: Vec<String> = texts
            .iter()
            .map(|index| document.nodes[*index].text.clone())
            .collect();
        let mut offsets = Vec::with_capacity(values.len());
        let mut offset = 0;
        for value in &values {
            offsets.push((offset, offset + value.len()));
            offset += value.len();
        }
        for start in matches.iter().rev().copied() {
            let end = start + find.len();
            let first = offsets
                .iter()
                .position(|(low, high)| *low <= start && start < *high)
                .context("match starts outside editable text")?;
            for (node, (low, high)) in offsets.iter().copied().enumerate() {
                if low >= end || high <= start {
                    continue;
                }
                let local_start = start.saturating_sub(low);
                let local_end = (end - low).min(high - low);
                values[node].replace_range(
                    local_start..local_end,
                    if node == first {
                        &arguments.replacement
                    } else {
                        ""
                    },
                );
            }
        }
        for (node, value) in texts.into_iter().zip(values) {
            let node = &document.nodes[node];
            if value != node.text {
                changes.push((node.start, node.end, text_xml(&value)));
            }
        }
        count += matches.len();
        if !arguments.replace_all {
            break;
        }
    }
    if count == 0 {
        return Ok(None);
    }
    Ok(Some((patch(xml, changes)?, count)))
}

/// Insert a paragraph before the zero-based `index`, or append before the body's
/// section properties.
pub(super) fn insert(
    arguments: &InsertParagraphArgs,
    document: &Document,
    xml: &str,
) -> Result<String> {
    let index = arguments.index.unwrap_or(document.paragraphs.len());
    ensure!(
        index <= document.paragraphs.len(),
        "paragraph index is out of bounds"
    );
    let paragraph = paragraph_xml(&arguments.text, arguments.style.as_deref())?;
    let body = &document.nodes[document.body];
    if body.self_closing() {
        return patch(
            xml,
            vec![(body.start, body.end, expand_empty(xml, body, &paragraph)?)],
        );
    }
    let offset = if let Some(target) = document.paragraphs.get(index) {
        document.editable(*target)?;
        document.nodes[*target].start
    } else {
        document
            .child(document.body, "sectPr")
            .map_or(body.close_start, |node| document.nodes[node].start)
    };
    patch(xml, vec![(offset, offset, paragraph)])
}

/// Remove the zero-based paragraph, returning the new XML and the deleted text.
pub(super) fn delete(document: &Document, xml: &str, index: usize) -> Result<(String, String)> {
    let paragraph = paragraph_at(document, index)?;
    document.editable(paragraph)?;
    document.unanchored(paragraph)?;
    ensure!(
        document
            .child(paragraph, "pPr")
            .is_none_or(|properties| document.child(properties, "sectPr").is_none()),
        "paragraph holds a section break (w:sectPr); deleting it would merge sections"
    );
    let node = &document.nodes[paragraph];
    let container = node.parent.context("paragraph has no container")?;
    ensure!(
        document
            .children(container)
            .filter(|(_, child)| child.is("p"))
            .nth(1)
            .is_some(),
        "cannot delete the only paragraph in its {} element",
        document.nodes[container].name
    );
    if document.nodes[container].is("tc") {
        // A table cell must end with a paragraph, for example after a nested table.
        let last_block = document
            .children(container)
            .filter(|&(child_index, child)| child_index != paragraph && !child.is("tcPr"))
            .last();
        ensure!(
            last_block.is_some_and(|(_, child)| child.is("p")),
            "a table cell must end with a paragraph"
        );
    }
    Ok((
        patch(xml, vec![(node.start, node.end, String::new())])?,
        document.text(paragraph),
    ))
}

/// A property element name and its XML.
type Property = (&'static str, String);

/// Sort properties into their schema order.
fn ordered(mut properties: Vec<Property>, order: &[&str]) -> Vec<Property> {
    properties.sort_by_key(|(name, _)| rank(order, name));
    properties
}

fn rank(order: &[&str], property: &str) -> usize {
    order
        .iter()
        .position(|name| *name == property)
        .unwrap_or(usize::MAX)
}

fn joined(properties: &[Property]) -> String {
    properties.iter().map(|(_, value)| value.as_str()).collect()
}

/// Set schema-ordered `properties` inside `parent`'s `name` container (`pPr` or
/// `rPr`), replacing same-named properties and keeping all others.
fn property_patches(
    document: &Document,
    xml: &str,
    parent: usize,
    name: &str,
    properties: &[Property],
    changes: &mut Vec<Patch>,
) -> Result<()> {
    if properties.is_empty() {
        return Ok(());
    }
    let order = if name == "pPr" {
        PARAGRAPH_PROPERTY_ORDER
    } else {
        RUN_PROPERTY_ORDER
    };
    let Some(container) = document.child(parent, name) else {
        let content = format!(
            "<w:{name} xmlns:w=\"{WORD_NS}\">{}</w:{name}>",
            joined(properties)
        );
        changes.push(prepend_child(xml, &document.nodes[parent], content)?);
        return Ok(());
    };
    let node = &document.nodes[container];
    if node.self_closing() {
        changes.push((
            node.start,
            node.end,
            expand_empty(xml, node, &joined(properties))?,
        ));
        return Ok(());
    }
    for (property, value) in properties {
        if let Some(existing) = document.child(container, property) {
            let existing = &document.nodes[existing];
            changes.push((existing.start, existing.end, value.clone()));
        } else {
            let offset = document
                .children(container)
                .find(|(_, child)| child.word && rank(order, &child.name) > rank(order, property))
                .map_or(node.close_start, |(_, child)| child.start);
            changes.push((offset, offset, value.clone()));
        }
    }
    Ok(())
}

/// Run properties requested by `format_paragraph`, in schema order.
fn requested_run_properties(arguments: &FormatParagraphArgs) -> Result<Vec<Property>> {
    let mut properties = Vec::new();
    for (name, value) in [("b", arguments.bold), ("i", arguments.italic)] {
        if let Some(value) = value {
            properties.push((name, property_xml(name, if value { "1" } else { "0" })));
        }
    }
    if let Some(value) = arguments.underline {
        properties.push((
            "u",
            property_xml("u", if value { "single" } else { "none" }),
        ));
    }
    if let Some(size) = arguments.font_size_pt {
        ensure!(
            size.is_finite() && (1.0..=1638.0).contains(&size) && (size * 2.0).fract() == 0.0,
            "font size must be 1–1638 points in half-point increments"
        );
        properties.push(("sz", property_xml("sz", &format!("{:.0}", size * 2.0))));
    }
    if let Some(font) = &arguments.font_family {
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

/// Paragraph properties requested by `format_paragraph`, in schema order.
fn requested_paragraph_properties(arguments: &FormatParagraphArgs) -> Result<Vec<Property>> {
    let mut properties = Vec::new();
    if let Some(alignment) = arguments.alignment {
        properties.push(("jc", property_xml("jc", alignment.val())));
    }
    if let Some(style) = &arguments.style {
        validate_text(style)?;
        properties.push(("pStyle", property_xml("pStyle", style)));
    }
    Ok(ordered(properties, PARAGRAPH_PROPERTY_ORDER))
}

/// Apply direct paragraph formatting and run formatting to every run.
pub(super) fn format(
    arguments: &FormatParagraphArgs,
    document: &Document,
    xml: &str,
) -> Result<String> {
    let paragraph = paragraph_at(document, arguments.index)?;
    document.editable(paragraph)?;
    let run_properties = requested_run_properties(arguments)?;
    let paragraph_properties = requested_paragraph_properties(arguments)?;
    ensure!(
        !run_properties.is_empty() || !paragraph_properties.is_empty(),
        "provide at least one formatting property"
    );
    let node = &document.nodes[paragraph];
    let mut changes = Vec::new();
    if node.self_closing() {
        let content = format!(
            "<w:pPr xmlns:w=\"{WORD_NS}\">{}</w:pPr><w:r xmlns:w=\"{WORD_NS}\"><w:rPr>{}</w:rPr><w:t/></w:r>",
            joined(&paragraph_properties),
            joined(&run_properties)
        );
        changes.push((node.start, node.end, expand_empty(xml, node, &content)?));
    } else {
        property_patches(
            document,
            xml,
            paragraph,
            "pPr",
            &paragraph_properties,
            &mut changes,
        )?;
        let runs: Vec<usize> = document
            .owned(paragraph)
            .filter_map(|(index, child)| (child.name == "r").then_some(index))
            .collect();
        for run in &runs {
            property_patches(document, xml, *run, "rPr", &run_properties, &mut changes)?;
        }
        if runs.is_empty() && !run_properties.is_empty() {
            changes.push((
                node.close_start,
                node.close_start,
                format!(
                    "<w:r xmlns:w=\"{WORD_NS}\"><w:rPr>{}</w:rPr><w:t/></w:r>",
                    joined(&run_properties)
                ),
            ));
        }
    }
    patch(xml, changes)
}
