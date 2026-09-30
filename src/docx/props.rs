//! Schema-ordered editing of Word property containers (`w:pPr`, `w:rPr`, and others).

use anyhow::Result;

use super::xml::{
    Document, PARAGRAPH_PROPERTY_ORDER, Patch, RUN_PROPERTY_ORDER, Tree, WORD_NS, expand_empty,
    prepend_child,
};

/// A property element name and its XML.
pub(super) type Property = (&'static str, String);

/// Sort properties into their schema order.
pub(super) fn ordered(mut properties: Vec<Property>, order: &[&str]) -> Vec<Property> {
    properties.sort_by_key(|(name, _)| rank(order, name));
    properties
}

fn rank(order: &[&str], property: &str) -> usize {
    order
        .iter()
        .position(|name| *name == property)
        .unwrap_or(usize::MAX)
}

/// The concatenated XML of `properties`.
pub(super) fn joined(properties: &[Property]) -> String {
    properties.iter().map(|(_, value)| value.as_str()).collect()
}

/// The schema child order of a property container.
pub(super) fn order_of(container: &str) -> &'static [&'static str] {
    match container {
        "pPr" => PARAGRAPH_PROPERTY_ORDER,
        "rPr" => RUN_PROPERTY_ORDER,
        "tblPr" => TABLE_PROPERTY_ORDER,
        "tcPr" => CELL_PROPERTY_ORDER,
        "trPr" => ROW_PROPERTY_ORDER,
        "sectPr" => SECTION_PROPERTY_ORDER,
        _ => &[],
    }
}

/// Set schema-ordered `properties` inside `parent`'s `name` container, replacing
/// same-named properties and keeping all others. The container is created as the
/// first child of `parent` when missing.
pub(super) fn property_patches(
    tree: &Tree,
    xml: &str,
    parent: usize,
    name: &str,
    properties: &[Property],
    changes: &mut Vec<Patch>,
) -> Result<()> {
    if properties.is_empty() {
        return Ok(());
    }
    let order = order_of(name);
    let Some(container) = tree.child(parent, name) else {
        let content = format!(
            "<w:{name} xmlns:w=\"{WORD_NS}\">{}</w:{name}>",
            joined(properties)
        );
        changes.push(prepend_child(xml, &tree.nodes[parent], content)?);
        return Ok(());
    };
    let node = &tree.nodes[container];
    if node.self_closing() {
        changes.push((
            node.start,
            node.end,
            expand_empty(xml, node, &joined(properties))?,
        ));
        return Ok(());
    }
    for (property, value) in properties {
        if let Some(existing) = tree.child(container, property) {
            let existing = &tree.nodes[existing];
            changes.push((existing.start, existing.end, value.clone()));
        } else {
            let offset = tree
                .children(container)
                .find(|(_, child)| child.word && rank(order, &child.name) > rank(order, property))
                .map_or(node.close_start, |(_, child)| child.start);
            changes.push((offset, offset, value.clone()));
        }
    }
    Ok(())
}

/// The runs (`w:r`) that belong directly to `paragraph`, including those inside
/// hyperlinks and smart tags but not nested paragraphs.
pub(super) fn runs(document: &Document, paragraph: usize) -> Vec<usize> {
    document
        .owned(paragraph)
        .filter_map(|(index, child)| (child.name == "r").then_some(index))
        .collect()
}

/// `CT_TblPr` child order.
pub(super) const TABLE_PROPERTY_ORDER: &[&str] = &[
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
    "tblPrChange",
];

/// `CT_TrPr` child order.
pub(super) const ROW_PROPERTY_ORDER: &[&str] = &[
    "cnfStyle",
    "divId",
    "gridBefore",
    "gridAfter",
    "wBefore",
    "wAfter",
    "cantSplit",
    "trHeight",
    "tblHeader",
    "tblCellSpacing",
    "jc",
    "hidden",
    "ins",
    "del",
    "trPrChange",
];

/// `CT_TcPr` child order.
pub(super) const CELL_PROPERTY_ORDER: &[&str] = &[
    "cnfStyle",
    "tcW",
    "gridSpan",
    "hMerge",
    "vMerge",
    "tcBorders",
    "shd",
    "noWrap",
    "tcMar",
    "textDirection",
    "tcFitText",
    "vAlign",
    "hideMark",
    "headers",
    "cellIns",
    "cellDel",
    "cellMerge",
    "tcPrChange",
];

/// `CT_SectPr` child order.
pub(super) const SECTION_PROPERTY_ORDER: &[&str] = &[
    "headerReference",
    "footerReference",
    "footnotePr",
    "endnotePr",
    "type",
    "pgSz",
    "pgMar",
    "paperSrc",
    "pgBorders",
    "lnNumType",
    "pgNumType",
    "cols",
    "formProt",
    "vAlign",
    "noEndnote",
    "titlePg",
    "textDirection",
    "bidi",
    "rtlGutter",
    "docGrid",
    "printerSettings",
    "sectPrChange",
];
