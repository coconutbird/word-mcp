#![expect(
    dead_code,
    reason = "SCAFFOLD: API for areas under construction; remove once all are used"
)]
//! OOXML part parsing into an offset-indexed element tree, and targeted patching.
//!
//! The parser records byte offsets for every element so edits replace only the
//! targeted spans; everything else in the source XML is preserved verbatim. [`Tree`]
//! handles any package part; [`Document`] adds the main-document view (body and
//! paragraphs) on top of it.

use std::{fmt::Write as _, ops::Deref};

use anyhow::{Context, Result, bail, ensure};
use quick_xml::{events::Event, name::ResolveResult, reader::NsReader};

/// The `WordprocessingML` main namespace.
pub(super) const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/// Largest accepted XML part.
pub(super) const MAX_XML: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_ELEMENTS: usize = 250_000;

/// One namespace-qualified attribute.
#[derive(Clone)]
pub(super) struct Attribute {
    /// Index into [`Tree::namespaces`]; `None` for an unqualified attribute.
    pub(super) ns: Option<u16>,
    /// Local attribute name.
    pub(super) name: String,
    /// Unescaped value.
    pub(super) value: String,
}

/// One XML element with the byte offsets of its tags in the source.
#[derive(Clone)]
pub(super) struct Node {
    /// Local element name.
    pub(super) name: String,
    /// Index into [`Tree::namespaces`]; `None` when the element has no namespace.
    pub(super) ns: Option<u16>,
    /// Whether the element is in the Word namespace.
    pub(super) word: bool,
    /// Offset of the opening `<`.
    pub(super) start: usize,
    /// Offset just past the opening tag.
    pub(super) open_end: usize,
    /// Offset of the closing tag; equals `end` for self-closing elements.
    pub(super) close_start: usize,
    /// Offset just past the element.
    pub(super) end: usize,
    pub(super) parent: Option<usize>,
    /// Unescaped character data directly inside this element.
    pub(super) text: String,
    /// Namespace-resolved Word `w:val` attribute of a Word element.
    pub(super) val: Option<String>,
    /// Every attribute except namespace declarations.
    pub(super) attributes: Vec<Attribute>,
}

impl Node {
    /// Whether this is the Word-namespace element `name`.
    pub(super) fn is(&self, name: &str) -> bool {
        self.word && self.name == name
    }

    /// Whether the element has no separate closing tag.
    pub(super) fn self_closing(&self) -> bool {
        self.open_end == self.end
    }
}

/// A parsed XML part: elements in document order, index 0 being the root.
pub(super) struct Tree {
    pub(super) nodes: Vec<Node>,
    /// Namespace URIs referenced by nodes and attributes.
    pub(super) namespaces: Vec<String>,
}

impl Tree {
    /// Parse any well-formed XML part within the size, depth, and element limits.
    #[expect(
        clippy::too_many_lines,
        reason = "One event loop keeps the offset bookkeeping in a single place"
    )]
    pub(super) fn parse(xml: &str) -> Result<Self> {
        ensure!(xml.len() <= MAX_XML, "XML part exceeds 16 MiB limit");
        let mut reader = NsReader::from_str(xml);
        reader.config_mut().check_end_names = true;
        let mut tree = Self {
            nodes: Vec::new(),
            namespaces: Vec::new(),
        };
        let mut stack: Vec<usize> = Vec::new();
        loop {
            let start = usize::try_from(reader.buffer_position())?;
            let (namespace, event) = reader.read_resolved_event()?;
            let word = is_word(&namespace);
            let ns = if matches!(event, Event::Start(_) | Event::Empty(_)) {
                tree.intern(&namespace)?
            } else {
                None
            };
            let end = usize::try_from(reader.buffer_position())?;
            match &event {
                Event::Start(element) | Event::Empty(element) => {
                    ensure!(stack.len() < MAX_DEPTH, "XML nesting exceeds 128 levels");
                    ensure!(
                        tree.nodes.len() < MAX_ELEMENTS,
                        "XML exceeds 250000 element limit"
                    );
                    let name = std::str::from_utf8(element.local_name().as_ref())?.to_owned();
                    // `ns` and `word` were resolved above, before the reader moved on.
                    let mut attributes = Vec::new();
                    let mut val = None;
                    for attribute in element.attributes() {
                        let attribute = attribute?;
                        let (attribute_ns, local) = reader.resolve_attribute(attribute.key);
                        if matches!(attribute_ns, ResolveResult::Unbound)
                            && attribute.key.as_ref().starts_with(b"xmlns")
                        {
                            continue;
                        }
                        let value = attribute.unescape_value()?.into_owned();
                        if word && is_word(&attribute_ns) && local.as_ref() == b"val" {
                            val = Some(value.clone());
                        }
                        attributes.push(Attribute {
                            ns: tree.intern(&attribute_ns)?,
                            name: std::str::from_utf8(local.as_ref())?.to_owned(),
                            value,
                        });
                    }
                    let index = tree.nodes.len();
                    tree.nodes.push(Node {
                        name,
                        ns,
                        word,
                        start,
                        open_end: end,
                        close_start: end,
                        end,
                        parent: stack.last().copied(),
                        text: String::new(),
                        val,
                        attributes,
                    });
                    if matches!(event, Event::Start(_)) {
                        stack.push(index);
                    }
                }
                Event::End(_) => {
                    let index = stack.pop().context("unexpected closing XML element")?;
                    tree.nodes[index].close_start = start;
                    tree.nodes[index].end = end;
                }
                Event::Text(value) => {
                    if let Some(index) = stack.last() {
                        tree.nodes[*index]
                            .text
                            .push_str(&quick_xml::escape::unescape(&value.xml_content()?)?);
                    }
                }
                Event::CData(value) => {
                    if let Some(index) = stack.last() {
                        tree.nodes[*index].text.push_str(&value.xml_content()?);
                    }
                }
                Event::DocType(_) => bail!("DOCTYPE declarations are not supported"),
                Event::Eof => break,
                Event::GeneralRef(value) => {
                    let entity = format!("&{};", value.decode()?);
                    let decoded = quick_xml::escape::unescape(&entity)
                        .context("unsupported XML entity reference")?;
                    if let Some(index) = stack.last() {
                        tree.nodes[*index].text.push_str(&decoded);
                    }
                }
                _ => {}
            }
        }
        ensure!(stack.is_empty(), "unclosed XML element");
        ensure!(
            !tree.nodes.is_empty()
                && tree
                    .nodes
                    .iter()
                    .filter(|node| node.parent.is_none())
                    .count()
                    == 1,
            "XML part must have exactly one root element"
        );
        Ok(tree)
    }

    fn intern(&mut self, namespace: &ResolveResult<'_>) -> Result<Option<u16>> {
        let ResolveResult::Bound(uri) = namespace else {
            return Ok(None);
        };
        let uri = std::str::from_utf8(uri.as_ref())?;
        if let Some(index) = self.namespaces.iter().position(|known| known == uri) {
            return Ok(Some(u16::try_from(index)?));
        }
        self.namespaces.push(uri.to_owned());
        Ok(Some(u16::try_from(self.namespaces.len() - 1)?))
    }

    fn namespace_index(&self, uri: &str) -> Option<u16> {
        self.namespaces
            .iter()
            .position(|known| known == uri)
            .and_then(|index| u16::try_from(index).ok())
    }

    /// Whether node `index` is the element `name` in namespace `uri`.
    pub(super) fn is(&self, index: usize, uri: &str, name: &str) -> bool {
        let node = &self.nodes[index];
        node.name == name && node.ns.is_some() && node.ns == self.namespace_index(uri)
    }

    /// The value of attribute `name` in namespace `uri` (`None` for unqualified).
    pub(super) fn attr(&self, index: usize, uri: Option<&str>, name: &str) -> Option<&str> {
        let ns = match uri {
            Some(uri) => Some(self.namespace_index(uri)?),
            None => None,
        };
        self.nodes[index]
            .attributes
            .iter()
            .find(|attribute| attribute.ns == ns && attribute.name == name)
            .map(|attribute| attribute.value.as_str())
    }

    /// The value of the Word-namespace attribute `w:name`.
    pub(super) fn word_attr(&self, index: usize, name: &str) -> Option<&str> {
        self.attr(index, Some(WORD_NS), name)
    }

    /// Every element `name` in namespace `uri`, in document order.
    pub(super) fn elements<'a>(
        &'a self,
        uri: &'a str,
        name: &'a str,
    ) -> impl Iterator<Item = usize> + 'a {
        let ns = self.namespace_index(uri);
        self.nodes
            .iter()
            .enumerate()
            .filter(move |(_, node)| ns.is_some() && node.ns == ns && node.name == name)
            .map(|(index, _)| index)
    }

    /// Strict descendants of `index`, in document order.
    pub(super) fn descendants(&self, index: usize) -> impl Iterator<Item = (usize, &Node)> {
        let close = self.nodes[index].close_start;
        self.nodes
            .iter()
            .enumerate()
            .skip(index + 1)
            .take_while(move |(_, node)| node.start < close)
    }

    /// Strict ancestors of `index`, nearest first.
    pub(super) fn ancestors(&self, index: usize) -> impl Iterator<Item = &Node> {
        std::iter::successors(self.nodes[index].parent, |index| self.nodes[*index].parent)
            .map(|index| &self.nodes[index])
    }

    /// The first Word child of `parent` named `name`.
    pub(super) fn child(&self, parent: usize, name: &str) -> Option<usize> {
        self.children(parent)
            .find_map(|(index, node)| node.is(name).then_some(index))
    }

    /// Direct children of `parent`, in document order.
    pub(super) fn children(&self, parent: usize) -> impl Iterator<Item = (usize, &Node)> {
        self.descendants(parent)
            .filter(move |(_, node)| node.parent == Some(parent))
    }

    /// The concatenated text of every Word `w:t` inside `index`.
    pub(super) fn word_text(&self, index: usize) -> String {
        let mut text = String::new();
        for (_, node) in self.descendants(index) {
            match node.name.as_str() {
                "t" if node.word => text.push_str(&node.text),
                "tab" if node.word => text.push('\t'),
                "br" | "cr" if node.word => text.push('\n'),
                "p" if node.word && !text.is_empty() => text.push('\n'),
                _ => {}
            }
        }
        text
    }
}

/// A parsed main document part.
pub(super) struct Document {
    tree: Tree,
    /// Every `w:p` element in document order, including table paragraphs.
    pub(super) paragraphs: Vec<usize>,
    /// The `w:body` element.
    pub(super) body: usize,
}

impl Deref for Document {
    type Target = Tree;

    fn deref(&self) -> &Tree {
        &self.tree
    }
}

impl Document {
    /// Parse and validate a main document part.
    pub(super) fn parse(xml: &str) -> Result<Self> {
        let tree = Tree::parse(xml)?;
        ensure!(tree.nodes[0].is("document"), "invalid Word document root");
        let mut body = None;
        let mut paragraphs = Vec::new();
        for (index, node) in tree.nodes.iter().enumerate() {
            if node.is("p") {
                paragraphs.push(index);
            } else if node.is("body") {
                ensure!(body.is_none(), "multiple document bodies");
                body = Some(index);
            }
        }
        let body = body.context("missing Word document body")?;
        ensure!(
            tree.nodes[body].parent == Some(0),
            "invalid Word document structure"
        );
        Ok(Self {
            tree,
            paragraphs,
            body,
        })
    }

    /// Word descendants of `paragraph` that belong to it rather than a nested paragraph.
    pub(super) fn owned(&self, paragraph: usize) -> impl Iterator<Item = (usize, &Node)> {
        self.descendants(paragraph)
            .filter(move |(index, node)| node.word && self.owner(*index) == Some(paragraph))
    }

    /// The nearest `w:p` containing or equal to `index`.
    pub(super) fn owner(&self, mut index: usize) -> Option<usize> {
        loop {
            let node = &self.nodes[index];
            if node.is("p") {
                return Some(index);
            }
            index = node.parent?;
        }
    }

    /// The zero-based paragraph number of the `w:p` element `index`.
    pub(super) fn paragraph_number(&self, index: usize) -> Option<usize> {
        self.paragraphs.binary_search(&index).ok()
    }

    /// The `w:t` elements carrying the paragraph's text.
    pub(super) fn texts(&self, paragraph: usize) -> Vec<usize> {
        self.owned(paragraph)
            .filter_map(|(index, node)| (node.name == "t").then_some(index))
            .collect()
    }

    /// The paragraph's logical text, with tabs and breaks as `\t` and `\n`.
    pub(super) fn text(&self, paragraph: usize) -> String {
        let mut text = String::new();
        for (_, node) in self.owned(paragraph) {
            match node.name.as_str() {
                "t" => text.push_str(&node.text),
                "tab" => text.push('\t'),
                "br" | "cr" => text.push('\n'),
                _ => {}
            }
        }
        text
    }

    /// The paragraph's `w:pStyle` style id.
    pub(super) fn style(&self, paragraph: usize) -> Option<&str> {
        let properties = self.child(paragraph, "pPr")?;
        self.nodes[self.child(properties, "pStyle")?].val.as_deref()
    }

    /// Whether the paragraph lies inside a table cell.
    pub(super) fn in_table(&self, paragraph: usize) -> bool {
        self.ancestors(paragraph).any(|node| node.is("tc"))
    }

    /// Refuse paragraphs whose structure plain-text edits could corrupt.
    pub(super) fn editable(&self, paragraph: usize) -> Result<()> {
        ensure!(
            !self
                .ancestors(paragraph)
                .any(|node| node.word && matches!(node.name.as_str(), "sdt" | "ins" | "del")),
            "paragraph belongs to an unsupported revision or content control"
        );
        for (_, node) in self.descendants(paragraph) {
            if node.word
                && matches!(
                    node.name.as_str(),
                    "fldChar"
                        | "fldSimple"
                        | "instrText"
                        | "del"
                        | "ins"
                        | "sdt"
                        | "drawing"
                        | "object"
                        | "p"
                        | "tab"
                        | "br"
                        | "cr"
                        | "sym"
                        | "altChunk"
                )
            {
                bail!(
                    "paragraph contains unsupported editing construct: {}",
                    node.name
                );
            }
        }
        Ok(())
    }

    /// Refuse paragraphs holding bookmark, comment or move-range anchors.
    pub(super) fn unanchored(&self, paragraph: usize) -> Result<()> {
        ensure!(
            !self.descendants(paragraph).any(|(_, node)| {
                node.word
                    && (node.name.contains("bookmark")
                        || node.name.contains("comment")
                        || node.name.starts_with("moveFrom")
                        || node.name.starts_with("moveTo"))
            }),
            "paragraph contains anchored ranges; use live Word editing"
        );
        Ok(())
    }
}

fn is_word(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == WORD_NS.as_bytes())
}

/// A replacement of the byte range `start..end` of the source XML.
pub(super) type Patch = (usize, usize, String);

/// Apply non-overlapping patches to the main document part and validate the result.
pub(super) fn patch(xml: &str, changes: Vec<Patch>) -> Result<String> {
    let result = apply(xml, changes)?;
    Document::parse(&result)?;
    Ok(result)
}

/// Apply non-overlapping patches to any XML part and validate that it still parses.
pub(super) fn patch_part(xml: &str, changes: Vec<Patch>) -> Result<String> {
    let result = apply(xml, changes)?;
    Tree::parse(&result)?;
    Ok(result)
}

fn apply(xml: &str, mut changes: Vec<Patch>) -> Result<String> {
    changes.sort_by_key(|change| change.0);
    for pair in changes.windows(2) {
        ensure!(pair[0].1 <= pair[1].0, "overlapping XML edits");
    }
    let mut result = xml.to_owned();
    for (start, end, replacement) in changes.into_iter().rev() {
        result.replace_range(start..end, &replacement);
    }
    Ok(result)
}

/// Escape text for XML content or attribute values.
pub(super) fn escaped(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

/// Refuse characters XML 1.0 cannot represent.
pub(super) fn validate_text(value: &str) -> Result<()> {
    ensure!(
        value
            .chars()
            .all(|character| matches!(character, '\t' | '\n' | '\r')
                || (character >= ' ' && !matches!(character, '\u{fffe}' | '\u{ffff}'))),
        "text contains an invalid XML control character"
    );
    Ok(())
}

/// A standalone paragraph holding `text` with an optional style id.
pub(super) fn paragraph_xml(text: &str, style: Option<&str>) -> Result<String> {
    validate_text(text)?;
    if let Some(style) = style {
        validate_text(style)?;
    }
    let properties = style.map_or_else(String::new, |style| {
        format!("<w:pPr><w:pStyle w:val=\"{}\"/></w:pPr>", escaped(style))
    });
    Ok(format!(
        "<w:p xmlns:w=\"{WORD_NS}\">{properties}<w:r>{}</w:r></w:p>",
        text_xml(text)
    ))
}

/// Run content for `text`, mapping newlines to `w:br` and tabs to `w:tab`.
pub(super) fn text_xml(text: &str) -> String {
    let normalized = text.replace("\r\n", "\n").replace('\r', "\n");
    let mut result = String::new();
    let mut start = 0;
    for (offset, character) in normalized.char_indices() {
        let name = match character {
            '\n' => "br",
            '\t' => "tab",
            _ => continue,
        };
        // Writing to a String cannot fail.
        let _ = write!(
            result,
            "<w:t xmlns:w=\"{WORD_NS}\" xml:space=\"preserve\">{}</w:t><w:{name} xmlns:w=\"{WORD_NS}\"/>",
            escaped(&normalized[start..offset])
        );
        start = offset + character.len_utf8();
    }
    let _ = write!(
        result,
        "<w:t xmlns:w=\"{WORD_NS}\" xml:space=\"preserve\">{}</w:t>",
        escaped(&normalized[start..])
    );
    result
}

/// A self-closing Word property element `<w:name w:val="value"/>`.
pub(super) fn property_xml(name: &str, value: &str) -> String {
    format!(
        "<w:{name} xmlns:w=\"{WORD_NS}\" w:val=\"{}\"/>",
        escaped(value)
    )
}

/// Rewrite the self-closing `node` as an open element containing `content`.
pub(super) fn expand_empty(xml: &str, node: &Node, content: &str) -> Result<String> {
    let opening = xml[node.start..node.end]
        .strip_suffix("/>")
        .context("expected self-closing XML element")?;
    let qualified_name = opening
        .trim_start_matches('<')
        .split_whitespace()
        .next()
        .context("missing XML name")?;
    Ok(format!("{opening}>{content}</{qualified_name}>"))
}

/// A patch inserting `content` as the first child of `node`, expanding it when
/// it is self-closing.
pub(super) fn prepend_child(xml: &str, node: &Node, content: String) -> Result<Patch> {
    Ok(if node.self_closing() {
        (node.start, node.end, expand_empty(xml, node, &content)?)
    } else {
        (node.open_end, node.open_end, content)
    })
}

/// A patch inserting `content` as the last child of `node`, expanding it when it is
/// self-closing.
pub(super) fn append_child(xml: &str, node: &Node, content: String) -> Result<Patch> {
    Ok(if node.self_closing() {
        (node.start, node.end, expand_empty(xml, node, &content)?)
    } else {
        (node.close_start, node.close_start, content)
    })
}

/// `CT_PPr` child order from the OOXML schema; Word rejects out-of-order children.
pub(super) const PARAGRAPH_PROPERTY_ORDER: &[&str] = &[
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
];

/// `CT_RPr` child order from the OOXML schema.
pub(super) const RUN_PROPERTY_ORDER: &[&str] = &[
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "webHidden",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
    "rPrChange",
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reject_doctype_depth_and_control_characters() {
        assert!(Document::parse("<!DOCTYPE document><w:document/>").is_err());
        let nested = format!("{}{}", "<a>".repeat(129), "</a>".repeat(129));
        assert!(Document::parse(&nested).is_err());
        assert!(validate_text("bad\u{ffff}").is_err());
    }

    #[test]
    fn records_namespaced_style_values_only() -> Result<()> {
        let xml = format!(
            "<q:document xmlns:q=\"{WORD_NS}\" xmlns:o=\"urn:other\"><q:body>\
             <q:p><q:pPr><q:pStyle q:val=\"Heading&amp;1\"/></q:pPr></q:p>\
             <q:p><q:pPr><q:pStyle val=\"Unqualified\" o:val=\"Foreign\"/></q:pPr></q:p>\
             </q:body></q:document>"
        );
        let document = Document::parse(&xml)?;
        assert_eq!(document.style(document.paragraphs[0]), Some("Heading&1"));
        assert_eq!(document.style(document.paragraphs[1]), None);
        Ok(())
    }
}
