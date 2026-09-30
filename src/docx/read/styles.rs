//! Read-side views of `styles.xml` and `numbering.xml`, property resolution through
//! the style hierarchy, and paragraph text extraction for any Word part.

use std::collections::HashMap;

use anyhow::Result;

use super::super::{
    ooxml::{ns, rel},
    package::Package,
    xml::{Tree, WORD_NS},
};

/// Longest `basedOn` chain followed; guards against cycles.
const MAX_CHAIN: usize = 32;

/// The part related to the main document by `kind`, parsed.
pub(super) fn related_tree(package: &Package, kind: &str) -> Result<Option<(String, Tree)>> {
    let Some(name) = package.document_part(kind)? else {
        return Ok(None);
    };
    let Some(xml) = package.text(&name)? else {
        return Ok(None);
    };
    let tree = Tree::parse(xml)?;
    Ok(Some((name, tree)))
}

/// Whether an `ST_OnOff` value means "on".
fn on(value: Option<&str>) -> bool {
    !matches!(value, Some("0" | "false" | "off"))
}

/// The nearest ancestor of `index` (or `index` itself) that is the Word element `name`,
/// searching no higher than `limit`.
pub(super) fn enclosing(tree: &Tree, mut index: usize, name: &str, limit: usize) -> Option<usize> {
    loop {
        if tree.nodes[index].is(name) {
            return Some(index);
        }
        if index == limit {
            return None;
        }
        index = tree.nodes[index].parent?;
    }
}

/// Word descendants of `paragraph` that belong to it rather than to a nested paragraph
/// (text boxes), in document order.
pub(super) fn own_nodes(tree: &Tree, paragraph: usize) -> impl Iterator<Item = usize> + '_ {
    let mut skip_until = 0;
    tree.descendants(paragraph)
        .filter_map(move |(index, node)| {
            if node.start < skip_until {
                return None;
            }
            if node.is("p") {
                skip_until = node.end;
                return None;
            }
            node.word.then_some(index)
        })
}

/// Whether node `index` is run content (a direct child of `w:r`).
pub(super) fn in_run(tree: &Tree, index: usize) -> bool {
    tree.nodes[index]
        .parent
        .is_some_and(|parent| tree.nodes[parent].is("r"))
}

/// The logical text of `paragraph` in any Word part: run text, with run-level tabs as
/// `\t` and breaks as `\n`; text of nested paragraphs and deleted text excluded.
pub(super) fn paragraph_text(tree: &Tree, paragraph: usize) -> String {
    let mut text = String::new();
    for index in own_nodes(tree, paragraph) {
        let node = &tree.nodes[index];
        match node.name.as_str() {
            "t" => text.push_str(&node.text),
            "tab" if in_run(tree, index) => text.push('\t'),
            "br" | "cr" if in_run(tree, index) => text.push('\n'),
            _ => {}
        }
    }
    text
}

/// One property container (`w:pPr`/`w:rPr`) contributing to a resolved value; earlier
/// sources win.
#[derive(Clone, Copy)]
pub(super) struct Source<'a> {
    pub(super) tree: &'a Tree,
    pub(super) node: usize,
}

/// The first source's child element `name`.
pub(super) fn element<'a>(sources: &[Source<'a>], name: &str) -> Option<(&'a Tree, usize)> {
    sources.iter().find_map(|source| {
        source
            .tree
            .child(source.node, name)
            .map(|index| (source.tree, index))
    })
}

/// The value of `val` on the first source's child element `name`.
pub(super) fn value<'a>(sources: &[Source<'a>], name: &str) -> Option<&'a str> {
    let (tree, index) = element(sources, name)?;
    tree.nodes[index].val.as_deref()
}

/// Word attribute `attribute` of the first source whose child `name` carries it.
pub(super) fn attribute<'a>(
    sources: &[Source<'a>],
    name: &str,
    attribute: &str,
) -> Option<&'a str> {
    sources.iter().find_map(|source| {
        let index = source.tree.child(source.node, name)?;
        source.tree.word_attr(index, attribute)
    })
}

/// A toggle property: present and not switched off.
pub(super) fn toggle(sources: &[Source<'_>], name: &str) -> Option<bool> {
    let (tree, index) = element(sources, name)?;
    Some(on(tree.nodes[index].val.as_deref()))
}

/// The styles part.
#[derive(Default)]
pub(super) struct Styles {
    tree: Option<Tree>,
    by_id: HashMap<String, usize>,
    default_paragraph: Option<String>,
}

impl Styles {
    pub(super) fn load(package: &Package) -> Result<Self> {
        let Some((_, tree)) = related_tree(package, rel::STYLES)? else {
            return Ok(Self::default());
        };
        let mut by_id = HashMap::new();
        let mut default_paragraph = None;
        for index in tree.elements(WORD_NS, "style") {
            let Some(id) = tree.word_attr(index, "styleId") else {
                continue;
            };
            if tree.word_attr(index, "type") == Some("paragraph")
                && tree
                    .word_attr(index, "default")
                    .is_some_and(|value| on(Some(value)))
            {
                default_paragraph.get_or_insert_with(|| id.to_owned());
            }
            by_id.entry(id.to_owned()).or_insert(index);
        }
        Ok(Self {
            tree: Some(tree),
            by_id,
            default_paragraph,
        })
    }

    /// The display name of style `id`.
    pub(super) fn name(&self, id: &str) -> Option<&str> {
        let tree = self.tree.as_ref()?;
        let name = tree.child(*self.by_id.get(id)?, "name")?;
        tree.nodes[name].val.as_deref()
    }

    /// The paragraph style id of a paragraph whose direct style is `direct`.
    pub(super) fn paragraph_style<'a>(&'a self, direct: Option<&'a str>) -> Option<&'a str> {
        direct.or(self.default_paragraph.as_deref())
    }

    /// Style `id` and its `basedOn` ancestors, as `(id, style element)`.
    pub(super) fn chain<'a>(&'a self, id: Option<&'a str>) -> Vec<(&'a str, usize)> {
        let mut chain = Vec::new();
        let (Some(tree), Some(mut id)) = (self.tree.as_ref(), id) else {
            return chain;
        };
        while chain.len() < MAX_CHAIN {
            let Some(&style) = self.by_id.get(id) else {
                break;
            };
            if chain.iter().any(|(_, known)| *known == style) {
                break;
            }
            chain.push((id, style));
            let Some(based) = tree
                .child(style, "basedOn")
                .and_then(|based| tree.nodes[based].val.as_deref())
            else {
                break;
            };
            id = based;
        }
        chain
    }

    /// Property sources of a style chain (`pPr` or `rPr` children), then document
    /// defaults (`docDefaults/pPrDefault/pPr` or `rPrDefault/rPr`).
    pub(super) fn sources(&self, id: Option<&str>, container: &str) -> Vec<Source<'_>> {
        let Some(tree) = self.tree.as_ref() else {
            return Vec::new();
        };
        let mut sources: Vec<Source<'_>> = self
            .chain(id)
            .into_iter()
            .filter_map(|(_, style)| tree.child(style, container))
            .map(|node| Source { tree, node })
            .collect();
        let default = if container == "pPr" {
            "pPrDefault"
        } else {
            "rPrDefault"
        };
        if let Some(node) = tree
            .elements(WORD_NS, "docDefaults")
            .next()
            .and_then(|defaults| tree.child(defaults, default))
            .and_then(|default| tree.child(default, container))
        {
            sources.push(Source { tree, node });
        }
        sources
    }

    /// The heading level implied by a style's id or name: 1-9 for "Heading N", 0 for
    /// the Title style.
    pub(super) fn named_level(&self, id: Option<&str>) -> Option<u8> {
        self.chain(id)
            .into_iter()
            .map(|(id, _)| id)
            .chain(id)
            .find_map(|id| name_level(id).or_else(|| self.name(id).and_then(name_level)))
    }
}

/// `Heading1`..`Heading9`, "heading 1".."heading 9", and "Title".
fn name_level(name: &str) -> Option<u8> {
    let compact: String = name
        .chars()
        .filter(|character| !character.is_whitespace())
        .collect::<String>()
        .to_ascii_lowercase();
    if compact == "title" {
        return Some(0);
    }
    let level: u8 = compact.strip_prefix("heading")?.parse().ok()?;
    (1..=9).contains(&level).then_some(level)
}

/// Paragraph property sources: direct `w:pPr`, then the paragraph style chain, then
/// document defaults.
pub(super) fn paragraph_sources<'a>(
    tree: &'a Tree,
    paragraph: usize,
    styles: &'a Styles,
    style: Option<&'a str>,
) -> Vec<Source<'a>> {
    let mut sources = Vec::new();
    if let Some(node) = tree.child(paragraph, "pPr") {
        sources.push(Source { tree, node });
    }
    sources.extend(styles.sources(style, "pPr"));
    sources
}

/// The direct `w:pStyle` of a paragraph in any part.
pub(super) fn direct_style(tree: &Tree, paragraph: usize) -> Option<&str> {
    let properties = tree.child(paragraph, "pPr")?;
    tree.nodes[tree.child(properties, "pStyle")?].val.as_deref()
}

/// The heading level of a paragraph: an explicit outline level (direct or from its
/// style chain; 0-8 map to levels 1-9, 9 means body text), else a Heading/Title style.
pub(super) fn heading_level(
    sources: &[Source<'_>],
    styles: &Styles,
    style: Option<&str>,
) -> Option<u8> {
    match value(sources, "outlineLvl").and_then(|level| level.parse::<u8>().ok()) {
        Some(level @ 0..=8) => Some(level + 1),
        Some(_) => None,
        None => styles.named_level(style),
    }
}

/// A paragraph's numbering reference: `(numId, ilvl)`, `None` when unnumbered.
pub(super) fn numbering_reference(sources: &[Source<'_>]) -> Option<(String, u8)> {
    let find = |name: &str| {
        sources.iter().find_map(|source| {
            let properties = source.tree.child(source.node, "numPr")?;
            let child = source.tree.child(properties, name)?;
            source.tree.nodes[child].val.as_deref()
        })
    };
    let id = find("numId")?;
    if id == "0" {
        return None;
    }
    let level = find("ilvl")
        .and_then(|level| level.parse().ok())
        .unwrap_or(0);
    Some((id.to_owned(), level))
}

/// One numbering level definition.
pub(super) struct Level {
    /// `w:numFmt`, such as `bullet` or `decimal`.
    pub(super) format: String,
    /// `w:lvlText`, such as `%1.`.
    pub(super) text: Option<String>,
    /// `w:start`.
    pub(super) start: u32,
}

/// The numbering part.
#[derive(Default)]
pub(super) struct Numbering {
    tree: Option<Tree>,
    nums: HashMap<String, usize>,
    abstracts: HashMap<String, usize>,
}

impl Numbering {
    pub(super) fn load(package: &Package) -> Result<Self> {
        let Some((_, tree)) = related_tree(package, rel::NUMBERING)? else {
            return Ok(Self::default());
        };
        let mut nums = HashMap::new();
        let mut abstracts = HashMap::new();
        for index in tree.elements(WORD_NS, "num") {
            if let Some(id) = tree.word_attr(index, "numId") {
                nums.entry(id.to_owned()).or_insert(index);
            }
        }
        for index in tree.elements(WORD_NS, "abstractNum") {
            if let Some(id) = tree.word_attr(index, "abstractNumId") {
                abstracts.entry(id.to_owned()).or_insert(index);
            }
        }
        Ok(Self {
            tree: Some(tree),
            nums,
            abstracts,
        })
    }

    /// Level `level` of numbering instance `id`, honoring level overrides.
    pub(super) fn level(&self, id: &str, level: u8) -> Option<Level> {
        let tree = self.tree.as_ref()?;
        let num = *self.nums.get(id)?;
        let wanted = level.to_string();
        let mut start = None;
        let mut definition = None;
        for (index, node) in tree.children(num) {
            if node.is("lvlOverride") && tree.word_attr(index, "ilvl") == Some(wanted.as_str()) {
                start = tree
                    .child(index, "startOverride")
                    .and_then(|child| tree.nodes[child].val.as_deref()?.parse().ok());
                definition = tree.child(index, "lvl");
            }
        }
        if definition.is_none() {
            let abstract_id = tree.nodes[tree.child(num, "abstractNumId")?]
                .val
                .as_deref()?;
            let abstract_num = *self.abstracts.get(abstract_id)?;
            definition = tree.children(abstract_num).find_map(|(index, node)| {
                (node.is("lvl") && tree.word_attr(index, "ilvl") == Some(wanted.as_str()))
                    .then_some(index)
            });
        }
        let definition = definition?;
        let property = |name: &str| {
            tree.child(definition, name)
                .and_then(|child| tree.nodes[child].val.clone())
        };
        Some(Level {
            format: property("numFmt").unwrap_or_else(|| "decimal".to_owned()),
            text: property("lvlText"),
            start: start
                .or_else(|| property("start").and_then(|start| start.parse().ok()))
                .unwrap_or(1),
        })
    }
}

/// External or anchor targets of the main document's relationships, by id.
pub(super) fn link_targets(package: &Package) -> Result<HashMap<String, String>> {
    Ok(package
        .relationships(super::super::package::DOCUMENT_PART)?
        .into_iter()
        .filter(|relationship| relationship.kind == rel::HYPERLINK)
        .map(|relationship| (relationship.id, relationship.target))
        .collect())
}

/// The target of a `w:hyperlink` element: its relationship target and/or `#anchor`.
pub(super) fn hyperlink_target(
    tree: &Tree,
    hyperlink: usize,
    links: &HashMap<String, String>,
) -> Option<String> {
    let target = tree
        .attr(hyperlink, Some(ns::R), "id")
        .and_then(|id| links.get(id))
        .cloned();
    let anchor = tree.word_attr(hyperlink, "anchor");
    match (target, anchor) {
        (Some(target), Some(anchor)) => Some(format!("{target}#{anchor}")),
        (Some(target), None) => Some(target),
        (None, Some(anchor)) => Some(format!("#{anchor}")),
        (None, None) => None,
    }
}

/// The target of a `HYPERLINK` field instruction, such as
/// `HYPERLINK "https://example.com" \l "anchor"`.
pub(super) fn field_target(instruction: &str) -> Option<String> {
    let rest = instruction.trim_start();
    let keyword = rest.get(..9)?;
    if !keyword.eq_ignore_ascii_case("HYPERLINK") {
        return None;
    }
    let mut tokens = Vec::new();
    let mut chars = rest[9..].chars().peekable();
    while let Some(&character) = chars.peek() {
        if character.is_whitespace() {
            chars.next();
        } else if character == '"' {
            chars.next();
            let token: String = chars.by_ref().take_while(|&c| c != '"').collect();
            tokens.push((token, true));
        } else {
            let mut token = String::new();
            while let Some(&c) = chars.peek() {
                if c.is_whitespace() {
                    break;
                }
                token.push(c);
                chars.next();
            }
            tokens.push((token, false));
        }
    }
    let mut url = None;
    let mut anchor = None;
    let mut iterator = tokens.into_iter();
    while let Some((token, quoted)) = iterator.next() {
        if !quoted && token.starts_with('\\') {
            if token.eq_ignore_ascii_case("\\l") {
                anchor = iterator.next().map(|(value, _)| value);
            } else if matches!(token.to_ascii_lowercase().as_str(), "\\o" | "\\t" | "\\m") {
                iterator.next();
            }
        } else if url.is_none() {
            url = Some(token);
        }
    }
    match (url, anchor) {
        (Some(url), Some(anchor)) => Some(format!("{url}#{anchor}")),
        (Some(url), None) => Some(url),
        (None, Some(anchor)) => Some(format!("#{anchor}")),
        (None, None) => None,
    }
}

#[cfg(test)]
mod tests {
    use super::{field_target, name_level};

    #[test]
    fn heading_names_and_field_links() {
        assert_eq!(name_level("Heading3"), Some(3));
        assert_eq!(name_level("heading 9"), Some(9));
        assert_eq!(name_level("Title"), Some(0));
        assert_eq!(name_level("Heading10"), None);
        assert_eq!(name_level("Normal"), None);
        assert_eq!(
            field_target(r#" HYPERLINK "https://example.com/a b" \o "tip" "#).as_deref(),
            Some("https://example.com/a b")
        );
        assert_eq!(
            field_target(r#"HYPERLINK \l "_Toc1""#).as_deref(),
            Some("#_Toc1")
        );
        assert_eq!(field_target("PAGE"), None);
    }
}
