//! `docx_read`: read saved documents without changing them.

mod markdown;
mod styles;

use std::{
    fs,
    path::{Path, PathBuf},
};

use anyhow::{Context, Result, bail, ensure};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::{Value, json};

use self::{
    markdown::{Context as MarkdownContext, notes, plain_text},
    styles::{
        Numbering, Source, Styles, attribute, enclosing, heading_level, in_run,
        numbering_reference, own_nodes, paragraph_sources, paragraph_text, related_tree, toggle,
        value,
    },
};
use super::{
    Definition,
    ooxml::{ns, rel},
    package::{DOCUMENT_PART, Package, read, write_atomic},
    xml::{Document, Tree, WORD_NS, escaped},
};
use crate::tool::{Effect, Output, parse, tool};

const NAME: &str = "docx_read";
/// Largest part text returned by `part`.
const MAX_PART_TEXT: usize = 1024 * 1024;

pub(super) const TOOL: Definition = Definition {
    name: NAME,
    tool: || {
        tool::<Arguments>(
            NAME,
            "Read a saved .docx without changing it: paragraphs (style, table membership), plain text, Markdown (headings, emphasis, lists, tables, links, footnotes), the heading outline, text search with context (body or every story), one paragraph's formatting and runs, package parts and raw part XML, or a logical HTML preview. Paragraph indices are zero-based and shared with the other docx tools.",
            Effect::ReadOnly,
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

/// Which stories `find` searches.
#[derive(Clone, Copy, Default, PartialEq, Eq, Deserialize, JsonSchema)]
#[serde(rename_all = "lowercase")]
enum Scope {
    /// Main-document paragraphs, including tables.
    #[default]
    Body,
    /// The body plus headers, footers, footnotes, endnotes, and comments.
    All,
}

#[derive(Deserialize, JsonSchema)]
#[serde(tag = "action", rename_all = "snake_case", deny_unknown_fields)]
enum Operation {
    /// Main-document paragraphs, including table paragraphs, in document order.
    Paragraphs {
        /// Zero-based index of the first paragraph to return.
        start: Option<usize>,
        /// Maximum number of paragraphs to return.
        limit: Option<usize>,
    },
    /// An escaped logical HTML preview (not a page-accurate rendering).
    Preview {
        /// Also write the preview to this `.html` or `.htm` file.
        output_path: Option<PathBuf>,
        /// Replace an existing file at `output_path`.
        #[serde(default)]
        overwrite: bool,
    },
    /// Headings in document order: paragraphs with an outline level (direct or from
    /// their style), or a Heading 1-9 / Title style. Level 0 is the Title style.
    Outline {},
    /// Literal text search. Offsets count Unicode characters within the paragraph text
    /// (tabs as `\t`, line breaks as `\n`); matches do not span paragraphs.
    Find {
        /// Text to find.
        #[schemars(length(min = 1))]
        text: String,
        /// Match letter case exactly.
        #[serde(default = "yes")]
        match_case: bool,
        /// Only match whole words.
        #[serde(default)]
        whole_word: bool,
        /// Maximum number of matches to return.
        #[serde(default = "default_max_results")]
        #[schemars(range(min = 1, max = 500))]
        max_results: usize,
        /// Characters of context to return before and after each match.
        #[serde(default = "default_context_chars")]
        #[schemars(range(max = 1000))]
        context_chars: usize,
        /// Search only the body (default) or every story.
        #[serde(default)]
        scope: Scope,
    },
    /// One paragraph's formatting: style, alignment, spacing, indents, numbering, and
    /// keep flags resolved through its style chain and document defaults, plus the
    /// direct formatting of each run.
    Formatting {
        /// Zero-based paragraph index.
        index: usize,
    },
    /// The body as Markdown: headings, bold/italic, bullet and numbered lists,
    /// tables as pipe tables, hyperlinks, line breaks, and footnotes.
    Markdown {
        /// Zero-based paragraph index where rendering starts; a table is included when
        /// its first paragraph is in range.
        start: Option<usize>,
        /// Maximum number of paragraphs whose blocks are rendered.
        #[schemars(range(min = 1))]
        limit: Option<usize>,
    },
    /// Without `name`, list the package parts with sizes and content types; with
    /// `name`, return that part's XML text (at most 1 MiB).
    Part {
        /// Part name such as `word/styles.xml`.
        #[schemars(length(min = 1))]
        name: Option<String>,
    },
    /// Plain text of the body: one line per paragraph, one line per table row with
    /// tab-separated cells.
    Text {},
}

fn yes() -> bool {
    true
}

fn default_max_results() -> usize {
    50
}

fn default_context_chars() -> usize {
    40
}

fn call(arguments: Value) -> Result<Output> {
    let Arguments { path, operation } = parse(NAME, arguments)?;
    let package = read(&path)?;
    let document = || -> Result<Document> { Ok(package.parse()?.1) };
    Ok(match operation {
        Operation::Paragraphs { start, limit } => paragraphs(&path, &document()?, start, limit),
        Operation::Preview {
            output_path,
            overwrite,
        } => preview(&path, &document()?, output_path.as_deref(), overwrite)?,
        Operation::Outline {} => outline(&path, &package, &document()?)?,
        Operation::Find {
            text,
            match_case,
            whole_word,
            max_results,
            context_chars,
            scope,
        } => {
            ensure!(!text.is_empty(), "text must not be empty");
            ensure!(
                (1..=500).contains(&max_results),
                "max_results must be between 1 and 500"
            );
            ensure!(context_chars <= 1000, "context_chars must be at most 1000");
            let search = Search {
                needle: fold(&text, match_case),
                match_case,
                whole_word,
                max_results,
                context_chars,
            };
            find(&path, &package, &document()?, &search, scope)?
        }
        Operation::Formatting { index } => formatting(&path, &package, &document()?, index)?,
        Operation::Markdown { start, limit } => {
            ensure!(limit != Some(0), "limit must be at least 1");
            markdown_body(&path, &package, &document()?, start, limit)?
        }
        Operation::Part { name } => part(&path, &package, name.as_deref())?,
        Operation::Text {} => {
            let document = document()?;
            let text = plain_text(&document);
            json!({
                "path": path,
                "text": text,
                "characters": text.chars().count(),
                "total_paragraphs": document.paragraphs.len(),
            })
        }
    }
    .into())
}

fn paragraphs(
    path: &Path,
    document: &Document,
    start: Option<usize>,
    limit: Option<usize>,
) -> Value {
    let start = start.unwrap_or(0);
    let mut texts = Vec::new();
    let paragraphs: Vec<Value> = document
        .paragraphs
        .iter()
        .enumerate()
        .skip(start)
        .take(limit.unwrap_or(usize::MAX))
        .map(|(index, paragraph)| {
            let text = document.text(*paragraph);
            let value = json!({
                "index": index,
                "text": text,
                "style": document.style(*paragraph),
                "in_table": document.in_table(*paragraph),
            });
            texts.push(text);
            value
        })
        .collect();
    json!({
        "path": path,
        "paragraphs": paragraphs,
        "text": texts.join("\n"),
        "total_paragraphs": document.paragraphs.len(),
        "start": start,
    })
}

/// Render escaped logical HTML, optionally saving it beside the document.
fn preview(
    path: &Path,
    document: &Document,
    output: Option<&Path>,
    overwrite: bool,
) -> Result<Value> {
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
    if let Some(output) = output {
        ensure!(
            output
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("html")
                    || extension.eq_ignore_ascii_case("htm")),
            "preview output must use .html or .htm extension"
        );
        ensure!(output != path, "preview cannot overwrite source document");
        if output.exists() {
            ensure!(
                fs::canonicalize(output)? != fs::canonicalize(path)?,
                "preview cannot overwrite source document"
            );
        }
        write_atomic(output, html.as_bytes(), None, overwrite)?;
    }
    Ok(json!({"path": path, "html": html, "output_path": output, "preview_type": "logical_html"}))
}

/// Headings with their zero-based paragraph index and level.
fn outline(path: &Path, package: &Package, document: &Document) -> Result<Value> {
    let styles = Styles::load(package)?;
    let headings: Vec<Value> = document
        .paragraphs
        .iter()
        .enumerate()
        .filter_map(|(index, &paragraph)| {
            let style = styles.paragraph_style(document.style(paragraph));
            let sources = paragraph_sources(document, paragraph, &styles, style);
            let level = heading_level(&sources, &styles, style)?;
            Some(json!({
                "index": index,
                "level": level,
                "text": paragraph_text(document, paragraph),
                "style": style,
            }))
        })
        .collect();
    Ok(json!({
        "path": path,
        "headings": headings,
        "total_paragraphs": document.paragraphs.len(),
    }))
}

/// A validated `find` request.
struct Search {
    /// The search text, case-folded unless `match_case`.
    needle: Vec<char>,
    match_case: bool,
    whole_word: bool,
    max_results: usize,
    context_chars: usize,
}

/// Characters of `text`, lower-cased one-to-one unless `match_case`, so offsets in
/// the folded text are offsets in the original.
fn fold(text: &str, match_case: bool) -> Vec<char> {
    text.chars()
        .map(|character| {
            if match_case {
                return character;
            }
            let mut lower = character.to_lowercase();
            match (lower.next(), lower.next()) {
                (Some(lower), None) => lower,
                _ => character,
            }
        })
        .collect()
}

fn is_word_char(character: char) -> bool {
    character.is_alphanumeric() || character == '_'
}

/// Non-overlapping matches of the search in `text`, as character ranges.
fn matches(search: &Search, text: &str) -> Vec<(usize, usize)> {
    let haystack = fold(text, search.match_case);
    let needle = &search.needle;
    let mut found = Vec::new();
    let mut start = 0;
    while needle.len() <= haystack.len() && start <= haystack.len() - needle.len() {
        let end = start + needle.len();
        if haystack[start..end] == needle[..]
            && (!search.whole_word
                || (start
                    .checked_sub(1)
                    .is_none_or(|before| !is_word_char(haystack[before]))
                    && haystack.get(end).is_none_or(|after| !is_word_char(*after))))
        {
            found.push((start, end));
            start = end;
        } else {
            start += 1;
        }
    }
    found
}

/// One searchable story: a part and the paragraphs to search in it.
struct Story<'a> {
    part: String,
    kind: &'static str,
    tree: &'a Tree,
    paragraphs: Vec<usize>,
}

fn find(
    path: &Path,
    package: &Package,
    document: &Document,
    search: &Search,
    scope: Scope,
) -> Result<Value> {
    let mut trees = Vec::new();
    if scope == Scope::All {
        for relationship in package.relationships(DOCUMENT_PART)? {
            let kind = match relationship.kind.as_str() {
                rel::HEADER => "header",
                rel::FOOTER => "footer",
                rel::FOOTNOTES => "footnote",
                rel::ENDNOTES => "endnote",
                rel::COMMENTS => "comment",
                _ => continue,
            };
            if relationship.external {
                continue;
            }
            let name = super::package::resolve(DOCUMENT_PART, &relationship.target);
            let Some(xml) = package.text(&name)? else {
                continue;
            };
            let tree = Tree::parse(xml).with_context(|| format!("cannot parse {name}"))?;
            trees.push((name, kind, tree));
        }
    }
    let mut stories = vec![Story {
        part: DOCUMENT_PART.to_owned(),
        kind: "body",
        tree: document,
        paragraphs: document.paragraphs.clone(),
    }];
    stories.extend(trees.iter().map(|(part, kind, tree)| Story {
        part: part.clone(),
        kind,
        tree,
        paragraphs: tree.elements(WORD_NS, "p").collect(),
    }));
    let mut results = Vec::new();
    let mut truncated = false;
    'stories: for story in &stories {
        for (number, &paragraph) in story.paragraphs.iter().enumerate() {
            let text = paragraph_text(story.tree, paragraph);
            let found = matches(search, &text);
            if found.is_empty() {
                continue;
            }
            let characters: Vec<char> = text.chars().collect();
            let slice = |from: usize, to: usize| characters[from..to].iter().collect::<String>();
            let id = match story.kind {
                "footnote" | "endnote" | "comment" => {
                    enclosing(story.tree, paragraph, story.kind, 0)
                        .and_then(|note| story.tree.word_attr(note, "id"))
                }
                _ => None,
            };
            for (start, end) in found {
                if results.len() == search.max_results {
                    truncated = true;
                    break 'stories;
                }
                let mut result = json!({
                    "part": story.part,
                    "story": story.kind,
                    "paragraph": number,
                    "offset": start,
                    "length": end - start,
                    "text": slice(start, end),
                    "before": slice(start.saturating_sub(search.context_chars), start),
                    "after": slice(end, (end + search.context_chars).min(characters.len())),
                });
                if let Some(id) = id {
                    result["id"] = json!(id);
                }
                results.push(result);
            }
        }
    }
    Ok(json!({
        "path": path,
        "count": results.len(),
        "truncated": truncated,
        "matches": results,
    }))
}

/// A twips measurement in points.
fn twips(value: &str) -> Option<f64> {
    value.parse::<f64>().ok().map(|twips| twips / 20.0)
}

/// One paragraph's resolved paragraph formatting and direct run formatting.
fn formatting(path: &Path, package: &Package, document: &Document, index: usize) -> Result<Value> {
    let paragraph = *document.paragraphs.get(index).with_context(|| {
        format!(
            "paragraph {index} is out of range; the document has {} paragraphs",
            document.paragraphs.len()
        )
    })?;
    let styles = Styles::load(package)?;
    let numbering = Numbering::load(package)?;
    let style = styles.paragraph_style(document.style(paragraph));
    let sources = paragraph_sources(document, paragraph, &styles, style);
    let measure = |name: &str, attributes: &[&str]| {
        attributes
            .iter()
            .find_map(|name_attribute| attribute(&sources, name, name_attribute))
            .and_then(twips)
    };
    let line = attribute(&sources, "spacing", "line").and_then(|line| line.parse::<f64>().ok());
    let (line_spacing, line_rule) = match (line, attribute(&sources, "spacing", "lineRule")) {
        (None, _) => (None, None),
        (Some(line), Some("exact")) => (Some(line / 20.0), Some("exact")),
        (Some(line), Some("atLeast")) => (Some(line / 20.0), Some("at_least")),
        (Some(line), _) => (Some(line / 240.0), Some("multiple")),
    };
    let numbering = numbering_reference(&sources).map(|(id, level)| {
        let definition = numbering.level(&id, level);
        json!({
            "num_id": id.parse::<i64>().map_or_else(|_| json!(id), |id| json!(id)),
            "ilvl": level,
            "format": definition.as_ref().map(|definition| definition.format.clone()),
            "level_text": definition.and_then(|definition| definition.text),
        })
    });
    let flag = |name: &str| toggle(&sources, name).unwrap_or(false);
    Ok(json!({
        "path": path,
        "index": index,
        "text": paragraph_text(document, paragraph),
        "in_table": document.in_table(paragraph),
        "style_id": style,
        "style_name": style.and_then(|style| styles.name(style)),
        "alignment": value(&sources, "jc"),
        "spacing": {
            "before_pt": measure("spacing", &["before"]),
            "after_pt": measure("spacing", &["after"]),
            "line": line_spacing,
            "line_rule": line_rule,
        },
        "indent": {
            "left_pt": measure("ind", &["left", "start"]),
            "right_pt": measure("ind", &["right", "end"]),
            "first_line_pt": measure("ind", &["firstLine"]),
            "hanging_pt": measure("ind", &["hanging"]),
        },
        "numbering": numbering,
        "keep_next": flag("keepNext"),
        "keep_lines": flag("keepLines"),
        "page_break_before": flag("pageBreakBefore"),
        "widow_control": flag("widowControl"),
        "outline_level": heading_level(&sources, &styles, style),
        "runs": runs(document, paragraph),
    }))
}

/// The paragraph's runs with text, and their direct character formatting.
fn runs(document: &Document, paragraph: usize) -> Vec<Value> {
    let tree: &Tree = document;
    own_nodes(tree, paragraph)
        .filter(|&index| tree.nodes[index].is("r"))
        .filter_map(|run| {
            let mut text = String::new();
            for (index, node) in tree.children(run) {
                match node.name.as_str() {
                    "t" if node.word => text.push_str(&node.text),
                    "tab" if node.word && in_run(tree, index) => text.push('\t'),
                    "br" | "cr" if node.word => text.push('\n'),
                    _ => {}
                }
            }
            if text.is_empty() {
                return None;
            }
            let sources: Vec<Source<'_>> = tree
                .child(run, "rPr")
                .map(|node| Source { tree, node })
                .into_iter()
                .collect();
            let strike = match (toggle(&sources, "strike"), toggle(&sources, "dstrike")) {
                (Some(true), _) | (_, Some(true)) => Some(true),
                (None, None) => None,
                _ => Some(false),
            };
            let font = ["ascii", "hAnsi", "eastAsia", "cs"]
                .iter()
                .find_map(|name| attribute(&sources, "rFonts", name));
            Some(json!({
                "text": text,
                "style": value(&sources, "rStyle"),
                "bold": toggle(&sources, "b"),
                "italic": toggle(&sources, "i"),
                "underline": value(&sources, "u"),
                "strike": strike,
                "font": font,
                "size_pt": value(&sources, "sz")
                    .and_then(|size| size.parse::<f64>().ok())
                    .map(|half_points| half_points / 2.0),
                "color": value(&sources, "color"),
                "highlight": value(&sources, "highlight"),
                "vertical_align": value(&sources, "vertAlign"),
            }))
        })
        .collect()
}

/// Render the body, or the blocks starting in a paragraph range, as Markdown.
fn markdown_body(
    path: &Path,
    package: &Package,
    document: &Document,
    start: Option<usize>,
    limit: Option<usize>,
) -> Result<Value> {
    let styles = Styles::load(package)?;
    let numbering = Numbering::load(package)?;
    let links = styles::link_targets(package)?;
    let footnotes = related_tree(package, rel::FOOTNOTES)?;
    let endnotes = related_tree(package, rel::ENDNOTES)?;
    let footnotes = notes(footnotes.as_ref().map(|(_, tree)| tree), "footnote");
    let endnotes = notes(endnotes.as_ref().map(|(_, tree)| tree), "endnote");
    let context = MarkdownContext {
        styles: &styles,
        numbering: &numbering,
        links: &links,
        footnotes: &footnotes,
        endnotes: &endnotes,
    };
    let start = start.unwrap_or(0);
    let limit = limit.unwrap_or(usize::MAX);
    let total = document.paragraphs.len();
    let rendered = markdown::markdown(document, &context, start, limit);
    // Paragraphs of a table that began before `start` are skipped, not pending.
    let end = rendered.end.max(start.saturating_add(limit).min(total));
    Ok(json!({
        "path": path,
        "markdown": rendered.markdown,
        "start": start,
        "end": end,
        "next_start": (end < total).then_some(end),
        "total_paragraphs": total,
    }))
}

/// Content types by lower-case part name (overrides) and extension (defaults).
fn content_type(package: &Package, name: &str) -> Result<Option<String>> {
    let Some(xml) = package.text("[Content_Types].xml")? else {
        return Ok(None);
    };
    let tree = Tree::parse(xml)?;
    let part_name = format!("/{name}");
    if let Some(kind) = tree
        .elements(ns::CONTENT_TYPES, "Override")
        .find(|&index| {
            tree.attr(index, None, "PartName")
                .is_some_and(|known| known.eq_ignore_ascii_case(&part_name))
        })
        .and_then(|index| tree.attr(index, None, "ContentType"))
    {
        return Ok(Some(kind.to_owned()));
    }
    let extension = name.rsplit_once('.').map_or("", |(_, extension)| extension);
    Ok(tree
        .elements(ns::CONTENT_TYPES, "Default")
        .find(|&index| {
            tree.attr(index, None, "Extension")
                .is_some_and(|known| known.eq_ignore_ascii_case(extension))
        })
        .and_then(|index| tree.attr(index, None, "ContentType"))
        .map(str::to_owned))
}

/// List the package parts, or return one part's text.
fn part(path: &Path, package: &Package, name: Option<&str>) -> Result<Value> {
    let Some(name) = name else {
        let parts = package
            .part_names()
            .map(|name| {
                Ok(json!({
                    "name": name,
                    "size": package.part(name).map_or(0, <[u8]>::len),
                    "content_type": content_type(package, name)?,
                }))
            })
            .collect::<Result<Vec<_>>>()?;
        return Ok(json!({"path": path, "parts": parts}));
    };
    let name = name.trim_start_matches('/');
    ensure!(!name.is_empty(), "name must not be empty");
    let bytes = package.part(name).with_context(|| {
        format!("the package has no part {name}; call part without a name to list parts")
    })?;
    let Some(text) = std::str::from_utf8(bytes)
        .ok()
        .filter(|text| !text.contains('\0'))
    else {
        bail!("{name} is a binary part; only UTF-8 text parts can be returned");
    };
    let text = text.strip_prefix('\u{feff}').unwrap_or(text);
    let mut cut = text.len().min(MAX_PART_TEXT);
    while !text.is_char_boundary(cut) {
        cut -= 1;
    }
    Ok(json!({
        "path": path,
        "name": name,
        "content_type": content_type(package, name)?,
        "size": bytes.len(),
        "text": &text[..cut],
        "truncated": cut < text.len(),
    }))
}

#[cfg(test)]
mod tests {
    use std::{fs, path::Path};

    use anyhow::{Context, Result};
    use serde_json::{Value, json};

    use super::super::{
        ooxml::{content_type, rel},
        package::{DOCUMENT_PART, Package},
        testing::{fixture, fixture_with, run},
        xml::WORD_NS,
    };

    const STYLES: &str = r#"<w:styles xmlns:w="WORD"><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240"/><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Chapter"><w:name w:val="Chapter"/><w:basedOn w:val="Heading1"/></w:style><w:style w:type="paragraph" w:styleId="berschrift2"><w:name w:val="heading 2"/></w:style><w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style></w:styles>"#;
    const NUMBERING: &str = r#"<w:numbering xmlns:w="WORD"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>"#;
    const FOOTNOTES: &str = r#"<w:footnotes xmlns:w="WORD"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Note about needle.</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
    const HEADER: &str =
        r#"<w:hdr xmlns:w="WORD"><w:p><w:r><w:t>Header needle</w:t></w:r></w:p></w:hdr>"#;
    const BODY: &str = concat!(
        r#"<q:p><q:pPr><q:pStyle q:val="Title"/></q:pPr><q:r><q:t>Report</q:t></q:r></q:p>"#,
        r#"<q:p><q:pPr><q:pStyle q:val="Heading1"/></q:pPr><q:r><q:t>Intro</q:t></q:r></q:p>"#,
        r#"<q:p><q:r><q:t xml:space="preserve">Plain </q:t></q:r><q:r><q:rPr><q:b/></q:rPr><q:t>bold</q:t></q:r><q:r><q:t xml:space="preserve"> and </q:t></q:r><q:r><q:rPr><q:b q:val="0"/><q:i/></q:rPr><q:t>italic*</q:t></q:r><q:r><q:rPr><q:rFonts q:ascii="Arial"/><q:strike/><q:color q:val="FF0000"/><q:sz q:val="28"/><q:highlight q:val="yellow"/><q:u q:val="single"/><q:vertAlign q:val="superscript"/></q:rPr><q:t>.</q:t></q:r><q:r><q:rPr><q:rStyle q:val="FootnoteReference"/></q:rPr><q:footnoteReference q:id="1"/></q:r></q:p>"#,
        r#"<q:p><q:pPr><q:pStyle q:val="ListBullet"/></q:pPr><q:r><q:t>first item</q:t></q:r></q:p>"#,
        r#"<q:p><q:pPr><q:numPr><q:ilvl q:val="1"/><q:numId q:val="1"/></q:numPr></q:pPr><q:r><q:t>nested item</q:t></q:r></q:p>"#,
        r#"<q:p><q:pPr><q:numPr><q:ilvl q:val="0"/><q:numId q:val="2"/></q:numPr></q:pPr><q:r><q:t>Step one</q:t></q:r></q:p>"#,
        r#"<q:p><q:pPr><q:numPr><q:ilvl q:val="0"/><q:numId q:val="2"/></q:numPr></q:pPr><q:r><q:t>steps two</q:t></q:r></q:p>"#,
        r#"<q:tbl><q:tr><q:tc><q:p><q:r><q:t>Name</q:t></q:r></q:p></q:tc><q:tc><q:p><q:r><q:t>Value</q:t></q:r></q:p></q:tc></q:tr><q:tr><q:tc><q:p><q:r><q:t>a|b</q:t></q:r></q:p></q:tc><q:tc><q:p><q:r><q:t>1</q:t></q:r></q:p><q:p><q:r><q:t>2</q:t></q:r></q:p></q:tc></q:tr></q:tbl>"#,
        r#"<q:p><q:r><q:t xml:space="preserve">See </q:t></q:r><q:hyperlink r:id="rId5"><q:r><q:t>the site</q:t></q:r></q:hyperlink><q:r><q:br/><q:t>next line</q:t></q:r></q:p>"#,
        r#"<q:p><q:pPr><q:pStyle q:val="Chapter"/></q:pPr><q:r><q:t>Deep</q:t></q:r></q:p>"#,
        r#"<q:p><q:pPr><q:pStyle q:val="berschrift2"/></q:pPr><q:r><q:t>Sub</q:t></q:r></q:p>"#,
        r#"<q:p><q:pPr><q:tabs><q:tab q:val="left" q:pos="720"/></q:tabs><q:jc q:val="center"/><q:ind q:left="720" q:hanging="360"/><q:outlineLvl q:val="2"/></q:pPr><q:r><q:t>Direct</q:t></q:r></q:p>"#,
        r#"<q:p><q:r><q:fldChar q:fldCharType="begin"/></q:r><q:r><q:instrText xml:space="preserve"> HYPERLINK "https://field.example" </q:instrText></q:r><q:r><q:fldChar q:fldCharType="separate"/></q:r><q:r><q:t>field link</q:t></q:r><q:r><q:fldChar q:fldCharType="end"/></q:r></q:p>"#,
    );
    const LINK: &str = "https://example.com/?a=1&b=2";

    /// A document with styles, numbering, footnotes, a header, and a hyperlink.
    fn rich(path: &Path) -> Result<()> {
        let styles = STYLES.replace("WORD", WORD_NS);
        let numbering = NUMBERING.replace("WORD", WORD_NS);
        let footnotes = FOOTNOTES.replace("WORD", WORD_NS);
        let header = HEADER.replace("WORD", WORD_NS);
        fixture_with(
            path,
            BODY,
            &[
                (
                    "word/styles.xml",
                    rel::STYLES,
                    content_type::STYLES,
                    &styles,
                ),
                (
                    "word/numbering.xml",
                    rel::NUMBERING,
                    content_type::NUMBERING,
                    &numbering,
                ),
                (
                    "word/footnotes.xml",
                    rel::FOOTNOTES,
                    content_type::FOOTNOTES,
                    &footnotes,
                ),
                (
                    "word/header1.xml",
                    rel::HEADER,
                    content_type::HEADER,
                    &header,
                ),
            ],
        )?;
        let mut package = Package::open(path)?;
        let id = package.add_relationship(DOCUMENT_PART, rel::HYPERLINK, LINK, true)?;
        assert_eq!(id, "rId5");
        fs::write(path, package.encode()?)?;
        Ok(())
    }

    fn read(path: &Path, operation: Value) -> Result<Value> {
        let mut arguments = json!({"path": path});
        arguments["operation"] = operation;
        run("docx_read", arguments)
    }

    #[test]
    fn markdown_renders_headings_emphasis_lists_tables_and_links() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("rich.docx");
        rich(&path)?;
        let before = fs::read(&path)?;
        let rendered = read(&path, json!({"action": "markdown"}))?;
        let expected = [
            "# Report",
            "",
            "# Intro",
            "",
            "Plain **bold** and *italic\\**.[^1]",
            "",
            "- first item",
            "    - nested item",
            "1. Step one",
            "2. steps two",
            "",
            "| Name | Value |",
            "| --- | --- |",
            "| a\\|b | 1<br>2 |",
            "",
            "See [the site](https://example.com/?a=1&b=2)  ",
            "next line",
            "",
            "# Deep",
            "",
            "## Sub",
            "",
            "### Direct",
            "",
            "[field link](https://field.example)",
            "",
            "[^1]: Note about needle.",
        ]
        .join("\n");
        assert_eq!(rendered["markdown"], expected);
        assert_eq!(rendered["next_start"], Value::Null);
        assert_eq!(rendered["total_paragraphs"], 17);
        // A slice: numbering continues from skipped items; tables start in range.
        let slice = read(&path, json!({"action": "markdown", "start": 6, "limit": 2}))?;
        assert_eq!(
            slice["markdown"],
            "2. steps two\n\n| Name | Value |\n| --- | --- |\n| a\\|b | 1<br>2 |"
        );
        assert_eq!(slice["end"], 12);
        assert_eq!(slice["next_start"], 12);
        // Starting inside a table skips it and still advances.
        let inside = read(&path, json!({"action": "markdown", "start": 8, "limit": 1}))?;
        assert_eq!(inside["markdown"], "");
        assert_eq!(inside["next_start"], 9);
        assert!(read(&path, json!({"action": "markdown", "limit": 0})).is_err());
        assert_eq!(fs::read(&path)?, before, "reads must not change the file");
        Ok(())
    }

    #[test]
    fn outline_uses_outline_levels_style_chains_and_names() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("outline.docx");
        rich(&path)?;
        let outline = read(&path, json!({"action": "outline"}))?;
        assert_eq!(
            outline["headings"],
            json!([
                {"index": 0, "level": 0, "text": "Report", "style": "Title"},
                {"index": 1, "level": 1, "text": "Intro", "style": "Heading1"},
                {"index": 13, "level": 1, "text": "Deep", "style": "Chapter"},
                {"index": 14, "level": 2, "text": "Sub", "style": "berschrift2"},
                {"index": 15, "level": 3, "text": "Direct", "style": "Normal"},
            ])
        );
        // Without styles.xml, Heading ids still count.
        let bare = directory.path().join("bare.docx");
        fixture(
            &bare,
            r#"<q:p><q:pPr><q:pStyle q:val="Heading2"/></q:pPr><q:r><q:t>h</q:t></q:r></q:p><q:p/>"#,
        )?;
        assert_eq!(
            read(&bare, json!({"action": "outline"}))?["headings"],
            json!([{"index": 0, "level": 2, "text": "h", "style": "Heading2"}])
        );
        Ok(())
    }

    #[test]
    fn find_reports_offsets_context_and_other_stories() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("find.docx");
        rich(&path)?;
        let found = read(
            &path,
            json!({"action": "find", "text": "item", "context_chars": 3}),
        )?;
        assert_eq!(
            found["matches"],
            json!([
                {"part": "word/document.xml", "story": "body", "paragraph": 3, "offset": 6, "length": 4, "text": "item", "before": "st ", "after": ""},
                {"part": "word/document.xml", "story": "body", "paragraph": 4, "offset": 7, "length": 4, "text": "item", "before": "ed ", "after": ""},
            ])
        );
        assert_eq!(found["truncated"], false);
        let case = read(&path, json!({"action": "find", "text": "STEP"}))?;
        assert_eq!(case["count"], 0);
        let folded = read(
            &path,
            json!({"action": "find", "text": "STEP", "match_case": false}),
        )?;
        assert_eq!(folded["count"], 2);
        let whole = read(
            &path,
            json!({"action": "find", "text": "step", "match_case": false, "whole_word": true}),
        )?;
        assert_eq!(whole["count"], 1);
        assert_eq!(whole["matches"][0]["paragraph"], 5);
        let limited = read(
            &path,
            json!({"action": "find", "text": "e", "max_results": 2}),
        )?;
        assert_eq!(limited["count"], 2);
        assert_eq!(limited["truncated"], true);
        let body = read(&path, json!({"action": "find", "text": "needle"}))?;
        assert_eq!(body["count"], 0);
        let all = read(
            &path,
            json!({"action": "find", "text": "needle", "scope": "all"}),
        )?;
        assert_eq!(
            all["matches"],
            json!([
                {"part": "word/footnotes.xml", "story": "footnote", "paragraph": 1, "offset": 12, "length": 6, "text": "needle", "before": " Note about ", "after": ".", "id": "1"},
                {"part": "word/header1.xml", "story": "header", "paragraph": 0, "offset": 7, "length": 6, "text": "needle", "before": "Header ", "after": ""},
            ])
        );
        assert!(read(&path, json!({"action": "find", "text": ""})).is_err());
        assert!(
            read(
                &path,
                json!({"action": "find", "text": "x", "max_results": 0})
            )
            .is_err()
        );
        assert!(read(&path, json!({"action": "find", "text": "x", "bogus": 1})).is_err());
        Ok(())
    }

    #[test]
    fn formatting_resolves_styles_numbering_and_runs() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("formatting.docx");
        rich(&path)?;
        let heading = read(&path, json!({"action": "formatting", "index": 1}))?;
        assert_eq!(heading["style_id"], "Heading1");
        assert_eq!(heading["style_name"], "heading 1");
        assert_eq!(heading["spacing"]["before_pt"], 12.0);
        assert_eq!(heading["spacing"]["after_pt"], 8.0);
        assert_eq!(heading["spacing"]["line"], 1.15);
        assert_eq!(heading["spacing"]["line_rule"], "multiple");
        assert_eq!(heading["keep_next"], true);
        assert_eq!(heading["keep_lines"], false);
        assert_eq!(heading["outline_level"], 1);
        assert_eq!(heading["numbering"], Value::Null);
        let direct = read(&path, json!({"action": "formatting", "index": 15}))?;
        assert_eq!(direct["text"], "Direct");
        assert_eq!(direct["alignment"], "center");
        assert_eq!(direct["indent"]["left_pt"], 36.0);
        assert_eq!(direct["indent"]["hanging_pt"], 18.0);
        assert_eq!(direct["indent"]["first_line_pt"], Value::Null);
        let numbered = read(&path, json!({"action": "formatting", "index": 3}))?;
        assert_eq!(
            numbered["numbering"],
            json!({"num_id": 1, "ilvl": 0, "format": "bullet", "level_text": "-"})
        );
        let runs = read(&path, json!({"action": "formatting", "index": 2}))?;
        let runs = &runs["runs"];
        assert_eq!(runs.as_array().context("runs")?.len(), 5);
        assert_eq!(runs[0]["text"], "Plain ");
        assert_eq!(runs[0]["bold"], Value::Null);
        assert_eq!(runs[1]["bold"], true);
        assert_eq!(runs[3]["bold"], false);
        assert_eq!(runs[3]["italic"], true);
        assert_eq!(
            runs[4],
            json!({"text": ".", "style": null, "bold": null, "italic": null, "underline": "single", "strike": true, "font": "Arial", "size_pt": 14.0, "color": "FF0000", "highlight": "yellow", "vertical_align": "superscript"})
        );
        let error = read(&path, json!({"action": "formatting", "index": 17}))
            .err()
            .context("out of range must fail")?;
        assert!(format!("{error:#}").contains("out of range"));
        Ok(())
    }

    #[test]
    fn text_and_parts() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("text.docx");
        rich(&path)?;
        let text = read(&path, json!({"action": "text"}))?;
        let expected = "Report\nIntro\nPlain bold and italic*.\nfirst item\nnested item\nStep one\nsteps two\nName\tValue\na|b\t1 2\nSee the site\nnext line\nDeep\nSub\nDirect\nfield link";
        assert_eq!(text["text"], expected);
        let parts = read(&path, json!({"action": "part"}))?;
        let listed = parts["parts"].as_array().context("parts")?;
        let styles = listed
            .iter()
            .find(|part| part["name"] == "word/styles.xml")
            .context("styles part listed")?;
        assert_eq!(styles["content_type"], content_type::STYLES);
        assert!(listed.iter().any(|part| part["name"] == "_rels/.rels"
            && part["content_type"] == content_type::RELATIONSHIPS));
        let xml = read(
            &path,
            json!({"action": "part", "name": "/word/numbering.xml"}),
        )?;
        assert_eq!(xml["text"], NUMBERING.replace("WORD", WORD_NS));
        assert_eq!(xml["truncated"], false);
        assert!(
            read(
                &path,
                json!({"action": "part", "name": "custom/preserved.bin"})
            )
            .is_err()
        );
        assert!(read(&path, json!({"action": "part", "name": "word/missing.xml"})).is_err());

        let large = directory.path().join("large.docx");
        let big = format!(
            "<w:hdr xmlns:w=\"{WORD_NS}\"><w:p><w:r><w:t>{}</w:t></w:r></w:p></w:hdr>",
            "é".repeat(600_000)
        );
        fixture_with(
            &large,
            "<q:p/>",
            &[("word/header1.xml", rel::HEADER, content_type::HEADER, &big)],
        )?;
        let truncated = read(
            &large,
            json!({"action": "part", "name": "word/header1.xml"}),
        )?;
        assert_eq!(truncated["truncated"], true);
        assert_eq!(truncated["size"], big.len());
        let returned = truncated["text"].as_str().context("text")?;
        assert!(returned.len() <= super::MAX_PART_TEXT && returned.len() > 1_000_000);
        Ok(())
    }

    #[test]
    fn unicode_entities_controls_and_preview_roundtrip() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("unicode.docx");
        let text = "é 🍎 <script>& \"quotes\"\nline\ttab";
        run(
            "docx_document",
            json!({"path": path, "operation": {"action": "create", "paragraphs": [text]}}),
        )?;
        let read = run(
            "docx_read",
            json!({"path": path, "operation": {"action": "paragraphs"}}),
        )?;
        assert_eq!(read["text"], text);
        let html = run(
            "docx_read",
            json!({"path": path, "operation": {"action": "preview"}}),
        )?;
        assert!(
            html["html"]
                .as_str()
                .context("missing HTML")?
                .contains("&lt;script&gt;&amp;")
        );
        Ok(())
    }

    #[test]
    fn paragraphs_report_style_and_table_membership() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("styles.docx");
        fixture(
            &path,
            "<q:p><q:pPr><q:pStyle q:val=\"Heading1\"/></q:pPr><q:r><q:t>title</q:t></q:r></q:p><q:tbl><q:tr><q:tc><q:p><q:pPr><q:pStyle q:val=\"TableText\"/></q:pPr><q:r><q:t>cell</q:t></q:r></q:p></q:tc></q:tr></q:tbl><q:p/>",
        )?;
        let read = run(
            "docx_read",
            json!({"path": path, "operation": {"action": "paragraphs", "start": 0}}),
        )?;
        assert_eq!(
            read["paragraphs"],
            json!([
                {"index": 0, "text": "title", "style": "Heading1", "in_table": false},
                {"index": 1, "text": "cell", "style": "TableText", "in_table": true},
                {"index": 2, "text": "", "style": null, "in_table": false},
            ])
        );
        Ok(())
    }
}
