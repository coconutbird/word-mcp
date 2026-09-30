//! Body blocks in reading order, rendered as plain text or Markdown.

use std::{collections::HashMap, fmt::Write as _};

use super::super::xml::{Document, Tree, WORD_NS};
use super::styles::{
    Numbering, Source, Styles, direct_style, enclosing, field_target, heading_level,
    hyperlink_target, in_run, numbering_reference, own_nodes, paragraph_sources, paragraph_text,
    toggle,
};

/// A top-level body block.
#[derive(Clone, Copy)]
pub(super) enum Block {
    Paragraph(usize),
    Table(usize),
}

/// The body's paragraphs and tables in reading order, looking through content
/// controls and custom XML wrappers.
pub(super) fn blocks(document: &Document) -> Vec<Block> {
    let mut blocks = Vec::new();
    collect(document, document.body, &mut blocks);
    blocks
}

fn collect(tree: &Tree, parent: usize, blocks: &mut Vec<Block>) {
    for (index, node) in tree.children(parent) {
        if !node.word {
            continue;
        }
        match node.name.as_str() {
            "p" => blocks.push(Block::Paragraph(index)),
            "tbl" => blocks.push(Block::Table(index)),
            "sdt" => {
                if let Some(content) = tree.child(index, "sdtContent") {
                    collect(tree, content, blocks);
                }
            }
            "customXml" => collect(tree, index, blocks),
            _ => {}
        }
    }
}

/// Zero-based paragraph numbers `first..end` covered by a block.
pub(super) fn paragraph_span(document: &Document, block: Block) -> Option<(usize, usize)> {
    match block {
        Block::Paragraph(node) => document
            .paragraph_number(node)
            .map(|number| (number, number + 1)),
        Block::Table(node) => {
            let mut numbers = document
                .descendants(node)
                .filter(|(_, child)| child.is("p"))
                .filter_map(|(index, _)| document.paragraph_number(index));
            let first = numbers.next()?;
            let last = numbers.last().unwrap_or(first);
            Some((first, last + 1))
        }
    }
}

/// Rows of a table, each a list of cells, each the cell's paragraph elements.
fn table_rows(tree: &Tree, table: usize) -> Vec<Vec<(usize, Vec<usize>)>> {
    tree.children(table)
        .filter(|(_, node)| node.is("tr"))
        .map(|(row, _)| {
            tree.children(row)
                .filter(|(_, node)| node.is("tc"))
                .map(|(cell, _)| {
                    let paragraphs = tree
                        .descendants(cell)
                        .filter(|(_, node)| node.is("p"))
                        .map(|(index, _)| index)
                        .collect();
                    (cell, paragraphs)
                })
                .collect()
        })
        .collect()
}

/// Plain text: one line per paragraph, one line per table row with tab-separated
/// cells (a cell's paragraphs joined by spaces).
pub(super) fn plain_text(document: &Document) -> String {
    let mut lines = Vec::new();
    for block in blocks(document) {
        match block {
            Block::Paragraph(node) => lines.push(paragraph_text(document, node)),
            Block::Table(table) => {
                for row in table_rows(document, table) {
                    let cells: Vec<String> = row
                        .iter()
                        .map(|(_, paragraphs)| {
                            paragraphs
                                .iter()
                                .map(|paragraph| paragraph_text(document, *paragraph))
                                .filter(|text| !text.is_empty())
                                .collect::<Vec<_>>()
                                .join(" ")
                        })
                        .collect();
                    lines.push(cells.join("\t"));
                }
            }
        }
    }
    lines.join("\n")
}

/// Note texts by id from a footnotes or endnotes part.
pub(super) fn notes(tree: Option<&Tree>, element: &str) -> HashMap<String, String> {
    let Some(tree) = tree else {
        return HashMap::new();
    };
    tree.elements(WORD_NS, element)
        .filter_map(|note| {
            let id = tree.word_attr(note, "id")?.to_owned();
            let text = tree
                .descendants(note)
                .filter(|(_, node)| node.is("p"))
                .map(|(paragraph, _)| paragraph_text(tree, paragraph))
                .filter(|text| !text.trim().is_empty())
                .collect::<Vec<_>>()
                .join(" ");
            Some((id, text.trim().to_owned()))
        })
        .collect()
}

/// Everything Markdown rendering needs besides the document.
pub(super) struct Context<'a> {
    pub(super) styles: &'a Styles,
    pub(super) numbering: &'a Numbering,
    pub(super) links: &'a HashMap<String, String>,
    pub(super) footnotes: &'a HashMap<String, String>,
    pub(super) endnotes: &'a HashMap<String, String>,
}

/// A rendered Markdown slice.
pub(super) struct Rendered {
    pub(super) markdown: String,
    /// Exclusive paragraph number reached.
    pub(super) end: usize,
}

/// A contiguous piece of paragraph content with uniform formatting.
struct Piece {
    text: String,
    bold: bool,
    italic: bool,
    link: Option<String>,
    /// Already Markdown (footnote references); never escaped or emphasized.
    raw: bool,
}

/// Render the blocks whose first paragraph lies in `start..start + limit`.
pub(super) fn markdown(
    document: &Document,
    context: &Context<'_>,
    start: usize,
    limit: usize,
) -> Rendered {
    let stop = start.saturating_add(limit);
    let mut rendered: Vec<(String, bool)> = Vec::new();
    let mut counters: HashMap<String, [Option<u32>; 9]> = HashMap::new();
    let mut references = Vec::new();
    let mut end = start;
    for block in blocks(document) {
        let Some((first, last)) = paragraph_span(document, block) else {
            continue;
        };
        if first < start || first >= stop {
            // Lists number across the whole document, so count skipped items too.
            if let Block::Paragraph(node) = block {
                list_marker(document, node, context, &mut counters);
            }
            continue;
        }
        end = end.max(last);
        match block {
            Block::Paragraph(node) => {
                if let Some(entry) =
                    paragraph_markdown(document, node, context, &mut counters, &mut references)
                {
                    rendered.push(entry);
                }
            }
            Block::Table(node) => {
                if let Some(table) = table_markdown(document, node, context, &mut references) {
                    rendered.push((table, false));
                }
            }
        }
    }
    let mut markdown = String::new();
    let mut previous_list = false;
    for (index, (block, list)) in rendered.iter().enumerate() {
        if index > 0 {
            markdown.push_str(if previous_list && *list { "\n" } else { "\n\n" });
        }
        markdown.push_str(block);
        previous_list = *list;
    }
    let mut definitions = String::new();
    for (label, text) in references {
        let _ = write!(definitions, "\n[^{label}]: {}", escape(&text, false));
    }
    if !definitions.is_empty() {
        markdown.push('\n');
        markdown.push_str(&definitions);
    }
    Rendered { markdown, end }
}

/// A paragraph's list marker and indentation, advancing the list counters.
fn list_marker(
    document: &Document,
    paragraph: usize,
    context: &Context<'_>,
    counters: &mut HashMap<String, [Option<u32>; 9]>,
) -> Option<String> {
    let style = context
        .styles
        .paragraph_style(direct_style(document, paragraph));
    let sources = paragraph_sources(document, paragraph, context.styles, style);
    let (id, level) = numbering_reference(&sources)?;
    let level = level.min(8);
    let definition = context.numbering.level(&id, level);
    let format = definition
        .as_ref()
        .map_or("decimal", |definition| definition.format.as_str());
    if format == "none" {
        return None;
    }
    let indent = "    ".repeat(usize::from(level));
    if format == "bullet" {
        return Some(format!("{indent}- "));
    }
    let slots = counters.entry(id).or_default();
    let slot = usize::from(level);
    let number = slots[slot].map_or_else(
        || definition.as_ref().map_or(1, |definition| definition.start),
        |previous| previous + 1,
    );
    slots[slot] = Some(number);
    for deeper in &mut slots[slot + 1..] {
        *deeper = None;
    }
    Some(format!("{indent}{number}. "))
}

fn paragraph_markdown(
    document: &Document,
    paragraph: usize,
    context: &Context<'_>,
    counters: &mut HashMap<String, [Option<u32>; 9]>,
    references: &mut Vec<(String, String)>,
) -> Option<(String, bool)> {
    let marker = list_marker(document, paragraph, context, counters);
    let rendered = inline(document, paragraph, context, references, false);
    let body = rendered.trim_end();
    if body.trim().is_empty() {
        return None;
    }
    let style = context
        .styles
        .paragraph_style(direct_style(document, paragraph));
    let sources = paragraph_sources(document, paragraph, context.styles, style);
    if let Some(level) = heading_level(&sources, context.styles, style) {
        let hashes = "#".repeat(usize::from(level.clamp(1, 6)));
        let text = body.trim().replace("  \n", " ");
        return Some((format!("{hashes} {text}"), false));
    }
    match marker {
        Some(marker) => Some((format!("{marker}{}", body.trim_start()), true)),
        None => Some((body.to_owned(), false)),
    }
}

fn table_markdown(
    document: &Document,
    table: usize,
    context: &Context<'_>,
    references: &mut Vec<(String, String)>,
) -> Option<String> {
    let mut rows: Vec<Vec<String>> = Vec::new();
    for row in table_rows(document, table) {
        let mut cells = Vec::new();
        for (cell, paragraphs) in row {
            let text = paragraphs
                .iter()
                .map(|paragraph| inline(document, *paragraph, context, references, true))
                .map(|text| text.trim().to_owned())
                .filter(|text| !text.is_empty())
                .collect::<Vec<_>>()
                .join("<br>");
            cells.push(text);
            let span = document
                .child(cell, "tcPr")
                .and_then(|properties| document.child(properties, "gridSpan"))
                .and_then(|span| document.nodes[span].val.as_deref()?.parse::<usize>().ok())
                .unwrap_or(1);
            cells.extend(std::iter::repeat_n(String::new(), span.saturating_sub(1)));
        }
        rows.push(cells);
    }
    let columns = rows.iter().map(Vec::len).max()?;
    if columns == 0 {
        return None;
    }
    let line = |cells: &[String]| {
        let mut line = String::from("|");
        for column in 0..columns {
            let _ = write!(line, " {} |", cells.get(column).map_or("", String::as_str));
        }
        line
    };
    let mut lines = vec![line(&rows[0]), format!("|{}", " --- |".repeat(columns))];
    lines.extend(rows[1..].iter().map(|row| line(row)));
    Some(lines.join("\n"))
}

/// A field being read: its instruction so far and, once separated, its link.
struct Field {
    instruction: String,
    separated: bool,
    link: Option<String>,
}

/// A paragraph's inline Markdown: emphasis, links, breaks, and note references.
fn inline(
    document: &Document,
    paragraph: usize,
    context: &Context<'_>,
    references: &mut Vec<(String, String)>,
    in_table: bool,
) -> String {
    let tree: &Tree = document;
    let mut pieces: Vec<Piece> = Vec::new();
    let mut fields: Vec<Field> = Vec::new();
    for index in own_nodes(tree, paragraph) {
        let node = &tree.nodes[index];
        match node.name.as_str() {
            "fldChar" => match tree.word_attr(index, "fldCharType") {
                Some("begin") => fields.push(Field {
                    instruction: String::new(),
                    separated: false,
                    link: None,
                }),
                Some("separate") => {
                    if let Some(field) = fields.last_mut() {
                        field.separated = true;
                        field.link = field_target(&field.instruction);
                    }
                }
                Some("end") => {
                    fields.pop();
                }
                _ => {}
            },
            "instrText" => {
                if let Some(field) = fields.last_mut().filter(|field| !field.separated) {
                    field.instruction.push_str(&node.text);
                }
            }
            "t" | "tab" | "br" | "cr" | "footnoteReference" | "endnoteReference" => {
                if !matches!(node.name.as_str(), "t") && !in_run(tree, index) {
                    continue;
                }
                if fields.iter().any(|field| !field.separated) {
                    continue;
                }
                let (text, raw) = match node.name.as_str() {
                    "t" => (node.text.clone(), false),
                    "tab" => ("\t".to_owned(), false),
                    "br" => {
                        if matches!(tree.word_attr(index, "type"), Some("page" | "column")) {
                            continue;
                        }
                        ("\n".to_owned(), false)
                    }
                    "cr" => ("\n".to_owned(), false),
                    name => {
                        let Some(id) = tree.word_attr(index, "id") else {
                            continue;
                        };
                        let (label, notes) = if name == "footnoteReference" {
                            (id.to_owned(), context.footnotes)
                        } else {
                            (format!("en{id}"), context.endnotes)
                        };
                        if let Some(text) = notes.get(id)
                            && !references.iter().any(|(known, _)| *known == label)
                        {
                            references.push((label.clone(), text.clone()));
                        }
                        (format!("[^{label}]"), true)
                    }
                };
                let run = enclosing(tree, index, "r", paragraph);
                let properties: Vec<Source<'_>> = run
                    .and_then(|run| tree.child(run, "rPr"))
                    .map(|node| Source { tree, node })
                    .into_iter()
                    .collect();
                let link = link_of(tree, index, paragraph, context)
                    .or_else(|| fields.iter().rev().find_map(|field| field.link.clone()));
                pieces.push(Piece {
                    text,
                    bold: toggle(&properties, "b").unwrap_or(false),
                    italic: toggle(&properties, "i").unwrap_or(false),
                    link,
                    raw,
                });
            }
            _ => {}
        }
    }
    let rendered = render(&pieces, in_table);
    if in_table {
        rendered.replace('\n', "<br>")
    } else {
        rendered.replace('\n', "  \n")
    }
}

/// The hyperlink (element or simple field) enclosing node `index`.
fn link_of(tree: &Tree, index: usize, paragraph: usize, context: &Context<'_>) -> Option<String> {
    let mut current = tree.nodes[index].parent?;
    while current != paragraph {
        let node = &tree.nodes[current];
        if node.is("hyperlink") {
            return hyperlink_target(tree, current, context.links);
        }
        if node.is("fldSimple")
            && let Some(target) = tree.word_attr(current, "instr").and_then(field_target)
        {
            return Some(target);
        }
        current = node.parent?;
    }
    None
}

/// Join pieces, merging equal neighbours, wrapping links and emphasis.
fn render(pieces: &[Piece], in_table: bool) -> String {
    let mut output = String::new();
    let mut index = 0;
    while index < pieces.len() {
        let link = &pieces[index].link;
        let group_end = pieces[index..]
            .iter()
            .position(|piece| piece.link != *link)
            .map_or(pieces.len(), |offset| index + offset);
        let mut inner = String::new();
        let mut run = index;
        while run < group_end {
            let first = &pieces[run];
            let same = pieces[run..group_end]
                .iter()
                .position(|piece| {
                    piece.raw != first.raw
                        || piece.raw
                        || piece.bold != first.bold
                        || piece.italic != first.italic
                })
                .map_or(group_end, |offset| run + offset.max(1));
            if first.raw {
                inner.push_str(&first.text);
            } else {
                let text: String = pieces[run..same]
                    .iter()
                    .map(|piece| piece.text.as_str())
                    .collect();
                inner.push_str(&emphasize(
                    &escape(&text, in_table),
                    first.bold,
                    first.italic,
                ));
            }
            run = same;
        }
        match link {
            Some(target) if !inner.trim().is_empty() => {
                let target = if target.contains([' ', '(', ')', '<', '>']) {
                    format!("<{}>", target.replace('<', "%3C").replace('>', "%3E"))
                } else {
                    target.clone()
                };
                let _ = write!(output, "[{inner}]({target})");
            }
            _ => output.push_str(&inner),
        }
        index = group_end;
    }
    output
}

/// Wrap the non-blank core of `text` in emphasis markers, keeping surrounding
/// whitespace outside them as Markdown requires.
fn emphasize(text: &str, bold: bool, italic: bool) -> String {
    let marker = match (bold, italic) {
        (true, true) => "***",
        (true, false) => "**",
        (false, true) => "*",
        (false, false) => return text.to_owned(),
    };
    let core = text.trim();
    if core.is_empty() {
        return text.to_owned();
    }
    let leading = &text[..text.len() - text.trim_start().len()];
    let trailing = &text[text.trim_end().len()..];
    format!("{leading}{marker}{core}{marker}{trailing}")
}

/// Escape Markdown punctuation that would otherwise change the meaning of text.
fn escape(text: &str, in_table: bool) -> String {
    let mut escaped = String::with_capacity(text.len());
    for character in text.chars() {
        if matches!(character, '\\' | '*' | '_' | '`' | '[' | ']') || (in_table && character == '|')
        {
            escaped.push('\\');
        }
        escaped.push(character);
    }
    escaped
}
