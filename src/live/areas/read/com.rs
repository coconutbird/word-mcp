//! Word automation behind `word_live_read`.

use anyhow::{Context, Result, ensure};
use serde_json::{Value, json};

use super::Target;
use crate::live::word::{IDispatch, check_range, item, paragraph_range, range};

/// `wdUndefined`: the value of a property that differs across a range.
const UNDEFINED: i32 = 9_999_999;
/// `wdStatisticPages`.
const STATISTIC_PAGES: i32 = 2;
/// `wdGoToPage` and `wdGoToAbsolute`.
const GO_TO_PAGE: i32 = 1;
const GO_TO_ABSOLUTE: i32 = 1;
/// `wdFindStop`.
const FIND_STOP: i32 = 0;
/// `wdColorAutomatic`.
const COLOR_AUTOMATIC: i32 = -16_777_216;
/// `wdMainTextStory`.
const MAIN_STORY: i32 = 1;
/// `wdListNoNumbering`.
const NO_LIST: i32 = 0;
/// Word measures "multiple" line spacing in points per 12-point line.
const LINE_POINTS: f64 = 12.0;

pub(super) fn text(document: &IDispatch, start: Option<i32>, end: Option<i32>) -> Result<Value> {
    let content = document.object("Content")?;
    let start = start.map_or_else(|| content.int("Start"), Ok)?;
    let end = end.map_or_else(|| content.int("End"), Ok)?;
    check_range(document, start, end)?;
    Ok(json!({
        "path": document.string("FullName")?,
        "text": range(document, start, end)?.string("Text")?,
        "start": start,
        "end": end,
        "saved": document.flag("Saved")?,
    }))
}

pub(super) fn paragraphs(document: &IDispatch, start: u32, limit: u32) -> Result<Value> {
    let paragraphs = document.object("Paragraphs")?;
    let total = paragraphs.int("Count")?;
    let first = i32::try_from(start).context("start is too large")?;
    let mut listed = Vec::new();
    let mut next = if first < total {
        Some(item(&paragraphs, first + 1)?)
    } else {
        None
    };
    let mut index = first;
    while let Some(paragraph) = next {
        if listed.len() == limit as usize {
            break;
        }
        let range = paragraph.object("Range")?;
        let text = range.string("Text")?;
        listed.push(json!({
            "index": index,
            "start": range.int("Start")?,
            "end": range.int("End")?,
            "style": style_name(&paragraph)?,
            "text": paragraph_text(&text),
        }));
        next = paragraph.call("Next", vec![])?.into_object()?;
        index += 1;
    }
    Ok(json!({
        "path": document.string("FullName")?,
        "total_paragraphs": total,
        "start": start,
        "paragraphs": listed,
    }))
}

/// The display name of a paragraph's (or range's) style.
pub(in crate::live) fn style_name(target: &IDispatch) -> Result<Option<String>> {
    let style = target.get("Style")?;
    if let Ok(name) = style.string() {
        return Ok(Some(name));
    }
    style
        .into_object()?
        .map(|style| style.string("NameLocal"))
        .transpose()
}

/// A paragraph's text without its paragraph or end-of-cell mark.
fn paragraph_text(text: &str) -> &str {
    text.strip_suffix("\r\u{7}")
        .or_else(|| text.strip_suffix('\r'))
        .unwrap_or(text)
}

/// Visit every paragraph of the main story with its zero-based index.
fn each_paragraph(
    document: &IDispatch,
    mut visit: impl FnMut(i32, &IDispatch) -> Result<()>,
) -> Result<()> {
    let mut next = document.object("Paragraphs")?.get("First")?.into_object()?;
    let mut index = 0;
    while let Some(paragraph) = next {
        visit(index, &paragraph)?;
        next = paragraph.call("Next", vec![])?.into_object()?;
        index += 1;
    }
    Ok(())
}

/// The text of every paragraph of the main story.
pub(super) fn paragraph_texts(document: &IDispatch) -> Result<Vec<String>> {
    let mut texts = Vec::new();
    each_paragraph(document, |_, paragraph| {
        let text = paragraph.object("Range")?.string("Text")?;
        texts.push(paragraph_text(&text).to_owned());
        Ok(())
    })?;
    Ok(texts)
}

/// The zero-based index of the main-story paragraph containing `position`.
fn paragraph_index(document: &IDispatch, position: i32) -> Result<i32> {
    let first = range(document, position, position)?
        .object("Paragraphs")?
        .object("First")?
        .object("Range")?
        .int("Start")?;
    if first <= 0 {
        return Ok(0);
    }
    range(document, 0, first)?
        .object("Paragraphs")?
        .int("Count")
}

/// A validated `find` request.
pub(super) struct Search {
    /// Find text, already escaped unless `wildcards`.
    pub(super) text: String,
    pub(super) match_case: bool,
    pub(super) whole_word: bool,
    pub(super) wildcards: bool,
    pub(super) max_results: u32,
    pub(super) context_chars: i32,
}

pub(super) fn find(document: &IDispatch, search: &Search) -> Result<Value> {
    let content = document.object("Content")?;
    let story_start = content.int("Start")?;
    let story_end = content.int("End")?;
    let target = range(document, story_start, story_end)?;
    let find = target.object("Find")?;
    find.call("ClearFormatting", vec![])?;
    find.put("Text", search.text.as_str().into())?;
    find.put("Forward", true.into())?;
    find.put("Wrap", FIND_STOP.into())?;
    find.put("Format", false.into())?;
    find.put("MatchCase", search.match_case.into())?;
    find.put("MatchWholeWord", search.whole_word.into())?;
    find.put("MatchWildcards", search.wildcards.into())?;
    find.put("MatchSoundsLike", false.into())?;
    find.put("MatchAllWordForms", false.into())?;
    let mut matches = Vec::new();
    let mut truncated = false;
    let mut previous = None;
    while find.call("Execute", vec![])?.flag()? {
        let start = target.int("Start")?;
        let end = target.int("End")?;
        // Word can re-find the same span (for example at a cell end); stop rather
        // than loop.
        if previous.is_some_and(|previous| previous >= (start, end)) || end > story_end {
            break;
        }
        previous = Some((start, end));
        if matches.len() == search.max_results as usize {
            truncated = true;
            break;
        }
        let before = (start - search.context_chars).max(story_start);
        let after = end.saturating_add(search.context_chars).min(story_end);
        matches.push(json!({
            "start": start,
            "end": end,
            "text": target.string("Text")?,
            "before": range(document, before, start)?.string("Text")?,
            "after": range(document, end, after)?.string("Text")?,
            "paragraph": paragraph_index(document, start)?,
        }));
    }
    Ok(json!({
        "path": document.string("FullName")?,
        "count": matches.len(),
        "truncated": truncated,
        "matches": matches,
    }))
}

pub(super) fn outline(document: &IDispatch) -> Result<Value> {
    let mut headings = Vec::new();
    let mut total = 0;
    each_paragraph(document, |index, paragraph| {
        total = index + 1;
        let level = paragraph.int("OutlineLevel")?;
        if (1..=9).contains(&level) {
            let range = paragraph.object("Range")?;
            let text = range.string("Text")?;
            headings.push(json!({
                "index": index,
                "level": level,
                "start": range.int("Start")?,
                "end": range.int("End")?,
                "style": style_name(paragraph)?,
                "text": paragraph_text(&text),
            }));
        }
        Ok(())
    })?;
    Ok(json!({
        "path": document.string("FullName")?,
        "headings": headings,
        "total_paragraphs": total,
    }))
}

pub(super) fn page_text(document: &IDispatch, first: u32, last: u32) -> Result<Value> {
    let total = document
        .call("ComputeStatistics", vec![STATISTIC_PAGES.into()])?
        .int()
        .context("Word page count")?;
    let first = i32::try_from(first)?;
    ensure!(
        first <= total,
        "page {first} does not exist; the document has {total} pages"
    );
    let last = i32::try_from(last)?.min(total);
    let page_start = |page: i32| -> Result<i32> {
        document
            .call(
                "GoTo",
                vec![GO_TO_PAGE.into(), GO_TO_ABSOLUTE.into(), page.into()],
            )?
            .into_object()?
            .context("Word returned no page range")?
            .int("Start")
    };
    let content_end = document.object("Content")?.int("End")?;
    let mut start = page_start(first)?;
    let mut pages = Vec::new();
    for page in first..=last {
        let end = if page == total {
            content_end
        } else {
            page_start(page + 1)?
        };
        pages.push(json!({
            "page": page,
            "start": start,
            "end": end,
            "text": range(document, start, end)?.string("Text")?,
        }));
        start = end;
    }
    Ok(json!({
        "path": document.string("FullName")?,
        "total_pages": total,
        "pages": pages,
    }))
}

/// An enumerated property, or `None` when it differs across the range.
fn defined(object: &IDispatch, name: &str) -> Result<Option<i32>> {
    let value = object.int(name)?;
    Ok((value != UNDEFINED).then_some(value))
}

/// A measurement in points, or `None` when it differs across the range.
fn measure(object: &IDispatch, name: &str) -> Result<Option<f64>> {
    let value = object
        .get(name)?
        .number()
        .with_context(|| format!("Word {name}"))?;
    Ok((value < f64::from(UNDEFINED) - 0.5).then_some(value))
}

/// A Word tri-state flag (`True`, `False`, or `wdUndefined`).
fn tristate(object: &IDispatch, name: &str) -> Result<Option<bool>> {
    Ok(defined(object, name)?.map(|value| value != 0))
}

/// `WdColor` (BGR) as `RRGGBB`, `auto`, or `None` when mixed or theme-only.
fn color(font: &IDispatch) -> Result<Option<String>> {
    let hex = |value: i32| {
        format!(
            "{:02X}{:02X}{:02X}",
            value & 0xFF,
            (value >> 8) & 0xFF,
            (value >> 16) & 0xFF
        )
    };
    Ok(match font.int("Color")? {
        UNDEFINED => None,
        COLOR_AUTOMATIC => Some("auto".to_owned()),
        value @ 0..=0x00FF_FFFF => Some(hex(value)),
        // Theme colors: ask for the resolved RGB value.
        _ => match font.object("TextColor")?.int("RGB")? {
            value @ 0..=0x00FF_FFFF => Some(hex(value)),
            _ => None,
        },
    })
}

fn alignment_name(value: i32) -> &'static str {
    match value {
        0 => "left",
        1 => "center",
        2 => "right",
        3 => "justify",
        4 => "distribute",
        _ => "other",
    }
}

fn underline_name(value: i32) -> &'static str {
    match value {
        0 => "none",
        1 => "single",
        2 => "words",
        3 => "double",
        4 => "dotted",
        6 => "thick",
        7 => "dash",
        9 => "dot_dash",
        10 => "dot_dot_dash",
        11 => "wavy",
        _ => "other",
    }
}

fn highlight_name(value: i32) -> &'static str {
    match value {
        0 => "none",
        1 => "black",
        2 => "blue",
        3 => "turquoise",
        4 => "bright_green",
        5 => "pink",
        6 => "red",
        7 => "yellow",
        8 => "white",
        9 => "dark_blue",
        10 => "teal",
        11 => "green",
        12 => "violet",
        13 => "dark_red",
        14 => "dark_yellow",
        15 => "gray_50",
        16 => "gray_25",
        _ => "other",
    }
}

/// Line spacing as docx reports it: lines for single/1.5/double/multiple, points
/// for exact and at-least.
fn line_spacing(format: &IDispatch) -> Result<(Option<f64>, Option<&'static str>)> {
    let Some(rule) = defined(format, "LineSpacingRule")? else {
        return Ok((None, None));
    };
    let Some(spacing) = measure(format, "LineSpacing")? else {
        return Ok((None, None));
    };
    Ok(match rule {
        3 => (Some(spacing), Some("at_least")),
        4 => (Some(spacing), Some("exact")),
        _ => (Some(spacing / LINE_POINTS), Some("multiple")),
    })
}

pub(super) fn formatting(document: &IDispatch, target: &Target) -> Result<Value> {
    let (target, paragraph) = match *target {
        Target::Range { start, end } => {
            check_range(document, start, end)?;
            (range(document, start, end)?, None)
        }
        Target::Paragraph(index) => (paragraph_range(document, index as usize)?, Some(index)),
    };
    let start = target.int("Start")?;
    let end = target.int("End")?;
    let format = target.object("ParagraphFormat")?;
    let font = target.object("Font")?;
    let (line, line_rule) = line_spacing(&format)?;
    let first_line = measure(&format, "FirstLineIndent")?;
    let list = target.object("ListFormat")?;
    let numbering = match defined(&list, "ListType")? {
        Some(NO_LIST) => None,
        _ => Some(json!({
            "list_string": list.string("ListString")?,
            "ilvl": defined(&list, "ListLevelNumber")?.map(|level| level - 1),
        })),
    };
    let vertical_align = match (
        tristate(&font, "Superscript")?,
        tristate(&font, "Subscript")?,
    ) {
        (Some(true), _) => Some("superscript"),
        (_, Some(true)) => Some("subscript"),
        (Some(false), Some(false)) => Some("baseline"),
        _ => None,
    };
    let name = font.string("Name")?;
    Ok(json!({
        "path": document.string("FullName")?,
        "start": start,
        "end": end,
        "paragraph": paragraph,
        "style_name": style_name(&format)?,
        "alignment": defined(&format, "Alignment")?.map(alignment_name),
        "spacing": {
            "before_pt": measure(&format, "SpaceBefore")?,
            "after_pt": measure(&format, "SpaceAfter")?,
            "line": line,
            "line_rule": line_rule,
        },
        "indent": {
            "left_pt": measure(&format, "LeftIndent")?,
            "right_pt": measure(&format, "RightIndent")?,
            "first_line_pt": first_line.filter(|indent| *indent >= 0.0),
            "hanging_pt": first_line.filter(|indent| *indent < 0.0).map(|indent| -indent),
        },
        "keep_next": tristate(&format, "KeepWithNext")?,
        "keep_lines": tristate(&format, "KeepTogether")?,
        "page_break_before": tristate(&format, "PageBreakBefore")?,
        "widow_control": tristate(&format, "WidowControl")?,
        "outline_level": defined(&format, "OutlineLevel")?,
        "numbering": numbering,
        "font": {
            "name": (!name.is_empty()).then_some(name),
            "size_pt": measure(&font, "Size")?,
            "bold": tristate(&font, "Bold")?,
            "italic": tristate(&font, "Italic")?,
            "underline": defined(&font, "Underline")?.map(underline_name),
            "strike": tristate(&font, "StrikeThrough")?,
            "color": color(&font)?,
            "highlight": defined(&target, "HighlightColorIndex")?.map(highlight_name),
            "vertical_align": vertical_align,
        },
    }))
}

fn story_name(story: i32) -> &'static str {
    match story {
        MAIN_STORY => "main",
        2 => "footnotes",
        3 => "endnotes",
        4 => "comments",
        5 => "text_frame",
        6..=11 => "header_footer",
        _ => "other",
    }
}

pub(super) fn selection(word: &IDispatch, document: &IDispatch) -> Result<Value> {
    let path = document.string("FullName")?;
    let active = match word.get("ActiveDocument") {
        Ok(active) => match active.into_object()? {
            Some(active) => active.string("FullName")? == path,
            None => false,
        },
        // No document window is active.
        Err(_) => false,
    };
    if !active {
        return Ok(json!({"path": path, "active": false, "selection": null}));
    }
    let selection = document.object("ActiveWindow")?.object("Selection")?;
    let start = selection.int("Start")?;
    let end = selection.int("End")?;
    let story = selection.int("StoryType")?;
    let paragraph = if story == MAIN_STORY {
        Some(paragraph_index(document, start)?)
    } else {
        None
    };
    Ok(json!({
        "path": path,
        "active": true,
        "selection": {
            "start": start,
            "end": end,
            // A collapsed selection's Text is the following character; report none.
            "text": if start == end { String::new() } else { selection.string("Text")? },
            "collapsed": start == end,
            "story": story_name(story),
            "paragraph": paragraph,
        },
    }))
}
