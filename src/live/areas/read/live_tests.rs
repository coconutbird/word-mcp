//! Live `word_live_read` tests against a private hidden Word instance.

use anyhow::{Context, Result};
use serde_json::{Value, json};

use crate::{
    live::word::{
        IDispatch, Session, Variant, paragraph_range, range,
        testing::{text, with_word},
    },
    tool::Output,
};

/// `wdStyleHeading1` and `wdStyleHeading2`, independent of Word's UI language.
const HEADING_1: i32 = -2;
const HEADING_2: i32 = -3;

fn read(session: &mut Session, path: &str, operation: Value) -> Result<Value> {
    let mut arguments = json!({"path": path});
    arguments["operation"] = operation;
    session.json("word_live_read", arguments)
}

fn set_style(document: &IDispatch, paragraph: usize, style: i32) -> Result<()> {
    paragraph_range(document, paragraph)?.put("Style", style.into())
}

#[test]
#[ignore = "Requires desktop Microsoft Word; starts a private hidden instance"]
fn reads_text_and_paragraph_positions() {
    with_word(|harness| {
        let (path, _document) = harness.document("read.docx", "alpha 🦀 beta\rsecond paragraph")?;
        let session = &mut harness.session;
        let whole = read(session, &path, json!({"action": "text"}))?;
        assert!(text(&whole).contains("alpha 🦀 beta"));
        let listed = read(session, &path, json!({"action": "paragraphs"}))?;
        assert_eq!(listed["total_paragraphs"], 2);
        assert_eq!(listed["paragraphs"][1]["text"], "second paragraph");
        let second = read(
            session,
            &path,
            json!({"action": "paragraphs", "start": 1, "limit": 1}),
        )?;
        assert_eq!(second["paragraphs"][0]["index"], 1);
        assert!(
            read(
                session,
                &path,
                json!({"action": "text", "start": 5, "end": 2})
            )
            .is_err()
        );
        Ok(())
    });
}

#[test]
#[ignore = "Requires desktop Microsoft Word; starts a private hidden instance"]
fn finds_outlines_and_reads_pages() {
    with_word(|harness| {
        // "Heading\r" is 8 units and "body alpha beta\r" 16, so paragraph 2 starts at 24.
        let (path, document) = harness.document(
            "find.docx",
            "Heading\rbody alpha beta\ralpha again\rSub heading",
        )?;
        set_style(&document, 0, HEADING_1)?;
        set_style(&document, 3, HEADING_2)?;
        let session = &mut harness.session;

        let found = read(
            session,
            &path,
            json!({"action": "find", "text": "alpha", "context_chars": 5}),
        )?;
        assert_eq!(
            found["matches"],
            json!([
                {"start": 13, "end": 18, "text": "alpha", "before": "body ", "after": " beta", "paragraph": 1},
                {"start": 24, "end": 29, "text": "alpha", "before": "beta\r", "after": " agai", "paragraph": 2},
            ])
        );
        assert_eq!(found["truncated"], false);
        let cased = read(session, &path, json!({"action": "find", "text": "ALPHA"}))?;
        assert_eq!(cased["count"], 0);
        let folded = read(
            session,
            &path,
            json!({"action": "find", "text": "ALPHA", "match_case": false}),
        )?;
        assert_eq!(folded["count"], 2);
        let partial = read(
            session,
            &path,
            json!({"action": "find", "text": "alph", "whole_word": true}),
        )?;
        assert_eq!(partial["count"], 0);
        let wild = read(
            session,
            &path,
            json!({"action": "find", "text": "a[l]p?a", "wildcards": true}),
        )?;
        assert_eq!(wild["count"], 2);
        let limited = read(
            session,
            &path,
            json!({"action": "find", "text": "alpha", "max_results": 1}),
        )?;
        assert_eq!(limited["count"], 1);
        assert_eq!(limited["truncated"], true);
        let special = read(session, &path, json!({"action": "find", "text": "^p"}))?;
        assert_eq!(special["count"], 0, "literal find must escape ^ codes");

        let outline = read(session, &path, json!({"action": "outline"}))?;
        assert_eq!(outline["total_paragraphs"], 4);
        let headings = outline["headings"].as_array().context("headings")?;
        assert_eq!(headings.len(), 2);
        assert_eq!(headings[0]["index"], 0);
        assert_eq!(headings[0]["level"], 1);
        assert_eq!(headings[0]["text"], "Heading");
        assert_eq!(headings[1]["index"], 3);
        assert_eq!(headings[1]["level"], 2);
        assert_eq!(headings[1]["start"], 36);

        paragraph_range(&document, 2)?
            .object("ParagraphFormat")?
            .put("PageBreakBefore", true.into())?;
        let first = read(session, &path, json!({"action": "page_text", "page": 1}))?;
        assert_eq!(first["total_pages"], 2);
        assert_eq!(
            first["pages"],
            json!([{"page": 1, "start": 0, "end": 24, "text": "Heading\rbody alpha beta\r"}])
        );
        let both = read(
            session,
            &path,
            json!({"action": "page_text", "page": 1, "end_page": 5}),
        )?;
        let pages = both["pages"].as_array().context("pages")?;
        assert_eq!(pages.len(), 2);
        assert_eq!(pages[1]["start"], 24);
        assert_eq!(pages[1]["text"], "alpha again\rSub heading\r");
        assert!(read(session, &path, json!({"action": "page_text", "page": 3})).is_err());
        Ok(())
    });
}

#[test]
#[ignore = "Requires desktop Microsoft Word; starts a private hidden instance"]
fn reports_range_and_paragraph_formatting() {
    with_word(|harness| {
        let (path, document) = harness.document(
            "formatting.docx",
            "Heading\rbody alpha beta\ralpha again\rlast",
        )?;
        set_style(&document, 0, HEADING_1)?;
        let alpha = range(&document, 13, 18)?;
        let font = alpha.object("Font")?;
        font.put("Bold", true.into())?;
        font.put("Size", 15.5.into())?;
        // wdColorRed is BGR 0x0000FF.
        font.put("Color", 255.into())?;
        font.put("Underline", 1.into())?;
        // wdYellow.
        alpha.put("HighlightColorIndex", 7.into())?;
        let body = paragraph_range(&document, 1)?.object("ParagraphFormat")?;
        body.put("Alignment", 1.into())?;
        body.put("SpaceBefore", 6.into())?;
        body.put("LeftIndent", 36.into())?;
        body.put("FirstLineIndent", (-18).into())?;
        paragraph_range(&document, 2)?
            .object("ListFormat")?
            .call("ApplyBulletDefault", vec![])?;
        let session = &mut harness.session;

        let styled = read(
            session,
            &path,
            json!({"action": "formatting", "start": 13, "end": 18}),
        )?;
        assert_eq!(styled["paragraph"], Value::Null);
        assert_eq!(
            styled["font"],
            json!({"name": styled["font"]["name"], "size_pt": 15.5, "bold": true, "italic": false, "underline": "single", "strike": false, "color": "FF0000", "highlight": "yellow", "vertical_align": "baseline"})
        );
        assert!(styled["font"]["name"].is_string());
        assert_eq!(styled["alignment"], "center");
        assert_eq!(styled["spacing"]["before_pt"], 6.0);
        assert_eq!(styled["indent"]["left_pt"], 36.0);
        assert_eq!(styled["indent"]["hanging_pt"], 18.0);
        assert_eq!(styled["indent"]["first_line_pt"], Value::Null);
        assert_eq!(styled["numbering"], Value::Null);

        let mixed = read(
            session,
            &path,
            json!({"action": "formatting", "start": 8, "end": 18}),
        )?;
        assert_eq!(mixed["font"]["bold"], Value::Null);
        assert_eq!(mixed["font"]["size_pt"], Value::Null);
        assert_eq!(mixed["font"]["highlight"], Value::Null);

        let heading = read(
            session,
            &path,
            json!({"action": "formatting", "paragraph": 0}),
        )?;
        assert_eq!(heading["paragraph"], 0);
        assert_eq!(heading["start"], 0);
        assert_eq!(heading["end"], 8);
        assert_eq!(heading["outline_level"], 1);
        assert_eq!(heading["keep_next"], true);
        assert!(heading["style_name"].is_string());
        assert_eq!(heading["spacing"]["line_rule"], "multiple");

        let listed = read(
            session,
            &path,
            json!({"action": "formatting", "paragraph": 2}),
        )?;
        assert_eq!(listed["numbering"]["ilvl"], 0);
        assert!(
            !listed["numbering"]["list_string"]
                .as_str()
                .unwrap_or_default()
                .is_empty()
        );
        assert!(
            read(
                session,
                &path,
                json!({"action": "formatting", "paragraph": 9}),
            )
            .is_err()
        );
        assert!(
            read(
                session,
                &path,
                json!({"action": "formatting", "start": 0, "end": 5000}),
            )
            .is_err()
        );
        Ok(())
    });
}

#[test]
#[ignore = "Requires desktop Microsoft Word; starts a private hidden instance"]
fn snapshots_diffs_and_selection() {
    with_word(|harness| {
        let (path, document) =
            harness.document("diff.docx", "Heading\rbody alpha beta\ralpha again\rlast")?;
        let (other_path, _other) = harness.document("other.docx", "other document")?;
        let session = &mut harness.session;

        assert!(read(session, &path, json!({"action": "diff"})).is_err());
        let snapshot = read(session, &path, json!({"action": "snapshot"}))?;
        assert_eq!(snapshot["paragraphs"], 4);
        let same = read(session, &path, json!({"action": "diff"}))?;
        assert_eq!(same["unchanged"], true);

        range(&document, 13, 18)?.put("Text", "ALPHA".into())?;
        let end = document.object("Content")?.int("End")?;
        range(&document, end - 1, end - 1)?.call("InsertAfter", vec!["\rnew tail".into()])?;
        let changed = read(session, &path, json!({"action": "diff"}))?;
        assert_eq!(changed["unchanged"], false);
        assert_eq!(changed["current_paragraphs"], 5);
        assert_eq!(
            changed["changes"],
            json!([
                {"type": "changed", "index": 1, "old_index": 1, "old_text": "body alpha beta", "text": "body ALPHA beta"},
                {"type": "inserted", "index": 4, "text": "new tail"},
            ])
        );
        // Snapshots are per document.
        assert!(read(session, &other_path, json!({"action": "diff"})).is_err());

        // other.docx was opened last, so it is the active document.
        let inactive = read(session, &path, json!({"action": "selection"}))?;
        assert_eq!(inactive["active"], false);
        assert_eq!(inactive["selection"], Value::Null);
        document.call("Activate", vec![])?;
        document
            .object("ActiveWindow")?
            .object("Selection")?
            .call("SetRange", vec![13.into(), 18.into()])?;
        let selected = read(session, &path, json!({"action": "selection"}))?;
        assert_eq!(
            selected["selection"],
            json!({"start": 13, "end": 18, "text": "ALPHA", "collapsed": false, "story": "main", "paragraph": 1})
        );
        document
            .object("ActiveWindow")?
            .object("Selection")?
            .call("SetRange", vec![24.into(), 24.into()])?;
        let caret = read(session, &path, json!({"action": "selection"}))?;
        assert_eq!(caret["selection"]["text"], "");
        assert_eq!(caret["selection"]["paragraph"], 2);
        Ok(())
    });
}

/// Word writes a document with a heading, bold run, bullet list, table, and hyperlink;
/// `docx_read markdown` must render it faithfully.
#[test]
#[ignore = "Requires desktop Microsoft Word; starts a private hidden instance"]
fn docx_markdown_reads_a_word_authored_document() {
    with_word(|harness| {
        let (path, document) = harness.document(
            "authored.docx",
            "Title text\rSome bold words\rfirst\rsecond\r\rafter table\rlink here",
        )?;
        set_style(&document, 0, HEADING_1)?;
        let bold = paragraph_range(&document, 1)?.int("Start")? + 5;
        range(&document, bold, bold + 4)?
            .object("Font")?
            .put("Bold", true.into())?;
        let link = paragraph_range(&document, 6)?.int("Start")?;
        let anchor = range(&document, link, link + 4)?;
        document.object("Hyperlinks")?.call(
            "Add",
            vec![Variant::object(&anchor), "https://example.com/docs".into()],
        )?;
        let list_start = paragraph_range(&document, 2)?.int("Start")?;
        let list_end = paragraph_range(&document, 3)?.int("End")?;
        range(&document, list_start, list_end)?
            .object("ListFormat")?
            .call("ApplyBulletDefault", vec![])?;
        let slot = paragraph_range(&document, 4)?.int("Start")?;
        let table = document
            .object("Tables")?
            .call(
                "Add",
                vec![
                    Variant::object(&range(&document, slot, slot)?),
                    2.into(),
                    2.into(),
                ],
            )?
            .into_object()?
            .context("no table")?;
        for (row, column, value) in [(1, 1, "Name"), (1, 2, "Value"), (2, 1, "a|b"), (2, 2, "1")] {
            table
                .call("Cell", vec![row.into(), column.into()])?
                .into_object()?
                .context("no cell")?
                .object("Range")?
                .put("Text", value.into())?;
        }
        document.call("Save", vec![])?;
        document.call("Close", vec![0.into()])?;

        let Output::Json(rendered) = crate::docx::call(
            "docx_read",
            json!({"path": path, "operation": {"action": "markdown"}}),
        )?
        else {
            anyhow::bail!("docx_read returned an image");
        };
        assert_eq!(
            rendered["markdown"],
            [
                "# Title text",
                "",
                "Some **bold** words",
                "",
                "- first",
                "- second",
                "",
                "| Name | Value |",
                "| --- | --- |",
                "| a\\|b | 1 |",
                "",
                "after table",
                "",
                "[link](https://example.com/docs) here",
            ]
            .join("\n"),
            "{rendered:#}"
        );
        Ok(())
    });
}
