//! Native OOXML document operations.

use std::{
    fmt::Write as _,
    fs,
    io::{Cursor, Read, Write},
    path::{Path, PathBuf},
};

use anyhow::{Context, Result, bail, ensure};
use quick_xml::{events::Event, name::ResolveResult, reader::NsReader};
use serde::Deserialize;
use serde_json::{Value, json};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

const WORD_NS: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MAX_ARCHIVE: u64 = 64 * 1024 * 1024;
const MAX_XML: usize = 16 * 1024 * 1024;

struct Entry {
    name: String,
    bytes: Vec<u8>,
    directory: bool,
}
struct Package {
    original: Vec<u8>,
    entries: Vec<Entry>,
    document: usize,
}

impl Package {
    fn open(path: &Path) -> Result<Self> {
        ensure!(
            fs::metadata(path)?.len() <= MAX_ARCHIVE,
            "DOCX exceeds 64 MiB archive limit"
        );
        let original = fs::read(path)?;
        let mut archive =
            ZipArchive::new(Cursor::new(&original)).context("invalid DOCX ZIP archive")?;
        ensure!(archive.len() <= 4096, "too many ZIP entries");
        let mut total = 0_u64;
        let mut entries = Vec::new();
        let mut document = None;
        for index in 0..archive.len() {
            let mut file = archive.by_index(index)?;
            ensure!(
                file.size() <= MAX_ARCHIVE - total,
                "DOCX uncompressed data exceeds 64 MiB limit"
            );
            ensure!(
                !entries
                    .iter()
                    .any(|entry: &Entry| entry.name == file.name()),
                "duplicate ZIP entry"
            );
            let mut bytes = Vec::new();
            file.by_ref()
                .take(MAX_ARCHIVE - total + 1)
                .read_to_end(&mut bytes)?;
            total = total
                .checked_add(bytes.len() as u64)
                .context("archive size overflow")?;
            ensure!(total <= MAX_ARCHIVE, "ZIP data exceeds size limit");
            if file.name() == "word/document.xml" {
                document = Some(index);
            }
            entries.push(Entry {
                name: file.name().into(),
                bytes,
                directory: file.is_dir(),
            });
        }
        let document = document.context("DOCX has no word/document.xml")?;
        Ok(Self {
            original,
            entries,
            document,
        })
    }

    fn xml(&self) -> Result<&str> {
        let bytes = &self.entries[self.document].bytes;
        ensure!(bytes.len() <= MAX_XML, "document XML exceeds 16 MiB limit");
        std::str::from_utf8(bytes).context("document XML must be UTF-8")
    }

    fn encode(&self) -> Result<Vec<u8>> {
        encode_entries(&self.entries)
    }
}

fn encode_entries(entries: &[Entry]) -> Result<Vec<u8>> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for entry in entries {
        if entry.directory {
            writer.add_directory(&entry.name, options)?;
        } else {
            writer.start_file(&entry.name, options)?;
            writer.write_all(&entry.bytes)?;
        }
    }
    Ok(writer.finish()?.into_inner())
}

#[derive(Clone)]
struct Node {
    name: String,
    word: bool,
    start: usize,
    open_end: usize,
    close_start: usize,
    end: usize,
    parent: Option<usize>,
    text: String,
}
struct Document {
    nodes: Vec<Node>,
    paragraphs: Vec<usize>,
    body: usize,
}

impl Document {
    fn parse(xml: &str) -> Result<Self> {
        ensure!(xml.len() <= MAX_XML, "document XML exceeds size limit");
        let mut reader = NsReader::from_str(xml);
        reader.config_mut().check_end_names = true;
        let mut nodes: Vec<Node> = Vec::new();
        let mut stack: Vec<usize> = Vec::new();
        let mut paragraphs = Vec::new();
        let mut body = None;
        loop {
            let start = usize::try_from(reader.buffer_position())?;
            let (namespace, event) = reader.read_resolved_event()?;
            let word = matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == WORD_NS);
            let end = usize::try_from(reader.buffer_position())?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    ensure!(stack.len() < 128, "XML nesting exceeds 128 levels");
                    ensure!(nodes.len() < 250_000, "XML exceeds 250000 element limit");
                    let name = std::str::from_utf8(element.local_name().as_ref())?.to_owned();
                    let index = nodes.len();
                    if word && name == "p" {
                        paragraphs.push(index);
                    }
                    if word && name == "body" {
                        ensure!(body.is_none(), "multiple document bodies");
                        body = Some(index);
                    }
                    nodes.push(Node {
                        name,
                        word,
                        start,
                        open_end: end,
                        close_start: end,
                        end,
                        parent: stack.last().copied(),
                        text: String::new(),
                    });
                    if matches!(event, Event::Start(_)) {
                        stack.push(index);
                    }
                }
                Event::End(_) => {
                    let index = stack.pop().context("unexpected closing XML element")?;
                    nodes[index].close_start = start;
                    nodes[index].end = end;
                }
                Event::Text(value) => {
                    if let Some(index) = stack.last() {
                        nodes[*index]
                            .text
                            .push_str(&quick_xml::escape::unescape(&value.xml_content()?)?);
                    }
                }
                Event::CData(value) => {
                    if let Some(index) = stack.last() {
                        nodes[*index].text.push_str(&value.xml_content()?);
                    }
                }
                Event::DocType(_) => bail!("DOCTYPE declarations are not supported"),
                Event::Eof => break,
                Event::GeneralRef(value) => {
                    let reference = value.decode()?;
                    let entity = format!("&{reference};");
                    let decoded = quick_xml::escape::unescape(&entity)
                        .context("unsupported XML entity reference")?;
                    if let Some(index) = stack.last() {
                        nodes[*index].text.push_str(&decoded);
                    }
                }
                _ => {}
            }
        }
        ensure!(stack.is_empty(), "unclosed XML element");
        let body = body.context("missing Word document body")?;
        ensure!(
            nodes
                .first()
                .is_some_and(|node| node.word && node.name == "document"),
            "invalid Word document root"
        );
        ensure!(
            nodes.iter().filter(|node| node.parent.is_none()).count() == 1
                && nodes[body].parent == Some(0),
            "invalid Word document structure"
        );
        Ok(Self {
            nodes,
            paragraphs,
            body,
        })
    }

    fn owner(&self, mut index: usize) -> Option<usize> {
        loop {
            let node = &self.nodes[index];
            if node.word && node.name == "p" {
                return Some(index);
            }
            index = node.parent?;
        }
    }

    fn texts(&self, paragraph: usize) -> Vec<usize> {
        self.nodes
            .iter()
            .enumerate()
            .skip(paragraph + 1)
            .take_while(|(_, node)| node.start < self.nodes[paragraph].close_start)
            .filter_map(|(index, node)| {
                (node.word && node.name == "t" && self.owner(index) == Some(paragraph))
                    .then_some(index)
            })
            .collect()
    }

    fn text(&self, paragraph: usize) -> String {
        let mut text = String::new();
        for (index, node) in self
            .nodes
            .iter()
            .enumerate()
            .skip(paragraph + 1)
            .take_while(|(_, node)| node.start < self.nodes[paragraph].close_start)
        {
            if !node.word || self.owner(index) != Some(paragraph) {
                continue;
            }
            match node.name.as_str() {
                "t" => text.push_str(&node.text),
                "tab" => text.push('\t'),
                "br" | "cr" => text.push('\n'),
                _ => {}
            }
        }
        text
    }

    fn editable(&self, paragraph: usize) -> Result<()> {
        let target = &self.nodes[paragraph];
        let mut parent = target.parent;
        while let Some(index) = parent {
            let node = &self.nodes[index];
            ensure!(
                !(node.word && matches!(node.name.as_str(), "sdt" | "ins" | "del")),
                "paragraph belongs to an unsupported revision or content control"
            );
            parent = node.parent;
        }
        for node in &self.nodes {
            if node.start <= target.start || node.end >= target.end {
                continue;
            }
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

    fn child(&self, parent: usize, name: &str) -> Option<usize> {
        self.nodes
            .iter()
            .position(|node| node.parent == Some(parent) && node.word && node.name == name)
    }
}

type Patch = (usize, usize, String);

fn patch(xml: &str, mut changes: Vec<Patch>) -> Result<String> {
    changes.sort_by_key(|change| change.0);
    for pair in changes.windows(2) {
        ensure!(pair[0].1 <= pair[1].0, "overlapping XML edits");
    }
    let mut result = xml.to_owned();
    for (start, end, replacement) in changes.into_iter().rev() {
        result.replace_range(start..end, &replacement);
    }
    Document::parse(&result)?;
    Ok(result)
}

fn escaped(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn validate_text(value: &str) -> Result<()> {
    ensure!(
        value
            .chars()
            .all(|character| matches!(character, '\t' | '\n' | '\r')
                || (character >= ' ' && !matches!(character, '\u{fffe}' | '\u{ffff}'))),
        "text contains an invalid XML control character"
    );
    Ok(())
}

fn paragraph_xml(text: &str, style: Option<&str>) -> Result<String> {
    validate_text(text)?;
    if let Some(style) = style {
        validate_text(style)?;
    }
    let properties = style.map_or_else(String::new, |style| {
        format!("<w:pPr><w:pStyle w:val=\"{}\"/></w:pPr>", escaped(style))
    });
    Ok(format!(
        "<w:p xmlns:w=\"{}\">{properties}<w:r>{}</w:r></w:p>",
        String::from_utf8_lossy(WORD_NS),
        text_xml(text)
    ))
}

fn text_xml(text: &str) -> String {
    let normalized = text.replace("\r\n", "\n").replace('\r', "\n");
    let namespace = String::from_utf8_lossy(WORD_NS);
    let mut result = String::new();
    let mut start = 0;
    for (offset, character) in normalized.char_indices() {
        if !matches!(character, '\n' | '\t') {
            continue;
        }
        let name = if character == '\n' { "br" } else { "tab" };
        // Writing to a String cannot fail.
        let _ = write!(
            result,
            "<w:t xmlns:w=\"{namespace}\" xml:space=\"preserve\">{}</w:t><w:{name} xmlns:w=\"{namespace}\"/>",
            escaped(&normalized[start..offset])
        );
        start = offset + character.len_utf8();
    }
    let _ = write!(
        result,
        "<w:t xmlns:w=\"{namespace}\" xml:space=\"preserve\">{}</w:t>",
        escaped(&normalized[start..])
    );
    result
}

fn expand_empty(xml: &str, node: &Node, content: &str) -> Result<String> {
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

fn check_lock(path: &Path) -> Result<()> {
    let name = path
        .file_name()
        .context("document path has no filename")?
        .to_string_lossy();
    let short: String = name.chars().skip(2).collect();
    for candidate in [format!("~${name}"), format!("~${short}")] {
        ensure!(
            !path.with_file_name(candidate).exists(),
            "Word lock file detected; close the document before saved-file edits"
        );
    }
    Ok(())
}

fn write_atomic(
    path: &Path,
    bytes: &[u8],
    original: Option<&[u8]>,
    overwrite: bool,
) -> Result<Option<PathBuf>> {
    check_lock(path)?;
    let parent = path
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut output = tempfile::NamedTempFile::new_in(parent)?;
    output.write_all(bytes)?;
    output.as_file().sync_all()?;
    let backup = if let Some(original) = original {
        ensure!(
            fs::read(path)? == original,
            "document changed during operation; retry after reading it again"
        );
        let mut backup = tempfile::Builder::new()
            .prefix(&format!(
                "{}.backup-",
                path.file_name()
                    .context("missing filename")?
                    .to_string_lossy()
            ))
            .suffix(".docx")
            .tempfile_in(parent)?;
        backup.write_all(original)?;
        backup.as_file().sync_all()?;
        let (_, backup_path) = backup.keep().context("could not retain document backup")?;
        Some(backup_path)
    } else {
        None
    };
    check_lock(path)?;
    if let Some(original) = original {
        ensure!(
            fs::read(path)? == original,
            "document changed before save; original retained in backup"
        );
    }
    if overwrite {
        output.persist(path).map_err(|error| error.error)?;
    } else {
        output
            .persist_noclobber(path)
            .map_err(|error| error.error)?;
    }
    Ok(backup)
}

fn save_edit(path: &Path, mut package: Package, xml: String) -> Result<Value> {
    package.entries[package.document].bytes = xml.into_bytes();
    let bytes = package.encode()?;
    let backup = write_atomic(path, &bytes, Some(&package.original), true)?;
    Ok(json!({"path":path,"backup_path":backup,"modified":true}))
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct Arguments {
    path: PathBuf,
    #[serde(default)]
    paragraphs: Vec<String>,
    #[serde(default)]
    overwrite: bool,
    start: Option<usize>,
    limit: Option<usize>,
    find: Option<String>,
    replacement: Option<String>,
    paragraph_index: Option<usize>,
    #[serde(default)]
    replace_all: bool,
    text: Option<String>,
    index: Option<usize>,
    style: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    font_size_pt: Option<f64>,
    font_family: Option<String>,
    alignment: Option<String>,
    output_path: Option<PathBuf>,
}

fn definition(
    name: &str,
    description: &str,
    properties: &Value,
    required: &[&str],
    read_only: bool,
) -> Value {
    json!({"name":name,"description":description,
        "inputSchema":{"type":"object","properties":properties,"required":required,"additionalProperties":false},
        "annotations":{"readOnlyHint":read_only,"destructiveHint":!read_only,"idempotentHint":read_only,"openWorldHint":false}})
}

/// Return the native DOCX tool definitions.
#[must_use]
pub fn tools() -> Vec<Value> {
    let path = json!({"type":"string","minLength":1,"description":"Local DOCX path"});
    let index = json!({"type":"integer","minimum":0});
    let string = json!({"type":"string"});
    let boolean = json!({"type":"boolean"});
    vec![
        definition(
            "create_document",
            "Create a DOCX package. Existing output requires overwrite:true.",
            &json!({"path":path,"paragraphs":{"type":"array","items":string},"overwrite":boolean}),
            &["path"],
            false,
        ),
        definition(
            "read_document",
            "Read main-document paragraphs, including tables, with zero-based pagination.",
            &json!({"path":path,"start":index,"limit":index}),
            &["path"],
            true,
        ),
        definition(
            "replace_text",
            "Replace literal text across plain runs, preserving surrounding formatting. Edits create a backup.",
            &json!({"path":path,"find":{"type":"string","minLength":1},"replacement":string,"paragraph_index":index,"replace_all":boolean}),
            &["path", "find", "replacement"],
            false,
        ),
        definition(
            "insert_paragraph",
            "Insert a paragraph before a zero-based index or append. Edits create a backup.",
            &json!({"path":path,"text":string,"index":index,"style":string}),
            &["path", "text"],
            false,
        ),
        definition(
            "format_paragraph",
            "Set direct paragraph/run formatting. Edits create a backup.",
            &json!({"path":path,"index":index,"bold":boolean,"italic":boolean,"underline":boolean,"font_size_pt":{"type":"number","minimum":1,"maximum":1638},"font_family":string,"alignment":{"type":"string","enum":["left","center","right","both"]},"style":string}),
            &["path", "index"],
            false,
        ),
        definition(
            "get_document_info",
            "Inspect paragraph count, character count and package parts.",
            &json!({"path":path}),
            &["path"],
            true,
        ),
        definition(
            "preview_document",
            "Return an escaped logical HTML preview; optional output requires overwrite:true if it exists.",
            &json!({"path":path,"output_path":string,"overwrite":boolean}),
            &["path"],
            false,
        ),
    ]
}

/// Dispatch a native DOCX operation.
///
/// # Errors
/// Returns an error for an unknown operation or an invalid document.
pub fn call(name: &str, args: Value) -> Result<Value> {
    let definition = tools()
        .into_iter()
        .find(|tool| tool["name"] == name)
        .context("unknown DOCX tool")?;
    let object = args.as_object().context("arguments must be an object")?;
    ensure!(
        object.values().all(|value| !value.is_null()),
        "arguments must not be null"
    );
    let properties = definition["inputSchema"]["properties"]
        .as_object()
        .context("invalid tool schema")?;
    for key in object.keys() {
        ensure!(properties.contains_key(key), "unexpected argument: {key}");
    }
    for key in definition["inputSchema"]["required"]
        .as_array()
        .context("invalid required schema")?
    {
        let key = key.as_str().context("invalid required property")?;
        ensure!(
            object.contains_key(key) && !object[key].is_null(),
            "missing required argument: {key}"
        );
    }
    let arguments: Arguments = serde_json::from_value(args).context("invalid tool arguments")?;
    ensure!(
        !arguments.path.as_os_str().is_empty(),
        "path must not be empty"
    );
    if name == "create_document" {
        return create(&arguments);
    }
    let package = Package::open(&arguments.path)?;
    let xml = package.xml()?;
    let document = Document::parse(xml)?;
    match name {
        "read_document" => {
            let start = arguments.start.unwrap_or(0);
            let paragraphs: Vec<Value> = document
                .paragraphs
                .iter()
                .enumerate()
                .skip(start)
                .take(arguments.limit.unwrap_or(usize::MAX))
                .map(|(index, paragraph)| json!({"index":index,"text":document.text(*paragraph)}))
                .collect();
            let text = paragraphs
                .iter()
                .filter_map(|paragraph| paragraph["text"].as_str())
                .collect::<Vec<_>>()
                .join("\n");
            Ok(
                json!({"path":arguments.path,"paragraphs":paragraphs,"text":text,"total_paragraphs":document.paragraphs.len(),"start":start}),
            )
        }
        "get_document_info" => {
            let characters: usize = document
                .paragraphs
                .iter()
                .map(|paragraph| document.text(*paragraph).chars().count())
                .sum();
            Ok(
                json!({"path":arguments.path,"paragraph_count":document.paragraphs.len(),"character_count":characters,"parts":package.entries.iter().map(|entry| &entry.name).collect::<Vec<_>>()}),
            )
        }
        "preview_document" => preview(&arguments, &document),
        "replace_text" => {
            let (new_xml, count) = replace(&arguments, &document, xml)?;
            if count == 0 {
                return Ok(json!({"path":arguments.path,"modified":false,"replacements":0}));
            }
            let mut result = save_edit(&arguments.path, package, new_xml)?;
            result["replacements"] = json!(count);
            Ok(result)
        }
        "insert_paragraph" => {
            let new_xml = insert(&arguments, &document, xml)?;
            save_edit(&arguments.path, package, new_xml)
        }
        "format_paragraph" => {
            let new_xml = format(&arguments, &document, xml)?;
            save_edit(&arguments.path, package, new_xml)
        }
        _ => bail!("unknown DOCX tool: {name}"),
    }
}

fn create(arguments: &Arguments) -> Result<Value> {
    let paragraphs = if arguments.paragraphs.is_empty() {
        vec![String::new()]
    } else {
        arguments.paragraphs.clone()
    };
    let body = paragraphs
        .iter()
        .map(|text| paragraph_xml(text, None))
        .collect::<Result<Vec<_>>>()?
        .join("");
    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"{}\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr></w:body></w:document>",
        String::from_utf8_lossy(WORD_NS)
    );
    Document::parse(&xml)?;
    let entries = vec![
        Entry { name: "[Content_Types].xml".into(), bytes: br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(), directory: false },
        Entry { name: "_rels/.rels".into(), bytes: br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(), directory: false },
        Entry { name: "word/document.xml".into(), bytes: xml.into_bytes(), directory: false },
    ];
    write_atomic(
        &arguments.path,
        &encode_entries(&entries)?,
        None,
        arguments.overwrite,
    )?;
    Ok(json!({"path":arguments.path,"created":true,"paragraph_count":paragraphs.len()}))
}

fn preview(arguments: &Arguments, document: &Document) -> Result<Value> {
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

fn selected_paragraph(arguments: &Arguments, document: &Document) -> Result<usize> {
    let index = arguments.index.context("index is required")?;
    document
        .paragraphs
        .get(index)
        .copied()
        .context("paragraph index is out of bounds")
}

fn replace(arguments: &Arguments, document: &Document, xml: &str) -> Result<(String, usize)> {
    let find = arguments.find.as_deref().context("find is required")?;
    let replacement = arguments
        .replacement
        .as_deref()
        .context("replacement is required")?;
    ensure!(!find.is_empty(), "find must not be empty");
    validate_text(replacement)?;
    if let Some(index) = arguments.paragraph_index {
        ensure!(
            index < document.paragraphs.len(),
            "paragraph index is out of bounds"
        );
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
        for node in document
            .nodes
            .iter()
            .skip(paragraph + 1)
            .take_while(|node| node.start < document.nodes[paragraph].close_start)
        {
            ensure!(
                !(node.word
                    && (node.name.contains("bookmark")
                        || node.name.contains("comment")
                        || node.name.starts_with("moveFrom")
                        || node.name.starts_with("moveTo"))),
                "replacement paragraph contains anchored ranges; use live Word editing"
            );
        }
        if !arguments.replace_all {
            matches.truncate(1);
        }
        let texts = document.texts(paragraph);
        let mut values: Vec<String> = texts
            .iter()
            .map(|index| document.nodes[*index].text.clone())
            .collect();
        let mut offsets = Vec::new();
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
                    if node == first { replacement } else { "" },
                );
            }
        }
        for (node, value) in texts.into_iter().zip(values) {
            if value == document.nodes[node].text {
                continue;
            }
            let node = &document.nodes[node];
            changes.push((node.start, node.end, text_xml(&value)));
        }
        count += matches.len();
        if !arguments.replace_all {
            break;
        }
    }
    Ok((patch(xml, changes)?, count))
}

fn insert(arguments: &Arguments, document: &Document, xml: &str) -> Result<String> {
    let index = arguments.index.unwrap_or(document.paragraphs.len());
    ensure!(
        index <= document.paragraphs.len(),
        "paragraph index is out of bounds"
    );
    let paragraph = paragraph_xml(
        arguments.text.as_deref().context("text is required")?,
        arguments.style.as_deref(),
    )?;
    let body = &document.nodes[document.body];
    if body.open_end == body.end {
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
            .map_or(document.nodes[document.body].close_start, |node| {
                document.nodes[node].start
            })
    };
    patch(xml, vec![(offset, offset, paragraph)])
}

fn property_xml(name: &str, value: &str) -> String {
    format!(
        "<w:{name} xmlns:w=\"{}\" w:val=\"{}\"/>",
        String::from_utf8_lossy(WORD_NS),
        escaped(value)
    )
}

/// `CT_PPr` child order from the OOXML schema; Word rejects out-of-order children.
const PARAGRAPH_PROPERTY_ORDER: &[&str] = &[
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
const RUN_PROPERTY_ORDER: &[&str] = &[
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

fn property_patches(
    document: &Document,
    xml: &str,
    parent: usize,
    name: &str,
    properties: &[(String, String)],
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
    let rank = |property: &str| {
        order
            .iter()
            .position(|name| *name == property)
            .unwrap_or(usize::MAX)
    };
    let mut properties: Vec<&(String, String)> = properties.iter().collect();
    properties.sort_by_key(|(name, _)| rank(name));
    if let Some(container) = document.child(parent, name) {
        let node = &document.nodes[container];
        if node.open_end == node.end {
            let children = properties
                .iter()
                .map(|(_, value)| value.as_str())
                .collect::<String>();
            changes.push((node.start, node.end, expand_empty(xml, node, &children)?));
            return Ok(());
        }
        for (property, value) in properties {
            if let Some(existing) = document.child(container, property) {
                let existing = &document.nodes[existing];
                changes.push((existing.start, existing.end, value.clone()));
            } else {
                let offset = document
                    .nodes
                    .iter()
                    .find(|child| {
                        child.parent == Some(container)
                            && child.word
                            && rank(&child.name) > rank(property)
                    })
                    .map_or(node.close_start, |child| child.start);
                changes.push((offset, offset, value.clone()));
            }
        }
    } else {
        let children = properties
            .iter()
            .map(|(_, value)| value.as_str())
            .collect::<String>();
        let content = format!(
            "<w:{name} xmlns:w=\"{}\">{children}</w:{name}>",
            String::from_utf8_lossy(WORD_NS)
        );
        let parent = &document.nodes[parent];
        if parent.open_end == parent.end {
            changes.push((
                parent.start,
                parent.end,
                expand_empty(xml, parent, &content)?,
            ));
        } else {
            changes.push((parent.open_end, parent.open_end, content));
        }
    }
    Ok(())
}

/// Run properties requested by `format_paragraph`, as (element name, XML) pairs.
fn requested_run_properties(arguments: &Arguments) -> Result<Vec<(String, String)>> {
    let mut run_properties = Vec::new();
    for (name, value) in [("b", arguments.bold), ("i", arguments.italic)] {
        if let Some(value) = value {
            run_properties.push((
                name.to_owned(),
                property_xml(name, if value { "1" } else { "0" }),
            ));
        }
    }
    if let Some(value) = arguments.underline {
        run_properties.push((
            "u".into(),
            property_xml("u", if value { "single" } else { "none" }),
        ));
    }
    if let Some(size) = arguments.font_size_pt {
        ensure!(
            size.is_finite() && (1.0..=1638.0).contains(&size) && (size * 2.0).fract() == 0.0,
            "font size must be 1–1638 points in half-point increments"
        );
        run_properties.push((
            "sz".into(),
            property_xml("sz", &format!("{:.0}", size * 2.0)),
        ));
    }
    if let Some(font) = &arguments.font_family {
        validate_text(font)?;
        ensure!(!font.is_empty(), "font family must not be empty");
        run_properties.push(("rFonts".into(), format!("<w:rFonts xmlns:w=\"{}\" w:ascii=\"{}\" w:hAnsi=\"{}\" w:eastAsia=\"{}\" w:cs=\"{}\"/>", String::from_utf8_lossy(WORD_NS), escaped(font), escaped(font), escaped(font), escaped(font))));
    }
    Ok(run_properties)
}

/// Paragraph properties requested by `format_paragraph`, as (element name, XML) pairs.
fn requested_paragraph_properties(arguments: &Arguments) -> Result<Vec<(String, String)>> {
    let mut paragraph_properties = Vec::new();
    if let Some(alignment) = &arguments.alignment {
        ensure!(
            matches!(alignment.as_str(), "left" | "center" | "right" | "both"),
            "invalid paragraph alignment"
        );
        paragraph_properties.push(("jc".into(), property_xml("jc", alignment)));
    }
    if let Some(style) = &arguments.style {
        validate_text(style)?;
        paragraph_properties.push(("pStyle".into(), property_xml("pStyle", style)));
    }
    Ok(paragraph_properties)
}

fn format(arguments: &Arguments, document: &Document, xml: &str) -> Result<String> {
    let paragraph = selected_paragraph(arguments, document)?;
    document.editable(paragraph)?;
    let run_properties = requested_run_properties(arguments)?;
    let paragraph_properties = requested_paragraph_properties(arguments)?;
    ensure!(
        !run_properties.is_empty() || !paragraph_properties.is_empty(),
        "provide at least one formatting property"
    );
    let node = &document.nodes[paragraph];
    let mut changes = Vec::new();
    if node.open_end == node.end {
        let ppr = paragraph_properties
            .iter()
            .map(|(_, value)| value.as_str())
            .collect::<String>();
        let rpr = run_properties
            .iter()
            .map(|(_, value)| value.as_str())
            .collect::<String>();
        let content = format!(
            "<w:pPr xmlns:w=\"{}\">{ppr}</w:pPr><w:r xmlns:w=\"{}\"><w:rPr>{rpr}</w:rPr><w:t/></w:r>",
            String::from_utf8_lossy(WORD_NS),
            String::from_utf8_lossy(WORD_NS)
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
            .nodes
            .iter()
            .enumerate()
            .skip(paragraph + 1)
            .take_while(|(_, child)| child.start < node.close_start)
            .filter_map(|(index, child)| {
                (child.word && child.name == "r" && document.owner(index) == Some(paragraph))
                    .then_some(index)
            })
            .collect();
        for run in &runs {
            property_patches(document, xml, *run, "rPr", &run_properties, &mut changes)?;
        }
        if runs.is_empty() && !run_properties.is_empty() {
            let properties = run_properties
                .iter()
                .map(|(_, value)| value.as_str())
                .collect::<String>();
            changes.push((
                node.close_start,
                node.close_start,
                format!(
                    "<w:r xmlns:w=\"{}\"><w:rPr>{properties}</w:rPr><w:t/></w:r>",
                    String::from_utf8_lossy(WORD_NS)
                ),
            ));
        }
    }
    patch(xml, changes)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture(path: &Path, body: &str) -> Result<()> {
        let xml = format!(
            "<q:document xmlns:q=\"{}\"><q:body>{body}</q:body></q:document>",
            String::from_utf8_lossy(WORD_NS)
        );
        let entries = [
            Entry {
                name: "word/document.xml".into(),
                bytes: xml.into_bytes(),
                directory: false,
            },
            Entry {
                name: "custom/preserved.bin".into(),
                bytes: vec![0, 255, 13, 10],
                directory: false,
            },
        ];
        fs::write(path, encode_entries(&entries)?)?;
        Ok(())
    }

    #[test]
    fn unicode_entities_controls_and_preview_roundtrip() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("unicode.docx");
        let text = "é 🍎 <script>& \"quotes\"\nline\ttab";
        call("create_document", json!({"path":path,"paragraphs":[text]}))?;
        let read = call("read_document", json!({"path":path}))?;
        assert_eq!(read["text"], text);
        let html = call("preview_document", json!({"path":path}))?;
        assert!(
            html["html"]
                .as_str()
                .context("missing HTML")?
                .contains("&lt;script&gt;&amp;")
        );
        assert!(call("create_document", json!({"path":path})).is_err());
        Ok(())
    }

    #[test]
    fn split_runs_preserve_format_parts_and_original_backup() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("split.docx");
        fixture(
            &path,
            "<q:p><q:r><q:rPr><q:b/></q:rPr><q:t>café </q:t></q:r><q:r><q:rPr><q:i/></q:rPr><q:t>🍎 today café 🍎</q:t></q:r></q:p>",
        )?;
        let original = fs::read(&path)?;
        let result = call(
            "replace_text",
            json!({"path":path,"find":"café 🍎","replacement":"A & B","replace_all":true}),
        )?;
        assert_eq!(result["replacements"], 2);
        let backup = result["backup_path"].as_str().context("missing backup")?;
        assert_eq!(fs::read(backup)?, original);
        let package = Package::open(&path)?;
        assert!(package.xml()?.contains("<q:b/>") && package.xml()?.contains("<q:i/>"));
        assert_eq!(
            package
                .entries
                .iter()
                .find(|entry| entry.name == "custom/preserved.bin")
                .context("missing part")?
                .bytes,
            vec![0, 255, 13, 10]
        );
        assert_eq!(
            call("read_document", json!({"path":path}))?["text"],
            "A & B today A & B"
        );
        Ok(())
    }

    #[test]
    fn formatting_preserves_properties_and_orders_new_ones() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("format.docx");
        fixture(
            &path,
            "<q:p><q:pPr><q:spacing q:after=\"120\"/><q:jc q:val=\"right\"/></q:pPr><q:r><q:rPr><q:color q:val=\"112233\"/></q:rPr><q:t>hello</q:t></q:r><q:r q:rsidR=\"01234567\"/></q:p><q:p q:rsidR=\"76543210\"/>",
        )?;
        call(
            "format_paragraph",
            json!({"path":path,"index":0,"bold":true,"font_size_pt":12.5,"style":"Heading1","alignment":"center"}),
        )?;
        call(
            "format_paragraph",
            json!({"path":path,"index":1,"italic":true}),
        )?;
        let package = Package::open(&path)?;
        let xml = package.xml()?;
        assert!(xml.contains("q:after=\"120\"") && xml.contains("q:val=\"112233\""));
        assert!(xml.contains("q:rsidR=\"01234567\"") && xml.contains("q:rsidR=\"76543210\""));
        assert!(
            xml.find("<w:pStyle").context("missing style")?
                < xml.find("<q:spacing").context("missing spacing")?
        );
        let document = Document::parse(xml)?;
        for (index, node) in document.nodes.iter().enumerate() {
            if node.word && node.name == "rPr" {
                assert!(
                    node.parent
                        .is_some_and(|parent| document.nodes[parent].name == "r"),
                    "run properties {index} have wrong parent"
                );
            }
        }
        Ok(())
    }

    #[test]
    fn insert_empty_body_and_table_indices() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("table.docx");
        fixture(
            &path,
            "<q:p><q:r><q:t>before</q:t></q:r></q:p><q:tbl><q:tr><q:tc><q:p><q:r><q:t>cell</q:t></q:r></q:p></q:tc></q:tr></q:tbl><q:sectPr/>",
        )?;
        call(
            "insert_paragraph",
            json!({"path":path,"index":1,"text":"inserted"}),
        )?;
        assert_eq!(
            call("read_document", json!({"path":path}))?["text"],
            "before\ninserted\ncell"
        );
        let empty = "<q:document xmlns:q=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><q:body/></q:document>";
        let document = Document::parse(empty)?;
        let arguments: Arguments = serde_json::from_value(json!({"path":path,"text":"new"}))?;
        let edited = insert(&arguments, &document, empty)?;
        let parsed = Document::parse(&edited)?;
        assert_eq!(parsed.nodes[parsed.paragraphs[0]].parent, Some(parsed.body));
        Ok(())
    }

    #[test]
    fn locks_and_unsupported_edits_leave_original_unchanged() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("locked.docx");
        fixture(&path, "<q:p><q:r><q:t>text</q:t></q:r></q:p>")?;
        let original = fs::read(&path)?;
        for name in ["~$locked.docx", "~$cked.docx"] {
            let lock = directory.path().join(name);
            fs::write(&lock, [])?;
            assert!(
                call(
                    "replace_text",
                    json!({"path":path,"find":"text","replacement":"changed"})
                )
                .is_err()
            );
            assert_eq!(fs::read(&path)?, original);
            fs::remove_file(lock)?;
        }
        for body in [
            "<q:sdt><q:sdtContent><q:p><q:r><q:t>text</q:t></q:r></q:p></q:sdtContent></q:sdt>",
            "<q:p><q:fldSimple q:instr=\"DATE\"><q:r><q:t>text</q:t></q:r></q:fldSimple></q:p>",
            "<q:p><q:bookmarkStart q:id=\"1\"/><q:r><q:t>text</q:t></q:r><q:bookmarkEnd q:id=\"1\"/></q:p>",
        ] {
            fixture(&path, body)?;
            let before = fs::read(&path)?;
            assert!(
                call(
                    "replace_text",
                    json!({"path":path,"find":"text","replacement":"changed"})
                )
                .is_err()
            );
            assert_eq!(fs::read(&path)?, before);
        }
        Ok(())
    }

    #[test]
    fn reject_doctype_depth_and_optimistic_conflict() -> Result<()> {
        assert!(Document::parse("<!DOCTYPE document><w:document/>").is_err());
        let nested = format!("{}{}", "<a>".repeat(129), "</a>".repeat(129));
        assert!(Document::parse(&nested).is_err());
        assert!(validate_text("bad\u{ffff}").is_err());
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("changed.docx");
        fs::write(&path, b"new content")?;
        assert!(write_atomic(&path, b"edit", Some(b"old content"), true).is_err());
        assert_eq!(fs::read(&path)?, b"new content");
        Ok(())
    }
}
