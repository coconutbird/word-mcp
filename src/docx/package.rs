#![expect(
    dead_code,
    reason = "SCAFFOLD: API for areas under construction; remove once all are used"
)]
//! DOCX ZIP packages: bounded loading, part and relationship editing, and safe
//! atomic saves.

use std::{
    fs,
    io::{Cursor, Read, Write},
    path::{Path, PathBuf},
};

use anyhow::{Context, Result, ensure};
use serde_json::{Value, json};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

use super::{
    ooxml::{content_type, ns},
    xml::{Document, MAX_XML, Tree, append_child, escaped, patch_part},
};

const MAX_ARCHIVE: u64 = 64 * 1024 * 1024;
const MAX_ENTRIES: usize = 4096;
/// Name of the main document part.
pub(super) const DOCUMENT_PART: &str = "word/document.xml";
/// Name of the content-types part.
const CONTENT_TYPES_PART: &str = "[Content_Types].xml";

/// One ZIP entry, kept byte-for-byte unless edited.
pub(super) struct Entry {
    pub(super) name: String,
    pub(super) bytes: Vec<u8>,
    pub(super) directory: bool,
}

/// One relationship from a `.rels` part.
#[derive(Clone, Debug)]
pub(super) struct Relationship {
    pub(super) id: String,
    pub(super) kind: String,
    pub(super) target: String,
    pub(super) external: bool,
}

/// A loaded DOCX package and the exact bytes it was read from.
pub(super) struct Package {
    /// The file contents at load time, for backups and conflict detection; empty for a
    /// package built in memory.
    pub(super) original: Vec<u8>,
    pub(super) entries: Vec<Entry>,
    modified: bool,
}

impl Package {
    /// Load a DOCX within the archive size and entry-count limits.
    pub(super) fn open(path: &Path) -> Result<Self> {
        let size = fs::metadata(path)
            .with_context(|| format!("cannot read {}", path.display()))?
            .len();
        ensure!(size <= MAX_ARCHIVE, "DOCX exceeds 64 MiB archive limit");
        let original = fs::read(path).with_context(|| format!("cannot read {}", path.display()))?;
        let mut archive =
            ZipArchive::new(Cursor::new(&original)).context("invalid DOCX ZIP archive")?;
        ensure!(archive.len() <= MAX_ENTRIES, "too many ZIP entries");
        let mut total = 0_u64;
        let mut entries: Vec<Entry> = Vec::new();
        for index in 0..archive.len() {
            let mut file = archive.by_index(index)?;
            ensure!(
                file.size() <= MAX_ARCHIVE - total,
                "DOCX uncompressed data exceeds 64 MiB limit"
            );
            ensure!(
                !entries.iter().any(|entry| entry.name == file.name()),
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
            entries.push(Entry {
                name: file.name().into(),
                bytes,
                directory: file.is_dir(),
            });
        }
        let package = Self {
            original,
            entries,
            modified: false,
        };
        ensure!(
            package.part(DOCUMENT_PART).is_some(),
            "DOCX has no word/document.xml"
        );
        Ok(package)
    }

    /// A package built in memory from `entries`, marked modified.
    pub(super) fn from_entries(entries: Vec<Entry>) -> Self {
        Self {
            original: Vec::new(),
            entries,
            modified: true,
        }
    }

    /// Whether any part changed since loading.
    pub(super) fn modified(&self) -> bool {
        self.modified
    }

    /// The bytes of part `name` (no leading slash).
    pub(super) fn part(&self, name: &str) -> Option<&[u8]> {
        self.entries
            .iter()
            .find(|entry| !entry.directory && entry.name == name)
            .map(|entry| entry.bytes.as_slice())
    }

    /// Part `name` as UTF-8 XML, or `None` when the part does not exist.
    pub(super) fn text(&self, name: &str) -> Result<Option<&str>> {
        let Some(bytes) = self.part(name) else {
            return Ok(None);
        };
        ensure!(bytes.len() <= MAX_XML, "{name} exceeds 16 MiB limit");
        let text = std::str::from_utf8(bytes).with_context(|| format!("{name} must be UTF-8"))?;
        // Tolerate a byte-order mark, which XML allows.
        Ok(Some(text.strip_prefix('\u{feff}').unwrap_or(text)))
    }

    /// The names of all non-directory parts.
    pub(super) fn part_names(&self) -> impl Iterator<Item = &str> {
        self.entries
            .iter()
            .filter(|entry| !entry.directory)
            .map(|entry| entry.name.as_str())
    }

    /// Replace part `name`, or add it at the end of the archive.
    pub(super) fn set(&mut self, name: &str, bytes: Vec<u8>) {
        self.modified = true;
        if let Some(entry) = self
            .entries
            .iter_mut()
            .find(|entry| !entry.directory && entry.name == name)
        {
            entry.bytes = bytes;
        } else {
            self.entries.push(Entry {
                name: name.to_owned(),
                bytes,
                directory: false,
            });
        }
    }

    /// Remove part `name`; returns whether it existed.
    pub(super) fn remove(&mut self, name: &str) -> bool {
        let before = self.entries.len();
        self.entries
            .retain(|entry| entry.directory || entry.name != name);
        let removed = self.entries.len() != before;
        self.modified |= removed;
        removed
    }

    /// The main document part as UTF-8.
    pub(super) fn xml(&self) -> Result<&str> {
        self.text(DOCUMENT_PART)?
            .context("DOCX has no word/document.xml")
    }

    /// The main document part and its parsed tree.
    pub(super) fn parse(&self) -> Result<(&str, Document)> {
        let xml = self.xml()?;
        Ok((xml, Document::parse(xml)?))
    }

    /// Replace the main document part after validating it.
    pub(super) fn set_document(&mut self, xml: String) -> Result<()> {
        Document::parse(&xml)?;
        self.set(DOCUMENT_PART, xml.into_bytes());
        Ok(())
    }

    /// Declare `part` (no leading slash) with an explicit content type.
    pub(super) fn set_content_type(&mut self, part: &str, kind: &str) -> Result<()> {
        let xml = self
            .text(CONTENT_TYPES_PART)?
            .context("DOCX has no [Content_Types].xml")?;
        let tree = Tree::parse(xml)?;
        let part_name = format!("/{part}");
        let existing = tree.elements(ns::CONTENT_TYPES, "Override").find(|&index| {
            tree.attr(index, None, "PartName")
                .is_some_and(|name| name.eq_ignore_ascii_case(&part_name))
        });
        let element = format!(
            "<Override xmlns=\"{}\" PartName=\"{}\" ContentType=\"{}\"/>",
            ns::CONTENT_TYPES,
            escaped(&part_name),
            escaped(kind)
        );
        let change = match existing {
            Some(index) if tree.attr(index, None, "ContentType") == Some(kind) => return Ok(()),
            Some(index) => (tree.nodes[index].start, tree.nodes[index].end, element),
            None => append_child(xml, &tree.nodes[0], element)?,
        };
        let updated = patch_part(xml, vec![change])?;
        self.set(CONTENT_TYPES_PART, updated.into_bytes());
        Ok(())
    }

    /// Declare a default content type for files with `extension`, unless one exists.
    pub(super) fn set_default_content_type(&mut self, extension: &str, kind: &str) -> Result<()> {
        let xml = self
            .text(CONTENT_TYPES_PART)?
            .context("DOCX has no [Content_Types].xml")?;
        let tree = Tree::parse(xml)?;
        if tree.elements(ns::CONTENT_TYPES, "Default").any(|index| {
            tree.attr(index, None, "Extension")
                .is_some_and(|known| known.eq_ignore_ascii_case(extension))
        }) {
            return Ok(());
        }
        let element = format!(
            "<Default xmlns=\"{}\" Extension=\"{}\" ContentType=\"{}\"/>",
            ns::CONTENT_TYPES,
            escaped(extension),
            escaped(kind)
        );
        let updated = patch_part(xml, vec![append_child(xml, &tree.nodes[0], element)?])?;
        self.set(CONTENT_TYPES_PART, updated.into_bytes());
        Ok(())
    }

    /// The relationships of `source` (`""` for the package root).
    pub(super) fn relationships(&self, source: &str) -> Result<Vec<Relationship>> {
        let Some(xml) = self.text(&relationships_part(source))? else {
            return Ok(Vec::new());
        };
        let tree = Tree::parse(xml)?;
        Ok(tree
            .elements(ns::PACKAGE_RELATIONSHIPS, "Relationship")
            .map(|index| Relationship {
                id: tree.attr(index, None, "Id").unwrap_or_default().to_owned(),
                kind: tree
                    .attr(index, None, "Type")
                    .unwrap_or_default()
                    .to_owned(),
                target: tree
                    .attr(index, None, "Target")
                    .unwrap_or_default()
                    .to_owned(),
                external: tree
                    .attr(index, None, "TargetMode")
                    .is_some_and(|mode| mode.eq_ignore_ascii_case("External")),
            })
            .collect())
    }

    /// Add a relationship from `source` and return its new id.
    pub(super) fn add_relationship(
        &mut self,
        source: &str,
        kind: &str,
        target: &str,
        external: bool,
    ) -> Result<String> {
        let part = relationships_part(source);
        let existing = self.relationships(source)?;
        let next = existing
            .iter()
            .filter_map(|relationship| relationship.id.strip_prefix("rId")?.parse::<u32>().ok())
            .max()
            .unwrap_or(0)
            + 1;
        let id = format!("rId{next}");
        let element = format!(
            "<Relationship xmlns=\"{}\" Id=\"{id}\" Type=\"{}\" Target=\"{}\"{}/>",
            ns::PACKAGE_RELATIONSHIPS,
            escaped(kind),
            escaped(target),
            if external {
                " TargetMode=\"External\""
            } else {
                ""
            }
        );
        let updated = if let Some(xml) = self.text(&part)? {
            let tree = Tree::parse(xml)?;
            patch_part(xml, vec![append_child(xml, &tree.nodes[0], element)?])?
        } else {
            self.set_default_content_type("rels", content_type::RELATIONSHIPS)?;
            format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{}\">{element}</Relationships>",
                ns::PACKAGE_RELATIONSHIPS
            )
        };
        self.set(&part, updated.into_bytes());
        Ok(id)
    }

    /// Remove relationship `id` from `source`; returns whether it existed.
    pub(super) fn remove_relationship(&mut self, source: &str, id: &str) -> Result<bool> {
        let part = relationships_part(source);
        let Some(xml) = self.text(&part)? else {
            return Ok(false);
        };
        let tree = Tree::parse(xml)?;
        let Some(index) = tree
            .elements(ns::PACKAGE_RELATIONSHIPS, "Relationship")
            .find(|&index| tree.attr(index, None, "Id") == Some(id))
        else {
            return Ok(false);
        };
        let node = &tree.nodes[index];
        let updated = patch_part(xml, vec![(node.start, node.end, String::new())])?;
        self.set(&part, updated.into_bytes());
        Ok(true)
    }

    /// The part that the main document relates to with relationship type `kind`.
    pub(super) fn document_part(&self, kind: &str) -> Result<Option<String>> {
        Ok(self
            .relationships(DOCUMENT_PART)?
            .into_iter()
            .find(|relationship| relationship.kind == kind && !relationship.external)
            .map(|relationship| resolve(DOCUMENT_PART, &relationship.target)))
    }

    /// The part related to the main document by `kind`, created from `initial` at
    /// `name` (with content type `content`) when it does not exist yet.
    pub(super) fn ensure_document_part(
        &mut self,
        kind: &str,
        name: &str,
        content: &str,
        initial: &str,
    ) -> Result<String> {
        if let Some(existing) = self.document_part(kind)?
            && self.part(&existing).is_some()
        {
            return Ok(existing);
        }
        Tree::parse(initial)?;
        self.set(name, initial.as_bytes().to_vec());
        self.set_content_type(name, content)?;
        let target = relative_target(DOCUMENT_PART, name);
        self.add_relationship(DOCUMENT_PART, kind, &target, false)?;
        Ok(name.to_owned())
    }

    /// Encode the package as a DOCX.
    pub(super) fn encode(&self) -> Result<Vec<u8>> {
        encode_entries(&self.entries)
    }

    /// Save over the file this package was loaded from, keeping a backup of the
    /// original, and return the backup's path.
    pub(super) fn save(&self, path: &Path) -> Result<Option<PathBuf>> {
        write_atomic(path, &self.encode()?, Some(&self.original), true)
    }
}

/// The `.rels` part holding relationships of `source` (`""` for the package root).
pub(super) fn relationships_part(source: &str) -> String {
    match source.rsplit_once('/') {
        Some((directory, file)) => format!("{directory}/_rels/{file}.rels"),
        None if source.is_empty() => "_rels/.rels".to_owned(),
        None => format!("_rels/{source}.rels"),
    }
}

/// The part name that `target`, relative to `source`, refers to.
pub(super) fn resolve(source: &str, target: &str) -> String {
    if let Some(absolute) = target.strip_prefix('/') {
        return absolute.to_owned();
    }
    let mut segments: Vec<&str> = source.split('/').collect();
    segments.pop();
    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            segment => segments.push(segment),
        }
    }
    segments.join("/")
}

/// A relationship target from `source` to part `name`.
pub(super) fn relative_target(source: &str, name: &str) -> String {
    let directory = source
        .rsplit_once('/')
        .map_or("", |(directory, _)| directory);
    if directory.is_empty() {
        return name.to_owned();
    }
    name.strip_prefix(&format!("{directory}/"))
        .map_or_else(|| format!("/{name}"), ToOwned::to_owned)
}

/// Encode entries as a Deflate-compressed ZIP in their given order.
pub(super) fn encode_entries(entries: &[Entry]) -> Result<Vec<u8>> {
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

/// Refuse to write while Word's owner lock file (`~$name`, or `~$` plus the name
/// without its first two characters) sits beside the document.
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

/// Write `bytes` to `path` through a synced temporary file.
///
/// With `original`, the file must still hold exactly those bytes before and after a
/// backup of them is written beside it; the backup path is returned.
pub(super) fn write_atomic(
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

/// Read the package at `path` for inspection.
pub(super) fn read(path: &Path) -> Result<Package> {
    Package::open(path)
}

/// Open the package at `path`, let `change` edit it, and save it (with a backup) when
/// anything changed. The JSON object `change` returns gains `path`, `modified`, and
/// `backup_path`.
pub(super) fn edit(
    path: &Path,
    change: impl FnOnce(&mut Package) -> Result<Value>,
) -> Result<Value> {
    let mut package = Package::open(path)?;
    let mut result = change(&mut package)?;
    let modified = package.modified();
    let backup = if modified { package.save(path)? } else { None };
    let object = result
        .as_object_mut()
        .context("edit result must be a JSON object")?;
    object.insert("path".into(), json!(path));
    object.insert("modified".into(), json!(modified));
    object.insert("backup_path".into(), json!(backup));
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn optimistic_conflict_leaves_newer_content() -> Result<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("changed.docx");
        fs::write(&path, b"new content")?;
        assert!(write_atomic(&path, b"edit", Some(b"old content"), true).is_err());
        assert_eq!(fs::read(&path)?, b"new content");
        Ok(())
    }

    #[test]
    fn relationship_targets_resolve_relative_to_their_source() {
        assert_eq!(
            relationships_part("word/document.xml"),
            "word/_rels/document.xml.rels"
        );
        assert_eq!(relationships_part(""), "_rels/.rels");
        assert_eq!(
            resolve("word/document.xml", "media/a.png"),
            "word/media/a.png"
        );
        assert_eq!(
            resolve("word/document.xml", "../customXml/x.xml"),
            "customXml/x.xml"
        );
        assert_eq!(
            resolve("word/document.xml", "/word/comments.xml"),
            "word/comments.xml"
        );
        assert_eq!(
            relative_target("word/document.xml", "word/comments.xml"),
            "comments.xml"
        );
        assert_eq!(
            relative_target("word/document.xml", "docProps/core.xml"),
            "/docProps/core.xml"
        );
    }
}
