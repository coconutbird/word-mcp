//! DOCX ZIP packages: bounded loading, re-encoding and safe atomic saves.

use std::{
    fs,
    io::{Cursor, Read, Write},
    path::{Path, PathBuf},
};

use anyhow::{Context, Result, ensure};
use serde_json::{Value, json};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

use super::xml::{Document, MAX_XML};

const MAX_ARCHIVE: u64 = 64 * 1024 * 1024;
const MAX_ENTRIES: usize = 4096;
/// Name of the main document part.
pub(super) const DOCUMENT_PART: &str = "word/document.xml";

/// One ZIP entry, kept byte-for-byte.
pub(super) struct Entry {
    pub(super) name: String,
    pub(super) bytes: Vec<u8>,
    pub(super) directory: bool,
}

/// A loaded DOCX package and the exact bytes it was read from.
pub(super) struct Package {
    /// The file contents at load time, for backups and conflict detection.
    pub(super) original: Vec<u8>,
    pub(super) entries: Vec<Entry>,
    /// Index of the main document part in `entries`.
    document: usize,
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
        let mut document = None;
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
            if file.name() == DOCUMENT_PART {
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

    /// The main document part as UTF-8.
    pub(super) fn xml(&self) -> Result<&str> {
        let bytes = &self.entries[self.document].bytes;
        ensure!(bytes.len() <= MAX_XML, "document XML exceeds 16 MiB limit");
        std::str::from_utf8(bytes).context("document XML must be UTF-8")
    }

    /// The main document part and its parsed tree.
    pub(super) fn parse(&self) -> Result<(&str, Document)> {
        let xml = self.xml()?;
        Ok((xml, Document::parse(xml)?))
    }
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

/// Replace the main document part and save with a backup of the original.
pub(super) fn save_edit(path: &Path, mut package: Package, xml: String) -> Result<Value> {
    package.entries[package.document].bytes = xml.into_bytes();
    let bytes = encode_entries(&package.entries)?;
    let backup = write_atomic(path, &bytes, Some(&package.original), true)?;
    Ok(json!({"path":path,"backup_path":backup,"modified":true}))
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
}
