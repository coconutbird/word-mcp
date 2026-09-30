//! Validation shared by the live tools; runs on the caller's thread before any work
//! reaches Word.

use std::path::{Path, PathBuf};

use anyhow::{Context, Result, ensure};

/// Word's limit on the length of a search string, in UTF-16 code units after escaping.
pub(super) const FIND_LIMIT: usize = 255;

/// Resolve an existing Word document to its canonical absolute path.
pub(super) fn document_path(path: &Path) -> Result<PathBuf> {
    ensure!(
        path.is_absolute(),
        "path must be absolute: {}",
        path.display()
    );
    ensure!(
        path.is_file(),
        "document does not exist: {}",
        path.display()
    );
    let extension = path
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    ensure!(
        matches!(
            extension.as_str(),
            "docx" | "docm" | "doc" | "dotx" | "dotm" | "rtf"
        ),
        "not a Word document: {}",
        path.display()
    );
    path.canonicalize()
        .with_context(|| format!("cannot resolve {}", path.display()))
}

/// The document path of a tool call that requires one.
pub(super) fn required_document(path: Option<&Path>, action: &str) -> Result<PathBuf> {
    document_path(path.with_context(|| format!("path is required for {action}"))?)
}

/// Validate an output file path: absolute, with one of `extensions`, in an existing
/// directory, and not an existing file unless `overwrite`. Returns it with a
/// canonical parent directory.
pub(super) fn output_path(path: &Path, extensions: &[&str], overwrite: bool) -> Result<PathBuf> {
    ensure!(
        path.is_absolute(),
        "output_path must be absolute: {}",
        path.display()
    );
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extensions
                .iter()
                .any(|allowed| extension.eq_ignore_ascii_case(allowed))),
        "output_path must end in .{}",
        extensions.join(" or .")
    );
    if path.exists() {
        ensure!(path.is_file(), "output_path is not a file");
        ensure!(
            overwrite,
            "{} already exists; pass overwrite=true to replace it",
            path.display()
        );
    }
    let parent = path
        .parent()
        .context("output_path has no parent directory")?;
    ensure!(parent.is_dir(), "output directory does not exist");
    Ok(parent
        .canonicalize()?
        .join(path.file_name().context("output_path has no file name")?))
}

/// Validate a Word range `start..end`; `allow_empty` permits `start == end`.
pub(super) fn check_positions(start: i32, end: i32, allow_empty: bool) -> Result<()> {
    ensure!(
        0 <= start && (start < end || (allow_empty && start == end)),
        "range must satisfy 0 <= start {} end",
        if allow_empty { "<=" } else { "<" }
    );
    Ok(())
}

/// Escape literal text for Word's Find, which treats `^` as a special-code prefix.
pub(super) fn word_find_text(text: &str) -> String {
    text.replace('^', "^^")
        .replace('\r', "^p")
        .replace(['\n', '\u{b}'], "^l")
        .replace('\t', "^t")
}

/// Validate literal search text against Word's Find limits.
pub(super) fn check_find(find: &str) -> Result<()> {
    ensure!(!find.is_empty(), "find must not be empty");
    ensure!(
        word_find_text(find).encode_utf16().count() <= FIND_LIMIT,
        "find exceeds Word's {FIND_LIMIT} UTF-16 unit search limit after escaping"
    );
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn escapes_word_find_special_codes() {
        assert_eq!(word_find_text("literal ^p\r\n\t"), "literal ^^p^p^l^t");
        assert_eq!(word_find_text("left\u{b}right"), "left^lright");
        assert!(check_find(&"^".repeat(128)).is_err());
    }

    #[test]
    fn rejects_relative_and_non_word_paths() {
        assert!(document_path(Path::new("test.docx")).is_err());
        assert!(check_positions(5, 2, true).is_err());
        assert!(check_positions(3, 3, false).is_err());
        assert!(check_positions(3, 3, true).is_ok());
    }
}
