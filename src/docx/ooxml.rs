#![expect(
    dead_code,
    reason = "SCAFFOLD: constants for areas under construction; remove once all are used"
)]
//! Namespaces, relationship types, and content types of the parts word-mcp reads or
//! writes.

/// Namespace URIs.
pub(super) mod ns {
    /// Office relationships (`r:id`, `r:embed`).
    pub(in crate::docx) const R: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    /// Package relationships (`.rels` parts).
    pub(in crate::docx) const PACKAGE_RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    /// `[Content_Types].xml`.
    pub(in crate::docx) const CONTENT_TYPES: &str =
        "http://schemas.openxmlformats.org/package/2006/content-types";
    /// Word 2010 extensions, such as `w14:paraId`.
    pub(in crate::docx) const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
    /// Word 2012 extensions (`w15:done`, `commentsEx`).
    pub(in crate::docx) const W15: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
    /// `DrawingML` word-processing drawing (`wp:inline`).
    pub(in crate::docx) const WP: &str =
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    /// `DrawingML` main (`a:graphic`).
    pub(in crate::docx) const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    /// `DrawingML` picture (`pic:pic`).
    pub(in crate::docx) const PIC: &str =
        "http://schemas.openxmlformats.org/drawingml/2006/picture";
    /// Core properties.
    pub(in crate::docx) const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    /// Dublin Core elements.
    pub(in crate::docx) const DC: &str = "http://purl.org/dc/elements/1.1/";
    /// Dublin Core terms.
    pub(in crate::docx) const DC_TERMS: &str = "http://purl.org/dc/terms/";
    /// XML Schema instance.
    pub(in crate::docx) const XSI: &str = "http://www.w3.org/2001/XMLSchema-instance";
}

/// Relationship types.
pub(super) mod rel {
    pub(in crate::docx) const OFFICE_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    pub(in crate::docx) const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub(in crate::docx) const NUMBERING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    pub(in crate::docx) const SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    pub(in crate::docx) const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub(in crate::docx) const FOOTNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    pub(in crate::docx) const ENDNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    pub(in crate::docx) const HEADER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    pub(in crate::docx) const FOOTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    pub(in crate::docx) const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub(in crate::docx) const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    /// Word 2012 comment extensions (resolved state).
    pub(in crate::docx) const COMMENTS_EXTENDED: &str =
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
    /// Core properties, referenced from the package root.
    pub(in crate::docx) const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
}

/// Content types.
pub(super) mod content_type {
    pub(in crate::docx) const DOCUMENT: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    pub(in crate::docx) const STYLES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    pub(in crate::docx) const NUMBERING: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
    pub(in crate::docx) const SETTINGS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    pub(in crate::docx) const COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    pub(in crate::docx) const COMMENTS_EXTENDED: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
    pub(in crate::docx) const FOOTNOTES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
    pub(in crate::docx) const ENDNOTES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
    pub(in crate::docx) const HEADER: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    pub(in crate::docx) const FOOTER: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
    pub(in crate::docx) const RELATIONSHIPS: &str =
        "application/vnd.openxmlformats-package.relationships+xml";
    pub(in crate::docx) const CORE_PROPERTIES: &str =
        "application/vnd.openxmlformats-package.core-properties+xml";
}
