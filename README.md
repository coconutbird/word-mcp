# word-mcp

A local **Rust 2024** MCP server for reading, editing, and viewing Word documents. Inspired by [word-mcp-live](https://github.com/ykarapazar/word-mcp-live), with two separate backends:

- **Saved DOCX:** create, read, replace text, insert paragraphs, format paragraphs, inspect metadata, and generate an HTML preview without installing Word.
- **Live Word on Windows:** open documents, read unsaved text, edit with optional tracked changes, undo, save, show the Word window, and export PDF through Word's layout engine.

This is an initial core implementation, not a port of every reference-project tool. The server uses the official Rust MCP SDK (`rmcp`) over stdio. It does not require Python, PowerShell automation, or a listening HTTP server.

## Build

Requires Rust 1.88 or later. Live tools additionally require Windows and desktop Microsoft Word.

The Windows COM backend uses [`cppvtable-com`](https://github.com/coconutbird/cppvtable) with its `windows-compat` feature, fetched from Git. `Cargo.lock` pins the exact commit; run `cargo update -p cppvtable-com` to move to a newer one.

```powershell
cargo build --release --locked
.\target\release\word-mcp.exe --help
.\target\release\word-mcp.exe --list-tools
```

## Connect an MCP client

Use the full path to the executable in your client's MCP server configuration:

```json
{
  "mcpServers": {
    "word": {
      "command": "C:/Users/dev/Documents/Git/coconutbird/word-mcp/target/release/word-mcp.exe",
      "args": []
    }
  }
}
```

Run the executable without arguments to start its stdio transport. Standard output contains only MCP messages; diagnostics go to standard error. Set `RUST_LOG=word_mcp=debug` for diagnostic logging.

## Tools

| Saved DOCX tool | Purpose |
| --- | --- |
| `create_document` | Create a DOCX document |
| `read_document` | Read main-document paragraphs, including table paragraphs |
| `get_document_info` | Inspect document structure and package metadata |
| `replace_text` | Replace literal text, including matches split across text runs |
| `insert_paragraph` | Insert a paragraph at a document position |
| `format_paragraph` | Set paragraph formatting |
| `preview_document` | Generate a logical HTML preview |

| Windows live tool | Purpose |
| --- | --- |
| `word_live_status` | Inspect Word availability and open documents without launching it |
| `word_live_open` | Open an existing document, optionally showing Word |
| `word_live_read` | Read the current main-story text, including unsaved edits |
| `word_live_replace_text` | Replace literal text within one undo record |
| `word_live_insert_text` | Insert text within one undo record |
| `word_live_save` | Save the specified open document |
| `word_live_export_pdf` | Export current layout to PDF |
| `word_live_undo` | Undo actions in the specified document |
| `word_live_view` | Show Word and activate the specified document |

Tool input schemas are the authoritative parameter reference. Inspect them with `--list-tools` or MCP `tools/list`.

## Saved-file example

```text
create_document {"path":"C:/Documents/report.docx","paragraphs":["Draft report","Revenue increased."]}
read_document {"path":"C:/Documents/report.docx"}
replace_text {"path":"C:/Documents/report.docx","find":"Draft","replacement":"Final","replace_all":true}
insert_paragraph {"path":"C:/Documents/report.docx","index":1,"text":"Reviewed and approved."}
format_paragraph {"path":"C:/Documents/report.docx","index":0,"bold":true,"font_size_pt":18}
preview_document {"path":"C:/Documents/report.docx","output_path":"C:/Documents/report.html"}
```

Creation and preview output require `overwrite:true` to replace existing files. `read_document` accepts `start` and `limit` for pagination. `replace_text` can restrict matching to `paragraph_index`. `format_paragraph` supports bold, italic, underline, font size/family, style, and alignment.

## Live editing example

Create or select a document on disk, then call:

```text
word_live_open {"path":"C:/Documents/report.docx"}
word_live_read {"path":"C:/Documents/report.docx"}
word_live_replace_text {"path":"C:/Documents/report.docx","find":"Draft","replacement":"Final","all":true,"tracked_changes":true}
word_live_view {"path":"C:/Documents/report.docx"}
word_live_export_pdf {"path":"C:/Documents/report.docx","output_path":"C:/Documents/report.pdf"}
word_live_save {"path":"C:/Documents/report.docx"}
```

Live tools target an exact absolute file path, never an implicitly selected document. Open the document with `word_live_open` first. Edits remain unsaved until `word_live_save`. Word range positions use UTF-16 coordinates and an exclusive end, not UTF-8 byte offsets. Live text operations cover the main story; headers, footers, and text boxes have separate Word stories and are outside this initial tool set.

Live replacement uses Word's native Find ranges so fields do not invalidate text positions. Searches are literal and case-sensitive, with Word's limit of 255 UTF-16 code units after escaping caret and control characters. A reused Word window is not hidden by `visible:false`; the result reports its actual visibility.

## Viewing and document preservation

`preview_document` produces an escaped HTML content preview. It is not a pagination or pixel-accurate Word renderer. Use `word_live_view` to see actual layout in Word, or `word_live_export_pdf` to create a layout-faithful PDF. These tools do not return screenshots or rendered page images to the MCP client.

Saved-file tools work on `.docx` packages. They preserve unrelated ZIP entries and edit targeted XML regions. Text, insertion, and formatting edits create backups and use atomic replacement; do not use them on a document currently open in Word. The tools reject detected Word lock files. Creation overwrites require explicit `overwrite:true`. Live tools operate through Word itself and preserve its document model. Offline paragraph indices are zero-based and include paragraphs inside tables in document order.

Offline editing is conservative around complex paragraphs, including fields, revision markup, content controls, and drawings. Unsupported edits return an error before saving. Use the live Word tools for edits that require Word's full document model.

The server executes with your local user permissions and can access files that user can access. It serializes document operations within one server process. Backups are local files and remain until you remove them. Word may display application dialogs that require user interaction. Do not retry a timed-out edit blindly: inspect the document first because a native COM call can finish after a client stops waiting.

## Development

```powershell
cargo fmt --all -- --check
cargo clippy --all-targets --locked -- -D warnings
cargo test --all-targets --locked
```

Tests cover DOCX preservation and editing, input validation, and the executable's MCP initialization, tool discovery, errors, and recovery. Word integration tests are opt-in so a routine test run does not launch or change an interactive Word session.

On a Windows machine with Word installed, run the isolated live integration test with:

```powershell
cargo test --lib -- --ignored --nocapture --test-threads=1
```

These tests create their own Word instances and temporary documents, exercise reading, editing, tracked changes, undo, saving, viewing, and PDF export, and close those instances afterwards. The interoperability test also opens a DOCX created by the Rust backend in Word and reads Word's saved output back through the Rust backend.

The live backend declares `IDispatch` with `cppvtable-com` and uses its `ComPtr` reference counting, plus Windows APIs for COM initialization and automation values. COM objects remain on one dedicated STA thread; requests cross the boundary as owned Rust values.
