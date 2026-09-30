//! Command-line entry point for the Word MCP stdio server.
use anyhow::{Result, bail};
use rmcp::ServiceExt;
use word_mcp::server::WordServer;

#[tokio::main]
async fn main() -> Result<()> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.as_slice() {
        [] => {}
        [arg] if arg == "--help" || arg == "-h" => {
            println!(
                "word-mcp {}\n\nLocal Word MCP server (Rust 2024).\n\nUSAGE:\n  word-mcp               Run MCP over stdio\n  word-mcp --list-tools  Print tool definitions as JSON\n  word-mcp --version    Print version\n\nSaved DOCX tools run without Word. Live tools require Windows and desktop Microsoft Word.\nLogs go to stderr; stdout is reserved for MCP. Set RUST_LOG for diagnostics.",
                env!("CARGO_PKG_VERSION")
            );
            return Ok(());
        }
        [arg] if arg == "--version" || arg == "-V" => {
            println!("word-mcp {}", env!("CARGO_PKG_VERSION"));
            return Ok(());
        }
        [arg] if arg == "--list-tools" => {
            println!(
                "{}",
                serde_json::to_string_pretty(WordServer::new().tool_definitions())?
            );
            return Ok(());
        }
        _ => bail!("unknown arguments; use --help"),
    }
    tracing_subscriber::fmt()
        .with_env_filter(
            tracing_subscriber::EnvFilter::try_from_default_env()
                .unwrap_or_else(|_| "word_mcp=info".into()),
        )
        .with_writer(std::io::stderr)
        .with_ansi(false)
        .init();
    let service = WordServer::new().serve(rmcp::transport::stdio()).await?;
    service.waiting().await?;
    Ok(())
}
