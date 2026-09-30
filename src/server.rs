//! MCP tool discovery, serialized dispatch, and error conversion.
use std::sync::{Arc, Mutex};

use rmcp::{
    ErrorData, RoleServer, ServerHandler,
    model::{
        CallToolRequestParams, CallToolResponse, CallToolResult, Implementation, ListToolsResult,
        PaginatedRequestParams, ServerCapabilities, ServerConfig, Tool,
    },
    service::RequestContext,
};
use serde_json::{Value, json};

use crate::{docx, live::LiveBackend};

const INSTRUCTIONS: &str = "Saved-file tools (create_document, read_document, replace_text, insert_paragraph, \
format_paragraph, delete_paragraph, get_document_info, preview_document) edit .docx files on disk without Word; \
they create backups and refuse documents that are open in Word. word_live_* tools drive desktop Microsoft Word \
on Windows: call word_live_open first, then read, edit, and save through Word, including unsaved changes. \
Paragraph indices of saved-file tools are zero-based; live tools use Word's UTF-16 range positions \
(see word_live_paragraphs). Live edits stay unsaved until word_live_save. preview_document gives a logical HTML \
preview; word_live_view and word_live_export_pdf show the real layout. Tools run with the current user's \
file permissions.";

/// The MCP server. Operations run one at a time so concurrent requests cannot
/// interleave edits to the same document.
#[derive(Clone)]
pub struct WordServer {
    tools: Arc<[Tool]>,
    backend: Arc<Mutex<LiveBackend>>,
}

impl Default for WordServer {
    fn default() -> Self {
        Self::new()
    }
}

impl WordServer {
    /// Construct the server without starting Word.
    #[must_use]
    pub fn new() -> Self {
        Self {
            tools: docx::tools()
                .into_iter()
                .chain(crate::live::tools())
                .collect(),
            backend: Arc::new(Mutex::new(LiveBackend::new())),
        }
    }

    /// The complete tool catalog.
    #[must_use]
    pub fn tool_definitions(&self) -> &[Tool] {
        &self.tools
    }

    /// Run one tool call to completion on the current thread.
    ///
    /// # Errors
    /// Returns an error for unknown tools, invalid arguments, document failures, or an
    /// unavailable Word. A failed live edit may be partially applied; its message says
    /// how to inspect or undo it.
    pub fn execute(&self, name: &str, arguments: Value) -> anyhow::Result<Value> {
        anyhow::ensure!(
            self.tools.iter().any(|tool| tool.name == name),
            "unknown tool: {name}"
        );
        // A panic while holding the lock leaves no partial state in the guard itself.
        let mut backend = self
            .backend
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if name.starts_with("word_live_") {
            backend.call(name, arguments)
        } else {
            docx::call(name, arguments)
        }
    }
}

impl ServerHandler for WordServer {
    fn get_info(&self) -> ServerConfig {
        ServerConfig::new(ServerCapabilities::builder().enable_tools().build())
            .with_server_info(Implementation::new(
                env!("CARGO_PKG_NAME"),
                env!("CARGO_PKG_VERSION"),
            ))
            .with_instructions(INSTRUCTIONS)
    }

    fn get_tool(&self, name: &str) -> Option<Tool> {
        self.tools.iter().find(|tool| tool.name == name).cloned()
    }

    fn list_tools(
        &self,
        request: Option<PaginatedRequestParams>,
        _context: RequestContext<RoleServer>,
    ) -> impl Future<Output = Result<ListToolsResult, ErrorData>> + Send + '_ {
        std::future::ready(if request.and_then(|params| params.cursor).is_some() {
            Err(ErrorData::invalid_params(
                "this server does not paginate tools",
                None,
            ))
        } else {
            Ok(ListToolsResult {
                tools: self.tools.to_vec(),
                ..Default::default()
            })
        })
    }

    async fn call_tool(
        &self,
        request: CallToolRequestParams,
        _context: RequestContext<RoleServer>,
    ) -> Result<CallToolResponse, ErrorData> {
        if self.get_tool(&request.name).is_none() {
            return Err(ErrorData::invalid_params(
                format!("unknown tool: {}", request.name),
                None,
            ));
        }
        let server = self.clone();
        let arguments = Value::Object(request.arguments.unwrap_or_default());
        let outcome =
            tokio::task::spawn_blocking(move || server.execute(&request.name, arguments)).await;
        let result = match outcome {
            Ok(Ok(value)) => CallToolResult::structured(value),
            Ok(Err(error)) => {
                CallToolResult::structured_error(json!({"error": format!("{error:#}")}))
            }
            Err(error) => CallToolResult::structured_error(
                json!({"error": format!("the tool panicked: {error}")}),
            ),
        };
        Ok(result.into())
    }
}
