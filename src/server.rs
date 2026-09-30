//! MCP tool discovery, serialized dispatch, and error conversion.
use std::sync::{Arc, Mutex};

use anyhow::Context;
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

/// Operations are serialized so concurrent MCP requests cannot overwrite each other.
#[derive(Clone)]
pub struct WordServer {
    tools: Arc<Vec<Tool>>,
    backend: Arc<Mutex<LiveBackend>>,
}

impl WordServer {
    /// Construct the server without launching or attaching to Word.
    ///
    /// # Errors
    /// Returns an error if a built-in tool definition is invalid.
    pub fn new() -> anyhow::Result<Self> {
        let tools = docx::tools()
            .into_iter()
            .chain(crate::live::tools())
            .map(serde_json::from_value)
            .collect::<Result<Vec<Tool>, _>>()
            .context("invalid built-in tool definition")?;
        Ok(Self {
            tools: Arc::new(tools),
            backend: Arc::new(Mutex::new(LiveBackend::new())),
        })
    }

    /// Return the complete MCP tool catalog.
    #[must_use]
    pub fn tool_definitions(&self) -> &[Tool] {
        &self.tools
    }

    /// Blocking API used by the MCP adapter; document errors are tool errors.
    ///
    /// # Errors
    /// Returns an error for unknown tools, invalid inputs, document failures,
    /// or an unavailable backend. A failed live edit may have partially applied;
    /// its error message explains when inspection or undo is needed.
    pub fn execute(&self, name: &str, arguments: Value) -> anyhow::Result<Value> {
        anyhow::ensure!(
            self.tools.iter().any(|tool| tool.name == name),
            "unknown tool: {name}"
        );
        anyhow::ensure!(
            arguments.is_object(),
            "tool arguments must be a JSON object"
        );
        let backend = self
            .backend
            .lock()
            .map_err(|_| anyhow::anyhow!("document worker lock poisoned; restart the server"))?;
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
            .with_server_info(Implementation::new(env!("CARGO_PKG_NAME"), env!("CARGO_PKG_VERSION")))
            .with_instructions("Read and edit saved .docx files with the document tools. Use word_live_* for documents open in Microsoft Word on Windows, including unsaved changes. Saved-file edits create backups and must not be used on open documents. preview_document produces a logical HTML preview; word_live_view shows actual Word layout and word_live_export_pdf exports it. Read tool schemas before calling: paragraph indices are zero-based, whereas live range offsets use Word coordinates. Live edits are not saved until word_live_save. Tools access local files with the current user's permissions.")
    }

    fn get_tool(&self, name: &str) -> Option<Tool> {
        self.tools.iter().find(|tool| tool.name == name).cloned()
    }

    fn list_tools(
        &self,
        request: Option<PaginatedRequestParams>,
        _context: RequestContext<RoleServer>,
    ) -> impl Future<Output = Result<ListToolsResult, ErrorData>> + Send + '_ {
        if request.and_then(|params| params.cursor).is_some() {
            return std::future::ready(Err(ErrorData::invalid_params(
                "This server does not use tool pagination",
                None,
            )));
        }
        std::future::ready(Ok(ListToolsResult {
            tools: self.tools.as_ref().clone(),
            ..Default::default()
        }))
    }

    async fn call_tool(
        &self,
        request: CallToolRequestParams,
        _context: RequestContext<RoleServer>,
    ) -> Result<CallToolResponse, ErrorData> {
        if self.get_tool(&request.name).is_none() {
            return Err(ErrorData::invalid_params(
                format!("Unknown tool: {}", request.name),
                None,
            ));
        }
        let server = self.clone();
        let arguments = Value::Object(request.arguments.unwrap_or_default());
        let result =
            tokio::task::spawn_blocking(move || server.execute(&request.name, arguments)).await;
        let result = match result {
            Ok(Ok(value)) => CallToolResult::structured(value),
            Ok(Err(error)) => {
                CallToolResult::structured_error(json!({"error": format!("{error:#}")}))
            }
            Err(error) => CallToolResult::structured_error(
                json!({"error": format!("document worker failed: {error}")}),
            ),
        };
        Ok(result.into())
    }
}
