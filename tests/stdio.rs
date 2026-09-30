//! Exercise the built executable over the same stdio transport used by MCP hosts.
use std::{
    io::{BufRead, BufReader, Write},
    process::{Child, ChildStdin, Command, Stdio},
    sync::mpsc::{self, Receiver},
    time::Duration,
};

use serde_json::{Value, json};

struct Client {
    child: Child,
    stdin: ChildStdin,
    lines: Receiver<String>,
    id: u64,
}

impl Client {
    fn new() -> Self {
        let mut child = Command::new(env!("CARGO_BIN_EXE_word-mcp"))
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::inherit())
            .spawn()
            .unwrap();
        let stdin = child.stdin.take().unwrap();
        let stdout = child.stdout.take().unwrap();
        let (tx, lines) = mpsc::channel();
        std::thread::spawn(move || {
            for line in BufReader::new(stdout).lines() {
                if tx.send(line.unwrap()).is_err() {
                    break;
                }
            }
        });
        let mut client = Self {
            child,
            stdin,
            lines,
            id: 0,
        };
        let response = client.request(
            "initialize",
            &json!({
                "protocolVersion":"2025-06-18", "capabilities":{},
                "clientInfo":{"name":"word-mcp-integration","version":"1.0"}
            }),
        );
        assert_eq!(response["result"]["serverInfo"]["name"], "word-mcp");
        assert!(response["result"]["capabilities"]["tools"].is_object());
        client.send(&json!({"jsonrpc":"2.0","method":"notifications/initialized"}));
        client
    }

    fn send(&mut self, value: &Value) {
        writeln!(self.stdin, "{value}").unwrap();
        self.stdin.flush().unwrap();
    }

    fn request(&mut self, method: &str, params: &Value) -> Value {
        self.id += 1;
        self.send(&json!({"jsonrpc":"2.0","id":self.id,"method":method,"params":params}));
        loop {
            let line = self
                .lines
                .recv_timeout(Duration::from_secs(30))
                .expect("MCP server did not reply within 30s");
            let response: Value =
                serde_json::from_str(&line).expect("stdout must contain only JSON-RPC");
            if response["id"] == self.id {
                return response;
            }
        }
    }

    fn tool(&mut self, name: &str, arguments: &Value) -> Value {
        self.request("tools/call", &json!({"name":name,"arguments":arguments}))
    }
}

impl Drop for Client {
    fn drop(&mut self) {
        let _ = self.child.kill();
        let _ = self.child.wait();
    }
}

#[test]
fn mcp_initialization_discovery_and_error_recovery() {
    let mut client = Client::new();
    let discovery = client.request("tools/list", &json!({}));
    let tools = discovery["result"]["tools"].as_array().unwrap();
    for name in [
        "create_document",
        "read_document",
        "replace_text",
        "preview_document",
        "word_live_read",
        "word_live_view",
        "word_live_export_pdf",
    ] {
        assert!(
            tools.iter().any(|tool| tool["name"] == name),
            "missing tool {name}"
        );
    }
    let unknown = client.tool("not_a_tool", &json!({}));
    assert_eq!(unknown["error"]["code"], -32602);
    let invalid = client.tool("read_document", &json!({"not_a_parameter":true}));
    assert_eq!(invalid["result"]["isError"], true);
    assert!(invalid["result"]["content"][0]["text"].is_string());
    let ping = client.request("ping", &json!({}));
    assert!(
        ping["result"].is_object(),
        "server must survive a failed tool call"
    );
}

#[test]
fn command_line_tool_catalog_is_valid_json() {
    let output = Command::new(env!("CARGO_BIN_EXE_word-mcp"))
        .arg("--list-tools")
        .output()
        .unwrap();
    assert!(output.status.success());
    let tools: Vec<Value> = serde_json::from_slice(&output.stdout).unwrap();
    let mut names = std::collections::HashSet::new();
    for tool in tools {
        assert!(names.insert(tool["name"].as_str().unwrap().to_owned()));
        assert_eq!(tool["inputSchema"]["type"], "object");
        assert_eq!(tool["inputSchema"]["additionalProperties"], false);
    }
}

fn successful_content(response: &Value) -> &Value {
    assert!(
        response.get("error").is_none(),
        "protocol error: {response}"
    );
    assert_ne!(
        response["result"]["isError"], true,
        "tool error: {response}"
    );
    &response["result"]["structuredContent"]
}

#[test]
fn saved_document_workflow_over_mcp() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("report.docx");
    let preview = directory.path().join("report.html");
    let mut client = Client::new();
    successful_content(&client.tool(
        "create_document",
        &json!({
            "path":path,"paragraphs":["Draft 🦀 report","Revenue <forecast> & actual"]
        }),
    ));
    assert!(path.is_file());
    let refusal = client.tool("create_document", &json!({"path":path}));
    assert_eq!(
        refusal["result"]["isError"], true,
        "creation must not overwrite by default"
    );
    let read = client.tool("read_document", &json!({"path":path}));
    assert!(
        successful_content(&read)
            .to_string()
            .contains("Draft 🦀 report")
    );
    successful_content(&client.tool(
        "replace_text",
        &json!({
            "path":path,"find":"Draft 🦀","replacement":"Final 🦀"
        }),
    ));
    successful_content(&client.tool(
        "insert_paragraph",
        &json!({
            "path":path,"index":1,"text":"Review complete"
        }),
    ));
    successful_content(&client.tool(
        "format_paragraph",
        &json!({
            "path":path,"index":0,"bold":true,"font_size_pt":18,"alignment":"center"
        }),
    ));
    let read = client.tool("read_document", &json!({"path":path}));
    let text = successful_content(&read).to_string();
    assert!(text.contains("Final 🦀 report"));
    assert!(text.contains("Review complete"));
    assert!(!text.contains("Draft 🦀 report"));
    successful_content(&client.tool("get_document_info", &json!({"path":path})));
    successful_content(&client.tool(
        "preview_document",
        &json!({"path":path,"output_path":preview}),
    ));
    let html = std::fs::read_to_string(&preview).unwrap();
    assert!(html.contains("Final 🦀 report"));
    assert!(html.contains("&lt;forecast&gt;"));
    let invalid = client.tool(
        "replace_text",
        &json!({"path":path,"find":"","replacement":"invalid"}),
    );
    assert_eq!(invalid["result"]["isError"], true);
}
