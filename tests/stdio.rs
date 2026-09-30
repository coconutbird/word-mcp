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
    for area in [
        "document",
        "read",
        "edit",
        "format",
        "table",
        "review",
        "layout",
        "references",
        "media",
        "controls",
    ] {
        for name in [format!("docx_{area}"), format!("word_live_{area}")] {
            assert!(
                tools.iter().any(|tool| tool["name"] == name.as_str()),
                "missing tool {name}"
            );
        }
    }
    let unknown = client.tool("not_a_tool", &json!({}));
    assert_eq!(unknown["error"]["code"], -32602);
    let invalid = client.tool("docx_read", &json!({"not_a_parameter":true}));
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
        let schema = &tool["inputSchema"];
        assert_eq!(schema["type"], "object");
        assert_eq!(schema["additionalProperties"], false);
        // Clients such as the Anthropic API reject combinators at the schema root.
        for keyword in ["oneOf", "anyOf", "allOf"] {
            assert!(
                schema.get(keyword).is_none(),
                "{} has a root {keyword}",
                tool["name"]
            );
        }
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
    let mut call = |tool: &str, operation: Value| {
        client.tool(tool, &json!({"path": path, "operation": operation}))
    };
    successful_content(&call(
        "docx_document",
        json!({"action": "create", "paragraphs": ["Draft 🦀 report", "Revenue <forecast> & actual"]}),
    ));
    assert!(path.is_file());
    let refusal = call("docx_document", json!({"action": "create"}));
    assert_eq!(
        refusal["result"]["isError"], true,
        "creation must not overwrite by default"
    );
    let read = call("docx_read", json!({"action": "paragraphs"}));
    assert!(
        successful_content(&read)
            .to_string()
            .contains("Draft 🦀 report")
    );
    successful_content(&call(
        "docx_edit",
        json!({"action": "replace", "find": "Draft 🦀", "replacement": "Final 🦀"}),
    ));
    successful_content(&call(
        "docx_edit",
        json!({"action": "insert_paragraph", "index": 1, "text": "Review complete"}),
    ));
    successful_content(&call(
        "docx_format",
        json!({"action": "paragraph", "index": 0, "bold": true, "font_size_pt": 18, "alignment": "center"}),
    ));
    let read = call("docx_read", json!({"action": "paragraphs"}));
    let text = successful_content(&read).to_string();
    assert!(text.contains("Final 🦀 report"));
    assert!(text.contains("Review complete"));
    assert!(!text.contains("Draft 🦀 report"));
    successful_content(&call("docx_document", json!({"action": "info"})));
    successful_content(&call(
        "docx_read",
        json!({"action": "preview", "output_path": preview}),
    ));
    let html = std::fs::read_to_string(&preview).unwrap();
    assert!(html.contains("Final 🦀 report"));
    assert!(html.contains("&lt;forecast&gt;"));
    let invalid = call(
        "docx_edit",
        json!({"action": "replace", "find": "", "replacement": "invalid"}),
    );
    assert_eq!(invalid["result"]["isError"], true);
}
