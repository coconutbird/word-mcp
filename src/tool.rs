//! Shared construction of MCP tool definitions, typed argument parsing, and results.
//!
//! Each tool's arguments are one Rust type. Its `JsonSchema` derive produces the
//! advertised input schema, and its `Deserialize` derive enforces the same contract, so
//! the schema and the parser cannot drift apart.
//!
//! Tools are grouped by area. A grouped tool's arguments are an object whose
//! `operation` property is an internally tagged enum (`{"action": "...", ...}`). The
//! schema root therefore stays a plain object, as MCP clients require, while each
//! action's fields are still validated strictly.
use anyhow::{Context, Result};
use rmcp::model::{JsonObject, Tool, ToolAnnotations};
use schemars::JsonSchema;
use serde::de::DeserializeOwned;
use serde_json::Value;

/// How a tool affects documents, reported to clients as MCP annotations.
#[derive(Clone, Copy)]
pub(crate) enum Effect {
    /// Reads only; repeating the call has no further effect.
    ReadOnly,
    /// Creates output or changes presentation without altering existing content.
    Additive,
    /// May modify or replace existing document content.
    Destructive,
}

/// The result of a tool call.
pub(crate) enum Output {
    /// A JSON object returned as structured content.
    Json(Value),
    /// A rendered image plus JSON details about it.
    #[expect(dead_code, reason = "SCAFFOLD: used by areas under construction")]
    Image {
        /// PNG-encoded bytes.
        png: Vec<u8>,
        /// Details such as the page number and pixel size.
        details: Value,
    },
}

impl From<Value> for Output {
    fn from(value: Value) -> Self {
        Self::Json(value)
    }
}

/// Build a tool whose input schema is generated from the argument type `T`.
///
/// # Panics
/// Panics if `T` does not produce an object schema. Argument types are fixed at
/// compile time, and the catalog tests exercise every definition.
pub(crate) fn tool<T: JsonSchema + 'static>(
    name: &'static str,
    description: &'static str,
    effect: Effect,
) -> Tool {
    let annotations = match effect {
        Effect::ReadOnly => ToolAnnotations::new()
            .read_only(true)
            .destructive(false)
            .idempotent(true),
        Effect::Additive => ToolAnnotations::new().read_only(false).destructive(false),
        Effect::Destructive => ToolAnnotations::new().read_only(false).destructive(true),
    };
    Tool::new(name, description, JsonObject::new())
        .with_input_schema::<T>()
        .with_annotations(annotations.open_world(false))
}

/// Deserialize tool arguments into their typed form.
///
/// # Errors
/// Returns an error naming the tool when the arguments violate its schema.
pub(crate) fn parse<T: DeserializeOwned>(name: &str, arguments: Value) -> Result<T> {
    serde_json::from_value(arguments).with_context(|| format!("invalid arguments for {name}"))
}

#[cfg(test)]
mod tests {
    use schemars::JsonSchema;
    use serde::Deserialize;
    use serde_json::json;

    use super::*;

    #[derive(Deserialize, JsonSchema)]
    #[serde(deny_unknown_fields)]
    struct Grouped {
        path: String,
        operation: Operation,
    }

    #[derive(Deserialize, JsonSchema)]
    #[serde(tag = "action", rename_all = "snake_case", deny_unknown_fields)]
    enum Operation {
        AddRow { table: usize },
        Delete {},
    }

    #[test]
    fn grouped_arguments_are_strict_and_keep_an_object_root() -> Result<()> {
        let parsed: Grouped = parse(
            "t",
            json!({"path": "a", "operation": {"action": "add_row", "table": 2}}),
        )?;
        assert_eq!(parsed.path, "a");
        assert!(matches!(parsed.operation, Operation::AddRow { table: 2 }));
        assert!(matches!(
            parse::<Grouped>("t", json!({"path": "a", "operation": {"action": "delete"}}))?
                .operation,
            Operation::Delete {}
        ));
        for invalid in [
            json!({"path": "a", "operation": {"action": "add_row", "table": 2, "extra": 1}}),
            json!({"path": "a", "operation": {"action": "delete", "table": 2}}),
            json!({"path": "a", "operation": {"action": "unknown"}}),
            json!({"path": "a", "operation": {"action": "add_row"}, "extra": 1}),
        ] {
            assert!(parse::<Grouped>("t", invalid.clone()).is_err(), "{invalid}");
        }
        let schema = Value::Object(
            tool::<Grouped>("t", "d", Effect::ReadOnly)
                .input_schema
                .as_ref()
                .clone(),
        );
        assert_eq!(schema["type"], "object");
        for keyword in ["oneOf", "anyOf", "allOf"] {
            assert!(schema.get(keyword).is_none(), "top-level {keyword}");
        }
        Ok(())
    }
}
