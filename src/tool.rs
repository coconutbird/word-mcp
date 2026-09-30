//! Shared construction of MCP tool definitions and typed argument parsing.
//!
//! Each tool's arguments are one Rust type. Its `JsonSchema` derive produces the
//! advertised input schema, and its `Deserialize` derive enforces the same contract, so
//! the schema and the parser cannot drift apart.
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
    /// Modifies or replaces existing document content.
    Destructive,
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
