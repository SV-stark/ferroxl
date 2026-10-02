//! The JSON-RPC 2.0 envelope spoken by the Model Context Protocol.
//!
//! MCP over stdio is one JSON object per line in each direction. This module owns the
//! framing: parsing a line into a request, building a success response and building a
//! failure. Everything protocol-shaped lives here so the tool handlers can stay ordinary
//! Rust functions.

use serde_json::{json, Value};

/// The protocol revision this server is written against.
pub const PROTOCOL_VERSION: &str = "2025-06-18";

/// Revisions this server will speak, newest first.
///
/// A client is told which one it got in the `initialize` result. The client proposes its
/// own revision and, if it is in this list, the negotiation succeeds with that value;
/// otherwise the server falls back to [`PROTOCOL_VERSION`], which is what the
/// specification requires.
pub const SUPPORTED_PROTOCOL_VERSIONS: &[&str] = &["2025-06-18", "2025-03-26", "2024-11-05"];

/// A decoded request or notification.
///
/// A notification is a message with no `id`; the specification says the server must not
/// answer one, which is why `id` is optional here rather than always present.
#[derive(Debug, Clone)]
pub struct Request {
    /// The correlation id, absent for a notification.
    pub id: Option<Value>,
    /// The method name, for example `tools/call`.
    pub method: String,
    /// The `params` object, empty when the sender omitted it.
    pub params: Value,
}

impl Request {
    /// Whether this message is a notification, which must never be answered.
    pub fn is_notification(&self) -> bool {
        self.id.is_none()
    }

    /// Read a string field from `params`.
    pub fn string_arg(&self, name: &str) -> Option<String> {
        self.params
            .get(name)
            .and_then(Value::as_str)
            .map(str::to_string)
    }
}

/// The JSON-RPC error codes this server can produce.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorCode {
    /// The payload was not valid JSON, or not a JSON-RPC object.
    ParseError,
    /// The object was well-formed JSON but not a valid JSON-RPC request.
    InvalidRequest,
    /// The method does not exist.
    MethodNotFound,
    /// The method exists but the arguments do not fit it.
    InvalidParams,
    /// The method is known but the server failed while carrying it out.
    InternalError,
}

impl ErrorCode {
    /// The numeric code sent on the wire.
    pub fn as_i64(self) -> i64 {
        match self {
            ErrorCode::ParseError => -32700,
            ErrorCode::InvalidRequest => -32600,
            ErrorCode::MethodNotFound => -32601,
            ErrorCode::InvalidParams => -32602,
            ErrorCode::InternalError => -32603,
        }
    }
}

/// A failure that is reported as a JSON-RPC error rather than as tool output.
#[derive(Debug, Clone)]
pub struct RpcFailure {
    /// The JSON-RPC error code.
    pub code: ErrorCode,
    /// A human-readable explanation.
    pub message: String,
}

impl RpcFailure {
    /// A `-32601 Method not found` failure.
    pub fn method_not_found(method: &str) -> Self {
        RpcFailure {
            code: ErrorCode::MethodNotFound,
            message: format!("unknown method {method:?}"),
        }
    }

    /// A `-32601` failure for a tool name that is not in the catalogue.
    ///
    /// The code is the same as an unknown method, but the message names the tool, because
    /// the request was well formed and only the name inside it was wrong. `tools/list` is
    /// where a caller finds the spelling it should have used.
    pub fn tool_not_found(tool: &str) -> Self {
        RpcFailure {
            code: ErrorCode::MethodNotFound,
            message: format!("no tool named {tool:?}; call tools/list for the catalogue"),
        }
    }

    /// A `-32602 Invalid params` failure describing a specific argument.
    pub fn invalid_params(message: impl Into<String>) -> Self {
        RpcFailure {
            code: ErrorCode::InvalidParams,
            message: message.into(),
        }
    }

    /// A `-32603 Internal error` failure.
    pub fn internal(message: impl Into<String>) -> Self {
        RpcFailure {
            code: ErrorCode::InternalError,
            message: message.into(),
        }
    }

    /// Render this failure as a JSON-RPC error object.
    pub fn to_value(&self) -> Value {
        json!({
            "code": self.code.as_i64(),
            "message": self.message,
        })
    }
}

/// Decode one line of input into a request.
///
/// The returned failure is already shaped as a JSON-RPC error; `id` is `null` because a
/// payload that could not be parsed cannot be correlated with anything.
pub fn parse(line: &str) -> Result<Request, RpcFailure> {
    let value: Value = serde_json::from_str(line).map_err(|e| RpcFailure {
        code: ErrorCode::ParseError,
        message: format!("invalid JSON: {e}"),
    })?;
    parse_value(&value)
}

/// Decode an already-parsed JSON value into a request.
pub fn parse_value(value: &Value) -> Result<Request, RpcFailure> {
    let object = value.as_object().ok_or_else(|| RpcFailure {
        code: ErrorCode::InvalidRequest,
        message: "a request must be a JSON object".to_string(),
    })?;
    if object.get("jsonrpc").and_then(Value::as_str) != Some("2.0") {
        return Err(RpcFailure {
            code: ErrorCode::InvalidRequest,
            message: "the jsonrpc member must be exactly \"2.0\"".to_string(),
        });
    }
    let method = object
        .get("method")
        .and_then(Value::as_str)
        .ok_or_else(|| RpcFailure {
            code: ErrorCode::InvalidRequest,
            message: "a request must carry a string method".to_string(),
        })?
        .to_string();
    // `id` may be absent (a notification) or explicitly null, but when present it must be
    // a string or a number so the reply can be correlated.
    let id = match object.get("id") {
        None | Some(Value::Null) => None,
        Some(value @ (Value::String(_) | Value::Number(_))) => Some(value.clone()),
        Some(_) => {
            return Err(RpcFailure {
                code: ErrorCode::InvalidRequest,
                message: "id must be a string or a number".to_string(),
            })
        }
    };
    Ok(Request {
        id,
        method,
        params: object.get("params").cloned().unwrap_or(Value::Null),
    })
}

/// Build a success response for `id`.
pub fn success(id: &Value, result: Value) -> Value {
    json!({ "jsonrpc": "2.0", "id": id, "result": result })
}

/// Build an error response for `id`.
pub fn failure(id: &Value, error: &RpcFailure) -> Value {
    json!({ "jsonrpc": "2.0", "id": id, "error": error.to_value() })
}

/// Build the `result` for a `tools/call`.
///
/// MCP distinguishes a tool that ran but failed (`isError: true` with the message in
/// `content`) from a protocol-level error, and the difference matters to the agent: the
/// former is something the model can reason about and retry.
pub fn tool_result(summary: &str, structured: Value) -> Value {
    let rendered =
        serde_json::to_string_pretty(&structured).unwrap_or_else(|_| summary.to_string());
    json!({
        "content": [
            { "type": "text", "text": summary },
            { "type": "text", "text": rendered },
        ],
        "isError": false,
        "structuredContent": structured,
    })
}

/// Build the `result` for a tool that ran and failed.
pub fn tool_error(message: &str) -> Value {
    json!({
        "content": [{ "type": "text", "text": message }],
        "isError": true,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_a_request_with_params() {
        let request =
            parse(r#"{"jsonrpc":"2.0","id":7,"method":"tools/call","params":{"name":"x"}}"#)
                .unwrap();
        assert_eq!(request.method, "tools/call");
        assert_eq!(request.id, Some(json!(7)));
        assert_eq!(request.string_arg("name").as_deref(), Some("x"));
        assert!(!request.is_notification());
    }

    #[test]
    fn a_message_without_an_id_is_a_notification() {
        let request = parse(r#"{"jsonrpc":"2.0","method":"notifications/initialized"}"#).unwrap();
        assert!(request.is_notification());
        assert_eq!(request.params, Value::Null);
    }

    #[test]
    fn an_explicit_null_id_is_also_a_notification() {
        let request = parse(r#"{"jsonrpc":"2.0","id":null,"method":"ping"}"#).unwrap();
        assert!(request.is_notification());
    }

    #[test]
    fn bad_payloads_are_rejected_with_the_right_code() {
        assert_eq!(parse("not json").unwrap_err().code, ErrorCode::ParseError);
        assert_eq!(
            parse("[1, 2, 3]").unwrap_err().code,
            ErrorCode::InvalidRequest
        );
        assert_eq!(
            parse(r#"{"jsonrpc":"1.0","id":1,"method":"ping"}"#)
                .unwrap_err()
                .code,
            ErrorCode::InvalidRequest
        );
        assert_eq!(
            parse(r#"{"jsonrpc":"2.0","id":1}"#).unwrap_err().code,
            ErrorCode::InvalidRequest
        );
        assert_eq!(
            parse(r#"{"jsonrpc":"2.0","id":{"a":1},"method":"ping"}"#)
                .unwrap_err()
                .code,
            ErrorCode::InvalidRequest
        );
    }

    #[test]
    fn error_codes_match_the_json_rpc_registry() {
        assert_eq!(ErrorCode::ParseError.as_i64(), -32700);
        assert_eq!(ErrorCode::InvalidRequest.as_i64(), -32600);
        assert_eq!(ErrorCode::MethodNotFound.as_i64(), -32601);
        assert_eq!(ErrorCode::InvalidParams.as_i64(), -32602);
        assert_eq!(ErrorCode::InternalError.as_i64(), -32603);
    }

    #[test]
    fn success_carries_the_correlation_id() {
        let response = success(&json!("abc"), json!({"ok": true}));
        assert_eq!(response["id"], json!("abc"));
        assert_eq!(response["jsonrpc"], json!("2.0"));
        assert_eq!(response["result"]["ok"], json!(true));
    }

    #[test]
    fn tool_results_carry_both_text_and_structured_content() {
        let result = tool_result("two rows", json!({"rows": [[1, 2]]}));
        assert_eq!(result["isError"], json!(false));
        assert_eq!(result["content"][0]["text"], json!("two rows"));
        assert_eq!(result["structuredContent"]["rows"][0][1], json!(2));
        // The second text block is the pretty-printed JSON of the structured content.
        assert!(result["content"][1]["text"]
            .as_str()
            .unwrap()
            .contains("\"rows\""));
    }

    #[test]
    fn tool_errors_are_flagged_rather_than_raised() {
        let result = tool_error("no such sheet");
        assert_eq!(result["isError"], json!(true));
        assert_eq!(result["content"][0]["text"], json!("no such sheet"));
    }

    #[test]
    fn the_protocol_version_is_one_of_the_supported_ones() {
        assert!(SUPPORTED_PROTOCOL_VERSIONS.contains(&PROTOCOL_VERSION));
    }
}
