//! `ferroxl-mcp` — a Model Context Protocol server exposing ferroxl's spreadsheet tools to
//! AI agents.
//!
//! The transport is stdio, as the specification requires for a locally launched server:
//! one JSON-RPC message per line on standard input, one per line on standard output.
//! Diagnostics go to standard error so they never corrupt the protocol stream.
//!
//! # Running
//!
//! ```text
//! cargo run -p ferroxl-mcp -- --root ./spreadsheets
//! ```
//!
//! `--root` (or the `FERROXL_ROOT` environment variable) bounds every path the tools may
//! touch. Without it the working directory is used. The bound is a convenience for the
//! agent, not a security sandbox: a caller that can reach this process can already do
//! whatever it likes.

mod handlers;
mod rpc;
mod tools;
mod values;
mod workspace;

#[cfg(test)]
mod testing;

use std::io::{BufRead, BufReader, Write};
use std::process::ExitCode;

use serde_json::{json, Value};

use rpc::Request;
use workspace::Workspace;

/// The program name reported to the client.
const SERVER_NAME: &str = "ferroxl-mcp";

fn main() -> ExitCode {
    let options = match Options::parse(std::env::args().skip(1)) {
        Ok(options) => options,
        Err(message) => {
            eprintln!("{SERVER_NAME}: {message}");
            eprintln!("{}", Options::usage());
            return ExitCode::from(2);
        }
    };
    if options.help {
        println!("{}", Options::usage());
        return ExitCode::SUCCESS;
    }
    if options.version {
        println!("{SERVER_NAME} {}", env!("CARGO_PKG_VERSION"));
        return ExitCode::SUCCESS;
    }
    let root = options
        .root
        .clone()
        .or_else(|| std::env::var("FERROXL_ROOT").ok())
        .unwrap_or_else(|| ".".to_string());
    let workspace = match Workspace::new(&root) {
        Ok(workspace) => workspace,
        Err(error) => {
            eprintln!("{SERVER_NAME}: cannot use {root:?} as a workspace: {error}");
            return ExitCode::FAILURE;
        }
    };
    eprintln!(
        "{SERVER_NAME} {} serving {}",
        env!("CARGO_PKG_VERSION"),
        workspace.root().display()
    );
    run(&workspace)
}

/// Serve until standard input reaches end of file.
fn run(workspace: &Workspace) -> ExitCode {
    let stdin = BufReader::new(std::io::stdin());
    let mut stdout = std::io::stdout().lock();
    for line in stdin.lines() {
        let line = match line {
            Ok(line) => line,
            Err(error) => {
                eprintln!("{SERVER_NAME}: cannot read standard input: {error}");
                return ExitCode::FAILURE;
            }
        };
        if line.trim().is_empty() {
            continue;
        }
        if let Some(response) = handle_line(workspace, &line) {
            if writeln!(stdout, "{response}").is_err() || stdout.flush().is_err() {
                // The client has gone; there is nothing useful left to do.
                return ExitCode::SUCCESS;
            }
        }
    }
    ExitCode::SUCCESS
}

/// Handle one line of input, returning the line to write back, if any.
///
/// A notification produces no response, which is what the specification requires.
pub fn handle_line(workspace: &Workspace, line: &str) -> Option<Value> {
    let request = match rpc::parse(line) {
        Ok(request) => request,
        Err(error) => {
            // A payload that will not parse cannot be correlated, so the id is null.
            return Some(rpc::failure(&Value::Null, &error));
        }
    };
    if request.is_notification() {
        // Notifications are fire-and-forget. `notifications/initialized` is the only one the
        // specification sends, and there is nothing to do for it.
        return None;
    }
    let id = request.id.clone().unwrap_or(Value::Null);
    match dispatch(workspace, &request) {
        Ok(result) => Some(rpc::success(&id, result)),
        Err(error) => Some(rpc::failure(&id, &error)),
    }
}

/// Route a request to its handler.
fn dispatch(workspace: &Workspace, request: &Request) -> Result<Value, rpc::RpcFailure> {
    match request.method.as_str() {
        "initialize" => Ok(initialize(request)),
        "ping" => Ok(json!({})),
        "tools/list" => Ok(json!({
            "tools": tools::catalogue().iter().map(tools::ToolSpec::to_json).collect::<Vec<_>>(),
        })),
        "tools/call" => call_tool(workspace, request),
        other => Err(rpc::RpcFailure::method_not_found(other)),
    }
}

/// The `initialize` result, negotiating the protocol version with the client.
fn initialize(request: &Request) -> Value {
    let proposed = request.string_arg("protocolVersion");
    let agreed = match proposed {
        Some(version) if rpc::SUPPORTED_PROTOCOL_VERSIONS.contains(&version.as_str()) => version,
        // The specification says to answer with the revision this server implements when
        // the client's is not one it knows.
        _ => rpc::PROTOCOL_VERSION.to_string(),
    };
    json!({
        "protocolVersion": agreed,
        "capabilities": {
            "tools": { "listChanged": false },
        },
        "serverInfo": {
            "name": SERVER_NAME,
            "title": "ferroxl spreadsheet tools",
            "version": env!("CARGO_PKG_VERSION"),
        },
        "instructions": "Read a workbook with list_sheets, describe_sheet and read_cells before \
                        editing it. Values beginning with '=' are stored as formulas, \
                        'YYYY-MM-DD' as a date, and a trailing '%' as a percentage. Every path \
                        is relative to the workspace root.",
    })
}

/// Run a tool and shape the outcome for the client.
fn call_tool(workspace: &Workspace, request: &Request) -> Result<Value, rpc::RpcFailure> {
    let name = request
        .string_arg("name")
        .ok_or_else(|| rpc::RpcFailure::invalid_params("tools/call needs a tool name"))?;
    let arguments = request
        .params
        .get("arguments")
        .cloned()
        .unwrap_or(Value::Null);
    match handlers::call(workspace, &name, &arguments) {
        Ok((summary, structured)) => Ok(rpc::tool_result(&summary, structured)),
        // A tool that ran and failed is reported in-band, so the model can read the
        // message and correct its call. An unknown tool or an inconsistent server is a
        // protocol-level error.
        Err(handlers::ToolFailure::Failed(message)) => Ok(rpc::tool_error(&message)),
        Err(handlers::ToolFailure::Unknown(name)) => Err(rpc::RpcFailure::tool_not_found(&name)),
        Err(handlers::ToolFailure::Broken(message)) => Err(rpc::RpcFailure::internal(message)),
    }
}

/// The command-line options.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct Options {
    root: Option<String>,
    help: bool,
    version: bool,
}

impl Options {
    /// Parse the arguments, rejecting anything unrecognised.
    fn parse<I, S>(args: I) -> Result<Self, String>
    where
        I: Iterator<Item = S>,
        S: Into<String>,
    {
        let mut options = Options::default();
        let mut args = args.peekable();
        while let Some(argument) = args.next() {
            let argument = argument.into();
            match argument.as_str() {
                "-h" | "--help" => options.help = true,
                "-V" | "--version" => options.version = true,
                "--root" => {
                    let value = args
                        .next()
                        .ok_or_else(|| "--root needs a directory".to_string())?;
                    options.root = Some(value.into());
                }
                other if other.starts_with("--root=") => {
                    options.root = Some(other.trim_start_matches("--root=").to_string())
                }
                other => return Err(format!("unrecognised argument {other:?}")),
            }
        }
        Ok(options)
    }

    /// The text printed for `--help`.
    fn usage() -> String {
        format!(
            "{SERVER_NAME} {} — a Model Context Protocol server for Excel workbooks\n\
             \n\
             USAGE:\n    {SERVER_NAME} [--root <directory>]\n\
             \n\
             OPTIONS:\n\
             \x20   --root <directory>  Bound every path the tools may touch. Defaults to\n\
             \x20                       $FERROXL_ROOT, then the working directory.\n\
             \x20   -h, --help          Print this help and exit.\n\
             \x20   -V, --version       Print the version and exit.\n\
             \n\
             The server speaks JSON-RPC 2.0 over stdio, one message per line.",
            env!("CARGO_PKG_VERSION")
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn workspace() -> Workspace {
        testing::empty_workspace("server")
    }

    fn ask(workspace: &Workspace, message: Value) -> Value {
        let response = handle_line(workspace, &message.to_string())
            .unwrap_or_else(|| panic!("expected a response to {message}"));
        response
    }

    #[test]
    fn initialize_reports_the_server_and_its_capabilities() {
        let workspace = workspace();
        let response = ask(
            &workspace,
            json!({
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize",
                "params": { "protocolVersion": "2025-03-26" },
            }),
        );
        assert_eq!(response["result"]["protocolVersion"], json!("2025-03-26"));
        assert_eq!(
            response["result"]["serverInfo"]["name"],
            json!("ferroxl-mcp")
        );
        assert!(response["result"]["capabilities"]["tools"].is_object());
    }

    #[test]
    fn an_unknown_protocol_version_falls_back_to_the_server_revision() {
        let workspace = workspace();
        let response = ask(
            &workspace,
            json!({
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize",
                "params": { "protocolVersion": "1999-01-01" },
            }),
        );
        assert_eq!(
            response["result"]["protocolVersion"],
            json!(rpc::PROTOCOL_VERSION)
        );
    }

    #[test]
    fn tools_list_advertises_every_tool_with_a_schema() {
        let workspace = workspace();
        let response = ask(
            &workspace,
            json!({ "jsonrpc": "2.0", "id": 2, "method": "tools/list" }),
        );
        let listed = response["result"]["tools"].as_array().expect("an array");
        assert_eq!(listed.len(), tools::catalogue().len());
        assert_eq!(listed[0]["name"], json!("list_sheets"));
        assert!(listed[0]["inputSchema"]["properties"]["path"].is_object());
    }

    #[test]
    fn ping_answers_with_an_empty_result() {
        let workspace = workspace();
        let response = ask(
            &workspace,
            json!({ "jsonrpc": "2.0", "id": 3, "method": "ping" }),
        );
        assert_eq!(response["result"], json!({}));
        assert_eq!(response["id"], json!(3));
    }

    #[test]
    fn an_unknown_method_is_a_protocol_error() {
        let workspace = workspace();
        let response = ask(
            &workspace,
            json!({ "jsonrpc": "2.0", "id": 4, "method": "teleport" }),
        );
        assert_eq!(response["error"]["code"], json!(-32601));
    }

    #[test]
    fn a_notification_gets_no_reply() {
        let workspace = workspace();
        let response = handle_line(
            &workspace,
            &json!({
                "jsonrpc": "2.0",
                "method": "notifications/initialized",
            })
            .to_string(),
        );
        assert!(response.is_none());
    }

    #[test]
    fn a_tool_call_returns_content_and_structured_output() {
        let workspace = testing::workspace();
        let response = ask(
            &workspace,
            json!({
                "jsonrpc": "2.0",
                "id": 5,
                "method": "tools/call",
                "params": {
                    "name": "list_sheets",
                    "arguments": { "path": "report.xlsx" },
                },
            }),
        );
        assert_eq!(response["result"]["isError"], json!(false));
        assert!(response["result"]["content"][0]["text"]
            .as_str()
            .unwrap()
            .contains("2 sheets"));
        assert_eq!(
            response["result"]["structuredContent"]["sheets"][0]["name"],
            json!("Data")
        );
    }

    #[test]
    fn a_tool_that_fails_reports_is_error_rather_than_a_protocol_error() {
        let workspace = testing::workspace();
        let response = ask(
            &workspace,
            json!({
                "jsonrpc": "2.0",
                "id": 6,
                "method": "tools/call",
                "params": {
                    "name": "read_cells",
                    "arguments": { "path": "report.xlsx", "sheet": "Ghost" },
                },
            }),
        );
        // The call succeeded at the protocol level; the tool is what failed.
        assert!(response.get("error").is_none());
        assert_eq!(response["result"]["isError"], json!(true));
        assert!(response["result"]["content"][0]["text"]
            .as_str()
            .unwrap()
            .contains("Ghost"));
    }

    #[test]
    fn calling_a_tool_without_a_name_is_invalid_params() {
        let workspace = workspace();
        let response = ask(
            &workspace,
            json!({
                "jsonrpc": "2.0",
                "id": 7,
                "method": "tools/call",
                "params": {},
            }),
        );
        assert_eq!(response["error"]["code"], json!(-32602));
    }

    #[test]
    fn unparseable_input_is_a_parse_error_with_a_null_id() {
        let workspace = workspace();
        let response = handle_line(&workspace, "{oh dear").unwrap();
        assert_eq!(response["error"]["code"], json!(-32700));
        assert_eq!(response["id"], json!(Value::Null));
    }

    #[test]
    fn string_ids_are_echoed_unchanged() {
        let workspace = workspace();
        let response = ask(
            &workspace,
            json!({ "jsonrpc": "2.0", "id": "abc", "method": "ping" }),
        );
        assert_eq!(response["id"], json!("abc"));
    }

    #[test]
    fn options_parse_the_documented_flags() {
        let parsed = Options::parse(["--root", "spreadsheets"].into_iter()).unwrap();
        assert_eq!(parsed.root.as_deref(), Some("spreadsheets"));
        let parsed = Options::parse(["--root=books".to_string()].into_iter()).unwrap();
        assert_eq!(parsed.root.as_deref(), Some("books"));
        assert!(
            Options::parse(["--help".to_string()].into_iter())
                .unwrap()
                .help
        );
        assert!(
            Options::parse(["-V".to_string()].into_iter())
                .unwrap()
                .version
        );
    }

    #[test]
    fn a_bare_root_flag_is_an_error() {
        let error = Options::parse(["--root"].into_iter()).unwrap_err();
        assert!(error.contains("--root needs a directory"), "{error}");
    }

    #[test]
    fn an_unknown_flag_is_rejected() {
        let error = Options::parse(["--turbo".to_string()].into_iter()).unwrap_err();
        assert!(error.contains("--turbo"), "{error}");
    }

    #[test]
    fn the_usage_text_names_the_server_and_its_options() {
        let usage = Options::usage();
        assert!(usage.contains("ferroxl-mcp"));
        assert!(usage.contains("--root"));
        assert!(usage.contains("stdio"));
    }
}
