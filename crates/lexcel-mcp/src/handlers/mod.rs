//! Tool implementations.
//!
//! Each handler takes the arguments object and returns either a summary line or a failure.
//! Failures are returned as `Err` rather than panicking, because the caller turns them
//! into an MCP `isError` result the model can read and act on.

mod edit;
mod format;
mod inspect;
mod media;

use lexcel::{LoadOptions, Workbook};

use crate::tools::Args;
use crate::workspace::Workspace;

/// Run the named tool.
///
/// An unknown name is a protocol error, not a tool error: the model asked for something
/// this server does not have, which is worth a `-32601` rather than a soft failure.
pub fn call(
    workspace: &Workspace,
    name: &str,
    arguments: &serde_json::Value,
) -> Result<ToolOutput, ToolFailure> {
    let spec = crate::tools::find(name).ok_or_else(|| ToolFailure::Unknown(name.to_string()))?;
    let args = Args::new(arguments);
    // The runtime agrees with the schema the model was shown, so a misspelled argument is
    // reported instead of quietly ignored.
    if let Err(message) = args.reject_unknown(spec) {
        return Err(ToolFailure::Failed(message));
    }
    let outcome = match name {
        "list_sheets" => inspect::list_sheets(workspace, &args),
        "describe_sheet" => inspect::describe_sheet(workspace, &args),
        "read_cells" => inspect::read_cells(workspace, &args),
        "read_formulas" => inspect::read_formulas(workspace, &args),
        "search_values" => inspect::search_values(workspace, &args),
        "summarize_range" => inspect::summarize_range(workspace, &args),
        "list_comments" => inspect::list_comments(workspace, &args),
        "list_named_ranges" => inspect::list_named_ranges(workspace, &args),
        "export_csv" => inspect::export_csv(workspace, &args),
        "create_workbook" => edit::create_workbook(workspace, &args),
        "add_sheet" => edit::add_sheet(workspace, &args),
        "remove_sheet" => edit::remove_sheet(workspace, &args),
        "rename_sheet" => edit::rename_sheet(workspace, &args),
        "merge_cells" => edit::merge_cells(workspace, &args),
        "unmerge_cells" => edit::unmerge_cells(workspace, &args),
        "freeze_panes" => edit::freeze_panes(workspace, &args),
        "set_auto_filter" => edit::set_auto_filter(workspace, &args),
        "add_named_range" => edit::add_named_range(workspace, &args),
        "set_cell" => edit::set_cell(workspace, &args),
        "write_cells" => edit::write_cells(workspace, &args),
        "append_row" => edit::append_row(workspace, &args),
        "clear_cells" => edit::clear_cells(workspace, &args),
        "set_column_width" => format::set_column_width(workspace, &args),
        "set_row_height" => format::set_row_height(workspace, &args),
        "set_header_footer" => format::set_header_footer(workspace, &args),
        "add_hyperlink" => format::add_hyperlink(workspace, &args),
        "style_cells" => format::style_cells(workspace, &args),
        "set_number_format" => format::set_number_format(workspace, &args),
        "add_data_validation" => format::add_data_validation(workspace, &args),
        "add_conditional_format" => format::add_conditional_format(workspace, &args),
        "add_chart" => media::add_chart(workspace, &args),
        "add_image" => media::add_image(workspace, &args),
        "add_comment" => media::add_comment(workspace, &args),
        // Reaching a tool that is advertised but has no handler is a bug in this table
        // rather than anything the client did.
        other => {
            return Err(ToolFailure::Broken(format!(
                "{other} is advertised but has no handler"
            )))
        }
    };
    outcome.map_err(ToolFailure::Failed)
}

/// Why a tool call did not produce a result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ToolFailure {
    /// The client asked for a tool this server does not have.
    Unknown(String),
    /// The tool ran and failed; the message is meant for the model to read and act on.
    Failed(String),
    /// The server is inconsistent: a tool is advertised with no handler behind it.
    Broken(String),
}

impl ToolFailure {
    /// The message to show the client.
    pub fn message(&self) -> String {
        match self {
            ToolFailure::Unknown(name) => format!("no tool named {name:?} is registered"),
            ToolFailure::Failed(message) | ToolFailure::Broken(message) => message.clone(),
        }
    }
}

impl std::fmt::Display for ToolFailure {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.message())
    }
}

/// The successful outcome of a tool: a human-readable line and the structured payload.
pub type ToolOutput = (String, serde_json::Value);

/// Shorthand for a handler result.
pub type Handled = Result<ToolOutput, String>;

/// Turn a lexcel error into the message a model can read.
///
/// lexcel's errors already carry a descriptive message, so `Display` is the whole
/// conversion; naming the helper keeps the call sites short.
pub fn err(error: impl std::fmt::Display) -> String {
    error.to_string()
}

/// Load a workbook for reading.
pub fn open(workspace: &Workspace, args: &Args) -> Result<Workbook, String> {
    let path = args.require_str("path")?;
    workspace.load(&path, LoadOptions::default())
}

/// Load a workbook for editing, along with the index of the sheet to act on.
///
/// The workbook is returned by value because the writer consumes it, which makes an
/// edit-then-save sequence a single linear function.
pub fn open_for_edit(workspace: &Workspace, args: &Args) -> Result<(Workbook, usize), String> {
    let workbook = open(workspace, args)?;
    let index = sheet_index(&workbook, args)?;
    Ok((workbook, index))
}

/// Load a workbook and resolve the index of a named sheet in one step.
pub fn open_at_sheet(
    workspace: &Workspace,
    args: &Args,
    name: &str,
) -> Result<(Workbook, usize), String> {
    let workbook = open(workspace, args)?;
    let index = workbook
        .get_index(name)
        .ok_or_else(|| unknown_sheet(&workbook, name))?;
    Ok((workbook, index))
}

/// Resolve a sheet name to its index, defaulting to the first sheet.
pub fn sheet_index(workbook: &Workbook, args: &Args) -> Result<usize, String> {
    match args.sheet() {
        Some(name) => workbook
            .get_index(&name)
            .ok_or_else(|| unknown_sheet(workbook, &name)),
        None => Ok(0),
    }
}

/// The message shown when a caller names a sheet that is not in the workbook.
pub fn unknown_sheet(workbook: &Workbook, name: &str) -> String {
    format!(
        "there is no sheet named {name:?}; the workbook has {}",
        workbook.get_sheet_names().join(", ")
    )
}

/// Expand a range argument, defaulting to the sheet's used range.
pub fn used_range(sheet: &lexcel::Worksheet, args: &Args) -> Result<String, String> {
    let Some(range) = args.opt_str("range") else {
        return sheet.calculate_dimension().map_err(|e| e.to_string());
    };
    Ok(range)
}

/// Iterate a rectangle of a sheet as `(coordinate, value)` pairs.
pub fn cells_in_range(
    sheet: &lexcel::Worksheet,
    range: &str,
) -> Result<Vec<(String, lexcel::CellValue)>, String> {
    let coordinates = sheet.range_coordinates(range).map_err(|e| e.to_string())?;
    Ok(coordinates
        .into_iter()
        .map(|coordinate| {
            let value = sheet
                .cell_value(&coordinate)
                .unwrap_or(lexcel::CellValue::None);
            (coordinate, value)
        })
        .collect())
}

/// Check that a colour looks like one Excel will accept.
///
/// Six digits are RGB and eight are ARGB; a leading `#` is tolerated and stripped. The
/// returned string is the bare hex, which is the form lexcel stores.
pub fn check_colour(name: &str, value: &str) -> Result<String, String> {
    let trimmed = value.trim().trim_start_matches('#');
    let well_formed =
        matches!(trimmed.len(), 6 | 8) && trimmed.chars().all(|c| c.is_ascii_hexdigit());
    if well_formed {
        return Ok(trimmed.to_ascii_uppercase());
    }
    Err(format!(
        "{name} must be 6 or 8 hexadecimal digits such as \"FF0000\" or \"FFFF0000\", not {value:?}"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn workbook() -> Workbook {
        let mut workbook = Workbook::new();
        workbook.create_sheet(Some("Data")).unwrap();
        workbook
    }

    #[test]
    fn a_sheet_name_defaults_to_the_first_sheet() {
        let workbook = workbook();
        let args = Args::new(&json!({}));
        assert_eq!(sheet_index(&workbook, &args).unwrap(), 0);
    }

    #[test]
    fn a_named_sheet_is_looked_up() {
        let workbook = workbook();
        let args = Args::new(&json!({ "sheet": "Data" }));
        assert_eq!(sheet_index(&workbook, &args).unwrap(), 1);
    }

    #[test]
    fn an_unknown_sheet_lists_the_sheet_that_do_exist() {
        let workbook = workbook();
        let args = Args::new(&json!({ "sheet": "Nope" }));
        let error = sheet_index(&workbook, &args).unwrap_err();
        assert!(error.contains("Nope"), "{error}");
        assert!(error.contains("Sheet1, Data"), "{error}");
    }

    #[test]
    fn colours_are_checked_for_length_and_digits() {
        assert_eq!(check_colour("font_color", "FFFF0000").unwrap(), "FFFF0000");
        assert_eq!(check_colour("font_color", "#FF0000").unwrap(), "FF0000");
        assert!(check_colour("font_color", "red").is_err());
        assert!(check_colour("font_color", "FF00").is_err());
        assert!(check_colour("font_color", "ZZZZZZZZ").is_err());
    }

    #[test]
    fn an_unknown_tool_is_reported_separately_from_a_failure() {
        let workspace = crate::testing::empty_workspace("dispatch");
        let error = call(&workspace, "teleport", &json!({})).unwrap_err();
        assert!(matches!(error, ToolFailure::Unknown(name) if name == "teleport"));
    }

    #[test]
    fn a_tool_failure_keeps_the_message_it_was_given() {
        let error = ToolFailure::Failed("no such spell".to_string());
        assert_eq!(error.message(), "no such spell");
    }
}
