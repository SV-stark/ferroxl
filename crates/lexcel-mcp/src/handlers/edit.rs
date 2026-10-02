//! Tools that change a workbook's values and structure.

use lexcel::{CellValue, NamedRange, Workbook};
use serde_json::{json, Value};

use super::{err, open_at_sheet, open_for_edit, Handled};
use crate::tools::Args;
use crate::values;
use crate::workspace::Workspace;

/// Create a new workbook.
pub fn create_workbook(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let resolved = workspace.resolve_spreadsheet(&path)?;
    if resolved.exists() && !args.opt_bool("overwrite", false) {
        return Err(format!(
            "{path} already exists; pass overwrite to replace it"
        ));
    }
    let requested: Vec<String> = match args.opt_array("sheets") {
        Some(items) => items
            .iter()
            .map(|item| match item.as_str() {
                Some(name) if !name.trim().is_empty() => Ok(name.to_string()),
                _ => Err("every entry in sheets must be a non-empty string".to_string()),
            })
            .collect::<Result<_, _>>()?,
        None => vec!["Sheet".to_string()],
    };
    let mut workbook = Workbook::new();
    // `Workbook::new` already made one sheet, so the first requested name renames it and
    // the rest are appended.
    for (position, name) in requested.iter().enumerate() {
        if position == 0 {
            workbook.worksheets[0]
                .set_title(name, &[])
                .map_err(err)?;
        } else {
            workbook.create_sheet(Some(name)).map_err(err)?;
        }
    }
    let names = workbook.get_sheet_names();
    workspace.save(workbook, &path)?;
    let summary = format!("created {path} with {}", names.join(", "));
    Ok((summary, json!({ "path": path, "sheets": names })))
}

/// Append a sheet.
pub fn add_sheet(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let title = args.require_str("title")?;
    let (mut workbook, _) = open_for_edit(workspace, args)?;
    let index = match args.opt_usize("index") {
        Some(position) => workbook
            .create_sheet_at(Some(&title), Some(position))
            .map_err(err)?,
        None => workbook.create_sheet(Some(&title)).map_err(err)?,
    };
    workspace.save(workbook, &path)?;
    let summary = format!("added sheet {title:?} at position {index}");
    Ok((summary, json!({ "path": path, "name": title, "index": index })))
}

/// Delete a sheet.
pub fn remove_sheet(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let name = args.require_str("sheet")?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    if workbook.worksheets.len() == 1 {
        return Err("a workbook must keep at least one sheet".to_string());
    }
    workbook.remove_sheet(index).map_err(err)?;
    let remaining = workbook.get_sheet_names();
    workspace.save(workbook, &path)?;
    let summary = format!("removed sheet {name:?}");
    Ok((summary, json!({ "path": path, "remaining": remaining })))
}

/// Rename a sheet.
pub fn rename_sheet(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let from = args.require_str("sheet")?;
    let to = args.require_str("title")?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    // `set_title` checks the name against the other sheets, not the one being renamed.
    let others: Vec<String> = workbook
        .get_sheet_names()
        .into_iter()
        .filter(|name| *name != from)
        .collect();
    workbook.worksheets[index]
        .set_title(&to, &others)
        .map_err(err)?;
    // openpyxl appends a counter rather than refusing a duplicate, so the effective name
    // is read back rather than assumed.
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!("renamed {from:?} to {name:?}");
    Ok((summary, json!({ "path": path, "name": name })))
}

/// Merge a range.
pub fn merge_cells(workspace: &Workspace, args: &Args) -> Handled {
    change_range(workspace, args, Merge::Merge)
}

/// Unmerge a range.
pub fn unmerge_cells(workspace: &Workspace, args: &Args) -> Handled {
    change_range(workspace, args, Merge::Unmerge)
}

/// Freeze the panes at a cell, or unfreeze when no cell is given.
pub fn freeze_panes(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    let requested = args.opt_str("cell");
    workbook.worksheets[index].set_freeze_panes(requested.as_deref().unwrap_or(""));
    let applied = workbook.worksheets[index].freeze_panes.clone();
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = match &applied {
        Some(cell) => format!("froze panes at {cell} on {name}"),
        None => format!("unfroze panes on {name}"),
    };
    Ok((summary, json!({ "path": path, "sheet": name, "cell": applied })))
}

/// Set or clear the autofilter.
pub fn set_auto_filter(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    let range = args.opt_str("range");
    workbook.worksheets[index]
        .auto_filter
        .set_reference(&range.clone().unwrap_or_default());
    let applied = workbook.worksheets[index]
        .auto_filter
        .reference()
        .map(str::to_string);
    workspace.save(workbook, &path)?;
    let summary = match &applied {
        Some(range) => format!("autofilter set to {range}"),
        None => "autofilter cleared".to_string(),
    };
    Ok((summary, json!({ "path": path, "range": applied })))
}

/// Define a named range.
pub fn add_named_range(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let name = args.require_str("name")?;
    let sheet = args.require_str("sheet")?;
    let range = args.require_str("range")?;
    let (mut workbook, index) = open_at_sheet(workspace, args, &sheet)?;
    if workbook.get_named_range(&name).is_some() {
        return Err(format!("there is already a named range called {name:?}"));
    }
    let scope = match args.opt_str("scope") {
        Some(scope) => Some(
            workbook
                .get_index(&scope)
                .ok_or_else(|| format!("the scope sheet {scope:?} does not exist"))?,
        ),
        None => None,
    };
    // The sheet is passed separately, so a sheet qualifier on the range would be written
    // twice. openpyxl's `create_named_range` has the same shape and takes a bare range.
    let bare = range.rsplit_once('!').map(|(_, tail)| tail).unwrap_or(&range);
    workbook.add_named_range(NamedRange::new(
        name.clone(),
        vec![(index, bare.to_string())],
        scope,
    ));
    workspace.save(workbook, &path)?;
    let summary = format!("defined {name:?} as {sheet}!{bare}");
    Ok((
        summary,
        json!({ "path": path, "name": name, "sheet": sheet, "range": bare }),
    ))
}

/// Write one cell.
pub fn set_cell(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let cell = args.require_str("cell")?;
    // `raw` is used because the schema types this as anything, including an explicit null.
    let raw = args
        .raw("value")
        .ok_or_else(|| "value is required".to_string())?;
    let value = values::from_json(raw).map_err(|e| e.0)?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    workbook.worksheets[index].set(&cell, value.clone()).map_err(err)?;
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!("wrote {} to {name}!{cell}", values::to_display(&value));
    Ok((
        summary,
        json!({
            "path": path,
            "sheet": name,
            "cell": cell,
            "value": values::to_json(&value),
        }),
    ))
}

/// Write a block of values.
pub fn write_cells(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let start = args.require_str("start")?;
    let rows = args.require_array("rows")?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;

    let (start_column, start_row) =
        lexcel::coordinate_from_string(&start).map_err(err)?;
    let start_column = lexcel::column_index_from_string(&start_column).map_err(err)?;
    let mut written = 0usize;
    for (row_offset, row) in rows.iter().enumerate() {
        let cells = row
            .as_array()
            .ok_or_else(|| "every entry in rows must itself be an array".to_string())?;
        for (column_offset, raw) in cells.iter().enumerate() {
            let value = values::from_json(raw).map_err(|e| e.0)?;
            let letter = lexcel::get_column_letter(start_column + column_offset as u32)
                .map_err(err)?;
            let coordinate = format!("{letter}{}", start_row + row_offset as u32);
            workbook.worksheets[index]
                .set(&coordinate, value)
                .map_err(err)?;
            written += 1;
        }
    }
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!("wrote {written} cells from {start} on {name}");
    Ok((
        summary,
        json!({ "path": path, "sheet": name, "start": start, "cells_written": written }),
    ))
}

/// Append a row of values below the data.
pub fn append_row(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let raw_values = args.require_array("values")?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    let parsed: Vec<CellValue> = raw_values
        .iter()
        .map(|raw| values::from_json(raw).map_err(|e| e.0))
        .collect::<Result<_, _>>()?;
    let row = workbook.worksheets[index].highest_row() + 1;
    workbook.worksheets[index]
        .write_row(row, 1, &parsed)
        .map_err(err)?;
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!(
        "appended {} value(s) to row {row} of {name}",
        parsed.len()
    );
    Ok((
        summary,
        json!({ "path": path, "sheet": name, "row": row, "values_written": parsed.len() }),
    ))
}

/// Clear the values in a range.
pub fn clear_cells(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let range = args.require_str("range")?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    let coordinates: Vec<String> = workbook.worksheets[index]
        .range_coordinates(&range)
        .map_err(err)?;
    for coordinate in &coordinates {
        workbook.worksheets[index]
            .set(coordinate, CellValue::None)
            .map_err(err)?;
    }
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!("cleared {} cells in {range} on {name}", coordinates.len());
    Ok((
        summary,
        json!({
            "path": path,
            "sheet": name,
            "range": range,
            "cells_cleared": coordinates.len(),
        }),
    ))
}

/// Whether a range is being merged or unmerged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Merge {
    /// Combine the cells into one.
    Merge,
    /// Split them back apart.
    Unmerge,
}

/// Apply a merge or unmerge to a range argument.
fn change_range(workspace: &Workspace, args: &Args, change: Merge) -> Handled {
    let path = args.require_str("path")?;
    let range = args.require_str("range")?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    match change {
        Merge::Merge => workbook.worksheets[index]
            .merge_cells(&range)
            .map_err(err)?,
        Merge::Unmerge => workbook.worksheets[index]
            .unmerge_cells(&range)
            .map_err(err)?,
    }
    let name = workbook.worksheets[index].title().to_string();
    let merged: Vec<String> = workbook.worksheets[index].merged_cells().to_vec();
    workspace.save(workbook, &path)?;
    let verb = match change {
        Merge::Merge => "merged",
        Merge::Unmerge => "unmerged",
    };
    let summary = format!("{verb} {range} on {name}");
    Ok((
        summary,
        json!({ "path": path, "sheet": name, "range": range, "merged_ranges": merged }),
    ))
}

/// The structured payload a test needs to see the file's state after an edit.
#[allow(dead_code)]
fn payload(path: &str, sheet: String, extra: Value) -> Value {
    json!({ "path": path, "sheet": sheet, "extra": extra })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::handlers::inspect;
    use crate::testing;

    #[test]
    fn create_refuses_to_clobber_without_permission() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({ "path": "report.xlsx" }));
        let error = create_workbook(&workspace, &args).unwrap_err();
        assert!(error.contains("already exists"), "{error}");

        let args = Args::new(&json!({ "path": "report.xlsx", "overwrite": true }));
        let (summary, _) = create_workbook(&workspace, &args).unwrap();
        assert!(summary.contains("created"), "{summary}");
    }

    #[test]
    fn create_names_the_sheets_it_was_asked_for() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "fresh.xlsx",
            "sheets": ["Summary", "Detail"],
        }));
        let (_, payload) = create_workbook(&workspace, &args).unwrap();
        assert_eq!(payload["sheets"], json!(["Summary", "Detail"]));
    }

    #[test]
    fn set_cell_round_trips_through_the_file() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "cell": "C1",
            "value": "=SUM(A1:B1)",
        }));
        set_cell(&workspace, &args).unwrap();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "range": "C1:C1",
        }));
        let (_, read_back) = inspect::read_cells(&workspace, &args).unwrap();
        assert_eq!(read_back["cells"][0]["value"]["type"], json!("formula"));
    }

    #[test]
    fn a_null_value_clears_the_cell() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "cell": "A1",
            "value": null,
        }));
        set_cell(&workspace, &args).unwrap();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "range": "A1:A1",
        }));
        let (_, read_back) = inspect::read_cells(&workspace, &args).unwrap();
        assert_eq!(read_back["cells"][0]["value"], Value::Null);
    }

    #[test]
    fn write_cells_fills_a_block() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "start": "C3",
            "rows": [[1, 2], [3, 4]],
        }));
        let (_, written) = write_cells(&workspace, &args).unwrap();
        assert_eq!(written["cells_written"], json!(4));

        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "range": "C3:D4",
        }));
        let (_, read_back) = inspect::export_csv(&workspace, &args).unwrap();
        assert_eq!(read_back["csv"], json!("1,2\r\n3,4"));
    }

    #[test]
    fn write_cells_rejects_ragged_input() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "start": "A1",
            "rows": [1, 2],
        }));
        let error = write_cells(&workspace, &args).unwrap_err();
        assert!(error.contains("array"), "{error}");
    }

    #[test]
    fn append_row_lands_below_the_data() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "values": [5, 6],
        }));
        let (_, appended) = append_row(&workspace, &args).unwrap();
        assert_eq!(appended["row"], json!(3));

        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "range": "A3:B3",
        }));
        let (_, read_back) = inspect::read_cells(&workspace, &args).unwrap();
        assert_eq!(read_back["cells"][0]["value"]["value"], json!(5.0));
    }

    #[test]
    fn sheets_can_be_added_renamed_and_removed() {
        let workspace = testing::workspace();
        add_sheet(
            &workspace,
            &Args::new(&json!({ "path": "report.xlsx", "title": "Extra" })),
        )
        .unwrap();
        rename_sheet(
            &workspace,
            &Args::new(&json!({
                "path": "report.xlsx",
                "sheet": "Extra",
                "title": "Renamed",
})),
        )
        .unwrap();
        let (_, removed) = remove_sheet(
            &workspace,
            &Args::new(&json!({ "path": "report.xlsx", "sheet": "Renamed" })),
        )
        .unwrap();
        assert_eq!(removed["remaining"], json!(["Data", "Numbers"]));
    }

    #[test]
    fn the_last_sheet_cannot_be_removed() {
        let workspace = testing::workspace();
        create_workbook(
            &workspace,
            &Args::new(&json!({ "path": "one-sheet.xlsx", "sheets": ["Only"] })),
        )
        .unwrap();
        let args = Args::new(&json!({ "path": "one-sheet.xlsx", "sheet": "Only" }));
        let error = remove_sheet(&workspace, &args).unwrap_err();
        assert!(error.contains("at least one sheet"), "{error}");
    }

    #[test]
    fn renaming_onto_a_taken_name_gets_a_counter_appended() {
        let workspace = testing::workspace();
        add_sheet(
            &workspace,
            &Args::new(&json!({ "path": "report.xlsx", "title": "Extra" })),
        )
        .unwrap();
        // openpyxl does not reject a duplicate; it appends a counter, so the workbook keeps
        // a unique name for every sheet.
        let (_, renamed) = rename_sheet(
            &workspace,
            &Args::new(&json!({
                "path": "report.xlsx",
                "sheet": "Extra",
                "title": "Numbers",
            })),
        )
        .unwrap();
        assert_ne!(renamed["name"], json!("Numbers"));
        assert!(
            renamed["name"]
                .as_str()
                .unwrap()
                .starts_with("Numbers"),
            "the new name should still read as Numbers, got {renamed}"
        );
    }

    #[test]
    fn merging_and_unmerging_are_reversible() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "range": "A1:B1",
        }));
        merge_cells(&workspace, &args).unwrap();
        let (_, described) = inspect::describe_sheet(&workspace, &args).unwrap();
        assert_eq!(described["merged_ranges"], json!(["A1:B1"]));

        unmerge_cells(&workspace, &args).unwrap();
        let (_, described) = inspect::describe_sheet(&workspace, &args).unwrap();
        assert_eq!(described["merged_ranges"], json!([]));
    }

    #[test]
    fn freeze_panes_sets_and_clears() {
        let workspace = testing::workspace();
        let (_, frozen) = freeze_panes(
            &workspace,
            &Args::new(&json!({
                "path": "report.xlsx",
                "sheet": "Numbers",
                "cell": "B2",
            })),
        )
        .unwrap();
        assert_eq!(frozen["cell"], json!("B2"));

        let (summary, cleared) = freeze_panes(
            &workspace,
            &Args::new(&json!({ "path": "report.xlsx", "sheet": "Numbers" })),
        )
        .unwrap();
        assert!(summary.contains("unfroze"), "{summary}");
        assert_eq!(cleared["cell"], Value::Null);
    }

    #[test]
    fn the_autofilter_can_be_set_and_cleared() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "range": "A1:B2",
        }));
        let (_, filtered) = set_auto_filter(&workspace, &args).unwrap();
        assert_eq!(filtered["range"], json!("A1:B2"));

        let args = Args::new(&json!({ "path": "report.xlsx", "sheet": "Numbers" }));
        let (summary, cleared) = set_auto_filter(&workspace, &args).unwrap();
        assert!(summary.contains("cleared"), "{summary}");
        assert_eq!(cleared["range"], Value::Null);
    }

    #[test]
    fn named_ranges_can_be_added_once() {
        let workspace = testing::workspace();
        add_named_range(
            &workspace,
            &Args::new(&json!({
                "path": "report.xlsx",
                "name": "Extra",
                "sheet": "Numbers",
                "range": "A1:A5",
})),
        )
        .unwrap();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "name": "Extra",
            "sheet": "Numbers",
            "range": "B1:B5",
        }));
        let error = add_named_range(&workspace, &args).unwrap_err();
        assert!(error.contains("already"), "{error}");
    }

    #[test]
    fn clear_cells_empties_a_range() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "range": "A1:B1",
        }));
        let (_, cleared) = clear_cells(&workspace, &args).unwrap();
        assert_eq!(cleared["cells_cleared"], json!(2));

        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "range": "A1:B1",
        }));
        let (_, read_back) = inspect::read_cells(&workspace, &args).unwrap();
        assert_eq!(read_back["cells"][0]["value"], Value::Null);
    }
}
