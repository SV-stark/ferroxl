//! Fixtures shared by the handler tests.
//!
//! Each fixture gets its own directory under the system temp folder so the tests can run
//! in parallel without one overwriting another's workbook.

#![cfg(test)]

use std::sync::atomic::{AtomicU32, Ordering};

use lexcel::{CellValue, Style, Workbook};

use crate::workspace::Workspace;

/// A counter that gives each test its own directory.
///
/// The tests run in parallel and several of them write to the same file names, so a shared
/// directory would let one test delete another's workbook mid-run.
static NEXT_ID: AtomicU32 = AtomicU32::new(0);

/// An empty directory that no other test is using.
///
/// Callers name it so a failure message says which fixture set produced it.
pub fn empty_workspace(label: &str) -> Workspace {
    let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
    let root = std::env::temp_dir().join(format!("lexcel-mcp-{label}-{id}"));
    let _ = std::fs::remove_dir_all(&root);
    Workspace::new(root).expect("a fresh directory")
}

/// A workspace containing the standard fixture workbooks.
pub fn workspace() -> Workspace {
    let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
    let root = std::env::temp_dir().join(format!("lexcel-mcp-fixtures-{id}"));
    let _ = std::fs::remove_dir_all(&root);
    let workspace = Workspace::new(&root).expect("fixture root");

    workspace.save(report_workbook(), "report.xlsx").expect("report");
    workspace
        .save(numbers_workbook(), "numbers.xlsx")
        .expect("numbers");
    workspace.save(named_workbook(), "names.xlsx").expect("names");
    workspace
}

/// A style that shows a percentage in two decimal places.
fn percent_style() -> Style {
    let mut style = Style::new();
    style.number_format.set_format_code("0.00%");
    style
}

/// The workbook most tests read: two sheets with a mix of value types.
///
/// The merged range is empty on purpose. openpyxl replaces the non-top-left cells of a
/// merge with blank merged cells, so merging over data would erase it.
pub fn report_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    let sheet = workbook.active_sheet_mut().expect("the default sheet");
    sheet.set_title("Data", &[]).expect("a valid title");
    sheet.set("A1", CellValue::text("alpha")).unwrap();
    sheet.set("B1", 10.0).unwrap();
    sheet.set("C1", 20.0).unwrap();
    sheet.set("A2", CellValue::text("beta")).unwrap();
    sheet
        .set("B2", CellValue::Formula("=SUM(B1:B1)".to_string()))
        .unwrap();
    sheet.set("A4", 0.125).unwrap();
    sheet.set_style("A4", percent_style()).unwrap();
    sheet.merge_cells("E1:F1").unwrap();

    workbook
        .create_sheet(Some("Numbers"))
        .expect("second sheet");
    workbook.worksheets[1].set("A1", 1.0).unwrap();
    workbook.worksheets[1].set("B1", 2.0).unwrap();
    workbook.worksheets[1].set("A2", 3.0).unwrap();
    workbook.worksheets[1].set("B2", 4.0).unwrap();
    workbook
}

/// A workbook whose only sheet is a plain 2x2 grid.
pub fn numbers_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    workbook.worksheets[0].set("A1", 1.0).unwrap();
    workbook.worksheets[0].set("B1", 2.0).unwrap();
    workbook.worksheets[0].set("A2", 3.0).unwrap();
    workbook.worksheets[0].set("B2", 4.0).unwrap();
    workbook
}

/// A workbook with a workbook-scoped named range.
///
/// `create_named_range` takes a bare range and prepends the sheet title itself, matching
/// openpyxl, so `"$A$1:$A$10"` is what belongs here.
pub fn named_workbook() -> Workbook {
    let mut workbook = numbers_workbook();
    let data = workbook.create_sheet(Some("Data")).expect("Data sheet");
    workbook
        .create_named_range("Totals", data, "$A$1:$A$10", None)
        .expect("named range");
    workbook
}
