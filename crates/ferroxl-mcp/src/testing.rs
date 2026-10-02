//! Fixtures shared by the handler tests.
//!
//! Each fixture gets its own directory under the system temp folder so the tests can run
//! in parallel without one overwriting another's workbook.

#![cfg(test)]

use std::sync::atomic::{AtomicU32, Ordering};

use ferroxl::{CellValue, Style, Workbook};

use crate::workspace::Workspace;

/// A counter that separates the fixtures created within one process.
static NEXT_ID: AtomicU32 = AtomicU32::new(0);

/// A directory name no other test is using.
///
/// The process id matters more than it looks. `cargo test` runs a binary's tests as threads
/// in a single process, so a counter alone is enough to keep them apart. `cargo nextest`
/// runs each test in its own process, where every counter starts again at zero — so all of
/// them would pick the same directory, wipe it on the way in, and delete each other's
/// workbooks mid-test. Mixing in the process id makes the name unique across both runners,
/// and cheap: the OS reclaims the directory when the process exits.
fn scratch(label: &str) -> std::path::PathBuf {
    let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
    std::env::temp_dir().join(format!("ferroxl-mcp-{label}-{}-{id}", std::process::id()))
}

/// An empty directory that no other test is using.
///
/// Callers name it so a failure message says which fixture set produced it.
pub fn empty_workspace(label: &str) -> Workspace {
    let root = scratch(label);
    let _ = std::fs::remove_dir_all(&root);
    Workspace::new(root).expect("a fresh directory")
}

/// A workspace containing the standard fixture workbooks.
pub fn workspace() -> Workspace {
    let root = scratch("fixtures");
    let _ = std::fs::remove_dir_all(&root);
    let workspace = Workspace::new(&root).expect("fixture root");

    workspace
        .save(report_workbook(), "report.xlsx")
        .expect("report");
    workspace
        .save(numbers_workbook(), "numbers.xlsx")
        .expect("numbers");
    workspace
        .save(named_workbook(), "names.xlsx")
        .expect("names");
    workspace
        .save(chain_workbook(), "chain.xlsx")
        .expect("chain");
    workspace
        .save(cycle_workbook(), "cycle.xlsx")
        .expect("cycle");
    workspace
}

/// A workbook with a dependency chain: `A4` totals `A1:A3` and `A5` doubles `A4`.
///
/// The tracing tools need a graph with a branch and a chain in it, not just a single
/// formula, or they would pass on inputs that say nothing about order.
pub fn chain_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    let sheet = workbook.active_sheet_mut().expect("the default sheet");
    for (coordinate, value) in [("A1", 1), ("A2", 2), ("A3", 3)] {
        sheet.set(coordinate, value).expect("value");
    }
    sheet
        .set("A4", CellValue::formula("SUM(A1:A3)"))
        .expect("A4");
    sheet.set("A5", CellValue::formula("A4*2")).expect("A5");
    sheet.set("B1", CellValue::formula("A1+A2")).expect("B1");
    workbook
}

/// A workbook whose three formulas reference each other in a loop.
pub fn cycle_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    let sheet = workbook.active_sheet_mut().expect("the default sheet");
    for (coordinate, formula) in [("A1", "B1"), ("B1", "C1"), ("C1", "A1")] {
        sheet
            .set(coordinate, CellValue::formula(formula))
            .expect("formula");
    }
    workbook
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
