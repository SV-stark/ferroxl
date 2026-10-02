//! Build a small workbook and read it back.
//!
//! Run with `cargo run --example build_and_read`. It writes `orders.xlsx` in the working
//! directory and then loads it, which is the shortest complete round trip through the
//! library: create a sheet, set cells, style a header, merge a range, freeze the panes,
//! save, and load.

use lexcel::{CellValue, Color, Fill, Font, Style, Workbook};

fn main() -> Result<(), lexcel::Error> {
    let path = "orders.xlsx";
    write(path)?;
    read(path)?;
    Ok(())
}

/// Build the workbook and write it to `path`.
fn write(path: &str) -> Result<(), lexcel::Error> {
    let mut workbook = Workbook::new();

    // `Workbook::new()` already made a sheet called `Sheet1`, and `create_sheet` appends
    // without making the new sheet active — openpyxl behaves the same way. So ask for the
    // sheet by the index `create_sheet` returns rather than reaching for `active_sheet`,
    // which would still be `Sheet1`.
    let orders = workbook.create_sheet(Some("Orders"))?;

    // Four columns, two data rows. A value that starts with `=` is a formula and one that
    // parses as a number is a number, which is the same rule the MCP server applies.
    let rows: [&[&str]; 3] = [
        &["Item", "Qty", "Price", "Total"],
        &["Bolt", "10", "1.5", "=B2*C2"],
        &["Nut", "25", "0.75", "=B3*C3"],
    ];
    let sheet = &mut workbook.worksheets[orders];
    for (row_offset, row) in rows.iter().enumerate() {
        for (column_offset, text) in row.iter().enumerate() {
            let letter = (b'A' + column_offset as u8) as char;
            let coordinate = format!("{letter}{}", row_offset + 1);
            sheet.set(&coordinate, parse(text))?;
        }
    }

    let header = Style {
        font: Font::new()
            .with_bold(true)
            .with_color(Color::new("FFFFFFFF".to_string())),
        fill: Fill::solid(Color::new("FF1F3864".to_string())),
        ..Style::default()
    };
    for letter in ["A", "B", "C", "D"] {
        sheet.set_style(&format!("{letter}1"), header.clone())?;
    }

    // `merge_cells` blanks every cell but the top-left one, so merge before writing the
    // caption or the caption disappears. It takes one coordinate at a time.
    sheet.merge_cells("A6:D6")?;
    sheet.set("A6", CellValue::text("Two line items"))?;
    sheet.set_freeze_panes("A2");
    for coordinate in ["B2", "C2", "B3", "C3"] {
        sheet.set_number_format(coordinate, "0.00")?;
    }

    workbook.save(path)?;
    println!("wrote {path}");
    Ok(())
}

/// Load the workbook and print what it holds.
fn read(path: &str) -> Result<(), lexcel::Error> {
    let workbook = lexcel::load_workbook(path, lexcel::LoadOptions::default())?;
    println!("sheets: {:?}", workbook.get_sheet_names());

    let sheet = workbook
        .get_sheet_by_name("Orders")
        .expect("the Orders sheet");
    println!("dimension: {}", sheet.calculate_dimension()?);
    println!("merged: {:?}", sheet.merged_cells());
    println!("frozen at: {:?}", sheet.freeze_panes);

    for row in sheet.range_values("A1:D3")? {
        let cells: Vec<String> = row.iter().map(|value| format!("{value:?}")).collect();
        println!("  {}", cells.join(" | "));
    }

    // The header style survives the round trip, and it is one style-table entry rather
    // than four.
    println!("A1 bold: {}", sheet.get_style("A1").font.bold);
    Ok(())
}

/// The value mapping shared with the MCP server: `=` is a formula, a bare number is a
/// number, and anything else is text.
fn parse(text: &str) -> CellValue {
    if let Some(expression) = text.strip_prefix('=') {
        CellValue::formula(expression)
    } else if let Ok(number) = text.parse::<f64>() {
        CellValue::number(number)
    } else {
        CellValue::text(text)
    }
}
