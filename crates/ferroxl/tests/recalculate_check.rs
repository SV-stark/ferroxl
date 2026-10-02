//! A formula cell's value has to survive the round trip, or `recalculate` has bought nothing.
//!
//! Checked against the written XML rather than ferroxl's own reader, because ferroxl reading
//! back what it wrote only proves it is self-consistent. `data_only=True` in openpyxl is the
//! closest equivalent to what an agent reading the file will see.

use ferroxl::cell::cell::CellValue;
use ferroxl::workbook::Workbook;
use std::io::{Cursor, Read};

fn sample() -> Vec<u8> {
    let mut workbook = Workbook::new();
    let sheet = workbook.active_sheet_mut().expect("sheet");
    for row in 1..=3 {
        sheet
            .set(&format!("A{row}"), CellValue::Number(row as f64))
            .expect("cell");
    }
    sheet
        .set("B1", CellValue::Text("hello".into()))
        .expect("cell");
    // One of each result kind, so all four `<v>` spellings are exercised.
    sheet
        .set("C1", CellValue::formula("=SUM(A1:A3)"))
        .expect("cell");
    sheet
        .set("C2", CellValue::formula("=UPPER(B1)"))
        .expect("cell");
    sheet.set("C3", CellValue::formula("=A1>0")).expect("cell");
    sheet.set("C4", CellValue::formula("=1/0")).expect("cell");
    // A formula this module will not evaluate: it must get no value at all.
    sheet
        .set("C5", CellValue::formula("=VLOOKUP(A1,A1:A3,1)"))
        .expect("cell");

    let report = workbook.recalculate();
    assert_eq!(report.computed_count(), 4, "four formulas evaluated");
    assert_eq!(report.unresolved.len(), 1, "one refused");
    assert_eq!(
        report.unresolved[0].reason,
        "unsupported function VLOOKUP()"
    );
    assert_eq!(
        workbook.calculation.full_calc_on_load,
        Some(true),
        "Excel has to be told to recompute, or a wrong value here would persist"
    );
    workbook.to_bytes().expect("saved")
}

fn sheet_xml(bytes: &[u8]) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes.to_vec())).expect("a zip");
    let mut entry = archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("the sheet part");
    let mut text = String::new();
    entry.read_to_string(&mut text).expect("utf-8");
    text
}

fn part_of(xml: &str, coordinate: &str) -> String {
    let at = xml
        .find(&format!("r=\"{coordinate}\""))
        .unwrap_or_else(|| panic!("no cell {coordinate} in {xml}"));
    let end = xml[at..]
        .find("</c>")
        .map(|n| at + n + 4)
        .unwrap_or(at + 200);
    xml[at..end].to_string()
}

#[test]
fn a_numeric_result_is_written_without_a_type_attribute() {
    let part = part_of(&sheet_xml(&sample()), "C1");
    assert!(part.contains("<v>6</v>"), "got {part}");
    assert!(
        !part.contains("t=\""),
        "a number is the default and needs no t: {part}"
    );
}

#[test]
fn a_text_result_is_marked_as_a_formula_string() {
    let part = part_of(&sheet_xml(&sample()), "C2");
    // `str`, not the shared-string `s`: a computed string was never put in the string table.
    assert!(part.contains("t=\"str\""), "got {part}");
    assert!(part.contains("<v>HELLO</v>"), "got {part}");
}

#[test]
fn a_boolean_result_is_written_as_one_or_zero() {
    let part = part_of(&sheet_xml(&sample()), "C3");
    assert!(part.contains("t=\"b\""), "got {part}");
    assert!(part.contains("<v>1</v>"), "got {part}");
}

#[test]
fn an_error_result_is_written_as_the_error_code() {
    let part = part_of(&sheet_xml(&sample()), "C4");
    assert!(part.contains("t=\"e\""), "got {part}");
    assert!(part.contains("<v>#DIV/0!</v>"), "got {part}");
}

#[test]
fn a_formula_that_could_not_be_evaluated_gets_no_value_at_all() {
    let part = part_of(&sheet_xml(&sample()), "C5");
    // An empty `<v/>` rather than a made-up number: this is the whole contract.
    assert!(
        part.contains("<v/>") || part.contains("<v></v>"),
        "got {part}"
    );
    assert!(
        !part.contains("t=\""),
        "no type is claimed for no value: {part}"
    );
}

#[test]
fn the_values_survive_ferroxls_own_reader() {
    let bytes = sample();
    let workbook = ferroxl::reader::excel::load_workbook_from_bytes(bytes, Default::default())
        .expect("loadable");
    let sheet = workbook.active_sheet().expect("sheet");
    assert_eq!(sheet.cached_value("C1"), Some(&CellValue::Number(6.0)));
    assert_eq!(
        sheet.cached_value("C2"),
        Some(&CellValue::Text("HELLO".to_string()))
    );
    assert_eq!(sheet.cached_value("C3"), Some(&CellValue::Bool(true)));
    assert_eq!(
        sheet.cached_value("C4"),
        Some(&CellValue::Error("#DIV/0!".to_string()))
    );
    assert_eq!(
        sheet.cached_value("C5"),
        None,
        "the refused formula has no value"
    );
}

#[test]
fn a_workbook_without_recalculate_still_writes_an_empty_value() {
    // The change must not make every existing formula cell claim a value it does not have.
    let mut workbook = Workbook::new();
    workbook
        .active_sheet_mut()
        .expect("sheet")
        .set("A1", CellValue::Number(1.0))
        .expect("cell");
    workbook
        .active_sheet_mut()
        .expect("sheet")
        .set("A2", CellValue::formula("=A1+1"))
        .expect("cell");
    let part = part_of(&sheet_xml(&workbook.to_bytes().expect("saved")), "A2");
    assert!(
        part.contains("<f>A1+1</f>"),
        "the formula is still there: {part}"
    );
    assert!(
        part.contains("<v></v>") || part.contains("<v/>"),
        "and still has no value: {part}"
    );
}

#[test]
fn recalculating_twice_is_stable() {
    // The snapshot is the reason: the second pass must not read the first pass's output.
    let mut workbook = Workbook::new();
    let sheet = workbook.active_sheet_mut().expect("sheet");
    sheet.set("A1", CellValue::Number(2.0)).expect("cell");
    sheet.set("A2", CellValue::formula("=A1*3")).expect("cell");
    let first = workbook.recalculate();
    let second = workbook.recalculate();
    assert_eq!(first.computed_count(), second.computed_count());
    assert_eq!(
        workbook.cached_value("Sheet1", "A2"),
        Some(&CellValue::Number(6.0))
    );
}
