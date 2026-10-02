//! An interrupted download is the commonest way a workbook arrives damaged: the central
//! directory is written before the end-of-central-directory record, so a download that stops
//! early usually leaves a recoverable archive.

use ferroxl::cell::cell::CellValue;
use ferroxl::reader::excel::load_workbook_from_bytes;
use ferroxl::workbook::Workbook;

fn sample() -> Vec<u8> {
    let mut workbook = Workbook::new();
    workbook.create_sheet(Some("Second")).expect("sheet");
    let sheet = workbook.active_sheet_mut().expect("active");
    sheet.set("A1", CellValue::number(42.0)).expect("cell");
    sheet
        .set("B2", CellValue::Text("kept".to_string()))
        .expect("cell");
    workbook.to_bytes().expect("saved")
}

fn sheet_names(bytes: Vec<u8>) -> Vec<String> {
    load_workbook_from_bytes(bytes, Default::default())
        .expect("loadable")
        .get_sheet_names()
}

#[test]
fn an_intact_workbook_loads() {
    assert_eq!(sheet_names(sample()), vec!["Sheet1", "Second"]);
}

#[test]
fn a_workbook_whose_trailing_record_was_cut_off_still_loads() {
    let full = sample();
    // Every cut inside the final 22-byte record. A cut of 30 would reach 8 bytes into the
    // central directory, which is a different failure entirely -- see the refusal test below.
    for cut in [1usize, 4, 8, 16, 21, 22] {
        let truncated = full[..full.len() - cut].to_vec();
        assert_eq!(
            sheet_names(truncated),
            vec!["Sheet1".to_string(), "Second".to_string()],
            "{cut} bytes cut off the end"
        );
    }
}

#[test]
fn the_recovered_cells_hold_their_values() {
    // Repairing the archive is not enough if the parts then read back empty: the test that
    // matters is whether the numbers are still there.
    let full = sample();
    let truncated = full[..full.len() - 9].to_vec();
    let workbook = load_workbook_from_bytes(truncated, Default::default()).expect("loadable");
    let sheet = workbook.active_sheet().expect("active");
    assert_eq!(
        sheet.cell_value("A1").map(|value| format!("{value:?}")),
        Some("Number(42.0)".to_string())
    );
    assert!(sheet.cell_value("B2").is_some(), "the string cell survived");
}

#[test]
fn a_workbook_truncated_further_is_refused_rather_than_half_read() {
    // Past the central directory there is nothing left to describe the archive. Refusing is
    // the honest outcome; returning a workbook missing sheets would not be.
    let full = sample();
    let truncated = full[..full.len() - 200].to_vec();
    assert!(load_workbook_from_bytes(truncated, Default::default()).is_err());
}

#[test]
fn a_file_that_is_not_a_zip_is_refused() {
    let error = load_workbook_from_bytes(b"<html>404</html>".to_vec(), Default::default())
        .expect_err("not a workbook");
    assert!(
        error.to_string().contains("truncated beyond recovery"),
        "unexpected error: {error}"
    );
}

#[test]
fn trailing_junk_was_never_a_problem() {
    // Recorded because the roadmap assumed it was: it is not, and the repair does not exist to
    // solve it. If this ever fails, the zip crate's tolerance changed and the archive module
    // needs to know.
    let mut bytes = sample();
    bytes.extend_from_slice(&[0u8; 4096]);
    assert_eq!(sheet_names(bytes), vec!["Sheet1", "Second"]);
}
