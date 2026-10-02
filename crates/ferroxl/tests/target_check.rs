//! Relationship targets are spelled several ways in the wild. Getting one wrong loses a whole
//! sheet silently, so these build real archives rather than testing the resolver in isolation.

use ferroxl::cell::cell::CellValue;
use ferroxl::reader::excel::load_workbook_from_bytes;
use ferroxl::workbook::Workbook;
use std::io::{Cursor, Read, Write};

/// A two-sheet workbook with its workbook relationships rewritten by `rewrite_rels`.
///
/// Rewriting through a real zip is the point: the resolver is only exercised properly when the
/// `[Content_Types].xml` and `.rels` parts disagree in the way a third-party generator's do.
fn workbook_with_rels(rewrite: impl Fn(&str) -> String) -> Vec<u8> {
    let mut workbook = Workbook::new();
    workbook.create_sheet(Some("Second")).expect("sheet");
    workbook
        .active_sheet_mut()
        .expect("active")
        .set("A1", CellValue::number(42.0))
        .expect("cell");
    let bytes = workbook.to_bytes().expect("saved");

    let mut source = zip::ZipArchive::new(Cursor::new(bytes)).expect("a zip");
    let mut out = Cursor::new(Vec::new());
    {
        let mut target = zip::ZipWriter::new(&mut out);
        for index in 0..source.len() {
            let mut entry = source.by_index(index).expect("entry");
            let name = entry.name().to_string();
            let mut data = Vec::new();
            entry.read_to_end(&mut data).expect("readable");
            if name == "xl/_rels/workbook.xml.rels" {
                let text = String::from_utf8(data).expect("rels are utf-8");
                data = rewrite(&text).into_bytes();
            }
            target
                .start_file(name, zip::write::SimpleFileOptions::default())
                .expect("writable");
            target.write_all(&data).expect("writable");
        }
        target.finish().expect("finished");
    }
    out.into_inner()
}

#[test]
fn the_usual_relative_target_is_found() {
    let bytes = workbook_with_rels(|rels| rels.to_string());
    let workbook = load_workbook_from_bytes(bytes, Default::default()).expect("loadable");
    assert_eq!(
        workbook.get_sheet_names(),
        vec!["Sheet1".to_string(), "Second".to_string()]
    );
    assert_eq!(
        workbook
            .active_sheet()
            .expect("active")
            .cell_value("A1")
            .map(|value| format!("{value:?}")),
        Some("Number(42.0)".to_string())
    );
}

#[test]
fn an_absolute_target_is_found() {
    // Before the resolver, `xl/` was prepended unconditionally and this produced
    // `xl//xl/worksheets/sheet1.xml`: no content-type match, no sheets, no error.
    let bytes =
        workbook_with_rels(|rels| rels.replace("Target=\"worksheets/", "Target=\"/xl/worksheets/"));
    let workbook = load_workbook_from_bytes(bytes, Default::default()).expect("loadable");
    assert_eq!(
        workbook.get_sheet_names(),
        vec!["Sheet1".to_string(), "Second".to_string()],
        "an absolute target names the same part as a relative one"
    );
}

#[test]
fn a_target_that_climbs_back_out_of_xl_is_found() {
    let bytes = workbook_with_rels(|rels| {
        rels.replace("Target=\"worksheets/", "Target=\"../xl/worksheets/")
    });
    let workbook = load_workbook_from_bytes(bytes, Default::default()).expect("loadable");
    assert_eq!(workbook.get_sheet_names().len(), 2);
}

#[test]
fn a_target_with_detours_is_found() {
    let bytes = workbook_with_rels(|rels| {
        rels.replace(
            "Target=\"worksheets/",
            "Target=\"/xl/./worksheets/../worksheets/",
        )
    });
    let workbook = load_workbook_from_bytes(bytes, Default::default()).expect("loadable");
    assert_eq!(workbook.get_sheet_names().len(), 2);
}
