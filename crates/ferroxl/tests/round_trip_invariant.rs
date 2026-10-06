//! The invariant: a workbook this library writes still opens, and its text is unchanged.
//!
//! Four bugs passed every name-matching test in this repository because each one produced a
//! well-formed file that silently meant something else. A chart or an image was deleted by the
//! next save, leaving a relationship pointing at a part that was not there, so the file could not
//! be opened at all. Every `&`, `<`, `>`, `"` and `'` was dropped from every element's text.
//! A cell holding an inline string came back empty. All four reported success, and all four were
//! invisible to a suite that checked names rather than files.
//!
//! So this file asserts the two properties that would have caught every one of them, on
//! workbooks built through the public API and then read back through the reader:
//!
//! 1. **the package is coherent** -- every relationship in it names a part that exists, and
//! 2. **the text survived** -- every character the writer was given is the character read back.
//!
//! Anything added to the writer should be added here.

use ferroxl::cell::cell::CellValue;
use ferroxl::charts::reference::{Reference, ReferenceDataType};
use ferroxl::charts::{BarChart, Series};
use ferroxl::comments::Comment;
use ferroxl::drawing::Image;
use ferroxl::reader::excel::load_workbook_from_bytes;
use ferroxl::workbook::Workbook;
use std::io::Cursor;

/// A 1x1 PNG, so the image path has real bytes to carry.
const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0xDA, 0x63, 0xF8, 0xCF, 0xC0, 0xF0,
    0x1F, 0x00, 0x05, 0x00, 0x01, 0xFF, 0x89, 0x99, 0x3D, 0x1D, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
    0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
];

/// Text chosen to cover every character XML has to escape, plus a non-ASCII one.
///
/// A string with an entity reference in it is the case that was broken twice over: the reader
/// discarded the reference, and an inline string carrying one came back empty.
const AWKWARD_TEXTS: &[&str] = &[
    "plain text",
    "Tom & Jerry",
    "a < b and c > d",
    "say \"hello\"",
    "it's fine",
    "5 < 6 && 7 > 6",
    "R&D / P&L",
    "<b>bold</b> & <i>italic</i>",
    "naïve café — ünïcode",
    "&amp; already escaped",
];

/// A workbook carrying a chart, an image, a comment, a header/footer and awkward text.
///
/// One workbook with everything, so one save has to satisfy both properties at once. The parts
/// that cannot be read back are the point: they are exactly the ones that used to be deleted.
fn decorated_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    workbook.create_sheet(Some("Notes")).expect("a sheet");

    {
        let sheet = workbook.active_sheet_mut().expect("an active sheet");
        sheet.set("A1", CellValue::text("Item")).expect("cell");
        sheet.set("A2", CellValue::number(10.0)).expect("cell");
        sheet.set("A3", CellValue::text("Bolt")).expect("cell");

        // Awkward text, one per row, so every escape is exercised in a real cell.
        for (offset, text) in AWKWARD_TEXTS.iter().enumerate() {
            let row = 10 + offset as u32;
            sheet
                .set(&format!("A{row}"), CellValue::text(*text))
                .expect("cell");
        }

        sheet
            .set("D2", CellValue::text("=SUM(B2:B2)"))
            .expect("cell");
        sheet.set_freeze_panes("A2");
        sheet.merge_cells("A5:C5").expect("merge");
        sheet
            .set_comment("D5", Some(Comment::new("a note & a <remark>", "tester")))
            .expect("comment");
        sheet.header_footer.set_header("&CR&D summary");

        // A chart and an image on separate sheets: two drawings, two parts.
        let mut chart = BarChart::new().into_chart();
        chart.add_series(Series::new(
            Reference::new(
                "Sheet1",
                (1, 0),
                Some((1, 0)),
                Some(ReferenceDataType::Numeric),
                None,
            )
            .expect("a reference"),
        ));
        sheet.charts.push(chart);

        let image = Image::from_png(PNG.to_vec()).expect("a png");
        workbook.worksheets[1].images.push(image);
    }

    workbook
}

/// Every relationship in `bytes` that names a part the package does not contain.
///
/// A dangling relationship is what made a workbook with a chart unopenable: the sheet's
/// `<drawing>` element and its relationship survived a save that emitted no drawing, so nothing
/// could resolve the target. `Target` is relative to the directory of the part declaring it, and
/// `_rels/.rels` sits at the package root.
fn dangling_relationships(bytes: &[u8]) -> Vec<String> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes.to_vec())).expect("a zip");
    let names: Vec<String> = (0..archive.len())
        .filter_map(|index| archive.by_index(index).ok().map(|e| e.name().to_string()))
        .collect();

    let mut dangling = Vec::new();
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).expect("entry");
        let name = entry.name().to_string();
        if !name.ends_with(".rels") {
            continue;
        }
        let mut text = String::new();
        std::io::Read::read_to_string(&mut entry, &mut text).expect("rels are utf-8");

        // `<dir>/_rels/<name>.rels` declares `<dir>/<name>`; `_rels/.rels` declares the package.
        let owner_directory = if name.starts_with("_rels/") {
            String::new()
        } else {
            let owner = name.replace("/_rels/", "/");
            match owner.strip_suffix(".rels").and_then(|o| o.rsplit_once('/')) {
                Some((directory, _)) => directory.to_string(),
                None => String::new(),
            }
        };

        for target in text
            .split("Target=\"")
            .skip(1)
            .map(|rest| rest.split('"').next().unwrap_or_default().to_string())
            .filter(|target| !target.starts_with("http") && !target.starts_with("mailto:"))
        {
            let resolved = if let Some(absolute) = target.strip_prefix('/') {
                absolute.to_string()
            } else {
                let mut segments: Vec<&str> = if owner_directory.is_empty() {
                    Vec::new()
                } else {
                    owner_directory.split('/').collect()
                };
                for piece in target.split('/') {
                    match piece {
                        ".." => {
                            segments.pop();
                        }
                        "." => {}
                        part => segments.push(part),
                    }
                }
                segments.join("/")
            };
            if !names.contains(&resolved) {
                dangling.push(format!("{name} -> {target} (as {resolved})"));
            }
        }
    }
    dangling
}

/// The invariant itself: save, then prove the package is coherent and the text survived.
#[test]
fn a_written_workbook_stays_openable_and_keeps_its_text() {
    let bytes = decorated_workbook().to_bytes().expect("saved");

    let dangling = dangling_relationships(&bytes);
    assert!(
        dangling.is_empty(),
        "every relationship must name a part that is in the package; these do not: {dangling:?}"
    );

    // The reader is the other half: if it cannot load the package, no consumer can.
    let workbook = load_workbook_from_bytes(bytes, Default::default()).expect("loadable");
    assert_eq!(
        workbook.get_sheet_names(),
        vec!["Sheet1".to_string(), "Notes".to_string()]
    );

    let sheet = workbook.active_sheet().expect("an active sheet");
    for (offset, expected) in AWKWARD_TEXTS.iter().enumerate() {
        let row = 10 + offset as u32;
        let coordinate = format!("A{row}");
        assert_eq!(
            text_of(sheet, &coordinate).as_deref(),
            Some(*expected),
            "{coordinate} should still read back as {expected:?} after a save"
        );
    }
}

/// A second save must not lose what the first one wrote.
///
/// The server saves per call, so this is the ordinary case rather than an edge: an agent edits a
/// workbook that already has a chart, and the drawing has to still be there afterwards. This is
/// the assertion that the chart deletion went without.
#[test]
fn a_second_save_does_not_lose_what_the_first_one_wrote() {
    let first = decorated_workbook().to_bytes().expect("saved");

    // Load it back -- which is where a chart or an image stops existing in the model -- edit one
    // unrelated cell, and save.
    let mut reloaded =
        load_workbook_from_bytes(first.clone(), Default::default()).expect("loadable");
    reloaded
        .active_sheet_mut()
        .expect("an active sheet")
        .set("Z9", CellValue::number(1.0))
        .expect("cell");
    let second = reloaded.to_bytes().expect("saved again");

    let dangling = dangling_relationships(&second);
    assert!(
        dangling.is_empty(),
        "a second save left relationships pointing at nothing: {dangling:?}"
    );

    let names = part_names(&second);
    assert!(
        names.contains(&"xl/drawings/drawing1.xml".to_string()),
        "the drawing survived the first save, so it must survive the second; the package has \
         {names:?}"
    );

    // And the text is still there too, which is the second thing a save can lose.
    let workbook = load_workbook_from_bytes(second, Default::default()).expect("loadable");
    let sheet = workbook.active_sheet().expect("an active sheet");
    assert_eq!(
        text_of(sheet, "A11").as_deref(),
        Some("Tom & Jerry"),
        "an unrelated edit must not cost a cell its ampersand"
    );
    assert_eq!(
        text_of(sheet, "A18").as_deref(),
        Some("naïve café — ünïcode"),
        "an unrelated edit must not cost a cell its non-ASCII text"
    );
}

/// Removing a sheet must not leave a trace of it, and must not cost the sheets that remain.
///
/// Two faults, both in the merge of a source's relationships with the writer's own. The dedup
/// compared target *strings*, so openpyxl's `/xl/worksheets/sheet1.xml` and the writer's
/// `worksheets/sheet1.xml` looked like two parts and both were written; the reader then built
/// two worksheets from one relationship set and reported a `Sheet2` that was not in the file.
/// And a preserved relationship to a sheet the writer renumbered away was kept, leaving the
/// package with a relationship naming a part that was not there -- which is the same
/// unopenable-file failure as the drawing bug above, reached through a different door.
///
/// The observable damage is that a removal followed by any ordinary edit *adds* a sheet: the
/// phantom is in the model by then, so the next save writes it, complete with a copy of its
/// neighbour's cells.
#[test]
fn removing_a_sheet_leaves_no_trace_and_costs_the_rest_nothing() {
    for victim in 0..3 {
        let mut source = Workbook::new();
        source.create_sheet(Some("B")).expect("a sheet");
        source.create_sheet(Some("C")).expect("a sheet");
        let keep = ["Sheet1", "B", "C"]
            .iter()
            .enumerate()
            .filter(|(index, _)| *index != victim)
            .map(|(_, name)| name.to_string())
            .collect::<Vec<_>>();

        let first = source.to_bytes().expect("saved");
        let mut reloaded = load_workbook_from_bytes(first, Default::default()).expect("loadable");
        reloaded
            .remove_sheet(victim)
            .expect("the sheet is there to remove");
        // An unrelated edit, which is what turns the phantom from a report into a file.
        reloaded
            .active_sheet_mut()
            .expect("an active sheet")
            .set("Z9", CellValue::text("edited"))
            .expect("cell");
        let second = reloaded.to_bytes().expect("saved again");

        let dangling = dangling_relationships(&second);
        assert!(
            dangling.is_empty(),
            "removing sheet {victim} left {dangling:?}"
        );

        let workbook = load_workbook_from_bytes(second, Default::default()).expect("loadable");
        assert_eq!(
            workbook.get_sheet_names(),
            keep,
            "removing sheet {victim} should leave exactly {keep:?}, with no sheet invented"
        );
        assert_eq!(
            text_of(workbook.active_sheet().expect("an active sheet"), "Z9").as_deref(),
            Some("edited"),
            "the edit that followed the removal has to be in the file"
        );
    }
}

/// A relationship into a part the writer does not own is the only thing keeping it reachable.
///
/// The counterpart to the test above, and the reason the filter cannot simply drop everything a
/// source wrote. A workbook with a drawing has to come back with the drawing *and* the
/// relationship naming it, or the part survives as anchors pointing at nothing.
#[test]
fn a_preserved_drawing_keeps_its_relationship_after_a_save() {
    // A drawing read from a file, not one built in memory: this is the case the writer cannot
    // reproduce, so the relationship has to be preserved rather than regenerated.
    let mut source = Workbook::new();
    source.create_sheet(Some("Other")).expect("a sheet");
    {
        let sheet = source.active_sheet_mut().expect("an active sheet");
        sheet.set("A1", CellValue::text("Item")).expect("cell");
        sheet
            .images
            .push(Image::from_png(PNG.to_vec()).expect("a png"));
    }
    let first = source.to_bytes().expect("saved");

    let mut reloaded = load_workbook_from_bytes(first, Default::default()).expect("loadable");
    reloaded
        .active_sheet_mut()
        .expect("an active sheet")
        .set("B2", CellValue::number(7.0))
        .expect("cell");
    let second = reloaded.to_bytes().expect("saved again");

    let names = part_names(&second);
    assert!(
        names.contains(&"xl/media/image1.png".to_string()),
        "the image itself must survive: {names:?}"
    );

    // And the sheet still says which drawing holds it.
    let mut archive = zip::ZipArchive::new(Cursor::new(second)).expect("a zip");
    let mut sheet_xml = String::new();
    {
        let mut sheet = archive
            .by_name("xl/worksheets/sheet1.xml")
            .expect("the first sheet");
        std::io::Read::read_to_string(&mut sheet, &mut sheet_xml).expect("utf-8");
    }
    // And the sheet still says which drawing holds it. Matched with or without a namespace
    // prefix, because the preserved element is written back under its braced namespace and the
    // serialiser spells that as `<s:drawing>` rather than `<drawing>`.
    assert!(
        sheet_xml.contains("<drawing") || sheet_xml.contains(":drawing"),
        "the sheet must still reference its drawing; it has {sheet_xml:?}"
    );
}

/// Every save has to be a fixed point of the text, not merely survive one.
///
/// A single check catches a regression that drops text once; this one catches anything that
/// degrades it a little more on each pass, which is the shape a partly-wrong escape handler has.
#[test]
fn text_is_stable_across_repeated_saves() {
    let mut bytes = decorated_workbook().to_bytes().expect("saved");
    for pass in 1..=3 {
        let mut workbook = load_workbook_from_bytes(bytes, Default::default()).expect("loadable");
        workbook
            .active_sheet_mut()
            .expect("an active sheet")
            .set(&format!("P{pass}"), CellValue::text("touch"))
            .expect("cell");
        bytes = workbook.to_bytes().expect("saved");

        let dangling = dangling_relationships(&bytes);
        assert!(dangling.is_empty(), "pass {pass}: {dangling:?}");

        let workbook =
            load_workbook_from_bytes(bytes.clone(), Default::default()).expect("loadable");
        let sheet = workbook.active_sheet().expect("an active sheet");
        for (offset, expected) in AWKWARD_TEXTS.iter().enumerate() {
            let coordinate = format!("A{}", 10 + offset as u32);
            assert_eq!(
                text_of(sheet, &coordinate).as_deref(),
                Some(*expected),
                "pass {pass}: {coordinate} degraded from {expected:?}"
            );
        }
    }
}

/// Inline strings have to survive a save, which is the shape openpyxl writes.
///
/// openpyxl puts a string in `<is><t>` when the workbook has no shared string table, and that is
/// the default for a file it has just created. The writer used to emit the text into `<v>` under
/// an `inlineStr` type -- a file no reader can get a value out of -- so every string cell in such
/// a workbook came back blank after one save.
/// A printed header has to survive a save, and so has the `&` inside its codes.
///
/// Two faults, in series. The reader's section parser compared whole `&`-delimited fields, so it
/// missed the marker when it was glued to its text -- `&Lleft` rather than `&L` then `left` --
/// which is the shape openpyxl and Excel both write. And with the entity references discarded
/// before the parser saw them, `&CR&D` arrived as `CR&D` and matched nothing either. Fixing
/// either alone still loses the header, which is why this asserts the whole thing: written by the
/// writer, read by the reader, kept through a further save.
#[test]
fn a_header_or_footer_survives_a_save() {
    // `&C` introduces the centre section, `&R` the right, and `&P`/`&N` are the page number and
    // page count -- an ampersand in the text and two in the codes.
    let header = "&CRegional & D&D summary";
    let footer = "&RPage &P of &N";

    let mut workbook = Workbook::new();
    {
        let sheet = workbook.active_sheet_mut().expect("an active sheet");
        sheet.set("A1", CellValue::text("x")).expect("cell");
        sheet.header_footer.set_header(header);
        sheet.header_footer.set_footer(footer);
    }

    let first = workbook.to_bytes().expect("saved");
    let reloaded = load_workbook_from_bytes(first, Default::default()).expect("loadable");
    let read = reloaded
        .active_sheet()
        .expect("an active sheet")
        .header_footer
        .center_header
        .text
        .clone();
    assert_eq!(
        read.as_deref(),
        Some("Regional & D&D summary"),
        "a centre header set as {header:?} should read back whole"
    );

    // And through a further save, which is where it used to go: one save kept it, the next did not.
    let mut again = reloaded;
    again
        .active_sheet_mut()
        .expect("an active sheet")
        .set("Z9", CellValue::number(1.0))
        .expect("cell");
    let second = again.to_bytes().expect("saved again");
    let after = load_workbook_from_bytes(second, Default::default()).expect("loadable");
    let footer_text = after
        .active_sheet()
        .expect("an active sheet")
        .header_footer
        .right_footer
        .text
        .clone();
    assert_eq!(
        footer_text.as_deref(),
        Some("Page &P of &N"),
        "the footer has to survive a save as well as the header"
    );
}

#[test]
fn an_inline_string_survives_a_save() {
    // `Workbook::new` writes through the shared string table, so the inline form has to be put
    // in the way openpyxl writes it: `<is><t>` rather than `<v>`.
    let with_inline = workbook_with_cell_body(
        r#"<row r="1" spans="1:1"><c r="A1" t="inlineStr"><is><t>inline &amp; text</t></is></c></row>"#,
    );

    let workbook = load_workbook_from_bytes(with_inline, Default::default()).expect("loadable");
    let sheet = workbook.active_sheet().expect("an active sheet");
    assert_eq!(
        text_of(sheet, "A1").as_deref(),
        Some("inline & text"),
        "an inline string carries its text in <is><t> and must read back whole"
    );

    // And it has to survive being written back out, which is where the two sides used to
    // disagree: the reader looked in `<v>` and the writer put the text there.
    let again = workbook.to_bytes().expect("saved again");
    let names = part_names(&again);
    let mut archive = zip::ZipArchive::new(Cursor::new(again)).expect("a zip");
    let mut sheet_xml = String::new();
    {
        let mut sheet = archive
            .by_name("xl/worksheets/sheet1.xml")
            .expect("the sheet");
        std::io::Read::read_to_string(&mut sheet, &mut sheet_xml).expect("utf-8");
    }
    assert!(
        sheet_xml.contains("<is>") && sheet_xml.contains("inline"),
        "an inline string must be written back as one, not into <v>; the sheet has {sheet_xml:?} \
         and the package has {names:?}"
    );
}

/// A one-cell workbook whose cell is replaced by `cell`, for the inline-string case above.
///
/// The splice replaces the first `<c>` element and the row it sits in, so the result is a
/// well-formed sheet rather than a cell dropped into an existing row.
fn workbook_with_cell_body(cell: &str) -> Vec<u8> {
    let mut workbook = Workbook::new();
    workbook
        .active_sheet_mut()
        .expect("an active sheet")
        .set("A1", CellValue::text("plain"))
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
            std::io::Read::read_to_end(&mut entry, &mut data).expect("readable");
            if name == "xl/worksheets/sheet1.xml" {
                let text = String::from_utf8(data).expect("utf-8");
                let row_start = text.find("<row ").expect("a row");
                let row_end = text[row_start..].find("</row>").expect("a row end")
                    + row_start
                    + "</row>".len();
                data = format!("{}{}{}", &text[..row_start], cell, &text[row_end..]).into_bytes();
            }
            target
                .start_file(name, zip::write::SimpleFileOptions::default())
                .expect("writable");
            std::io::Write::write_all(&mut target, &data).expect("writable");
        }
        target.finish().expect("finished");
    }
    out.into_inner()
}

/// Every part name in the package.
fn part_names(bytes: &[u8]) -> Vec<String> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes.to_vec())).expect("a zip");
    (0..archive.len())
        .filter_map(|index| archive.by_index(index).ok().map(|e| e.name().to_string()))
        .collect()
}

/// The text of a cell, or `None` if it holds nothing.
///
/// `cell_value` is the public accessor and returns the whole `CellValue`, so comparing against a
/// `&str` is what the other tests do; this is the same thing with the wrapping in one place.
fn text_of(sheet: &ferroxl::worksheet::Worksheet, coordinate: &str) -> Option<String> {
    match sheet.cell_value(coordinate)? {
        CellValue::Text(text) => Some(text),
        other => Some(format!("{other:?}")),
    }
}
