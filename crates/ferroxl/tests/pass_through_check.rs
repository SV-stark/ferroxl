//! The test that matters: does an enterprise workbook come back with its parts intact?
//!
//! Built by hand rather than from a fixture file, because the parts under test are precisely
//! the ones no fixture in this repository has. A pivot table needs three parts, three content
//! types, two relationships and two referencing elements to be real, and a fixture that is
//! missing any of those would test a weaker thing than it looks like.

use ferroxl::reader::excel::load_workbook_from_bytes;
use ferroxl::workbook::Workbook;
use std::io::{Cursor, Read, Write};

const PKG_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SHEET_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// A workbook carrying a pivot table, a slicer cache, an ActiveX control's VML, threaded
/// comments, a query table and a custom XML part -- none of which ferroxl models.
///
/// Every one of them is a part, a content type, a relationship and a referencing element, which
/// is the full set that has to travel together for any of it to still work.
fn enterprise_workbook() -> Vec<u8> {
    let mut buffer = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut buffer);
        let options: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default();

        let put = |zip: &mut zip::ZipWriter<&mut Cursor<Vec<u8>>>, name: &str, body: String| {
            zip.start_file(name.to_string(), options).expect("writable");
            zip.write_all(body.as_bytes()).expect("writable");
        };

        put(
            &mut zip,
            "[Content_Types].xml",
            r#"<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="{REL_NS}"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.activeX"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/queryTables/queryTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>
  <Override PartName="/customXml/item1.xml" ContentType="application/xml"/>
</Types>"#.to_string(),
        );

        put(
            &mut zip,
            "_rels/.rels",
            format!(
                r#"<?xml version="1.0"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="{OFFICE_REL}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
            ),
        );

        // The pivot cache lives at rId2, which is exactly the id the writer will hand to
        // `sharedStrings.xml` -- so preservation has to move it and the reference with it.
        put(
            &mut zip,
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<?xml version="1.0"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="{OFFICE_REL}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="{OFFICE_REL}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
  <Relationship Id="rId3" Type="{OFFICE_REL}/queryTable" Target="queryTables/queryTable1.xml"/>
</Relationships>"#
            ),
        );

        put(
            &mut zip,
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0"?>
<workbook xmlns="{SHEET_NS}" xmlns:r="{OFFICE_REL}">
  <sheets><sheet name="Report" sheetId="1" r:id="rId1"/></sheets>
  <pivotCaches><pivotCache cacheId="0" r:id="rId2"/></pivotCaches>
</workbook>"#,
            ),
        );

        put(
            &mut zip,
            "xl/styles.xml",
            format!(
                r#"<?xml version="1.0"?><styleSheet xmlns="{SHEET_NS}">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>"#,
            ),
        );

        put(
            &mut zip,
            "xl/sharedStrings.xml",
            format!(r#"<?xml version="1.0"?><sst xmlns="{SHEET_NS}" count="0"/>"#),
        );

        put(
            &mut zip,
            "xl/theme/theme1.xml",
            r#"<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"/>"#
                .to_string(),
        );

        // The sheet names the pivot table, an ActiveX control through its VML, and a legacy
        // drawing -- all children the writer rebuilds, so all three are the preserved ones.
        put(
            &mut zip,
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<?xml version="1.0"?>
<worksheet xmlns="{SHEET_NS}" xmlns:r="{OFFICE_REL}">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1"><v>42</v></c><c r="B1"><v>100</v></c></row>
    <row r="2"><c r="A2"><v>7</v></c></row>
  </sheetData>
  <legacyDrawing r:id="rId1"/>
  <pivotTableParts count="1"><pivotTablePart r:id="rId2"/></pivotTableParts>
</worksheet>"#,
            ),
        );

        put(
            &mut zip,
            "xl/worksheets/_rels/sheet1.xml.rels",
            format!(
                r#"<?xml version="1.0"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="{OFFICE_REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
  <Relationship Id="rId2" Type="{OFFICE_REL}/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#,
            ),
        );

        put(
            &mut zip,
            "xl/pivotTables/pivotTable1.xml",
            format!(
                r#"<?xml version="1.0"?><pivotTableDefinition xmlns="{SHEET_NS}" name="PT1"/>"#,
            ),
        );
        put(
            &mut zip,
            "xl/pivotCache/pivotCacheDefinition1.xml",
            format!(r#"<?xml version="1.0"?><pivotCacheDefinition xmlns="{SHEET_NS}"/>"#),
        );
        put(
            &mut zip,
            "xl/queryTables/queryTable1.xml",
            format!(r#"<?xml version="1.0"?><queryTable xmlns="{SHEET_NS}" name="QT1"/>"#),
        );
        put(
            &mut zip,
            "xl/threadedComments/threadedComment1.xml",
            r#"<?xml version="1.0"?><ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"/>"#.to_string(),
        );
        put(
            &mut zip,
            "xl/ctrlProps/ctrlProp1.xml",
            r#"<?xml version="1.0"?><formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/>"#.to_string(),
        );
        put(
            &mut zip,
            "xl/drawings/vmlDrawing1.vml",
            r#"<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="s1"/></xml>"#.to_string(),
        );
        put(
            &mut zip,
            "customXml/item1.xml",
            r#"<?xml version="1.0"?><root><value>kept</value></root>"#.to_string(),
        );
        put(
            &mut zip,
            "customXml/itemProps1.xml",
            r#"<?xml version="1.0"?><ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"/>"#.to_string(),
        );

        zip.finish().expect("finished");
    }
    buffer.into_inner()
}

fn part_names(bytes: &[u8]) -> Vec<String> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes.to_vec())).expect("a readable zip");
    (0..archive.len())
        .map(|index| archive.by_index(index).expect("entry").name().to_string())
        .collect()
}

fn part_text(bytes: &[u8], name: &str) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes.to_vec())).expect("a readable zip");
    let mut entry = archive.by_name(name).expect("the part is there");
    let mut text = String::new();
    entry.read_to_string(&mut text).expect("utf-8");
    text
}

fn round_trip(bytes: Vec<u8>) -> Vec<u8> {
    let workbook = load_workbook_from_bytes(bytes, Default::default()).expect("loadable");
    workbook.to_bytes().expect("saved")
}

/// The `r:id` of the first element whose local name is `element`.
///
/// Namespace-prefix agnostic on purpose. Re-serialising the element model renames the default
/// namespace to an `s:` prefix, so `<pivotCache>` comes back as `<s:pivotCache>`. Asserting on
/// the literal spelling would make these tests fail over a difference that means nothing.
fn r_id_of(xml: &str, element: &str) -> Option<String> {
    let at = xml.find(&format!("{element} "))?;
    let rest = &xml[at..];
    let marker = "r:id=\"";
    let start = rest.find(marker)? + marker.len();
    let end = rest[start..].find('"')? + start;
    Some(rest[start..end].to_string())
}

/// Whether `xml` has an element whose local name is `element`.
fn has_element(xml: &str, element: &str) -> bool {
    xml.contains(&format!("{element} "))
}

/// The parts that must survive. Everything here is something the writer cannot produce.
const MUST_SURVIVE: [&str; 8] = [
    "xl/pivotTables/pivotTable1.xml",
    "xl/pivotCache/pivotCacheDefinition1.xml",
    "xl/queryTables/queryTable1.xml",
    "xl/threadedComments/threadedComment1.xml",
    "xl/ctrlProps/ctrlProp1.xml",
    "xl/drawings/vmlDrawing1.vml",
    "customXml/item1.xml",
    "customXml/itemProps1.xml",
];

#[test]
fn every_unmodelled_part_survives_the_round_trip() {
    let before = part_names(&enterprise_workbook());
    assert!(before.contains(&"xl/pivotTables/pivotTable1.xml".to_string()));

    let after = round_trip(enterprise_workbook());
    let after = part_names(&after);
    for expected in MUST_SURVIVE {
        assert!(
            after.contains(&expected.to_string()),
            "{expected} was deleted; the file now has {after:?}"
        );
    }
}

#[test]
fn a_preserved_parts_bytes_are_unchanged() {
    let before = enterprise_workbook();
    let after = round_trip(before.clone());

    let mut original = zip::ZipArchive::new(Cursor::new(before)).expect("zip");
    let mut rewritten = zip::ZipArchive::new(Cursor::new(after)).expect("zip");

    let mut original_entry = original
        .by_name("xl/pivotTables/pivotTable1.xml")
        .expect("the part");
    let mut original_bytes = Vec::new();
    original_entry
        .read_to_end(&mut original_bytes)
        .expect("readable");

    let mut rewritten_entry = rewritten
        .by_name("xl/pivotTables/pivotTable1.xml")
        .expect("the part");
    let mut rewritten_bytes = Vec::new();
    rewritten_entry
        .read_to_end(&mut rewritten_bytes)
        .expect("readable");

    assert_eq!(
        original_bytes, rewritten_bytes,
        "the pivot table's bytes came back different"
    );
}

#[test]
fn the_reference_from_the_workbook_to_its_pivot_cache_still_resolves() {
    // This is the whole point of preserving relationships. The part being present proves nothing
    // on its own; what matters is that `<pivotCache r:id>` still names a relationship that
    // exists and still points at the cache.
    let after = round_trip(enterprise_workbook());
    let workbook_xml = part_text(&after, "xl/workbook.xml");
    let rels = part_text(&after, "xl/_rels/workbook.xml.rels");

    assert!(
        has_element(&workbook_xml, "pivotCache"),
        "the reference itself was dropped: {workbook_xml}"
    );
    let id = r_id_of(&workbook_xml, "pivotCache").expect("a pivotCache with an r:id");
    assert!(
        rels.contains(&format!("Id=\"{id}\"")),
        "r:id {id} names no relationship: {rels}"
    );
    assert!(
        rels.contains("pivotCache/pivotCacheDefinition1.xml"),
        "and it does not point at the cache: {rels}"
    );
}

#[test]
fn the_sheet_reference_to_its_pivot_table_still_resolves() {
    let after = round_trip(enterprise_workbook());
    let sheet = part_text(&after, "xl/worksheets/sheet1.xml");
    let rels = part_text(&after, "xl/worksheets/_rels/sheet1.xml.rels");

    assert!(has_element(&sheet, "pivotTableParts"), "{sheet}");
    let id = r_id_of(&sheet, "pivotTablePart").expect("a pivotTablePart with an r:id");
    assert!(
        rels.contains(&format!("Id=\"{id}\"")),
        "r:id {id} names no relationship: {rels}"
    );
    assert!(
        rels.contains("../pivotTables/pivotTable1.xml"),
        "the ActiveX control's VML and the pivot table share the part, so both must be in it: {rels}"
    );
    assert!(
        rels.contains("vmlDrawing1.vml"),
        "and the legacy drawing's relationship has to survive too: {rels}"
    );
}

#[test]
fn the_activex_controls_legacy_drawing_is_still_referenced() {
    let after = round_trip(enterprise_workbook());
    let sheet = part_text(&after, "xl/worksheets/sheet1.xml");
    assert!(
        has_element(&sheet, "legacyDrawing"),
        "the reference to the control's VML was dropped: {sheet}"
    );
}

#[test]
fn every_preserved_part_has_a_content_type() {
    // A part with no declared type is a package Excel refuses to open, which would make
    // preserving it worse than dropping it.
    let after = round_trip(enterprise_workbook());
    let content_types = part_text(&after, "[Content_Types].xml");
    let names = part_names(&after);

    for name in &names {
        let extension = name.rsplit('.').next().unwrap_or_default();
        let declared = content_types.contains(&format!("PartName=\"/{name}\""))
            || content_types.contains(&format!("Extension=\"{extension}\""));
        assert!(declared, "{name} has no content type: {content_types}");
    }
}

#[test]
fn the_workbook_still_opens_and_its_cells_are_untouched() {
    // Preservation must not come at the cost of the parts the writer does model.
    let after = round_trip(enterprise_workbook());
    let workbook = load_workbook_from_bytes(after, Default::default()).expect("loadable");
    assert_eq!(workbook.get_sheet_names(), vec!["Report".to_string()]);
    let sheet = workbook.active_sheet().expect("a sheet");
    assert_eq!(
        sheet.cell_value("A1").map(|value| format!("{value:?}")),
        Some("Number(42.0)".to_string())
    );
    assert_eq!(
        sheet.cell_value("B1").map(|value| format!("{value:?}")),
        Some("Number(100.0)".to_string())
    );
}

#[test]
fn a_workbook_with_nothing_to_preserve_is_unchanged_by_it() {
    let mut workbook = Workbook::new();
    workbook
        .active_sheet_mut()
        .expect("sheet")
        .set("A1", ferroxl::CellValue::number(1.0))
        .expect("cell");
    assert!(workbook.preserved.is_empty());
    let before = workbook.to_bytes().expect("saved");
    let names_before = part_names(&before);
    let after = round_trip(before);
    assert_eq!(
        names_before,
        part_names(&after),
        "a workbook with nothing preserved should not gain or lose parts"
    );
}

#[test]
fn the_original_file_was_actually_missing_what_this_feature_adds() {
    // A guard on the test suite itself: without this, a reader could believe these tests pass
    // for the wrong reason.
    let names = part_names(&enterprise_workbook());
    assert!(names.contains(&"xl/pivotTables/pivotTable1.xml".to_string()));
    assert!(names.contains(&"xl/drawings/vmlDrawing1.vml".to_string()));
    assert!(!names.contains(&"xl/charts/chart1.xml".to_string()));
}

// -- A chart, written and then re-written -------------------------------------------------

/// A workbook carrying a chart, which is the case that used to destroy itself.
///
/// ferroxl writes a chart but does not read one back, so the second save has no chart in the
/// model and produces no drawing for it. The sheet's `<drawing>` element and its relationship
/// were preserved regardless, which left the relationship naming a part that was no longer in
/// the package -- a file no reader can open. Any edit at all was enough to reach it.
fn workbook_with_a_chart() -> Vec<u8> {
    use ferroxl::charts::reference::{Reference, ReferenceDataType};
    use ferroxl::charts::{BarChart, Series};

    let mut workbook = Workbook::new();
    let sheet = workbook.active_sheet_mut().expect("a sheet");
    sheet
        .set("A1", ferroxl::CellValue::text("Units"))
        .expect("cell");
    sheet
        .set("A2", ferroxl::CellValue::number(10.0))
        .expect("cell");
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
    workbook.to_bytes().expect("saved")
}

#[test]
fn a_chart_survives_a_round_trip() {
    let names = part_names(&round_trip(workbook_with_a_chart()));
    for part in [
        "xl/drawings/drawing1.xml",
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/charts/chart1.xml",
    ] {
        assert!(
            names.contains(&part.to_string()),
            "{part} was deleted; the file now has {names:?}"
        );
    }
}

#[test]
fn a_chart_survives_an_unrelated_edit() {
    // The whole failure: a second save that touches nothing about the chart.
    let mut workbook =
        load_workbook_from_bytes(workbook_with_a_chart(), Default::default()).expect("loadable");
    workbook
        .active_sheet_mut()
        .expect("a sheet")
        .set("Z9", ferroxl::CellValue::number(1.0))
        .expect("cell");
    let names = part_names(&workbook.to_bytes().expect("saved"));

    for part in [
        "xl/drawings/drawing1.xml",
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/charts/chart1.xml",
    ] {
        assert!(
            names.contains(&part.to_string()),
            "{part} was deleted by an unrelated edit; the file now has {names:?}"
        );
    }
}

#[test]
fn nothing_in_a_workbook_with_a_chart_points_at_a_part_that_is_not_there() {
    // The general form of the defect, and the assertion worth having: every relationship in the
    // package has to name something that exists. A dangling one is what made the file unopenable,
    // and checking each relationship says so directly rather than through a reader's stack trace.
    let mut workbook =
        load_workbook_from_bytes(workbook_with_a_chart(), Default::default()).expect("loadable");
    workbook
        .active_sheet_mut()
        .expect("a sheet")
        .set("Z9", ferroxl::CellValue::number(1.0))
        .expect("cell");
    let after = workbook.to_bytes().expect("saved");

    let names = part_names(&after);
    let mut archive = zip::ZipArchive::new(Cursor::new(after)).expect("a readable zip");
    let rels: Vec<(String, String)> = (0..archive.len())
        .filter_map(|index| {
            let entry = archive.by_index(index).ok()?;
            let name = entry.name().to_string();
            if !name.ends_with(".rels") {
                return None;
            }
            let mut text = String::new();
            let mut entry = entry;
            entry.read_to_string(&mut text).expect("utf-8");
            Some((name, text))
        })
        .collect();

    for (rels_path, text) in rels {
        let owner_directory = if rels_path.starts_with("_rels/") {
            // The package root. `_rels/.rels` describes the package, so its targets are already
            // relative to the root and there is no containing directory to add.
            String::new()
        } else {
            match ferroxl::reader::preserved::part_for_rels(&rels_path).rsplit_once('/') {
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
            let resolved = if target.starts_with('/') {
                // A package-absolute target needs no resolving.
                target.trim_start_matches('/').to_string()
            } else {
                // Otherwise the target is relative to the directory of the part that declares
                // it, so `../drawings/drawing1.xml` from `xl/worksheets/_rels/` is
                // `xl/drawings/drawing1.xml`.
                let mut segments: Vec<&str> = if owner_directory.is_empty() {
                    Vec::new()
                } else {
                    owner_directory.split('/').collect()
                };
                for segment in target.split('/') {
                    match segment {
                        ".." => {
                            segments.pop();
                        }
                        "." => {}
                        part => segments.push(part),
                    }
                }
                segments.join("/")
            };
            assert!(
                names.contains(&resolved),
                "{rels_path} names {resolved}, which is not in the package {names:?}"
            );
        }
    }
}
