//! Assembling the xlsx package (`openpyxl/writer/excel.py`).

use std::io::{Cursor, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::exceptions::{Error, Result};
use crate::workbook::Workbook;
use crate::worksheet::Worksheet;
use crate::writer::charts::write_any_chart;
use crate::writer::comments::{write_comments, write_comments_vml};
use crate::writer::drawings::{write_drawing, write_drawing_rels, write_shapes};
use crate::writer::dump_worksheet::DumpWorksheet;
use crate::writer::strings::{create_string_table, write_string_table, StringTable};
use crate::writer::styles::{build_style_tables, write_style_table, StyleTables};
use crate::writer::table::write_table;
use crate::writer::theme::write_theme;
use crate::writer::workbook::{
    write_content_types, write_properties_app, write_properties_core, write_root_rels,
    write_workbook, write_workbook_rels,
};
use crate::writer::worksheet::{write_worksheet, write_worksheet_rels};
use crate::xml::constants::{
    ARC_APP, ARC_CONTENT_TYPES, ARC_CORE, ARC_ROOT_RELS, ARC_SHARED_STRINGS, ARC_STYLE, ARC_THEME,
    ARC_WORKBOOK, ARC_WORKBOOK_RELS, PACKAGE_DRAWINGS, PACKAGE_IMAGES, PACKAGE_WORKSHEETS,
    PACKAGE_XL,
};

/// The compression level used for archive entries.
const COMPRESSION: zip::CompressionMethod = zip::CompressionMethod::Deflated;

/// Writes a workbook to an xlsx package.
///
/// The writer collects the shared strings and style tables once, then emits the package
/// parts, tracking the drawing, chart, image and comment ids so each part gets a unique
/// name.
#[derive(Debug)]
pub struct ExcelWriter {
    /// The workbook being written.
    pub workbook: Workbook,
    /// The deduplicated style tables.
    pub style_tables: StyleTables,
    /// The shared string table.
    pub string_table: StringTable,
}

/// The counters that number the drawings, charts, images and comment parts.
///
/// They run across the whole workbook rather than per sheet, because a sheet's charts are
/// numbered globally in the package.
#[derive(Debug, Default)]
struct PartIds {
    drawing_id: u32,
    chart_id: u32,
    image_id: u32,
    shape_id: usize,
    comments_id: u32,
    table_id: u32,
}

impl PartIds {
    /// Counters start at 1, because part names in the package are 1-based.
    fn first() -> Self {
        PartIds {
            drawing_id: 1,
            chart_id: 1,
            image_id: 1,
            shape_id: 1,
            comments_id: 1,
            table_id: 1,
        }
    }
}

impl ExcelWriter {
    /// Build a writer for a workbook, collecting its shared tables.
    pub fn new(workbook: Workbook) -> Self {
        // Cells that hold nothing are dropped before the tables are built, so styles on
        // discarded cells do not leak into styles.xml.
        let mut workbook = workbook;
        for sheet in &mut workbook.worksheets {
            sheet.garbage_collect();
        }
        collect_differential_styles(&mut workbook);
        let string_table = create_string_table(&workbook.worksheets);
        let style_tables = build_style_tables(&workbook);
        ExcelWriter {
            workbook,
            style_tables,
            string_table,
        }
    }

    /// Write every part of the package, streaming each sheet's rows.
    ///
    /// The parts are written in the same order as [`write_data`](Self::write_data); the
    /// only difference is that a worksheet part is written into its zip entry as it is
    /// produced rather than being built in memory first.
    pub fn write_dump(&self, archive: &mut ZipWriter<Cursor<Vec<u8>>>) -> Result<()> {
        self.write_package_parts(archive)?;
        self.write_sheets_dump(archive)
    }

    /// The parts that are not worksheets: content types, rels, properties, theme, styles,
    /// the workbook itself and the shared strings.
    ///
    /// None of these grows with the size of a sheet's data, so both writers share them.
    fn write_package_parts(&self, archive: &mut ZipWriter<Cursor<Vec<u8>>>) -> Result<()> {
        let preserved = &self.workbook.preserved;
        writestr(
            archive,
            ARC_CONTENT_TYPES,
            crate::writer::preserved::merge_content_types(
                &write_content_types(&self.workbook),
                preserved,
            )
            .as_bytes(),
        )?;
        let (root_rels, _) = crate::writer::preserved::merge_rels(
            &write_root_rels(&self.workbook),
            preserved,
            "_rels/.rels",
        );
        writestr(archive, ARC_ROOT_RELS, root_rels.as_bytes())?;
        let (workbook_rels, workbook_id_map) = crate::writer::preserved::merge_rels(
            &write_workbook_rels(&self.workbook),
            preserved,
            ARC_WORKBOOK_RELS,
        );
        writestr(archive, ARC_WORKBOOK_RELS, workbook_rels.as_bytes())?;
        writestr(
            archive,
            ARC_APP,
            write_properties_app(&self.workbook).as_bytes(),
        )?;
        writestr(
            archive,
            ARC_CORE,
            write_properties_core(&self.workbook.properties).as_bytes(),
        )?;

        // A loaded theme is re-emitted verbatim so charts keep their colours.
        match &self.workbook.loaded_theme {
            Some(theme) => writestr(archive, ARC_THEME, theme)?,
            None => writestr(archive, ARC_THEME, &write_theme())?,
        }

        writestr(
            archive,
            ARC_STYLE,
            write_style_table(&self.workbook).as_bytes(),
        )?;
        // `<pivotCaches>` and `<externalReferences>` are how a preserved part is reached from
        // the workbook, so they travel with it -- and their `r:id` values follow any remapping
        // the merge above had to do.
        let workbook_xml = crate::writer::preserved::append_children(
            &write_workbook(&self.workbook),
            "workbook",
            preserved.workbook_children(),
            &workbook_id_map,
        );
        writestr(archive, ARC_WORKBOOK, workbook_xml.as_bytes())?;
        writestr(
            archive,
            ARC_SHARED_STRINGS,
            write_string_table(&self.string_table).as_bytes(),
        )?;

        if let Some(vba) = &self.workbook.vba_archive {
            copy_vba_archive(archive, vba)?;
        }

        // Last, so nothing here can shadow a part the writer produced. `.rels` parts are merged
        // into the writer's rather than written as they stand, because an id the writer has
        // reused has to move and a `r:id` naming it has to move with it.
        crate::writer::preserved::write_parts(archive, preserved)?;

        Ok(())
    }

    /// Write every part of the package into `archive`, buffering each worksheet.
    pub fn write_data(&self, archive: &mut ZipWriter<Cursor<Vec<u8>>>) -> Result<()> {
        self.write_package_parts(archive)?;
        self.write_worksheets(archive)
    }

    fn write_worksheets(&self, archive: &mut ZipWriter<Cursor<Vec<u8>>>) -> Result<()> {
        let mut ids = PartIds::first();
        for (index, sheet) in self.workbook.worksheets.iter().enumerate() {
            let xml = self.sheet_xml(sheet, index)?;
            writestr(
                archive,
                &format!("{PACKAGE_WORKSHEETS}/sheet{}.xml", index + 1),
                xml.as_bytes(),
            )?;
            self.write_sheet_parts(archive, sheet, index, &mut ids)?;
        }
        Ok(())
    }

    /// A sheet's part, with the preserved children the writer is not producing itself.
    ///
    /// A loaded sheet's `<drawing>`, `<legacyDrawing>` and `<tableParts>` are among the
    /// preserved children: ferroxl does not read charts, legacy drawings or tables back, so for
    /// a loaded workbook it produces none of its own and the originals are the only copy. For a
    /// sheet being built from scratch it produces all three, and the preserved ones are dropped
    /// instead -- two `<tableParts>` in one part is a file Excel reports as corrupt.
    fn sheet_xml(&self, sheet: &Worksheet, index: usize) -> Result<String> {
        let xml = write_worksheet(sheet, &self.string_table, &self.style_tables)?;
        let mut produced: Vec<&str> = Vec::new();
        if !sheet.charts.is_empty() || !sheet.images.is_empty() {
            produced.push("drawing");
            produced.push("picture");
        }
        if !sheet.tables.is_empty() {
            produced.push("tableParts");
        }
        if sheet.comment_count() > 0 {
            produced.push("legacyDrawing");
            produced.push("legacyDrawingHF");
        }
        let rels_path = format!("{PACKAGE_WORKSHEETS}/_rels/sheet{}.xml.rels", index + 1);
        // The sheet's own relationship ids are the ones its preserved children name, so the
        // remapping has to be the one that will actually be written out.
        let generated = write_worksheet_rels(sheet, 1, 1, &[]);
        let (_, id_map) =
            crate::writer::preserved::merge_rels(&generated, &self.workbook.preserved, &rels_path);
        let kept: Vec<crate::xml::functions::Element> = self
            .workbook
            .preserved
            .worksheet_children_excluding(index, &produced)
            .into_iter()
            .cloned()
            .collect();
        Ok(crate::writer::preserved::append_children(
            &xml,
            "worksheet",
            &kept,
            &id_map,
        ))
    }

    /// Write every worksheet, streaming each sheet's rows into its zip entry.
    fn write_sheets_dump(&self, archive: &mut ZipWriter<Cursor<Vec<u8>>>) -> Result<()> {
        let mut ids = PartIds::first();
        for (index, sheet) in self.workbook.worksheets.iter().enumerate() {
            let name = format!("{PACKAGE_WORKSHEETS}/sheet{}.xml", index + 1);
            archive
                .start_file(name, SimpleFileOptions::default())
                .map_err(|e| Error::Io(e.to_string()))?;
            // The dump borrows the archive for the length of one sheet, so it is scoped to
            // the sheet rather than to the loop.
            {
                let mut dump = DumpWorksheet::new(
                    &mut *archive,
                    sheet,
                    &self.string_table,
                    &self.style_tables,
                );
                dump.start()?;
                dump.write_rows()?;
                dump.finish()?;
            }
            self.write_sheet_parts(archive, sheet, index, &mut ids)?;
        }
        Ok(())
    }

    /// Write everything a sheet carries besides its cells.
    fn write_sheet_parts(
        &self,
        archive: &mut ZipWriter<Cursor<Vec<u8>>>,
        sheet: &Worksheet,
        index: usize,
        ids: &mut PartIds,
    ) -> Result<()> {
        // Table parts come first so the sheet's relationship ids are settled before the rels
        // are written: a `<tablePart>` in the sheet names a relationship in sheetN.xml.rels,
        // and Excel reports the file as corrupt if the two disagree.
        // The relationship ids come from the sheet, so this loop and the sheet's own
        // `<tableParts>` cannot disagree about which part is which.
        let table_relationship_ids: Vec<(String, u32)> = sheet
            .table_relationship_ids()
            .into_iter()
            .zip(sheet.tables.iter())
            .map(|(id, _)| (id, ids.table_id))
            .collect();
        for (table, (_, table_id)) in sheet.tables.iter().zip(&table_relationship_ids) {
            writestr(
                archive,
                &format!("{PACKAGE_XL}/tables/table{table_id}.xml"),
                write_table(&numbered(table, *table_id)).as_bytes(),
            )?;
            ids.table_id += 1;
        }

        let has_drawings = !sheet.charts.is_empty() || !sheet.images.is_empty();
        let rels_path = format!("{PACKAGE_WORKSHEETS}/_rels/sheet{}.xml.rels", index + 1);
        let preserved = &self.workbook.preserved;
        let has_preserved = preserved
            .relationships(&rels_path)
            .is_some_and(|list| !list.is_empty());
        let writer_has_own = has_drawings
            || !sheet.relationships.is_empty()
            || sheet.comment_count() > 0
            || !table_relationship_ids.is_empty();
        if writer_has_own || has_preserved {
            let rels = if writer_has_own {
                crate::writer::preserved::merge_rels(
                    &write_worksheet_rels(
                        sheet,
                        ids.drawing_id,
                        ids.comments_id,
                        &table_relationship_ids,
                    ),
                    preserved,
                    &rels_path,
                )
                .0
            } else {
                // The writer would skip this part entirely, which is exactly the case for a
                // loaded sheet whose only relationship is to a pivot table: without the
                // preserved relationship there is nowhere for it to live.
                crate::writer::preserved::rels_from_preserved(preserved, &rels_path)
                    .unwrap_or_default()
            };
            writestr(archive, &rels_path, rels.as_bytes())?;
        }

        if has_drawings {
            writestr(
                archive,
                &format!("{PACKAGE_DRAWINGS}/drawing{}.xml", ids.drawing_id),
                write_drawing(sheet).as_bytes(),
            )?;
            writestr(
                archive,
                &format!(
                    "{PACKAGE_DRAWINGS}/_rels/drawing{}.xml.rels",
                    ids.drawing_id
                ),
                write_drawing_rels(sheet, ids.chart_id, ids.image_id).as_bytes(),
            )?;
            ids.drawing_id += 1;

            for chart in &sheet.charts {
                let axes = chart.axes.clone();
                let xml = write_any_chart(chart, axes.is_some(), axes.as_ref())?;
                writestr(
                    archive,
                    &format!("{PACKAGE_XL}/charts/chart{}.xml", ids.chart_id),
                    xml.as_bytes(),
                )?;
                if !chart.shapes.is_empty() {
                    writestr(
                        archive,
                        &format!("{PACKAGE_XL}/charts/_rels/chart{}.xml.rels", ids.chart_id),
                        crate::writer::charts::write_chart_rels(ids.drawing_id).as_bytes(),
                    )?;
                    writestr(
                        archive,
                        &format!("{PACKAGE_DRAWINGS}/drawing{}.xml", ids.drawing_id),
                        write_shapes(&chart.shapes, ids.shape_id).as_bytes(),
                    )?;
                    ids.shape_id += chart.shapes.len();
                    ids.drawing_id += 1;
                }
                ids.chart_id += 1;
            }

            for image in &sheet.images {
                writestr(
                    archive,
                    &format!(
                        "{PACKAGE_IMAGES}/image{}.{}",
                        ids.image_id,
                        image.extension()
                    ),
                    &image.data,
                )?;
                ids.image_id += 1;
            }
        }

        if sheet.comment_count() > 0 {
            writestr(
                archive,
                &format!("{PACKAGE_XL}/comments{}.xml", ids.comments_id),
                write_comments(sheet).as_bytes(),
            )?;
            writestr(
                archive,
                &format!("{PACKAGE_DRAWINGS}/commentsDrawing{}.vml", ids.comments_id),
                write_comments_vml(sheet)?.as_bytes(),
            )?;
            ids.comments_id += 1;
        }
        Ok(())
    }

    /// Serialise the package to bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
        self.write_data(&mut archive)?;
        let cursor = archive.finish().map_err(|e| Error::Io(e.to_string()))?;
        Ok(cursor.into_inner())
    }
}

/// A copy of `table` whose part id is `id`.
///
/// Part names are global to the package, so two sheets can each define a table called `Sales`
/// and both need distinct `tableN.xml` entries. The id a caller set is therefore advisory:
/// the writer is the only thing that knows the running number.
fn numbered(table: &crate::worksheet::table::Table, id: u32) -> crate::worksheet::table::Table {
    let mut copy = table.clone();
    copy.id = id;
    copy
}

fn writestr(archive: &mut ZipWriter<Cursor<Vec<u8>>>, name: &str, data: &[u8]) -> Result<()> {
    archive
        .start_file(
            name,
            SimpleFileOptions::default().compression_method(COMPRESSION),
        )
        .map_err(|e| Error::Io(e.to_string()))?;
    archive
        .write_all(data)
        .map_err(|e| Error::Io(e.to_string()))?;
    Ok(())
}

/// Copy the parts of a preserved VBA archive into the output package.
fn copy_vba_archive(archive: &mut ZipWriter<Cursor<Vec<u8>>>, vba: &[u8]) -> Result<()> {
    let mut source = zip::ZipArchive::new(Cursor::new(vba.to_vec()))
        .map_err(|e| Error::BadZipFile(e.to_string()))?;
    for index in 0..source.len() {
        let mut entry = source
            .by_index(index)
            .map_err(|e| Error::BadZipFile(e.to_string()))?;
        let name = entry.name().to_string();
        if !crate::xml::constants::ARC_VBA
            .iter()
            .any(|prefix| name.starts_with(prefix))
        {
            continue;
        }
        let mut data = Vec::new();
        {
            use std::io::Read;
            entry
                .read_to_end(&mut data)
                .map_err(|e| Error::Io(e.to_string()))?;
        }
        writestr(archive, &name, &data)?;
    }
    Ok(())
}

/// Move every conditional format's differential style into the workbook's `<dxfs>` list.
///
/// A `dxf` lives on the rule while it is being built and belongs in `styles.xml` once the
/// package is assembled, with the rule left pointing at it by index. Nothing else performs
/// that move, so a rule carrying a font colour, fill or bold on a match would otherwise be
/// written with an empty `<dxfs>` — the rule would match and highlight nothing, which looks
/// like a rule that silently does nothing.
///
/// Runs before the style tables are built because both halves depend on it: the worksheets
/// need the assigned `dxfId`, and `styles.xml` needs the collected styles.
fn collect_differential_styles(workbook: &mut Workbook) {
    // A loaded workbook already has its `<dxfs>`; new rules append to that list, which is
    // what keeps the loaded styles' indices valid.
    let mut collected = workbook
        .style_properties
        .as_ref()
        .map(|properties| properties.dxf_list.clone())
        .unwrap_or_default();
    let before = collected.len();
    for sheet in &mut workbook.worksheets {
        sheet
            .conditional_formatting
            .collect_dxf_styles(&mut collected);
    }
    if collected.len() == before {
        // Nothing new, so leave the workbook exactly as it was read.
        return;
    }
    let properties = workbook
        .style_properties
        .get_or_insert_with(Default::default);
    properties.dxf_list = collected;
}

/// Save a workbook to a file path.
pub fn save_workbook(workbook: Workbook, path: impl AsRef<std::path::Path>) -> Result<()> {
    let bytes = save_virtual_workbook(workbook)?;
    std::fs::write(path, bytes)?;
    Ok(())
}

/// Save a workbook, returning the package bytes.
///
/// This is the form a web service wants: the bytes can be streamed straight into a response
/// body without touching the filesystem.
pub fn save_virtual_workbook(workbook: Workbook) -> Result<Vec<u8>> {
    ExcelWriter::new(workbook).to_bytes()
}

/// Save a workbook to an arbitrary writer.
pub fn save_workbook_to(workbook: Workbook, mut sink: impl Write) -> Result<()> {
    let bytes = save_virtual_workbook(workbook)?;
    sink.write_all(&bytes)?;
    Ok(())
}

/// Write a workbook, streaming each sheet's rows instead of buffering the part.
///
/// This is openpyxl's `save_dump`. The difference from [`save_workbook_to`] is where the
/// memory goes: that one builds each worksheet's whole XML in a `String` before handing it
/// to the archive, this one writes the head, then each row as it is produced, then the tail.
/// Peak memory is set by the largest row rather than by the largest sheet, which is the
/// difference that matters for a hundred-thousand-row export.
///
/// The bytes are identical either way — `dump_and_buffered_agree` asserts that — because a
/// second writer that produced different XML would be a liability, not a feature.
pub fn save_dump(workbook: Workbook, mut sink: impl Write) -> Result<()> {
    let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
    // The zip entry has to be started before the rows can be written into it, so the row
    // loop owns the archive rather than the other way round.
    let writer = ExcelWriter::new(workbook);
    writer.write_dump(&mut archive)?;
    let cursor = archive.finish().map_err(|e| Error::Io(e.to_string()))?;
    sink.write_all(&cursor.into_inner())?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::CellValue;
    use crate::charts::BarChart;
    use crate::formatting::rules::CellIsRule;
    use crate::styles::style::Style;
    use zip::ZipArchive;

    fn read_parts(bytes: &[u8]) -> Vec<String> {
        let mut archive = ZipArchive::new(Cursor::new(bytes.to_vec())).expect("valid zip");
        (0..archive.len())
            .map(|index| archive.by_index(index).expect("entry").name().to_string())
            .collect()
    }

    fn part_contents(bytes: &[u8], name: &str) -> String {
        let mut archive = ZipArchive::new(Cursor::new(bytes.to_vec())).expect("valid zip");
        let mut entry = archive.by_name(name).unwrap_or_else(|_| {
            panic!("missing part {name}");
        });
        let mut text = String::new();
        {
            use std::io::Read;
            entry.read_to_string(&mut text).expect("utf-8 part");
        }
        text
    }

    fn sample_workbook() -> Workbook {
        let mut workbook = Workbook::new();
        workbook.create_sheet(Some("Data")).unwrap();
        workbook.worksheets[0]
            .set("A1", CellValue::text("hello"))
            .unwrap();
        workbook.worksheets[0].set("B1", 42.0).unwrap();
        workbook.worksheets[1].set("A1", 1.0).unwrap();
        workbook
    }

    /// A `dxf` lives on the rule while it is built and belongs in `<dxfs>` once the package
    /// is assembled. Nothing else performs that move, so without it a conditional format's
    /// font colour, fill and bold are written nowhere and the rule highlights nothing.
    #[test]
    fn a_conditional_formats_differential_style_reaches_the_stylesheet() {
        let mut workbook = Workbook::new();
        let mut font = crate::styles::Font::new();
        font.bold = true;
        let rule = CellIsRule::new(Some("greaterThan"), Some("5"), false)
            .to_rule()
            .with_dxf(crate::formatting::DxfStyle {
                font: Some(font),
                border: None,
                fill: None,
            });
        workbook.worksheets[0]
            .conditional_formatting
            .add("B1:B3", rule);

        let bytes = save_virtual_workbook(workbook).unwrap();
        let styles = part_contents(&bytes, "xl/styles.xml");
        assert!(
            styles.contains("<dxfs count=\"1\">"),
            "the dxf was not collected: {styles}"
        );
        assert!(
            styles.contains("<dxf>"),
            "the dxf's contents must be written, not just counted: {styles}"
        );
        let sheet = part_contents(&bytes, "xl/worksheets/sheet1.xml");
        assert!(
            sheet.contains("dxfId=\"0\""),
            "the sheet's rule must point at the collected dxf: {sheet}"
        );
    }

    /// A loaded workbook's `<dxfs>` are already-indexed, so a new rule has to append to
    /// that list rather than replace it -- otherwise every loaded differential style ends up
    /// pointing at the wrong index.
    #[test]
    fn a_new_rule_appends_to_a_loaded_dxf_list() {
        let mut workbook = Workbook::new();
        let loaded = crate::formatting::DxfStyle::default();
        workbook.style_properties = Some(crate::formatting::StyleProperties {
            color_index: Vec::new(),
            dxf_list: vec![loaded],
        });
        let rule = CellIsRule::new(Some("greaterThan"), Some("5"), false)
            .to_rule()
            .with_dxf(crate::formatting::DxfStyle {
                font: Some(crate::styles::Font::new()),
                border: None,
                fill: None,
            });
        workbook.worksheets[0]
            .conditional_formatting
            .add("B1:B3", rule);

        let bytes = save_virtual_workbook(workbook).unwrap();
        let styles = part_contents(&bytes, "xl/styles.xml");
        assert!(
            styles.contains("<dxfs count=\"2\">"),
            "the loaded dxf was replaced rather than kept: {styles}"
        );
        assert!(
            part_contents(&bytes, "xl/worksheets/sheet1.xml").contains("dxfId=\"1\""),
            "the new rule must take the index after the loaded one"
        );
    }

    #[test]
    fn writes_every_required_part() {
        let bytes = save_virtual_workbook(sample_workbook()).unwrap();
        let parts = read_parts(&bytes);
        for expected in [
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/_rels/workbook.xml.rels",
            "docProps/app.xml",
            "docProps/core.xml",
            "xl/theme/theme1.xml",
            "xl/styles.xml",
            "xl/workbook.xml",
            "xl/sharedStrings.xml",
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/sheet2.xml",
        ] {
            assert!(parts.contains(&expected.to_string()), "missing {expected}");
        }
    }

    #[test]
    fn package_is_a_valid_zip() {
        let bytes = save_virtual_workbook(sample_workbook()).unwrap();
        let archive = ZipArchive::new(Cursor::new(bytes)).expect("valid zip");
        assert!(archive.len() >= 11);
    }

    #[test]
    fn worksheet_content_is_written() {
        let bytes = save_virtual_workbook(sample_workbook()).unwrap();
        let sheet = part_contents(&bytes, "xl/worksheets/sheet1.xml");
        assert!(sheet.contains("ref=\"A1:B1\""));
        assert!(!sheet.contains("hello"), "strings live in sharedStrings");
    }

    #[test]
    fn shared_strings_carry_cell_text() {
        let bytes = save_virtual_workbook(sample_workbook()).unwrap();
        let strings = part_contents(&bytes, "xl/sharedStrings.xml");
        assert!(strings.contains("hello"));
        assert!(strings.contains("uniqueCount=\"1\""));
    }

    #[test]
    fn styles_are_deduplicated_across_sheets() {
        let mut workbook = Workbook::new();
        let mut bold = Style::new();
        bold.font.bold = true;
        workbook.worksheets[0]
            .set("A1", CellValue::text("a"))
            .unwrap();
        workbook.worksheets[0]
            .set_style("A1", bold.clone())
            .unwrap();
        workbook.create_sheet(Some("S2")).unwrap();
        workbook.worksheets[1]
            .set("A1", CellValue::text("b"))
            .unwrap();
        workbook.worksheets[1].set_style("A1", bold).unwrap();

        let bytes = save_virtual_workbook(workbook).unwrap();
        let styles = part_contents(&bytes, "xl/styles.xml");
        let root = crate::xml::functions::fromstring(styles.as_bytes()).unwrap();
        // Two fonts: the default plus one bold.
        assert_eq!(root.find("fonts").unwrap().get("count"), Some("2"));
        assert_eq!(root.find("cellXfs").unwrap().get("count"), Some("2"));
    }

    #[test]
    fn charts_and_drawings_add_parts() {
        let mut workbook = Workbook::new();
        workbook.worksheets[0].charts.push(BarChart::new().0.base);
        let bytes = save_virtual_workbook(workbook).unwrap();
        let parts = read_parts(&bytes);
        assert!(parts.contains(&"xl/charts/chart1.xml".to_string()));
        assert!(parts.contains(&"xl/drawings/drawing1.xml".to_string()));
        assert!(parts.contains(&"xl/drawings/_rels/drawing1.xml.rels".to_string()));
        assert!(parts.contains(&"xl/worksheets/_rels/sheet1.xml.rels".to_string()));

        let chart = part_contents(&bytes, "xl/charts/chart1.xml");
        assert!(chart.contains("barChart"));
        let drawing = part_contents(&bytes, "xl/drawings/drawing1.xml");
        assert!(drawing.contains("graphicFrame"));
    }

    #[test]
    fn images_add_media_parts() {
        let mut workbook = Workbook::new();
        workbook.worksheets[0].add_image(crate::drawing::Image::new(
            vec![1, 2, 3],
            "png",
            (10, 10),
        ));
        let bytes = save_virtual_workbook(workbook).unwrap();
        let parts = read_parts(&bytes);
        assert!(parts.contains(&"xl/media/image1.png".to_string()));
    }

    #[test]
    fn comments_add_their_parts() {
        let mut workbook = Workbook::new();
        workbook.worksheets[0]
            .set_comment("A1", Some(crate::comments::Comment::new("note", "me")))
            .unwrap();
        let bytes = save_virtual_workbook(workbook).unwrap();
        let parts = read_parts(&bytes);
        assert!(parts.contains(&"xl/comments1.xml".to_string()));
        assert!(parts.contains(&"xl/drawings/commentsDrawing1.vml".to_string()));

        let comments = part_contents(&bytes, "xl/comments1.xml");
        assert!(comments.contains("note"));
        let sheet = part_contents(&bytes, "xl/worksheets/sheet1.xml");
        assert!(
            sheet.contains("legacyDrawing"),
            "the VML part must be linked"
        );
    }

    #[test]
    fn loaded_theme_is_preserved() {
        let mut workbook = Workbook::new();
        workbook.loaded_theme = Some(b"<theme-custom/>".to_vec());
        let bytes = save_virtual_workbook(workbook).unwrap();
        let theme = part_contents(&bytes, "xl/theme/theme1.xml");
        assert_eq!(theme, "<theme-custom/>");
    }

    #[test]
    fn garbage_collection_runs_before_tables_are_built() {
        let mut workbook = Workbook::new();
        workbook.worksheets[0].set("A1", CellValue::None).unwrap();
        // A styled but empty cell survives garbage collection.
        let mut style = Style::new();
        style.font.italic = true;
        workbook.worksheets[0].set_style("B1", style).unwrap();

        let writer = ExcelWriter::new(workbook);
        assert_eq!(writer.style_tables.style_list.len(), 1);
    }

    #[test]
    fn saving_to_a_sink_matches_saving_to_bytes() {
        let bytes = save_virtual_workbook(sample_workbook()).unwrap();
        let mut sink = Vec::new();
        save_workbook_to(sample_workbook(), &mut sink).unwrap();
        assert_eq!(sink, bytes);
    }

    #[test]
    fn saving_to_a_path_round_trips_through_the_reader() {
        let dir = std::env::temp_dir().join("ferroxl-writer-test");
        std::fs::create_dir_all(&dir).expect("temp dir");
        let path = dir.join("roundtrip.xlsx");
        save_workbook(sample_workbook(), &path).expect("save");

        let options = crate::reader::LoadOptions::default();
        let reloaded = crate::reader::load_workbook(&path, options).expect("reload");
        assert_eq!(reloaded.get_sheet_names(), vec!["Sheet1", "Data"]);
        assert_eq!(
            reloaded.worksheets[0].cell_value("A1"),
            Some(CellValue::text("hello"))
        );
        assert_eq!(
            reloaded.worksheets[0].cell_value("B1"),
            Some(CellValue::Number(42.0))
        );
        assert_eq!(
            reloaded.worksheets[1].cell_value("A1"),
            Some(CellValue::Number(1.0))
        );
        let _ = std::fs::remove_file(&path);
    }
}
