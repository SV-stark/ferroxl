//! Assembling the xlsx package (`openpyxl/writer/excel.py`).

use std::io::{Cursor, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::exceptions::{Error, Result};
use crate::workbook::Workbook;
use crate::writer::charts::write_any_chart;
use crate::writer::comments::{write_comments, write_comments_vml};
use crate::writer::drawings::{write_drawing, write_drawing_rels, write_shapes};
use crate::writer::strings::{create_string_table, write_string_table, StringTable};
use crate::writer::styles::{build_style_tables, write_style_table, StyleTables};
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

impl ExcelWriter {
    /// Build a writer for a workbook, collecting its shared tables.
    pub fn new(workbook: Workbook) -> Self {
        // Cells that hold nothing are dropped before the tables are built, so styles on
        // discarded cells do not leak into styles.xml.
        let mut workbook = workbook;
        for sheet in &mut workbook.worksheets {
            sheet.garbage_collect();
        }
        let string_table = create_string_table(&workbook.worksheets);
        let style_tables = build_style_tables(&workbook);
        ExcelWriter {
            workbook,
            style_tables,
            string_table,
        }
    }

    /// Write every part of the package into `archive`.
    pub fn write_data(&self, archive: &mut ZipWriter<Cursor<Vec<u8>>>) -> Result<()> {
        writestr(
            archive,
            ARC_CONTENT_TYPES,
            write_content_types(&self.workbook).as_bytes(),
        )?;
        writestr(
            archive,
            ARC_ROOT_RELS,
            write_root_rels(&self.workbook).as_bytes(),
        )?;
        writestr(
            archive,
            ARC_WORKBOOK_RELS,
            write_workbook_rels(&self.workbook).as_bytes(),
        )?;
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
        writestr(
            archive,
            ARC_WORKBOOK,
            write_workbook(&self.workbook).as_bytes(),
        )?;
        writestr(
            archive,
            ARC_SHARED_STRINGS,
            write_string_table(&self.string_table).as_bytes(),
        )?;

        if let Some(vba) = &self.workbook.vba_archive {
            copy_vba_archive(archive, vba)?;
        }

        self.write_worksheets(archive)
    }

    fn write_worksheets(&self, archive: &mut ZipWriter<Cursor<Vec<u8>>>) -> Result<()> {
        let mut drawing_id = 1u32;
        let mut chart_id = 1u32;
        let mut image_id = 1u32;
        let mut shape_id = 1usize;
        let mut comments_id = 1u32;

        for (index, sheet) in self.workbook.worksheets.iter().enumerate() {
            let xml = write_worksheet(sheet, &self.string_table, &self.style_tables)?;
            writestr(
                archive,
                &format!("{PACKAGE_WORKSHEETS}/sheet{}.xml", index + 1),
                xml.as_bytes(),
            )?;

            let has_drawings = !sheet.charts.is_empty() || !sheet.images.is_empty();
            if has_drawings || !sheet.relationships.is_empty() || sheet.comment_count() > 0 {
                let rels = write_worksheet_rels(sheet, drawing_id, comments_id);
                writestr(
                    archive,
                    &format!("{PACKAGE_WORKSHEETS}/_rels/sheet{}.xml.rels", index + 1),
                    rels.as_bytes(),
                )?;
            }

            if has_drawings {
                writestr(
                    archive,
                    &format!("{PACKAGE_DRAWINGS}/drawing{drawing_id}.xml"),
                    write_drawing(sheet).as_bytes(),
                )?;
                writestr(
                    archive,
                    &format!("{PACKAGE_DRAWINGS}/_rels/drawing{drawing_id}.xml.rels"),
                    write_drawing_rels(sheet, chart_id, image_id).as_bytes(),
                )?;
                drawing_id += 1;

                for chart in &sheet.charts {
                    let axes = chart.axes.clone();
                    let xml = write_any_chart(chart, axes.is_some(), axes.as_ref())?;
                    writestr(
                        archive,
                        &format!("{PACKAGE_XL}/charts/chart{chart_id}.xml"),
                        xml.as_bytes(),
                    )?;
                    if !chart.shapes.is_empty() {
                        writestr(
                            archive,
                            &format!("{PACKAGE_XL}/charts/_rels/chart{chart_id}.xml.rels"),
                            crate::writer::charts::write_chart_rels(drawing_id).as_bytes(),
                        )?;
                        writestr(
                            archive,
                            &format!("{PACKAGE_DRAWINGS}/drawing{drawing_id}.xml"),
                            write_shapes(&chart.shapes, shape_id).as_bytes(),
                        )?;
                        shape_id += chart.shapes.len();
                        drawing_id += 1;
                    }
                    chart_id += 1;
                }

                for image in &sheet.images {
                    writestr(
                        archive,
                        &format!("{PACKAGE_IMAGES}/image{image_id}.{}", image.extension()),
                        &image.data,
                    )?;
                    image_id += 1;
                }
            }

            if sheet.comment_count() > 0 {
                writestr(
                    archive,
                    &format!("{PACKAGE_XL}/comments{comments_id}.xml"),
                    write_comments(sheet).as_bytes(),
                )?;
                writestr(
                    archive,
                    &format!("{PACKAGE_DRAWINGS}/commentsDrawing{comments_id}.vml"),
                    write_comments_vml(sheet)?.as_bytes(),
                )?;
                comments_id += 1;
            }
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::CellValue;
    use crate::charts::BarChart;
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
        let dir = std::env::temp_dir().join("lexcel-writer-test");
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
