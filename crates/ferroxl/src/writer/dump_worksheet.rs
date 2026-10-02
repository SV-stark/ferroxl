//! The streaming writer (`openpyxl/writer/dump_worksheet.py`).
//!
//! openpyxl has two writers. The default builds a worksheet's XML as a tree in memory; the
//! `dump_worksheet` path streams rows out as they are produced instead, so peak memory is
//! set by the largest single row rather than by the whole sheet. That matters for a
//! hundred-thousand-row export and for nothing else.
//!
//! This module is the Rust equivalent. [`DumpWorksheet`] writes the head of a part, then
//! streams one row at a time, then the tail. The row serialisation is shared with the
//! in-memory writer, so the two produce byte-identical output — a test asserts exactly
//! that, because a streaming writer that quietly produced different XML would be worse than
//! no streaming writer.
//!
//! What this does not do is what `DumpWorksheet` does not do either: the worksheet is
//! already in memory by the time it is written. openpyxl's saving optimisation is about the
//! *output* stream, not about reading a sheet lazily. Lazy reading is a separate feature,
//! and [`crate::reader::LoadOptions`] is where it belongs.

use std::io::Write;

use crate::exceptions::{Error, Result};
use crate::worksheet::Worksheet;
use crate::writer::strings::StringTable;
use crate::writer::styles::StyleTables;
use crate::writer::worksheet::{group_rows, write_row};
use crate::xml::functions::XmlWriter;

/// Streams one worksheet's XML part to a sink.
///
/// The three phases mirror the structure of the part and must be run in order:
///
/// ```no_run
/// # use std::io::Write;
/// # use ferroxl::{DumpWorksheet, Workbook};
/// # fn run<W: Write>(mut sink: W, workbook: &Workbook) -> ferroxl::Result<()> {
/// # let sheet = workbook.active_sheet()?;
/// # let tables = ferroxl::writer::style_tables_for(workbook);
/// # let strings = ferroxl::writer::create_string_table(&workbook.worksheets);
/// let mut dump = DumpWorksheet::new(&mut sink, sheet, &strings, &tables);
/// dump.start()?;
/// dump.write_rows()?;
/// dump.finish()?;
/// # Ok(())
/// # }
/// ```
pub struct DumpWorksheet<'a, W: Write> {
    sink: W,
    worksheet: &'a Worksheet,
    string_table: &'a StringTable,
    style_tables: &'a StyleTables,
    max_column: u32,
}

impl<'a, W: Write> DumpWorksheet<'a, W> {
    /// Wrap a sink and the tables the rows are resolved against.
    pub fn new(
        sink: W,
        worksheet: &'a Worksheet,
        string_table: &'a StringTable,
        style_tables: &'a StyleTables,
    ) -> Self {
        DumpWorksheet {
            max_column: worksheet.highest_column(),
            sink,
            worksheet,
            string_table,
            style_tables,
        }
    }

    /// Write the worksheet element, its properties and the opening `<sheetData>`.
    ///
    /// This half is bounded by the sheet's configuration — its columns, its view — not by
    /// how many cells it holds, so it is built in memory like the rest of the part.
    pub fn start(&mut self) -> Result<()> {
        let head = crate::writer::worksheet::write_worksheet_head_to_string(
            self.worksheet,
            self.style_tables,
        )?;
        self.write(&head)
    }

    /// Write every row, one at a time.
    ///
    /// Rows are grouped before the loop, because a worksheet stores cells by coordinate and
    /// `"A10"` sorts before `"A2"`, so iteration order is not row order. The grouping is a
    /// vector of references, which is why what this saves is the XML rather than the cells.
    pub fn write_rows(&mut self) -> Result<()> {
        let rows = group_rows(self.worksheet);
        // Destructured so the sink and the worksheet can be borrowed at once, which the
        // loop needs and `&mut self` alone will not allow.
        let DumpWorksheet {
            sink,
            worksheet,
            string_table,
            style_tables,
            max_column,
        } = self;
        for (index, cells) in rows {
            let mut row = XmlWriter::new();
            write_row(
                &mut row,
                index,
                &cells,
                worksheet,
                string_table,
                style_tables,
                *max_column,
            )?;
            sink.write_all(row.as_str().as_bytes())
                .map_err(|e| Error::Io(e.to_string()))?;
        }
        Ok(())
    }

    /// Close `sheetData` and the worksheet element, and hand the sink back.
    pub fn finish(mut self) -> Result<W> {
        let tail = crate::writer::worksheet::write_worksheet_tail_to_string(self.worksheet)?;
        self.write(&tail)?;
        Ok(self.sink)
    }

    fn write(&mut self, text: &str) -> Result<()> {
        self.sink
            .write_all(text.as_bytes())
            .map_err(|e| Error::Io(e.to_string()))
    }
}
#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::CellValue;
    use crate::styles::style::Style;
    use crate::styles::Color;
    use crate::styles::{Fill, Font};
    use crate::writer::excel::save_virtual_workbook;
    use crate::writer::strings::create_string_table;
    use crate::writer::styles::build_style_tables;
    use crate::writer::worksheet::write_worksheet;
    use std::io::Write;

    /// A sheet with enough variety that a streaming writer could plausibly get part of it
    /// wrong: text, numbers, formulas, a blank, a style, a merge, a freeze pane, a column
    /// width, and rows that are not contiguous.
    fn sheet() -> Worksheet {
        let mut worksheet = Worksheet::new("Data").expect("a valid title");
        worksheet.set("A1", CellValue::text("Item")).unwrap();
        worksheet.set("B1", CellValue::text("Qty")).unwrap();
        for row in 2..=12u32 {
            let coordinate = format!("A{row}");
            worksheet
                .set(&coordinate, CellValue::text(format!("row {row}")))
                .unwrap();
            worksheet
                .set(&format!("B{row}"), CellValue::number(row as f64 * 1.5))
                .unwrap();
            worksheet
                .set(&format!("C{row}"), CellValue::formula(format!("=B{row}*2")))
                .unwrap();
        }
        // A row past the first, so the streaming path has to emit several.
        worksheet.set("E20", CellValue::text("tail")).unwrap();

        let header = Style {
            font: Font::new().with_bold(true),
            fill: Fill::solid(Color::new("FFDDDDDD".to_string())),
            ..Style::default()
        };
        worksheet.set_style("A1", header.clone()).unwrap();
        worksheet.set_style("B1", header).unwrap();
        worksheet.merge_cells("A15:C15").unwrap();
        worksheet.set("A15", CellValue::text("merged")).unwrap();
        worksheet.set_freeze_panes("B2");
        worksheet.column_dimensions.insert(
            "A".to_string(),
            crate::worksheet::ColumnDimension::from_index(1).unwrap(),
        );
        worksheet
    }

    /// The whole point of the streaming writer: it must produce the same bytes as the
    /// in-memory one. If it did not, every file it wrote would be a subtly different format.
    #[test]
    fn dump_and_buffered_agree() {
        let worksheet = sheet();
        let string_table = create_string_table(std::slice::from_ref(&worksheet));
        let tables = build_style_tables(&worksheet_workbook(&worksheet));

        let buffered = write_worksheet(&worksheet, &string_table, &tables).expect("buffered");

        let mut streamed = Vec::new();
        {
            let mut dump = DumpWorksheet::new(&mut streamed, &worksheet, &string_table, &tables);
            dump.start().expect("head");
            dump.write_rows().expect("rows");
            dump.finish().expect("tail").flush().expect("flush");
        }

        let streamed = String::from_utf8(streamed).expect("utf-8");
        assert!(
            streamed == buffered,
            "{}",
            describe_divergence("worksheet part", &buffered, &streamed),
        );
    }

    /// Where two renderings first part company, with a little context around it.
    ///
    /// A plain `assert_eq!` on two worksheet XML strings prints both in full, which is a
    /// megabyte of output that says nothing about where they differ.
    fn describe_divergence(label: &str, expected: &str, actual: &str) -> String {
        if expected == actual {
            return format!("{label} matches");
        }
        let at = expected
            .bytes()
            .zip(actual.bytes())
            .position(|(a, b)| a != b)
            .unwrap_or_else(|| expected.len().min(actual.len()));
        let start = at.saturating_sub(60);
        format!(
            "{label} differs at byte {at} (expected {} bytes, got {})\n  expected: ...{}\n  actual:   ...{}",
            expected.len(),
            actual.len(),
            &expected[start..(start + 120).min(expected.len())],
            &actual[start..(start + 120).min(actual.len())],
        )
    }

    /// And the streaming path has to produce a package whose parts are identical to the
    /// buffered one's.
    ///
    /// The parts are compared rather than the raw archive: the zip container legitimately
    /// differs, because `start_file` and `writestr!` do not choose the same compression,
    /// and a byte comparison of two zips would be asserting an accident of the writer
    /// rather than a property of the format.
    #[test]
    fn dump_and_buffered_agree_at_the_package_level() {
        use std::collections::BTreeMap;
        use std::io::Read;
        use zip::ZipArchive;

        fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
            let mut archive = ZipArchive::new(std::io::Cursor::new(bytes.to_vec())).expect("zip");
            let mut out = BTreeMap::new();
            for index in 0..archive.len() {
                let mut entry = archive.by_index(index).expect("entry");
                let name = entry.name().to_string();
                let mut body = Vec::new();
                entry.read_to_end(&mut body).expect("read");
                out.insert(name, body);
            }
            out
        }

        let buffered =
            parts(&save_virtual_workbook(worksheet_workbook(&sheet())).expect("buffered"));
        let workbook = worksheet_workbook(&sheet());
        let mut streamed = Vec::new();
        crate::writer::save_dump(workbook, &mut streamed).expect("dumped");
        let streamed = parts(&streamed);

        assert_eq!(
            buffered.keys().collect::<Vec<_>>(),
            streamed.keys().collect::<Vec<_>>(),
            "the same parts, by name"
        );
        for (name, expected) in &buffered {
            assert_eq!(
                expected,
                streamed
                    .get(name)
                    .unwrap_or_else(|| panic!("{name} missing")),
                "{}",
                describe_divergence(
                    name,
                    &String::from_utf8_lossy(expected),
                    &String::from_utf8_lossy(streamed.get(name).unwrap())
                )
            );
        }
    }

    /// The rows come out in order even though the cells are stored by coordinate and
    /// `"A10"` sorts before `"A2"`.
    #[test]
    fn rows_come_out_in_numeric_order() {
        let worksheet = sheet();
        let grouped = group_rows(&worksheet);
        let indices: Vec<u32> = grouped.iter().map(|(index, _)| *index).collect();
        let mut sorted = indices.clone();
        sorted.sort_unstable();
        assert_eq!(indices, sorted, "rows must ascend");
        assert!(
            indices.contains(&20),
            "the row past the first block is present"
        );
    }

    /// A workbook of one sheet, so the helpers above can build a style table for a sheet
    /// they are holding on its own.
    fn worksheet_workbook(worksheet: &Worksheet) -> crate::workbook::Workbook {
        let mut workbook = crate::workbook::Workbook::empty();
        workbook.add_sheet(worksheet.clone(), None).expect("sheet");
        workbook
    }
}
