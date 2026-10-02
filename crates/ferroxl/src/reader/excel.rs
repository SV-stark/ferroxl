//! The top-level reader (`openpyxl/reader/excel.py`).
//!
//! [`load_workbook`] opens a zip archive, resolves the workbook parts and reads each sheet
//! in turn.

use std::collections::HashMap;
use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use zip::ZipArchive;

use crate::cell::cell::CellContext;
use crate::exceptions::{Error, Result};
use crate::reader::comments::{comments_file_path, read_comments};
use crate::reader::strings::read_string_table;
use crate::reader::style::read_style_table;
use crate::reader::workbook::{
    detect_worksheets, read_content_types, read_excel_base_date, read_named_ranges, read_rels,
    read_sheets, read_workbook_settings, DetectedSheet,
};
use crate::reader::worksheet::{read_worksheet, WorksheetParseContext};
use crate::workbook::Workbook;
use crate::xml::constants::{
    ARC_CORE, ARC_SHARED_STRINGS, ARC_STYLE, ARC_THEME, ARC_WORKBOOK, PACKAGE_WORKSHEET_RELS,
};
use crate::xml::functions::fromstring;

/// Options controlling how a workbook is loaded.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct LoadOptions {
    /// Guess cell types from the text rather than trusting the stored type.
    ///
    /// This makes a text cell such as `"50%"` load as the number `0.5`.
    pub guess_types: bool,
    /// Report a formula cell's cached value instead of the formula itself.
    pub data_only: bool,
    /// Preserve the package's VBA project and its parts.
    pub keep_vba: bool,
}

impl LoadOptions {
    /// Default options: trust the stored types and keep formulas.
    pub fn new() -> Self {
        LoadOptions::default()
    }

    /// Guess cell types from their text.
    pub fn guessing_types(mut self) -> Self {
        self.guess_types = true;
        self
    }

    /// Load cached values instead of formulas.
    pub fn values_only(mut self) -> Self {
        self.data_only = true;
        self
    }

    /// Preserve the VBA project.
    pub fn keeping_vba(mut self) -> Self {
        self.keep_vba = true;
        self
    }
}

/// Load a workbook from a file path or from bytes already in memory.
///
/// A path is written as `Some(path)`, raw package bytes as `None`.
pub fn load_workbook(source: impl Into<WorkbookSource>, options: LoadOptions) -> Result<Workbook> {
    match source.into() {
        WorkbookSource::Path(path) => load_from_path(&path, options),
        WorkbookSource::Bytes(bytes) => load_from_bytes(&bytes, options),
    }
}

/// Load a workbook from raw package bytes.
pub fn load_workbook_from_bytes(bytes: Vec<u8>, options: LoadOptions) -> Result<Workbook> {
    load_from_bytes(&bytes, options)
}

/// Where a workbook is being loaded from.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum WorkbookSource {
    /// A filesystem path.
    Path(std::path::PathBuf),
    /// Package bytes already in memory.
    Bytes(Vec<u8>),
}

impl From<&str> for WorkbookSource {
    fn from(value: &str) -> Self {
        WorkbookSource::Path(std::path::PathBuf::from(value))
    }
}

impl From<&Path> for WorkbookSource {
    fn from(value: &Path) -> Self {
        WorkbookSource::Path(value.to_path_buf())
    }
}

impl From<&PathBuf> for WorkbookSource {
    fn from(value: &PathBuf) -> Self {
        WorkbookSource::Path(value.clone())
    }
}

impl From<PathBuf> for WorkbookSource {
    fn from(value: PathBuf) -> Self {
        WorkbookSource::Path(value)
    }
}

impl From<String> for WorkbookSource {
    fn from(value: String) -> Self {
        WorkbookSource::Path(std::path::PathBuf::from(value))
    }
}

impl From<Vec<u8>> for WorkbookSource {
    fn from(value: Vec<u8>) -> Self {
        WorkbookSource::Bytes(value)
    }
}

impl From<&[u8]> for WorkbookSource {
    fn from(value: &[u8]) -> Self {
        WorkbookSource::Bytes(value.to_vec())
    }
}

impl From<&Vec<u8>> for WorkbookSource {
    fn from(value: &Vec<u8>) -> Self {
        WorkbookSource::Bytes(value.clone())
    }
}

fn load_from_path(path: &Path, options: LoadOptions) -> Result<Workbook> {
    let bytes = fs::read(path)?;
    load_from_bytes(&bytes, options)
}

fn load_from_bytes(bytes: &[u8], options: LoadOptions) -> Result<Workbook> {
    let mut archive = ZipArchive::new(Cursor::new(bytes.to_vec()))
        .map_err(|e| Error::InvalidFile(e.to_string()))?;
    let names: Vec<String> = (0..archive.len())
        .map(|index| {
            archive
                .by_index(index)
                .map(|entry| entry.name().to_string())
                .unwrap_or_default()
        })
        .collect();

    let mut workbook = Workbook::empty();
    workbook.guess_types = options.guess_types;
    workbook.data_only = options.data_only;
    if options.keep_vba {
        workbook.vba_archive = Some(bytes.to_vec());
    }

    // `xl/workbook.xml` is the manifest; without it there is nothing to load. openpyxl
    // raises a `KeyError` from its manifest reader in the same situation.
    if !names.iter().any(|name| name == ARC_WORKBOOK) {
        return Err(Error::InvalidFile(format!(
            "the archive has no {} part",
            ARC_WORKBOOK
        )));
    }

    // Core properties and the date system are optional; a missing part is not an error.
    if let Some(data) = read_part(&mut archive, ARC_CORE) {
        if let Ok(properties) = crate::reader::workbook::read_properties_core(&data) {
            workbook.properties = properties;
        }
    }
    if let Some(data) = read_part(&mut archive, ARC_WORKBOOK) {
        if let Ok(Some(active)) = read_workbook_settings(&data) {
            workbook.active_sheet_index = active;
        }
        workbook.properties.excel_base_date = read_excel_base_date(&data)?;
    }
    if let Some(theme) = read_part(&mut archive, ARC_THEME) {
        workbook.loaded_theme = Some(theme);
    }

    let string_table = read_part(&mut archive, ARC_SHARED_STRINGS)
        .map(|data| read_string_table(&data))
        .unwrap_or_default();

    let style_table = match read_part(&mut archive, ARC_STYLE) {
        Some(data) => read_style_table(&data)?,
        None => crate::reader::style::StyleTable::default(),
    };
    // The colour index is needed while parsing sheets, so it is cloned before the table is
    // split into its parts.
    let color_index = style_table.color_index.clone();
    let dxf_list = style_table.dxf_list.clone();
    workbook.style_properties = Some(crate::formatting::StyleProperties {
        color_index: color_index.clone(),
        dxf_list,
    });
    let styles = style_table.table;

    // Resolve which archive parts hold which worksheet.
    let sheets = detect_parts(&mut archive, &names)?;
    let context = CellContext {
        base_date: workbook.properties.excel_base_date,
        guess_types: options.guess_types,
    };

    for (index, sheet) in sheets.iter().enumerate() {
        let part_path = format!("xl/{}", sheet.path);
        let Some(data) = read_part(&mut archive, &part_path) else {
            continue;
        };
        let parse_context = WorksheetParseContext {
            string_table: &string_table,
            style_table: &styles,
            color_index: &color_index,
            guess_types: options.guess_types,
            data_only: options.data_only,
        };
        let worksheet = read_worksheet(&data, &sheet.title, index, &parse_context)?;
        let mut worksheet = worksheet;
        worksheet.context = context;

        // Comments live in a separate part, referenced from the sheet's relationships.
        if let Some(comments) = comments_for(&mut archive, &names, &part_path) {
            read_comments(&mut worksheet, &comments)?;
        }

        // So do tables. Without this a saved workbook's tables would be dropped on the next
        // save, which is the round-trip loss the writer is careful to avoid.
        for part in tables_for(&mut archive, &names, &part_path) {
            let xml = String::from_utf8_lossy(&part).to_string();
            let table = crate::writer::table::read_table(&xml)?;
            worksheet.tables.add(table)?;
        }
        workbook.worksheets.push(worksheet);
    }

    workbook.reindex_sheets();
    if workbook.active_sheet_index >= workbook.worksheets.len() {
        workbook.active_sheet_index = 0;
    }

    if let Some(data) = read_part(&mut archive, ARC_WORKBOOK) {
        let titles = workbook.get_sheet_names();
        workbook.named_ranges =
            read_named_ranges(&data, &|title: &str| titles.iter().position(|t| t == title))?;
    }
    Ok(workbook)
}

fn detect_parts(
    archive: &mut ZipArchive<Cursor<Vec<u8>>>,
    names: &[String],
) -> Result<Vec<DetectedSheet>> {
    let Some(content_types) = read_part(archive, crate::xml::constants::ARC_CONTENT_TYPES) else {
        return Ok(Vec::new());
    };
    let content_types = read_content_types(&content_types)?;
    let Some(workbook_data) = read_part(archive, ARC_WORKBOOK) else {
        return Ok(Vec::new());
    };
    let rels_data = read_part(archive, crate::xml::constants::ARC_WORKBOOK_RELS)
        .ok_or_else(|| Error::InvalidFile("workbook relationships are missing".into()))?;
    let rels: HashMap<usize, String> = read_rels(&rels_data)?;
    let sheets = read_sheets(&workbook_data)?;
    let detected = detect_worksheets(&content_types, &rels, &sheets);
    // Skip relationships that point at parts the archive does not contain.
    Ok(detected
        .into_iter()
        .filter(|sheet| {
            names
                .iter()
                .any(|name| name == &format!("xl/{}", sheet.path))
        })
        .collect())
}

/// Every table part a worksheet's relationships point at, in order.
///
/// A sheet can have several, and `<tableParts>` names them by relationship id, so the
/// relationships part is the authority on which part is which rather than a guess from the
/// file names. Returning the parts in relationship order is what makes that mapping
/// recoverable.
fn tables_for(
    archive: &mut ZipArchive<Cursor<Vec<u8>>>,
    names: &[String],
    worksheet_part: &str,
) -> Vec<Vec<u8>> {
    if !worksheet_part.starts_with(crate::xml::constants::PACKAGE_WORKSHEETS) {
        return Vec::new();
    }
    let Some(codename) = worksheet_part.rsplit('/').next() else {
        return Vec::new();
    };
    let rels_part = format!("{PACKAGE_WORKSHEET_RELS}/{codename}.rels");
    let Some(rels_data) = read_part(archive, &rels_part) else {
        return Vec::new();
    };
    let Some(root) = fromstring(&rels_data).ok() else {
        return Vec::new();
    };
    let mut parts = Vec::new();
    for node in root.children() {
        if node.get("Type") != Some(crate::writer::workbook::TABLE_REL_TYPE) {
            continue;
        }
        let Some(target) = node.get("Target") else {
            continue;
        };
        // Resolve `../tables/table1.xml` relative to `xl/worksheets/`.
        let resolved = format!("{}/{target}", crate::xml::constants::PACKAGE_WORKSHEETS);
        let mut segments: Vec<&str> = Vec::new();
        for segment in resolved.split('/') {
            match segment {
                ".." => {
                    segments.pop();
                }
                "." | "" => {}
                other => segments.push(other),
            }
        }
        let normalised = segments.join("/");
        if !names.iter().any(|name| name == &normalised) {
            continue;
        }
        if let Some(data) = read_part(archive, &normalised) {
            parts.push(data);
        }
    }
    parts
}

fn comments_for(
    archive: &mut ZipArchive<Cursor<Vec<u8>>>,
    names: &[String],
    worksheet_part: &str,
) -> Option<Vec<u8>> {
    let codename = worksheet_part.rsplit('/').next()?;
    let rels_part = format!("{PACKAGE_WORKSHEET_RELS}/{codename}.rels");
    let rels_data = read_part(archive, &rels_part)?;
    let comments_part = comments_file_path(worksheet_part, &rels_data, names)?;
    read_part(archive, &comments_part)
}

/// Read a part from the archive, returning `None` when it is absent.
fn read_part(archive: &mut ZipArchive<Cursor<Vec<u8>>>, name: &str) -> Option<Vec<u8>> {
    let mut entry = archive.by_name(name).ok()?;
    let mut data = Vec::with_capacity(entry.size() as usize);
    use std::io::Read;
    entry.read_to_end(&mut data).ok()?;
    Some(data)
}

/// Read a workbook's raw package bytes, for callers that want to re-emit them verbatim.
pub fn package_bytes(path: &Path) -> Result<Vec<u8>> {
    Ok(fs::read(path)?)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::CellValue;
    use crate::comments::Comment;
    use crate::workbook::Workbook;

    fn build(workbook: Workbook) -> Vec<u8> {
        crate::writer::save_virtual_workbook(workbook).expect("save")
    }

    fn simple_workbook() -> Workbook {
        let mut workbook = Workbook::new();
        workbook.worksheets[0]
            .set("A1", CellValue::text("hello"))
            .unwrap();
        workbook.worksheets[0].set("A2", 42.0).unwrap();
        workbook.worksheets[0].set("A3", true).unwrap();
        workbook.worksheets[0]
            .set("A4", CellValue::formula("SUM(A2:A2)"))
            .unwrap();
        workbook
    }

    #[test]
    fn round_trips_values() {
        let bytes = build(simple_workbook());
        let workbook = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        assert_eq!(workbook.get_sheet_names(), vec!["Sheet1"]);
        let sheet = &workbook.worksheets[0];
        assert_eq!(sheet.cell_value("A1"), Some(CellValue::text("hello")));
        assert_eq!(sheet.cell_value("A2"), Some(CellValue::Number(42.0)));
        assert_eq!(sheet.cell_value("A3"), Some(CellValue::Bool(true)));
    }

    #[test]
    fn formulas_load_as_formulas() {
        let bytes = build(simple_workbook());
        let workbook = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        assert_eq!(
            workbook.worksheets[0].cell_value("A4"),
            Some(CellValue::Formula("=SUM(A2:A2)".into()))
        );
    }

    #[test]
    fn multiple_sheets_round_trip() {
        let mut workbook = Workbook::new();
        workbook.create_sheet(Some("Second")).unwrap();
        workbook.create_sheet(Some("Third")).unwrap();
        workbook.worksheets[2].set("A1", 7.0).unwrap();
        let bytes = build(workbook);
        let reloaded = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        assert_eq!(
            reloaded.get_sheet_names(),
            vec!["Sheet1", "Second", "Third"]
        );
        assert_eq!(
            reloaded.worksheets[2].cell_value("A1"),
            Some(CellValue::Number(7.0))
        );
    }

    #[test]
    fn styles_round_trip() {
        let mut workbook = Workbook::new();
        workbook.worksheets[0]
            .set("A1", CellValue::text("styled"))
            .unwrap();
        let mut style = crate::styles::Style::new();
        style.font.bold = true;
        style.font.name = "Arial".to_string();
        style.font.size = 16.0;
        style.set_number_format_code("0.00%");
        style.alignment = crate::styles::Alignment::new()
            .with_horizontal("center")
            .with_wrap_text(true);
        workbook.worksheets[0].set_style("A1", style).unwrap();

        let bytes = build(workbook);
        let reloaded = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        let style = reloaded.worksheets[0].get_style("A1");
        assert!(style.font.bold);
        assert_eq!(style.font.name, "Arial");
        assert_eq!(style.font.size, 16.0);
        assert_eq!(style.number_format.format_code(), "0.00%");
        assert_eq!(style.alignment.horizontal, "center");
        assert!(style.alignment.wrap_text);
    }

    #[test]
    fn dates_round_trip_through_the_number_format() {
        use chrono::NaiveDate;
        let mut workbook = Workbook::new();
        let date = NaiveDate::from_ymd_opt(2010, 1, 18).unwrap();
        workbook.worksheets[0]
            .set("A1", CellValue::Date(date))
            .unwrap();

        let bytes = build(workbook);
        let reloaded = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        // A file only stores the serial, so the date format is what brings the value back
        // as a temporal type. openpyxl always reconstructs a `datetime` on read, never a
        // `date`, so that is what a reloaded cell reports.
        assert_eq!(
            reloaded.worksheets[0].cell_value("A1"),
            Some(CellValue::DateTime(date.and_time(chrono::NaiveTime::MIN)))
        );
        assert_eq!(reloaded.worksheets[0].number_format("A1"), "yyyy-mm-dd");
    }

    #[test]
    fn merges_round_trip() {
        let mut workbook = Workbook::new();
        workbook.worksheets[0].merge_cells("A1:C3").unwrap();
        let bytes = build(workbook);
        let reloaded = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        assert_eq!(
            reloaded.worksheets[0].merged_cells(),
            &["A1:C3".to_string()]
        );
    }

    #[test]
    fn comments_round_trip() {
        let mut workbook = Workbook::new();
        workbook.worksheets[0]
            .set_comment("B2", Some(Comment::new("a note", "Alice")))
            .unwrap();
        let bytes = build(workbook);
        let reloaded = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        assert_eq!(reloaded.worksheets[0].comment_count(), 1);
        let comment = reloaded.worksheets[0]
            .get_cell("B2")
            .unwrap()
            .comment
            .as_ref()
            .unwrap();
        assert_eq!(comment.text(), "a note");
        assert_eq!(comment.author(), "Alice");
    }

    #[test]
    fn named_ranges_round_trip() {
        let mut workbook = Workbook::new();
        workbook
            .create_named_range("MyRef", 0, "a1:b3", None)
            .unwrap();
        let bytes = build(workbook);
        let reloaded = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        let found = reloaded.get_named_range("MyRef").expect("named range");
        match found {
            crate::namedrange::DefinedName::Range(range) => {
                assert_eq!(range.destinations, vec![(0, "$A$1:$B$3".to_string())]);
            }
            other => panic!("expected a range, got {other:?}"),
        }
    }

    #[test]
    fn value_names_round_trip() {
        let mut workbook = Workbook::new();
        workbook.add_named_value("MyValue", "9.99", None);
        let bytes = build(workbook);
        let reloaded = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        match reloaded.get_named_range("MyValue").unwrap() {
            crate::namedrange::DefinedName::Value(value) => assert_eq!(value.value, "9.99"),
            other => panic!("expected a value, got {other:?}"),
        }
    }

    #[test]
    fn row_and_column_dimensions_round_trip() {
        let mut workbook = Workbook::new();
        // A `<row>` element is only written for a row that holds a cell, so the row
        // dimension needs one to survive the round trip. This matches openpyxl, which
        // groups the row dimension map by the cells it finds.
        workbook.worksheets[0].set("A3", 1.0).unwrap();
        workbook.worksheets[0]
            .row_dimensions
            .insert(3, crate::worksheet::RowDimension::new(3).with_height(30.0));
        workbook.worksheets[0].column_dimensions.insert(
            "B".to_string(),
            crate::worksheet::ColumnDimension::new("B").with_width(25.0),
        );
        let bytes = build(workbook);
        let reloaded = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        let sheet = &reloaded.worksheets[0];
        assert_eq!(sheet.row_dimensions[&3].height, 30.0);
        assert_eq!(sheet.column_dimensions["B"].width, 25.0);
    }

    #[test]
    fn autofilter_round_trips() {
        let mut workbook = Workbook::new();
        workbook.worksheets[0]
            .set("A1", CellValue::text("x"))
            .unwrap();
        workbook.worksheets[0].auto_filter.set_reference("a1:c5");
        let bytes = build(workbook);
        let reloaded = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        assert_eq!(
            reloaded.worksheets[0].auto_filter.reference(),
            Some("A1:C5")
        );
    }

    #[test]
    fn data_validation_round_trips() {
        let mut workbook = Workbook::new();
        workbook.worksheets[0]
            .set("A1", CellValue::text("x"))
            .unwrap();
        let mut validation = crate::datavalidation::DataValidation::new(
            crate::datavalidation::ValidationType::Whole,
            Some(crate::datavalidation::ValidationOperator::GreaterThan),
            Some("5"),
            None,
            true,
        );
        validation.add_cell("A1");
        workbook.worksheets[0].add_data_validation(validation);
        let bytes = build(workbook);
        let reloaded = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        let loaded = &reloaded.worksheets[0].data_validations[0];
        assert_eq!(
            loaded.validation_type,
            crate::datavalidation::ValidationType::Whole
        );
        assert_eq!(loaded.formula1, "5");
        assert_eq!(loaded.cells, vec!["A1".to_string()]);
    }

    #[test]
    fn guess_types_converts_text_to_numbers() {
        let mut workbook = Workbook::new();
        // Write the value as text so guessing has something to do.
        workbook.worksheets[0]
            .cell_mut("A1")
            .unwrap()
            .set_explicit_value(CellValue::text("50%"), crate::cell::DataType::SharedString)
            .unwrap();
        let bytes = build(workbook);

        let guessed =
            load_workbook_from_bytes(bytes.clone(), LoadOptions::default().guessing_types())
                .unwrap();
        let cell = guessed.worksheets[0].get_cell("A1").unwrap();
        assert_eq!(cell.data_type, crate::cell::DataType::Numeric);

        let literal = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        assert_eq!(
            literal.worksheets[0].get_cell("A1").unwrap().data_type,
            crate::cell::DataType::SharedString
        );
    }

    #[test]
    fn empty_cells_are_not_reported_as_missing() {
        let mut workbook = Workbook::new();
        workbook.worksheets[0].set("A1", 1.0).unwrap();
        workbook.worksheets[0].set("C3", 2.0).unwrap();
        let bytes = build(workbook);
        let reloaded = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap();
        // The gap cell is absent, not an error.
        assert!(reloaded.worksheets[0].get_cell("B2").is_none());
        assert_eq!(
            reloaded.worksheets[0].cell_value("C3"),
            Some(CellValue::Number(2.0))
        );
    }

    #[test]
    fn invalid_input_is_reported_as_an_invalid_file() {
        let error =
            load_workbook_from_bytes(b"not a zip".to_vec(), LoadOptions::default()).unwrap_err();
        assert!(matches!(error, Error::InvalidFile(_)));
    }

    #[test]
    fn workbook_without_workbook_part_is_reported() {
        // A minimal but structurally valid zip that lacks xl/workbook.xml.
        let mut archive = zip::ZipWriter::new(Cursor::new(Vec::new()));
        {
            use std::io::Write;
            archive
                .start_file(
                    "[Content_Types].xml",
                    zip::write::SimpleFileOptions::default(),
                )
                .unwrap();
            archive.write_all(b"<Types/>").unwrap();
        }
        let bytes = archive.finish().unwrap().into_inner();
        let error = load_workbook_from_bytes(bytes, LoadOptions::default()).unwrap_err();
        assert!(matches!(error, Error::InvalidFile(_)));
    }

    #[test]
    fn sources_convert_from_common_types() {
        assert!(matches!(
            WorkbookSource::from("a.xlsx"),
            WorkbookSource::Path(_)
        ));
        assert!(matches!(
            WorkbookSource::from(vec![0u8, 1]),
            WorkbookSource::Bytes(_)
        ));
        assert!(matches!(
            WorkbookSource::from(Path::new("a.xlsx")),
            WorkbookSource::Path(_)
        ));
    }

    #[test]
    fn options_builders_compose() {
        let options = LoadOptions::new()
            .guessing_types()
            .values_only()
            .keeping_vba();
        assert!(options.guess_types);
        assert!(options.data_only);
        assert!(options.keep_vba);
    }

    #[test]
    fn package_bytes_reads_a_file() {
        let dir = std::env::temp_dir().join("ferroxl-package-test");
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("bytes.xlsx");
        std::fs::write(&path, b"payload").unwrap();
        assert_eq!(package_bytes(&path).unwrap(), b"payload");
        let _ = std::fs::remove_file(&path);
    }
}
