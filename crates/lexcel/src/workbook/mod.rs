//! The workbook (`openpyxl/workbook/workbook.py`).
//!
//! `Workbook` is the top-level container: it owns the worksheets, the document properties
//! and security settings, and the defined names. Saving is delegated to
//! [`crate::writer::excel::save_workbook`].

use chrono::{NaiveDateTime, NaiveTime};

use crate::cell::cell::CellContext;
use crate::date_time::{BaseDate, CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900};
use crate::exceptions::{Error, Result};
use crate::formatting::StyleProperties;
use crate::namedrange::{DefinedName, NamedRange};
use crate::styles::style::Style;
use crate::worksheet::Worksheet;

/// High-level document properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentProperties {
    /// The document creator.
    pub creator: String,
    /// Who last modified the document.
    pub last_modified_by: String,
    /// Creation timestamp.
    pub created: NaiveDateTime,
    /// Modification timestamp.
    pub modified: NaiveDateTime,
    /// Document title.
    pub title: String,
    /// Document subject.
    pub subject: String,
    /// Document description.
    pub description: String,
    /// Document keywords.
    pub keywords: String,
    /// Document category.
    pub category: String,
    /// The company recorded in the extended properties.
    pub company: String,
    /// The workbook's date system.
    pub excel_base_date: BaseDate,
}

impl Default for DocumentProperties {
    fn default() -> Self {
        let now = NaiveDateTime::default();
        DocumentProperties {
            creator: "Unknown".to_string(),
            last_modified_by: "Unknown".to_string(),
            created: now,
            modified: now,
            title: "Untitled".to_string(),
            subject: String::new(),
            description: String::new(),
            keywords: String::new(),
            category: String::new(),
            company: "Microsoft Corporation".to_string(),
            excel_base_date: BaseDate::Windows1900,
        }
    }
}

impl DocumentProperties {
    /// A new property set with openpyxl's defaults.
    pub fn new() -> Self {
        DocumentProperties::default()
    }
}

/// Security information about the document.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DocumentSecurity {
    /// Whether revision tracking is locked.
    pub lock_revision: bool,
    /// Whether workbook structure is locked.
    pub lock_structure: bool,
    /// Whether the workbook window is locked.
    pub lock_windows: bool,
    /// The revision password.
    pub revision_password: String,
    /// The workbook password.
    pub workbook_password: String,
}

impl DocumentSecurity {
    /// An unlocked document.
    pub fn new() -> Self {
        DocumentSecurity::default()
    }
}

/// The container for all other parts of the document.
#[derive(Debug, Clone)]
pub struct Workbook {
    /// The worksheets, in order.
    pub worksheets: Vec<Worksheet>,
    /// Index of the active worksheet.
    pub active_sheet_index: usize,
    /// The defined names.
    pub named_ranges: Vec<DefinedName>,
    /// Document properties.
    pub properties: DocumentProperties,
    /// Document security.
    pub security: DocumentSecurity,
    /// The workbook-level default style.
    pub style: Style,
    /// Whether to use the streaming writer.
    pub optimized_write: bool,
    /// Whether the workbook was loaded for read-only streaming.
    pub optimized_read: bool,
    /// Whether cell types are guessed rather than read from the file.
    pub guess_types: bool,
    /// Whether formula cells report their cached value instead of the formula.
    pub data_only: bool,
    /// The text encoding used when reading and writing.
    pub encoding: String,
    /// The raw theme bytes read from a source workbook, re-emitted verbatim.
    pub loaded_theme: Option<Vec<u8>>,
    /// The workbook's raw bytes, kept for VBA projects.
    pub vba_archive: Option<Vec<u8>>,
    /// Style properties read from a source workbook.
    pub style_properties: Option<StyleProperties>,
    /// Drawings belonging to any sheet.
    pub drawings: Vec<usize>,
    /// Workbook-level relationships.
    pub relationships: Vec<crate::worksheet::relationship::Relationship>,
}

impl Default for Workbook {
    fn default() -> Self {
        Workbook::new()
    }
}

impl Workbook {
    /// A new workbook with one empty sheet, as openpyxl creates.
    pub fn new() -> Self {
        let context = CellContext::default();
        let mut workbook = Workbook {
            worksheets: Vec::new(),
            active_sheet_index: 0,
            named_ranges: Vec::new(),
            properties: DocumentProperties::new(),
            security: DocumentSecurity::new(),
            style: Style::new(),
            optimized_write: false,
            optimized_read: false,
            guess_types: false,
            data_only: false,
            encoding: "utf-8".to_string(),
            loaded_theme: None,
            vba_archive: None,
            style_properties: None,
            drawings: Vec::new(),
            relationships: Vec::new(),
        };
        if let Ok(mut sheet) = Worksheet::with_title("", 0) {
            sheet.context = context;
            workbook.worksheets.push(sheet);
        }
        workbook
    }

    /// A new workbook with no sheets.
    ///
    /// The streaming reader uses this, because the sheets are discovered from the archive.
    pub fn empty() -> Self {
        let mut workbook = Workbook::new();
        workbook.worksheets.clear();
        workbook
    }

    /// Create a worksheet at an optional index and return its index.
    pub fn create_sheet(&mut self, title: Option<&str>) -> Result<usize> {
        self.create_sheet_at(title, None)
    }

    /// Create a worksheet, optionally inserting it at `index`.
    pub fn create_sheet_at(&mut self, title: Option<&str>, index: Option<usize>) -> Result<usize> {
        if self.optimized_read {
            return Err(Error::ReadOnlyWorkbook(
                "Cannot create new sheet in a read-only workbook".into(),
            ));
        }
        let existing: Vec<String> = self.get_sheet_names();
        let position = self.worksheets.len();
        let mut sheet = Worksheet::with_title(title.unwrap_or(""), position)?;
        sheet.context = self.cell_context();
        let proposed = sheet.title().to_string();
        sheet.set_title(&proposed, &existing)?;
        let insert_at = index
            .unwrap_or(self.worksheets.len())
            .min(self.worksheets.len());
        self.worksheets.insert(insert_at, sheet);
        self.reindex_sheets();
        Ok(insert_at)
    }

    /// Insert an existing worksheet at an optional index.
    pub fn add_sheet(&mut self, worksheet: Worksheet, index: Option<usize>) -> Result<()> {
        if self.optimized_read {
            return Err(Error::ReadOnlyWorkbook(
                "Cannot add sheet to a read-only workbook".into(),
            ));
        }
        let existing: Vec<String> = self.get_sheet_names();
        let mut worksheet = worksheet;
        let proposed = worksheet.title().to_string();
        worksheet.set_title(&proposed, &existing)?;
        let insert_at = index
            .unwrap_or(self.worksheets.len())
            .min(self.worksheets.len());
        self.worksheets.insert(insert_at, worksheet);
        self.reindex_sheets();
        Ok(())
    }

    /// Remove a worksheet by index.
    pub fn remove_sheet(&mut self, index: usize) -> Result<()> {
        if self.optimized_read {
            return Err(Error::ReadOnlyWorkbook(
                "Cannot remove sheet from a read-only workbook".into(),
            ));
        }
        if index >= self.worksheets.len() {
            return Err(Error::Key(format!(
                "Worksheet index {index} does not exist"
            )));
        }
        self.worksheets.remove(index);
        self.reindex_sheets();
        Ok(())
    }

    /// The active worksheet index.
    pub fn active(&self) -> usize {
        self.active_sheet_index
    }

    /// Set the active worksheet index.
    pub fn set_active(&mut self, index: usize) -> Result<()> {
        if index >= self.worksheets.len() {
            return Err(Error::Key(format!(
                "Worksheet index {index} does not exist"
            )));
        }
        self.active_sheet_index = index;
        Ok(())
    }

    /// The active worksheet.
    pub fn active_sheet(&self) -> Result<&Worksheet> {
        self.worksheets
            .get(self.active_sheet_index)
            .ok_or_else(|| Error::Key("No active worksheet".to_string()))
    }

    /// The active worksheet, mutably.
    pub fn active_sheet_mut(&mut self) -> Result<&mut Worksheet> {
        let index = self.active_sheet_index;
        self.worksheets
            .get_mut(index)
            .ok_or_else(|| Error::Key("No active worksheet".to_string()))
    }

    /// Reassign worksheet indices after the sheet list changes.
    ///
    /// Called internally after any add or remove; exposed because the reader builds the
    /// sheet list directly.
    pub fn reindex_sheets(&mut self) {
        for (index, sheet) in self.worksheets.iter_mut().enumerate() {
            sheet.index = index;
        }
        if self.active_sheet_index >= self.worksheets.len() {
            self.active_sheet_index = self.worksheets.len().saturating_sub(1);
        }
    }

    /// Look a worksheet up by title.
    pub fn get_sheet_by_name(&self, name: &str) -> Option<&Worksheet> {
        self.worksheets.iter().find(|sheet| sheet.title() == name)
    }

    /// Look a worksheet up by title, mutably.
    pub fn get_sheet_by_name_mut(&mut self, name: &str) -> Option<&mut Worksheet> {
        self.worksheets
            .iter_mut()
            .find(|sheet| sheet.title() == name)
    }

    /// The index of a worksheet by title.
    pub fn get_index(&self, name: &str) -> Option<usize> {
        self.worksheets
            .iter()
            .position(|sheet| sheet.title() == name)
    }

    /// The titles of all worksheets, in order.
    pub fn get_sheet_names(&self) -> Vec<String> {
        self.worksheets
            .iter()
            .map(|sheet| sheet.title().to_string())
            .collect()
    }

    /// Whether a worksheet with the given title exists.
    pub fn contains(&self, name: &str) -> bool {
        self.worksheets.iter().any(|sheet| sheet.title() == name)
    }

    /// The workbook's date system.
    pub fn excel_base_date(&self) -> BaseDate {
        self.properties.excel_base_date
    }

    /// The cell context new cells should use.
    pub fn cell_context(&self) -> CellContext {
        CellContext {
            base_date: self.properties.excel_base_date,
            guess_types: self.guess_types,
        }
    }

    /// Mark the workbook as read-only for streaming.
    pub fn set_optimized_read(&mut self) {
        self.optimized_read = true;
    }

    /// Write the workbook to `path` as an Office Open XML package.
    ///
    /// This is the Rust spelling of `Workbook.save` in the Python original. The workbook
    /// is consumed because the writer takes ownership of it, which keeps a single
    /// definition of what "saving" does.
    pub fn save(self, path: impl AsRef<std::path::Path>) -> Result<()> {
        crate::writer::save_workbook(self, path)
    }

    /// Serialise the workbook to the package bytes without touching the filesystem.
    ///
    /// Useful for a web service that streams the download straight into a response.
    pub fn to_bytes(self) -> Result<Vec<u8>> {
        crate::writer::save_virtual_workbook(self)
    }

    /// The serial offset for the current date system.
    pub fn base_date_offset(&self) -> f64 {
        match self.properties.excel_base_date {
            BaseDate::Windows1900 => CALENDAR_WINDOWS_1900,
            BaseDate::Mac1904 => CALENDAR_MAC_1904,
        }
    }

    /// Register a defined name.
    pub fn add_named_range(&mut self, named_range: NamedRange) {
        self.named_ranges.push(DefinedName::Range(named_range));
    }

    /// Register a value-bearing defined name.
    pub fn add_named_value(&mut self, name: &str, value: &str, scope: Option<usize>) {
        self.named_ranges.push(DefinedName::Value(
            crate::namedrange::NamedRangeContainingValue::new(name, value, scope),
        ));
    }

    /// All defined names.
    pub fn get_named_ranges(&self) -> &[DefinedName] {
        &self.named_ranges
    }

    /// Look a defined name up by name.
    pub fn get_named_range(&self, name: &str) -> Option<&DefinedName> {
        self.named_ranges.iter().find(|entry| entry.name() == name)
    }

    /// Remove a defined name.
    pub fn remove_named_range(&mut self, name: &str) -> Result<()> {
        let before = self.named_ranges.len();
        self.named_ranges.retain(|entry| entry.name() != name);
        if self.named_ranges.len() == before {
            return Err(Error::Key(format!("No such named range: {name}")));
        }
        Ok(())
    }

    /// Create a named range covering `range` on the sheet at `sheet_index`.
    pub fn create_named_range(
        &mut self,
        name: &str,
        sheet_index: usize,
        range: &str,
        scope: Option<usize>,
    ) -> Result<()> {
        if sheet_index >= self.worksheets.len() {
            return Err(Error::Key(format!(
                "Worksheet index {sheet_index} does not exist"
            )));
        }
        self.add_named_range(NamedRange::new(
            name,
            vec![(sheet_index, range.to_string())],
            scope,
        ));
        Ok(())
    }

    /// Register a print-title named range for a sheet.
    pub fn add_print_title(
        &mut self,
        sheet_index: usize,
        n: u32,
        rows_or_cols: &str,
    ) -> Result<()> {
        let range = Worksheet::print_title_range(n, rows_or_cols)?;
        self.create_named_range("_xlnm.Print_Titles", sheet_index, &range, Some(sheet_index))
    }

    /// Append rows to a sheet, one entry per row.
    pub fn append_rows(
        &mut self,
        sheet_index: usize,
        rows: &[Vec<crate::cell::cell::CellValue>],
    ) -> Result<()> {
        let sheet = self
            .worksheets
            .get_mut(sheet_index)
            .ok_or_else(|| Error::Key(format!("Worksheet index {sheet_index} does not exist")))?;
        let mut row_number = sheet.highest_row();
        for row in rows {
            row_number += 1;
            sheet.write_row(row_number, 1, row)?;
        }
        Ok(())
    }

    /// The core properties' creation time as a time-only value, for tests.
    pub fn created_time(&self) -> NaiveTime {
        self.properties.created.time()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::cell::CellValue;
    use chrono::NaiveDate;

    #[test]
    fn new_workbook_has_one_sheet() {
        let wb = Workbook::new();
        assert_eq!(wb.worksheets.len(), 1);
        assert_eq!(wb.get_sheet_names(), vec!["Sheet1".to_string()]);
        assert_eq!(wb.active(), 0);
    }

    #[test]
    fn empty_workbook_has_no_sheets() {
        let wb = Workbook::empty();
        assert!(wb.worksheets.is_empty());
    }

    #[test]
    fn create_sheet_titles_uniquely() {
        let mut wb = Workbook::new();
        wb.create_sheet(None).unwrap();
        wb.create_sheet(None).unwrap();
        assert_eq!(wb.get_sheet_names(), vec!["Sheet1", "Sheet2", "Sheet3"]);
        wb.create_sheet(Some("Data")).unwrap();
        assert!(wb.contains("Data"));
        assert_eq!(wb.get_index("Data"), Some(3));
    }

    #[test]
    fn create_sheet_at_index() {
        let mut wb = Workbook::new();
        wb.create_sheet_at(Some("First"), Some(0)).unwrap();
        assert_eq!(wb.get_sheet_names()[0], "First");
        // Indices stay contiguous after insertion.
        for (index, sheet) in wb.worksheets.iter().enumerate() {
            assert_eq!(sheet.index, index);
        }
    }

    #[test]
    fn remove_sheet_reindexes() {
        let mut wb = Workbook::new();
        wb.create_sheet(Some("A")).unwrap();
        wb.create_sheet(Some("B")).unwrap();
        wb.remove_sheet(0).unwrap();
        assert_eq!(wb.get_sheet_names(), vec!["A", "B"]);
        assert_eq!(wb.worksheets[0].index, 0);
        assert!(wb.remove_sheet(9).is_err());
    }

    #[test]
    fn read_only_workbook_rejects_mutation() {
        let mut wb = Workbook::new();
        wb.set_optimized_read();
        assert!(wb.create_sheet(None).is_err());
        assert!(wb.remove_sheet(0).is_err());
        assert!(wb.add_sheet(Worksheet::new("X").unwrap(), None).is_err());
    }

    #[test]
    fn active_sheet_accessors() {
        let mut wb = Workbook::new();
        wb.create_sheet(Some("Second")).unwrap();
        wb.set_active(1).unwrap();
        assert_eq!(wb.active(), 1);
        assert_eq!(wb.active_sheet().unwrap().title(), "Second");
        wb.active_sheet_mut().unwrap().set_freeze_panes("B2");
        assert_eq!(wb.worksheets[1].freeze_panes.as_deref(), Some("B2"));
        assert!(wb.set_active(99).is_err());
    }

    #[test]
    fn named_ranges_round_trip() {
        let mut wb = Workbook::new();
        wb.create_named_range("MyRef", 0, "$A$1", None).unwrap();
        let found = wb.get_named_range("MyRef").unwrap();
        assert_eq!(found.name(), "MyRef");
        assert_eq!(found.scope(), None);
        match found {
            DefinedName::Range(r) => assert_eq!(r.destinations, vec![(0, "$A$1".to_string())]),
            other => panic!("expected a range, got {other:?}"),
        }
        assert!(wb.remove_named_range("MyRef").is_ok());
        assert!(wb.remove_named_range("MyRef").is_err());
    }

    #[test]
    fn value_names_are_supported() {
        let mut wb = Workbook::new();
        wb.add_named_value("MyValue", "9.99", None);
        assert_eq!(wb.get_named_range("MyValue").unwrap().name(), "MyValue");
        match wb.get_named_range("MyValue").unwrap() {
            DefinedName::Value(v) => assert_eq!(v.value, "9.99"),
            other => panic!("expected a value, got {other:?}"),
        }
    }

    #[test]
    fn print_titles_register_scoped_names() {
        let mut wb = Workbook::new();
        wb.create_sheet(Some("S2")).unwrap();
        wb.add_print_title(0, 2, "rows").unwrap();
        wb.add_print_title(1, 3, "cols").unwrap();
        let ranges = wb.get_named_ranges();
        assert_eq!(ranges.len(), 2);
        match &ranges[0] {
            DefinedName::Range(r) => {
                assert_eq!(r.name, "_xlnm.Print_Titles");
                assert_eq!(r.destinations[0].1, "$1:$2");
                assert_eq!(r.scope, Some(0));
            }
            other => panic!("expected a range, got {other:?}"),
        }
        match &ranges[1] {
            DefinedName::Range(r) => assert_eq!(r.destinations[0].1, "$A:$C"),
            other => panic!("expected a range, got {other:?}"),
        }
    }

    #[test]
    fn appending_rows_writes_them() {
        let mut wb = Workbook::new();
        wb.append_rows(
            0,
            &[
                vec![CellValue::text("a"), CellValue::Number(1.0)],
                vec![CellValue::text("b"), CellValue::Number(2.0)],
            ],
        )
        .unwrap();
        let sheet = &wb.worksheets[0];
        assert_eq!(sheet.cell_value("A2"), Some(CellValue::text("a")));
        assert_eq!(sheet.cell_value("B3"), Some(CellValue::Number(2.0)));
        assert!(wb.append_rows(9, &[]).is_err());
    }

    #[test]
    fn base_date_propagates_to_cells() {
        let mut wb = Workbook::new();
        wb.properties.excel_base_date = BaseDate::Mac1904;
        wb.get_sheet_by_name_mut("Sheet1").unwrap().context = wb.cell_context();
        wb.worksheets[0]
            .set(
                "A1",
                CellValue::Date(NaiveDate::from_ymd_opt(1904, 1, 1).unwrap()),
            )
            .unwrap();
        // The stored serial is the Mac calendar's zero, and reading it back gives the date.
        assert_eq!(
            *wb.worksheets[0].get_cell("A1").unwrap().internal_value(),
            CellValue::Number(0.0)
        );
        assert_eq!(
            wb.worksheets[0].cell_value("A1"),
            Some(CellValue::Date(
                NaiveDate::from_ymd_opt(1904, 1, 1).unwrap()
            ))
        );
        assert_eq!(wb.base_date_offset(), CALENDAR_MAC_1904);
    }

    #[test]
    fn guess_types_propagates() {
        let mut wb = Workbook::new();
        wb.guess_types = true;
        assert!(wb.cell_context().guess_types);
        // The context is read before the sheet is borrowed mutably.
        let context = wb.cell_context();
        let sheet = wb.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.context = context;
        sheet.set("A1", CellValue::text("50%")).unwrap();
        assert_eq!(sheet.cell_value("A1"), Some(CellValue::Number(0.5)));
    }

    #[test]
    fn default_properties() {
        let props = DocumentProperties::new();
        assert_eq!(props.creator, "Unknown");
        assert_eq!(props.title, "Untitled");
        assert_eq!(props.company, "Microsoft Corporation");
        assert_eq!(props.excel_base_date, BaseDate::Windows1900);
        assert!(!DocumentSecurity::new().lock_structure);
    }
}
