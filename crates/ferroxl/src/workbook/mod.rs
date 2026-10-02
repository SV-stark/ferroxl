//! The workbook (`openpyxl/workbook/workbook.py`).
//!
//! `Workbook` is the top-level container: it owns the worksheets, the document properties
//! and security settings, and the defined names. Saving is delegated to
//! [`crate::writer::excel::save_workbook`].

use chrono::{NaiveDateTime, NaiveTime};

use crate::cell::cell::{CellContext, CellValue};
use crate::date_time::{BaseDate, CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900};
use crate::exceptions::{Error, Result};
use crate::formatting::StyleProperties;
use crate::formula::eval::{Recalculation, Unresolved, ValueSource};
use crate::namedrange::{DefinedName, NamedRange};
use crate::styles::style::Style;
use crate::worksheet::Worksheet;
use std::collections::BTreeMap;

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

/// `<calcPr>`: how and when Excel recalculates.
///
/// The one setting here that changes what a reader sees rather than how a file looks is
/// `full_calc_on_load`. ferroxl writes formulas without cached results, so a workbook opened
/// by anything other than Excel shows no values at all until something recalculates it. That
/// makes the flag load-bearing rather than cosmetic, and it is why it is settable here rather
/// than being the fixed literal it was.
// iterate_delta is an f64, so this cannot be Eq or Hash without the manual bit
// comparison GradientStop needs. It is compared as a value instead.
#[derive(Debug, Clone, PartialEq)]
pub struct CalcProperties {
    /// The engine version the file was last calculated by. Excel uses it to decide whether to
    /// recalculate; openpyxl writes 124519.
    pub calc_id: u32,
    /// `manual`, `auto` or `autoNoTable`.
    pub calc_mode: Option<String>,
    /// Recalculate the whole workbook when it is opened.
    ///
    /// `Some(true)` by default, which is what makes a workbook written by ferroxl show values
    /// when Excel opens it.
    pub full_calc_on_load: Option<bool>,
    /// Whether references are written `A1` or `R1C1`.
    pub ref_mode: Option<String>,
    /// Whether circular references are iterated rather than refused.
    pub iterate: Option<bool>,
    /// How many times to iterate before giving up.
    pub iterate_count: Option<u32>,
    /// The change between iterations that counts as converged.
    pub iterate_delta: Option<f64>,
    /// Whether to use full precision rather than Excel's 15 digits.
    pub full_precision: Option<bool>,
    /// Whether the last calculation finished. `Some(false)` marks a workbook as needing one.
    pub calc_completed: Option<bool>,
    /// Whether to recalculate when saving.
    pub calc_on_save: Option<bool>,
    /// Whether concurrent calculation is enabled.
    pub concurrent_calc: Option<bool>,
    /// How many calculation threads to allow while calculating manually.
    pub concurrent_manual_count: Option<u32>,
    /// Whether to force a full calculation.
    pub force_full_calc: Option<bool>,
}

impl Default for CalcProperties {
    /// Excel's own defaults, and openpyxl's: automatic, recalculating on load.
    fn default() -> Self {
        CalcProperties {
            calc_id: 124_519,
            calc_mode: Some("auto".to_string()),
            full_calc_on_load: Some(true),
            ref_mode: None,
            iterate: None,
            iterate_count: None,
            iterate_delta: None,
            full_precision: None,
            calc_completed: None,
            calc_on_save: None,
            concurrent_calc: None,
            concurrent_manual_count: None,
            force_full_calc: None,
        }
    }
}

impl CalcProperties {
    /// Manual calculation: nothing recalculates until asked.
    pub fn manual() -> Self {
        CalcProperties {
            calc_mode: Some("manual".to_string()),
            ..CalcProperties::default()
        }
    }

    /// The attributes for `<calcPr>`, in the order openpyxl declares them.
    pub fn attributes(&self) -> Vec<(String, String)> {
        let mut attrs = vec![("calcId".to_string(), self.calc_id.to_string())];
        let mut put = |name: &str, value: Option<String>| {
            if let Some(value) = value {
                attrs.push((name.to_string(), value));
            }
        };
        put("calcMode", self.calc_mode.clone());
        put("fullCalcOnLoad", self.full_calc_on_load.map(flag));
        put("refMode", self.ref_mode.clone());
        put("iterate", self.iterate.map(flag));
        put("iterateCount", self.iterate_count.map(|v| v.to_string()));
        put(
            "iterateDelta",
            self.iterate_delta.map(crate::xml::functions::safe_string),
        );
        put("fullPrecision", self.full_precision.map(flag));
        put("calcCompleted", self.calc_completed.map(flag));
        put("calcOnSave", self.calc_on_save.map(flag));
        put("concurrentCalc", self.concurrent_calc.map(flag));
        put(
            "concurrentManualCount",
            self.concurrent_manual_count.map(|v| v.to_string()),
        );
        put("forceFullCalc", self.force_full_calc.map(flag));
        attrs
    }
}

/// A boolean as OOXML writes it: `1` and `0`.
fn flag(value: bool) -> String {
    if value {
        "1".to_string()
    } else {
        "0".to_string()
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
    /// The workbook's named styles, which a cell references by name.
    ///
    /// This is what makes `cell.style = "Good"` possible. It was absent entirely, so a
    /// workbook using named styles lost both the names and the formatting behind them.
    pub named_styles: crate::styles::named_style::NamedStyleList,
    /// How and when Excel recalculates this workbook.
    pub calculation: CalcProperties,
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
            calculation: CalcProperties::default(),
            named_styles: crate::styles::named_style::NamedStyleList::new(),
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
    /// Replace the calculation properties.
    ///
    /// [`CalcProperties::default`] sets `full_calc_on_load`, which is the one setting that
    /// changes what a reader sees rather than how the file looks: ferroxl writes formulas
    /// without cached results, so without it anything other than Excel opens the workbook and
    /// finds no values at all.
    pub fn set_calculation_properties(&mut self, calculation: CalcProperties) -> &mut Self {
        self.calculation = calculation;
        self
    }

    /// Apply a named style to a cell, registering the style if the workbook does not have it.
    ///
    /// This is the method that makes a named style mean what it says in two places at once.
    /// [`Worksheet::apply_named_style`] only copies the formatting onto the cell, which is
    /// enough for the cell to *look* right -- but a style nobody has heard of does not appear
    /// in Excel's style gallery, does not show up in `wb.named_styles`, and leaves
    /// `cell.style` reading back empty in openpyxl. Registering it means all three work.
    ///
    /// A name Excel already defines resolves to that built-in, so the two ways of asking for
    /// `"Good"` agree:
    ///
    /// ```
    /// # use ferroxl::workbook::Workbook;
    /// # use ferroxl::cell::cell::CellValue;
    /// let mut workbook = Workbook::new();
    /// workbook.active_sheet_mut().unwrap().set("A1", CellValue::number(1.0)).unwrap();
    /// workbook.apply_named_style(0, "A1", "Good").unwrap();
    /// assert!(workbook.named_styles.get("Good").is_some());
    /// ```
    ///
    /// # Errors
    ///
    /// An unknown name is an error rather than a default. Silently falling back to `Normal`
    /// would produce a file that opens, looks deliberate, and is wrong -- which is the worst
    /// of the three outcomes.
    pub fn apply_named_style(
        &mut self,
        sheet_index: usize,
        coordinate: &str,
        name: &str,
    ) -> Result<&'static str> {
        use crate::styles::named_style::{NamedStyle, NamedStyleList};

        // Resolve the sheet first. Registering the style and *then* discovering the index is
        // out of range would leave the workbook changed by a call that reported failure.
        if !self.worksheets.get(sheet_index).is_some() {
            return Err(Error::Key(format!("no sheet at index {sheet_index}")));
        }

        if self.named_styles.get(name).is_none() {
            // `Normal` is pinned to index 0 because `xfId="0"` means Normal; a workbook whose
            // index 0 is something else has every style reference pointing at the wrong thing.
            if name.eq_ignore_ascii_case("normal") && !self.named_styles.is_empty() {
                return Err(Error::Key(
                    "the workbook already has named styles, so index 0 is not free for \"Normal\""
                        .to_string(),
                ));
            }
            let style = NamedStyle::builtin(name).ok_or_else(|| {
                crate::styles::named_style::unknown_style(name, &self.named_styles.names())
            })?;
            if self.named_styles.is_empty() && self.named_styles.index_of("Normal").is_none() {
                self.named_styles = NamedStyleList::new();
                self.named_styles.add(NamedStyle::new("Normal"))?;
            }
            self.named_styles.add(style)?;
        }

        let style = self
            .named_styles
            .get(name)
            .ok_or_else(|| {
                crate::styles::named_style::unknown_style(name, &self.named_styles.names())
            })?
            .to_style();
        let sheet = self
            .worksheets
            .get_mut(sheet_index)
            .ok_or_else(|| Error::Key(format!("no sheet at index {sheet_index}")))?;
        sheet.set_style(coordinate, style)?;
        // The style is registered now, so the index is real rather than a guess.
        Ok(crate::styles::named_style::builtin_style(name).unwrap_or(""))
    }

    /// The named styles in this workbook, by name.
    pub fn named_style_names(&self) -> Vec<&str> {
        self.named_styles.names()
    }

    /// Register a named style, adding it to the workbook's style gallery.
    ///
    /// Returns an error if the name is taken, rather than replacing it: overwriting `Good`
    /// would quietly restyle every cell that already referenced it.
    pub fn add_named_style(&mut self, style: crate::styles::named_style::NamedStyle) -> Result<()> {
        use crate::styles::named_style::{NamedStyle, NamedStyleList};
        if self.named_styles.is_empty() {
            self.named_styles = NamedStyleList::new();
            self.named_styles.add(NamedStyle::new("Normal"))?;
        }
        self.named_styles.add(style)
    }

    /// Evaluate every formula in the workbook and record the values.
    ///
    /// This is opt-in and it is never silent:
    ///
    /// - A formula that cannot be evaluated gets **no** cached value, and appears in
    ///   [`Recalculation::unresolved`] with the reason. Nothing is guessed.
    /// - `calcPr/@fullCalcOnLoad` is set, so Excel recomputes the whole workbook when the file
    ///   is opened. These values are a convenience for readers that are not Excel; a mistake
    ///   here cannot survive being opened and saved by a human.
    /// - Formulas are evaluated in one pass over reading order, which means a formula reading
    ///   another formula's cell sees it as blank. That is a visible gap, not a wrong number,
    ///   and [`Worksheet::trace_precedents`] provides the order to do it properly.
    ///
    /// # Examples
    ///
    /// ```
    /// # use ferroxl::{CellValue, Workbook};
    /// let mut workbook = Workbook::new();
    /// let sheet = workbook.active_sheet_mut().unwrap();
    /// sheet.set("A1", CellValue::number(2.0)).unwrap();
    /// sheet.set("A2", CellValue::number(3.0)).unwrap();
    /// sheet.set("A3", CellValue::formula("=SUM(A1:A2)")).unwrap();
    ///
    /// let report = workbook.recalculate();
    /// assert_eq!(report.computed_count(), 1);
    /// assert_eq!(workbook.active_sheet().unwrap().cached_value("A3"), Some(&CellValue::Number(5.0)));
    /// ```
    pub fn recalculate(&mut self) -> Recalculation {
        let mut report = Recalculation::default();
        // Every sheet's values are visible to every formula, so the whole workbook is the
        // scope. Building it once keeps a formula over ten thousand rows from re-reading it.
        let index: BTreeMap<String, usize> = self
            .worksheets
            .iter()
            .enumerate()
            .map(|(at, sheet)| (sheet.title.clone(), at))
            .collect();
        let snapshot: Vec<BTreeMap<String, CellValue>> = self
            .worksheets
            .iter()
            .map(|sheet| {
                sheet
                    .cells()
                    .map(|cell| (cell.coordinate(), cell.internal_value().clone()))
                    .collect()
            })
            .collect();

        let source = Snapshot {
            values: &snapshot,
            titles: &index,
        };

        for sheet in self.worksheets.iter_mut() {
            let title = sheet.title.clone();
            let formulas: Vec<(String, String)> = sheet
                .cells()
                .filter_map(|cell| match cell.internal_value() {
                    CellValue::Formula(text) => Some((cell.coordinate(), text.to_string())),
                    _ => None,
                })
                .collect();

            let mut computed = BTreeMap::new();
            for (coordinate, formula) in formulas {
                match crate::formula::eval::evaluate(&formula, &title, &source) {
                    Ok(value) => {
                        sheet.set_cached_value(&coordinate, value.clone());
                        computed.insert(coordinate, value);
                    }
                    Err(reason) => {
                        // Any value this cell had is now stale, so it goes rather than staying
                        // to be read as though it were current.
                        sheet.clear_cached_value(&coordinate);
                        report.unresolved.push(Unresolved {
                            sheet: title.clone(),
                            coordinate,
                            formula,
                            reason,
                        });
                    }
                }
            }
            if !computed.is_empty() {
                report.computed.insert(title, computed);
            }
        }

        // Excel recomputes on open, so a value this method got wrong is corrected before anyone
        // sees it in Excel -- and never written back.
        self.calculation.full_calc_on_load = Some(true);
        report
    }

    /// The value computed for a formula cell, if `recalculate` has produced one.
    pub fn cached_value(&self, sheet: &str, coordinate: &str) -> Option<&CellValue> {
        let at = self.get_index(sheet)?;
        self.worksheets.get(at)?.cached_value(coordinate)
    }

    /// The sheet with this title, if there is one.
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

/// Reads cells for [`Workbook::recalculate`].
///
/// A snapshot rather than the live workbook, because evaluation must not see a value this same
/// pass has already written -- otherwise the order formulas happen to be visited in would
/// decide their answers, which is the one thing a spreadsheet must never do.
struct Snapshot<'a> {
    values: &'a [BTreeMap<String, CellValue>],
    titles: &'a BTreeMap<String, usize>,
}

impl ValueSource for Snapshot<'_> {
    fn cell(&self, sheet: &str, coordinate: &str) -> CellValue {
        let at = match self.titles.get(sheet) {
            Some(at) => *at,
            // A reference to a sheet that is not there is `#REF!`, which is what Excel shows.
            None => return CellValue::Error("#REF!".to_string()),
        };
        self.values
            .get(at)
            .and_then(|cells| cells.get(coordinate))
            .cloned()
            .unwrap_or(CellValue::None)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
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

    #[test]
    fn the_defaults_match_excel_and_openpyxl() {
        let calc = CalcProperties::default();
        assert_eq!(calc.calc_id, 124_519);
        assert_eq!(calc.calc_mode.as_deref(), Some("auto"));
        // The one that matters: ferroxl writes formulas with no cached result, so without
        // this a workbook opens in Excel showing blanks.
        assert_eq!(calc.full_calc_on_load, Some(true));
    }

    #[test]
    fn manual_calculation_is_expressible() {
        let calc = CalcProperties::manual();
        assert_eq!(calc.calc_mode.as_deref(), Some("manual"));
        // Turning calculation off without turning recalculation-on-load off is how a workbook
        // ends up showing nothing at all, so the default has to survive the switch.
        assert_eq!(calc.full_calc_on_load, Some(true));
    }

    #[test]
    fn the_attributes_come_out_in_declaration_order() {
        let calc = CalcProperties {
            calc_mode: Some("manual".to_string()),
            iterate: Some(true),
            iterate_count: Some(100),
            iterate_delta: Some(0.001),
            force_full_calc: Some(true),
            ..CalcProperties::default()
        };
        let attrs = calc.attributes();
        let names: Vec<&str> = attrs.iter().map(|(k, _)| k.as_str()).collect();
        assert_eq!(
            names,
            [
                "calcId",
                "calcMode",
                "fullCalcOnLoad",
                "iterate",
                "iterateCount",
                "iterateDelta",
                "forceFullCalc",
            ]
        );
        assert!(attrs.contains(&("iterate".to_string(), "1".to_string())));
    }

    #[test]
    fn a_workbook_carries_its_calculation_properties() {
        let mut workbook = Workbook::new();
        assert_eq!(workbook.calculation.calc_mode.as_deref(), Some("auto"));
        workbook.set_calculation_properties(CalcProperties::manual());
        assert_eq!(workbook.calculation.calc_mode.as_deref(), Some("manual"));
    }

    #[test]
    fn the_calc_properties_survive_a_round_trip() {
        let mut workbook = Workbook::new();
        workbook.set_calculation_properties(CalcProperties {
            calc_mode: Some("manual".to_string()),
            full_calc_on_load: Some(false),
            iterate: Some(true),
            iterate_count: Some(100),
            ..CalcProperties::default()
        });
        let bytes = workbook.to_bytes().expect("saved");
        let loaded = crate::reader::excel::load_workbook_from_bytes(bytes, Default::default())
            .expect("loaded");
        assert_eq!(loaded.calculation.calc_mode.as_deref(), Some("manual"));
        assert_eq!(loaded.calculation.full_calc_on_load, Some(false));
        assert_eq!(loaded.calculation.iterate, Some(true));
        assert_eq!(loaded.calculation.iterate_count, Some(100));
    }

    #[cfg(test)]
    mod named_style_tests {
        use crate::cell::cell::CellValue;
        use crate::workbook::Workbook;

        #[test]
        fn applying_a_built_in_registers_it_so_it_is_readable_back() {
            let mut workbook = Workbook::new();
            let sheet = workbook.active_sheet_mut().expect("sheet");
            sheet.set("A1", CellValue::number(1.0)).expect("cell");
            workbook
                .apply_named_style(0, "A1", "Good")
                .expect("Good is built in");

            // The point of registering: the name has to be findable afterwards, or the style is
            // invisible to everything except the cell that happens to use it.
            let bytes = workbook.to_bytes().expect("saved");
            let loaded = crate::reader::excel::load_workbook_from_bytes(bytes, Default::default())
                .expect("loaded");
            assert_eq!(loaded.named_style_names(), vec!["Normal", "Good"]);
            let good = loaded.named_styles.get("Good").expect("Good survived");
            assert_eq!(good.builtin_id.as_deref(), Some("26"));
            // `Normal` has to stay at index 0 -- `xfId="0"` resolves to it.
            assert_eq!(loaded.named_styles.index_of("Normal"), Some(0));
        }

        #[test]
        fn applying_twice_does_not_register_it_twice() {
            let mut workbook = Workbook::new();
            workbook
                .active_sheet_mut()
                .expect("sheet")
                .set("A1", CellValue::number(1.0))
                .expect("cell");
            workbook.apply_named_style(0, "A1", "Good").expect("first");
            workbook.apply_named_style(0, "A1", "Good").expect("second");
            assert_eq!(workbook.named_style_names(), vec!["Normal", "Good"]);
        }

        #[test]
        fn a_workbook_defined_style_can_be_applied_by_name() {
            // The sheet-level method can only see built-ins, so this is the one path that makes a
            // style the caller defined actually usable.
            let mut workbook = Workbook::new();
            workbook
                .add_named_style(
                    crate::styles::named_style::NamedStyle::new("Band")
                        .with_number_format("0.00%")
                        .with_font(crate::styles::fonts::Font::new().with_bold(true)),
                )
                .expect("registered");
            workbook
                .active_sheet_mut()
                .expect("sheet")
                .set("A1", CellValue::number(0.5))
                .expect("cell");
            workbook
                .apply_named_style(0, "A1", "Band")
                .expect("Band is defined");

            let cell = workbook.worksheets[0].get_style("A1");
            assert!(cell.font.bold, "the style's formatting reaches the cell");
            assert_eq!(cell.number_format.format_code(), "0.00%");
        }

        #[test]
        fn an_unknown_name_is_an_error_rather_than_a_silent_normal() {
            let mut workbook = Workbook::new();
            workbook
                .active_sheet_mut()
                .expect("sheet")
                .set("A1", CellValue::number(1.0))
                .expect("cell");
            assert!(workbook.apply_named_style(0, "A1", "Nonsense").is_err());
            // Nothing was registered and nothing was formatted: the error leaves no trace that a
            // later reader could mistake for intent.
            // Nothing at all was registered -- not even `Normal`. The writer synthesises `Normal` for
            // an empty list, so leaving it empty is both the smaller change and the accurate one.
            assert!(workbook.named_style_names().is_empty());
            assert_eq!(
                workbook.worksheets[0].get_style("A1"),
                crate::styles::Style::default()
            );
        }

        #[test]
        fn a_missing_sheet_is_an_error_not_a_panic() {
            let mut workbook = Workbook::new();
            assert!(workbook.apply_named_style(7, "A1", "Good").is_err());
        }
    }
}
