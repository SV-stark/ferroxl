//! The worksheet (`openpyxl/worksheet/worksheet.py`).
//!
//! A worksheet owns its cells, styles, dimensions and ancillary settings. Cells are created
//! on first access and can be dropped again by [`Worksheet::garbage_collect`], matching the
//! memory-conscious design of the original.

use std::collections::BTreeMap;

use crate::cell::cell::{Cell, CellContext, CellValue, DataType};
use crate::cell::formula::FormulaStore;
use crate::cell::utils::{column_index_from_string, coordinate_from_string, get_column_letter};
use crate::comments::Comment;
use crate::datavalidation::DataValidation;
use crate::drawing::Image;
use crate::exceptions::{Error, Result};
use crate::formatting::ConditionalFormatting;
use crate::styles::style::Style;
use crate::styles::Style as StyleType;
use crate::units::{points_to_pixels, DEFAULT_COLUMN_WIDTH, DEFAULT_ROW_HEIGHT};
use crate::worksheet::dimensions::{ColumnDimension, RowDimension};
use crate::worksheet::filters::AutoFilter;
use crate::worksheet::header_footer::HeaderFooter;
use crate::worksheet::page::{PageMargins, PageSetup};
use crate::worksheet::protection::SheetProtection;
use crate::worksheet::relationship::{Relationship, RelationshipType};

use std::num::NonZeroU32;

/// A worksheet.
#[derive(Debug, Clone)]
pub struct Worksheet {
    /// The sheet title.
    pub title: String,
    /// Row display properties, keyed by 1-based row number.
    pub row_dimensions: BTreeMap<u32, RowDimension>,
    /// Column display properties, keyed by column letters.
    pub column_dimensions: BTreeMap<String, ColumnDimension>,
    /// Manual page breaks.
    pub page_breaks: Vec<u32>,
    cells: BTreeMap<String, Cell>,
    styles: BTreeMap<String, StyleType>,
    /// Charts attached to this sheet.
    pub charts: Vec<crate::charts::Chart>,
    /// Images attached to this sheet.
    pub images: Vec<Image>,
    /// Relationships owned by this sheet.
    pub relationships: Vec<Relationship>,
    /// Data validations.
    pub data_validations: Vec<DataValidation>,
    /// The selected cell.
    pub selected_cell: String,
    /// The active cell.
    pub active_cell: String,
    /// `visible`, `hidden` or `veryHidden`.
    pub sheet_state: String,
    /// Page layout.
    pub page_setup: PageSetup,
    /// Page margins.
    pub page_margins: PageMargins,
    /// Headers and footers.
    pub header_footer: HeaderFooter,
    /// Sheet protection.
    pub protection: SheetProtection,
    /// Whether gridlines are shown.
    pub show_gridlines: bool,
    /// Whether gridlines are printed.
    pub print_gridlines: bool,
    /// Whether summaries appear below.
    pub show_summary_below: bool,
    /// Whether summaries appear to the right.
    pub show_summary_right: bool,
    /// The auto filter.
    pub auto_filter: AutoFilter,
    /// The frozen pane's top-left cell.
    pub freeze_panes: Option<String>,
    /// The paper size, when set directly on the sheet.
    pub paper_size: Option<String>,
    /// Shared formula storage.
    pub formula_store: FormulaStore,
    /// The ambient cell context (base date, guess-types).
    pub context: CellContext,
    /// Conditional formatting rules.
    pub conditional_formatting: ConditionalFormatting,
    /// Merged ranges.
    merged_cells: Vec<String>,
    /// Comment count, maintained by the writer.
    comment_count: usize,
    /// The raw worksheet XML, preserved for VBA workbooks.
    pub xml_source: Option<Vec<u8>>,
    /// The worksheet index within its workbook.
    pub index: usize,
}

impl Worksheet {
    /// Page break kinds, mirroring the `BREAK_*` constants.
    pub const BREAK_NONE: u32 = 0;
    /// Row break.
    pub const BREAK_ROW: u32 = 1;
    /// Column break.
    pub const BREAK_COLUMN: u32 = 2;

    /// Visible sheet state.
    pub const SHEETSTATE_VISIBLE: &'static str = "visible";
    /// Hidden sheet state.
    pub const SHEETSTATE_HIDDEN: &'static str = "hidden";
    /// Very-hidden sheet state.
    pub const SHEETSTATE_VERYHIDDEN: &'static str = "veryHidden";

    /// Portrait orientation.
    pub const ORIENTATION_PORTRAIT: &'static str = "portrait";
    /// Landscape orientation.
    pub const ORIENTATION_LANDSCAPE: &'static str = "landscape";

    /// The paper-size codes accepted by [`Worksheet::set_printer_settings`].
    pub const PAPERSIZES: [(&'static str, &str); 11] = [
        ("PAPERSIZE_LETTER", "1"),
        ("PAPERSIZE_LETTER_SMALL", "2"),
        ("PAPERSIZE_TABLOID", "3"),
        ("PAPERSIZE_LEDGER", "4"),
        ("PAPERSIZE_LEGAL", "5"),
        ("PAPERSIZE_STATEMENT", "6"),
        ("PAPERSIZE_EXECUTIVE", "7"),
        ("PAPERSIZE_A3", "8"),
        ("PAPERSIZE_A4", "9"),
        ("PAPERSIZE_A4_SMALL", "10"),
        ("PAPERSIZE_A5", "11"),
    ];

    /// Create a sheet titled `title`.
    ///
    /// An empty title is replaced by `Sheet<n>`, where `n` counts the existing sheets plus
    /// one.
    pub fn new(title: &str) -> Result<Self> {
        Worksheet::with_title(title, 0)
    }

    /// Create a sheet titled `title` at the given workbook index.
    pub fn with_title(title: &str, index: usize) -> Result<Self> {
        let title = if title.is_empty() {
            format!("Sheet{}", index + 1)
        } else {
            title.to_string()
        };
        let worksheet = Worksheet {
            title: String::new(),
            row_dimensions: BTreeMap::new(),
            column_dimensions: BTreeMap::new(),
            page_breaks: Vec::new(),
            cells: BTreeMap::new(),
            styles: BTreeMap::new(),
            charts: Vec::new(),
            images: Vec::new(),
            relationships: Vec::new(),
            data_validations: Vec::new(),
            selected_cell: "A1".to_string(),
            active_cell: "A1".to_string(),
            sheet_state: Worksheet::SHEETSTATE_VISIBLE.to_string(),
            page_setup: PageSetup::default(),
            page_margins: PageMargins::default(),
            header_footer: HeaderFooter::new(),
            protection: SheetProtection::new(),
            show_gridlines: true,
            print_gridlines: false,
            show_summary_below: true,
            show_summary_right: true,
            auto_filter: AutoFilter::new(),
            freeze_panes: None,
            paper_size: None,
            formula_store: FormulaStore::new(),
            context: CellContext::default(),
            conditional_formatting: ConditionalFormatting::new(),
            merged_cells: Vec::new(),
            comment_count: 0,
            xml_source: None,
            index,
        };
        let mut worksheet = worksheet;
        worksheet.set_title(&title, &[])?;
        Ok(worksheet)
    }

    /// The sheet title.
    pub fn title(&self) -> &str {
        &self.title
    }

    /// Set the sheet title, validating it against Excel's rules.
    ///
    /// `existing` lists the titles already in use so the title can be de-duplicated.
    pub fn set_title(&mut self, value: &str, existing: &[String]) -> Result<()> {
        if value.contains(['\\', '*', '?', ':', '/', '[', ']']) {
            return Err(Error::SheetTitle(
                "Invalid character found in sheet title".into(),
            ));
        }
        let value = self.unique_sheet_name(value, existing);
        if value.chars().count() > 31 {
            return Err(Error::SheetTitle(
                "Maximum 31 characters allowed in sheet title".into(),
            ));
        }
        self.title = value;
        Ok(())
    }

    /// Append a counter to a duplicate title, matching Excel's behaviour.
    fn unique_sheet_name(&self, value: &str, existing: &[String]) -> String {
        if !existing.iter().any(|title| title == value) {
            return value.to_string();
        }
        // Find the highest trailing digit already used for this title.
        let mut highest = 0u32;
        for title in existing {
            let Some(rest) = title.strip_prefix(value) else {
                continue;
            };
            if rest.is_empty() {
                continue;
            }
            if let Ok(count) = rest.parse::<u32>() {
                highest = highest.max(count);
            }
        }
        format!("{value}{}", highest + 1)
    }

    /// Read a cell by coordinate, creating it if needed.
    pub fn cell(&mut self, coordinate: &str) -> Result<&Cell> {
        self.create_cell(coordinate)?;
        Ok(self.cells.get(coordinate).expect("cell just created"))
    }

    /// Mutably read a cell by coordinate, creating it if needed.
    pub fn cell_mut(&mut self, coordinate: &str) -> Result<&mut Cell> {
        self.create_cell(coordinate)?;
        Ok(self.cells.get_mut(coordinate).expect("cell just created"))
    }

    /// Look a cell up without creating it.
    pub fn get_cell(&self, coordinate: &str) -> Option<&Cell> {
        self.cells.get(coordinate)
    }

    /// Look a cell up mutably without creating it.
    pub fn get_cell_mut(&mut self, coordinate: &str) -> Option<&mut Cell> {
        self.cells.get_mut(coordinate)
    }

    fn create_cell(&mut self, coordinate: &str) -> Result<()> {
        if self.cells.contains_key(coordinate) {
            return Ok(());
        }
        let (column, row) = coordinate_from_string(coordinate)?;
        let cell = Cell::new(&column, row);
        self.cells.insert(coordinate.to_string(), cell);
        self.column_dimensions
            .entry(column.clone())
            .or_insert_with(|| ColumnDimension::new(&column));
        self.row_dimensions
            .entry(row)
            .or_insert_with(|| RowDimension::new(row));
        Ok(())
    }

    /// Read the cell at a 1-based `(row, column)` position.
    pub fn cell_at(&mut self, row: u32, column: u32) -> Result<&Cell> {
        let coordinate = format!("{}{}", get_column_letter(column)?, row);
        self.cell(&coordinate)
    }

    /// Set a cell's value, applying any number format the caster requires.
    pub fn set_cell_value(&mut self, coordinate: &str, value: impl Into<CellValue>) -> Result<()> {
        let context = self.context;
        self.create_cell(coordinate)?;
        let cell = self.cells.get_mut(coordinate).expect("cell just created");
        let format = cell.set_value(value, context);
        if let Some(code) = format {
            self.get_style_mut(coordinate)?
                .set_number_format_code(&code);
        }
        Ok(())
    }

    /// Read a cell's value, converting date-formatted numerics back into date types.
    pub fn cell_value(&self, coordinate: &str) -> Option<CellValue> {
        let cell = self.cells.get(coordinate)?;
        let format = self.styles.get(coordinate).map(|s| s.number_format_code());
        Some(cell.display_value(format))
    }

    /// Set a cell's value from a display value such as a string or number.
    pub fn set(&mut self, coordinate: &str, value: impl Into<CellValue>) -> Result<()> {
        self.set_cell_value(coordinate, value)
    }

    /// An unordered iterator over the sheet's cells.
    pub fn cells(&self) -> impl Iterator<Item = &Cell> {
        self.cells.values()
    }

    /// The number of stored cells.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Delete cells that hold neither a value, a comment nor a non-default style.
    pub fn garbage_collect(&mut self) {
        let defaults = crate::styles::defaults();
        let to_delete: Vec<String> = self
            .cells
            .iter()
            .filter(|(coordinate, cell)| {
                !cell.merged
                    && cell.internal_value().is_empty()
                    && cell.comment.is_none()
                    // A cell is dropped when it is unstyled, or when its style is the
                    // default — either way the style carries no information.
                    && match self.styles.get(*coordinate) {
                        None => true,
                        Some(style) => crate::styles::same_visual_style(style, &defaults),
                    }
            })
            .map(|(coordinate, _)| coordinate.clone())
            .collect();
        for coordinate in to_delete {
            self.cells.remove(&coordinate);
        }
    }

    /// The style for a coordinate, creating a default style if needed.
    ///
    /// A style loaded from the file is marked `static`; the first mutable access replaces it
    /// with a private copy so the shared table entry cannot be changed.
    pub fn get_style(&self, coordinate: &str) -> Style {
        self.styles
            .get(coordinate)
            .cloned()
            .unwrap_or_else(Style::new)
    }

    /// The style for a coordinate, without the `static` unwrapping.
    pub fn get_style_read_only(&self, coordinate: &str) -> Option<&Style> {
        self.styles.get(coordinate)
    }

    /// The style for mutation, creating or unwrapping as needed.
    pub fn get_style_mut(&mut self, coordinate: &str) -> Result<&mut Style> {
        // Python's `_cells` and `_styles` are populated together: reading a style for a
        // coordinate that has no cell yet creates an empty one, which is why the writer can
        // attach a style to a blank cell. Row-level ("12") and column-level ("C") styles do
        // not create a cell.
        if coordinate_from_string(coordinate).is_ok() {
            self.create_cell(coordinate)?;
        }
        let entry = self.styles.entry(coordinate.to_string()).or_default();
        if entry.is_static {
            *entry = entry.copy_style();
        }
        Ok(entry)
    }

    /// Whether a style has been recorded for the coordinate.
    pub fn has_style(&self, coordinate: &str) -> bool {
        self.styles.contains_key(coordinate)
    }

    /// Attach a style to a coordinate.
    pub fn set_style(&mut self, coordinate: &str, style: Style) -> Result<()> {
        if coordinate_from_string(coordinate).is_ok() {
            self.create_cell(coordinate)?;
        }
        self.styles.insert(coordinate.to_string(), style);
        Ok(())
    }

    /// Every style recorded on this sheet, including row-level styles.
    ///
    /// The writer needs these to build the shared style tables.
    pub fn style_values(&self) -> Vec<Style> {
        self.styles.values().cloned().collect()
    }

    /// The number format code applied to a coordinate.
    pub fn number_format(&self, coordinate: &str) -> String {
        self.styles
            .get(coordinate)
            .map(|s| s.number_format_code().to_string())
            .unwrap_or_else(|| "General".to_string())
    }

    /// Set the number format code for a coordinate.
    pub fn set_number_format(&mut self, coordinate: &str, code: &str) -> Result<()> {
        self.get_style_mut(coordinate)?.set_number_format_code(code);
        Ok(())
    }

    /// The highest row index that holds data.
    pub fn highest_row(&self) -> u32 {
        self.row_dimensions.keys().copied().max().unwrap_or(1)
    }

    /// The largest column index currently stored.
    pub fn highest_column(&self) -> u32 {
        self.column_dimensions
            .keys()
            .filter_map(|key| column_index_from_string(key).ok())
            .max()
            .unwrap_or(1)
    }

    /// The minimum bounding range for all cells containing data.
    pub fn calculate_dimension(&self) -> Result<String> {
        Ok(format!(
            "A1:{}{}",
            get_column_letter(self.highest_column())?,
            self.highest_row()
        ))
    }

    /// Resolve a range string to its inclusive bounds.
    pub fn range_bounds(
        &self,
        range_string: &str,
    ) -> Result<crate::worksheet::iter_worksheet::RangeBounds> {
        crate::worksheet::iter_worksheet::get_range_boundaries(range_string, 0, 1)
    }

    /// The coordinates in a range, in row-major order.
    pub fn range_coordinates(&self, range_string: &str) -> Result<Vec<String>> {
        let bounds = self.range_bounds(range_string)?;
        let mut out = Vec::new();
        for row in bounds.min_row..=bounds.max_row {
            // `max_col` is an exclusive bound, so the range covers `min_col..max_col`.
            for column in bounds.min_col..bounds.max_col {
                out.push(format!("{}{}", get_column_letter(column)?, row));
            }
        }
        Ok(out)
    }

    /// The values in a range, row by row.
    pub fn range_values(&self, range_string: &str) -> Result<Vec<Vec<CellValue>>> {
        let bounds = self.range_bounds(range_string)?;
        let mut rows = Vec::new();
        for row in bounds.min_row..=bounds.max_row {
            let mut values = Vec::new();
            // `max_col` is an exclusive bound, so the range covers `min_col..max_col`.
            for column in bounds.min_col..bounds.max_col {
                let coordinate = format!("{}{}", get_column_letter(column)?, row);
                values.push(self.cell_value(&coordinate).unwrap_or(CellValue::None));
            }
            rows.push(values);
        }
        Ok(rows)
    }

    /// Every row of the sheet's used range, row by row, left to right.
    ///
    /// The row-major counterpart of [`cells`](Self::cells), and what most callers want
    /// first: openpyxl calls this `ws.iter_rows()` and almost every script that reads a
    /// workbook is a loop over it.
    ///
    /// Rows come back as whole rows even where the sheet is sparse, so the outer and inner
    /// iterators always line up — a row that is entirely empty is a row of `None` rather
    /// than a row missing from the output. Skipping blanks is a filter, and leaving it to
    /// the caller means the shape of the result is never in question.
    pub fn iter_rows(&self) -> Vec<Vec<Option<&Cell>>> {
        let (min_col, min_row, max_col, max_row) = self.used_bounds();
        self.iter_rows_within(min_row, max_row, min_col, max_col)
    }

    /// Rows from `min_row` to `max_row`, inclusive, and `min_col` to `max_col`, inclusive.
    ///
    /// The bounds are inclusive on both ends, which is how a caller thinks about `A1:B5`
    /// and is also what openpyxl's keyword arguments mean. That differs from
    /// [`RangeBounds`](crate::worksheet::iter_worksheet::RangeBounds), whose `max_col` is
    /// exclusive because it comes out of a range string.
    pub fn iter_rows_within(
        &self,
        min_row: u32,
        max_row: u32,
        min_col: u32,
        max_col: u32,
    ) -> Vec<Vec<Option<&Cell>>> {
        (min_row..=max_row)
            .map(|row| {
                (min_col..=max_col)
                    .map(|column| self.get_cell_by_index(column, row))
                    .collect()
            })
            .collect()
    }

    /// Every column of the sheet's used range, column by column, top to bottom.
    ///
    /// Column-major iteration. The transpose of [`iter_rows`](Self::iter_rows), and the one
    /// that reads a time series down a column or a record across a header row.
    ///
    /// A whole column of a million-row sheet is a million cells in one vector, which is
    /// larger than the row-major form's working set. That is inherent to the shape rather
    /// than to this implementation, and it is why both are collected rather than streamed:
    /// the borrow of `self` outlives them, so a lazy iterator would need a lifetime the
    /// signature cannot express without a closure-based `for_each`.
    pub fn iter_cols(&self) -> Vec<Vec<Option<&Cell>>> {
        let (min_col, min_row, max_col, max_row) = self.used_bounds();
        self.iter_cols_within(min_col, max_col, min_row, max_row)
    }

    /// Columns from `min_col` to `max_col`, inclusive, each `min_row` to `max_row`,
    /// inclusive.
    pub fn iter_cols_within(
        &self,
        min_col: u32,
        max_col: u32,
        min_row: u32,
        max_row: u32,
    ) -> Vec<Vec<Option<&Cell>>> {
        (min_col..=max_col)
            .map(|column| {
                (min_row..=max_row)
                    .map(|row| self.get_cell_by_index(column, row))
                    .collect()
            })
            .collect()
    }

    /// The sheet's used range as inclusive `(min_col, min_row, max_col, max_row)`.
    ///
    /// An empty sheet is `1, 1, 1, 1`, which yields exactly one empty cell rather than
    /// nothing. That is the same shape a sheet with one blank cell has, which is the right
    /// ambiguity: a caller iterating an empty sheet should not have to special-case it.
    fn used_bounds(&self) -> (u32, u32, u32, u32) {
        (
            1,
            1,
            self.highest_column().max(1),
            self.highest_row().max(1),
        )
    }

    /// The cell at a 1-based column and row, or `None` if it is empty.
    fn get_cell_by_index(&self, column: u32, row: u32) -> Option<&Cell> {
        let letters = get_column_letter(column).ok()?;
        self.get_cell(&format!("{letters}{row}"))
    }

    /// Write a row of values starting at `start_column` on `row`.
    pub fn write_row(&mut self, row: u32, start_column: u32, values: &[CellValue]) -> Result<()> {
        for (offset, value) in values.iter().enumerate() {
            let coordinate = format!(
                "{}{}",
                get_column_letter(start_column + offset as u32)?,
                row
            );
            self.set_cell_value(&coordinate, value.clone())?;
        }
        Ok(())
    }

    /// Append a row of values below the current data.
    ///
    /// `columns` maps a zero-based or letter key to a value; passing an empty slice appends
    /// nothing.
    pub fn append(&mut self, values: &[(u32, CellValue)]) -> Result<()> {
        // openpyxl uses `len(row_dimensions)` as the row index, which counts the distinct
        // rows that have been touched rather than the highest row used. The port reproduces
        // that rule so a sequence of appends matches.
        let row = self.row_dimensions.len() as u32 + 1;
        for (column, value) in values {
            let coordinate = format!("{}{}", get_column_letter(*column + 1)?, row);
            self.set_cell_value(&coordinate, value.clone())?;
        }
        Ok(())
    }

    /// Set the printer settings, validating the orientation.
    pub fn set_printer_settings(&mut self, paper_size: &str, orientation: &str) -> Result<()> {
        self.page_setup.paper_size = Some(paper_size.to_string());
        if orientation != Worksheet::ORIENTATION_PORTRAIT
            && orientation != Worksheet::ORIENTATION_LANDSCAPE
        {
            return Err(Error::value(format!(
                "Values should be {} or {}",
                Worksheet::ORIENTATION_PORTRAIT,
                Worksheet::ORIENTATION_LANDSCAPE
            )));
        }
        self.page_setup.orientation = Some(orientation.to_string());
        Ok(())
    }

    /// Add a relationship for this sheet, returning its index.
    pub fn create_relationship(&mut self, relationship_type: RelationshipType) -> usize {
        self.relationships
            .push(Relationship::new(relationship_type, None));
        let index = self.relationships.len() - 1;
        self.relationships[index].id = Some(format!("rId{}", index + 1));
        index
    }

    /// Add a data-validation rule.
    pub fn add_data_validation(&mut self, validation: DataValidation) {
        self.data_validations.push(validation);
    }

    /// Attach an image.
    pub fn add_image(&mut self, image: Image) {
        self.images.push(image);
    }

    /// Merge a range, blanking every cell but the top-left one.
    pub fn merge_cells(&mut self, range_string: &str) -> Result<()> {
        let parts: Vec<&str> = range_string.split(':').collect();
        if parts.len() != 2 {
            return Err(Error::InsufficientCoordinates(
                "Range must be a cell range (e.g. A1:E1)".into(),
            ));
        }
        let range_string = range_string.replace('$', "");
        let bounds = self.range_bounds(&range_string)?;
        for row in bounds.min_row..=bounds.max_row {
            // `max_col` is an exclusive bound, so `A1:D1` covers columns 1 through 4.
            for column in bounds.min_col..bounds.max_col {
                if row == bounds.min_row && column == bounds.min_col {
                    continue;
                }
                let coordinate = format!("{}{}", get_column_letter(column)?, row);
                self.create_cell(&coordinate)?;
                let cell = self.cells.get_mut(&coordinate).expect("cell just created");
                cell.set_value(CellValue::None, self.context);
                cell.merged = true;
            }
        }
        if !self.merged_cells.contains(&range_string) {
            self.merged_cells.push(range_string);
        }
        Ok(())
    }

    /// Remove a merge.
    pub fn unmerge_cells(&mut self, range_string: &str) -> Result<()> {
        let range_string = range_string.replace('$', "");
        if !self.merged_cells.contains(&range_string) {
            return Err(Error::InsufficientCoordinates(format!(
                "Cell range {range_string} not known as merged."
            )));
        }
        self.merged_cells.retain(|item| item != &range_string);
        let bounds = self.range_bounds(&range_string)?;
        for row in bounds.min_row..=bounds.max_row {
            // `max_col` is an exclusive bound, so `A1:D1` covers columns 1 through 4.
            for column in bounds.min_col..bounds.max_col {
                if row == bounds.min_row && column == bounds.min_col {
                    continue;
                }
                let coordinate = format!("{}{}", get_column_letter(column)?, row);
                self.create_cell(&coordinate)?;
                self.cells
                    .get_mut(&coordinate)
                    .expect("cell just created")
                    .merged = false;
            }
        }
        Ok(())
    }

    /// The merged ranges.
    pub fn merged_cells(&self) -> &[String] {
        &self.merged_cells
    }

    /// Freeze panes at the given top-left cell; `A1` clears the freeze.
    pub fn set_freeze_panes(&mut self, top_left_cell: &str) {
        let normalised = top_left_cell.to_uppercase();
        self.freeze_panes = if normalised.is_empty() || normalised == "A1" {
            None
        } else {
            Some(normalised)
        };
    }

    /// The comment count used by the writer.
    pub fn comment_count(&self) -> usize {
        self.comment_count
    }

    /// Recompute the comment count from the stored cells.
    pub fn recount_comments(&mut self) {
        self.comment_count = self.cells.values().filter(|c| c.comment.is_some()).count();
    }

    /// Attach a comment to a cell.
    pub fn set_comment(&mut self, coordinate: &str, comment: Option<Comment>) -> Result<()> {
        self.create_cell(coordinate)?;
        self.cells
            .get_mut(coordinate)
            .expect("cell just created")
            .comment = comment;
        self.recount_comments();
        Ok(())
    }

    /// The pixel anchor of a cell, measured from the top-left of the sheet.
    pub fn cell_anchor(&self, coordinate: &str) -> Result<(i64, i64)> {
        let (column, row) = coordinate_from_string(coordinate)?;
        let left_columns = column_index_from_string(&column)? - 1;
        let dpi = NonZeroU32::new(96).expect("96 dpi");
        let default_width = points_to_pixels(DEFAULT_COLUMN_WIDTH, dpi);
        let mut left_anchor = 0i64;
        for index in 1..=left_columns {
            let letter = get_column_letter(index)?;
            match self.column_dimensions.get(&letter) {
                Some(dimension) if dimension.width > 0.0 => {
                    left_anchor += points_to_pixels(dimension.width, dpi);
                }
                _ => left_anchor += default_width,
            }
        }
        let default_height = points_to_pixels(DEFAULT_ROW_HEIGHT, dpi);
        let mut top_anchor = 0i64;
        for index in 1..=row.saturating_sub(1) {
            match self.row_dimensions.get(&index) {
                Some(dimension) if dimension.height > 0.0 => {
                    top_anchor += points_to_pixels(dimension.height, dpi);
                }
                _ => top_anchor += default_height,
            }
        }
        Ok((left_anchor, top_anchor))
    }

    /// Which cell lies under the given pixel position.
    pub fn point_pos(&self, left: i64, top: i64) -> Result<(String, u32)> {
        let dpi = NonZeroU32::new(96).expect("96 dpi");
        let default_width = points_to_pixels(DEFAULT_COLUMN_WIDTH, dpi);
        let default_height = points_to_pixels(DEFAULT_ROW_HEIGHT, dpi);
        let mut current_col = 1u32;
        let mut left_pos = 0i64;
        let mut letter;
        loop {
            letter = get_column_letter(current_col)?;
            current_col += 1;
            match self.column_dimensions.get(&letter) {
                Some(dimension) if dimension.width > 0.0 => {
                    left_pos += points_to_pixels(dimension.width, dpi);
                }
                _ => left_pos += default_width,
            }
            if left_pos > left {
                break;
            }
            if current_col > crate::cell::utils::MAX_COLUMN_INDEX {
                break;
            }
        }
        let mut current_row = 1u32;
        let mut top_pos = 0i64;
        let mut row;
        loop {
            row = current_row;
            current_row += 1;
            match self.row_dimensions.get(&row) {
                Some(dimension) if dimension.height > 0.0 => {
                    top_pos += points_to_pixels(dimension.height, dpi);
                }
                _ => top_pos += default_height,
            }
            if top_pos > top {
                break;
            }
            if current_row > crate::xml::constants::MAX_ROW {
                break;
            }
        }
        Ok((letter, row))
    }

    /// Set a cell's hyperlink, creating the relationship and, when empty, the display value.
    pub fn set_hyperlink(&mut self, coordinate: &str, target: &str) -> Result<String> {
        let relationship_index = match self
            .cells
            .get(coordinate)
            .and_then(|cell| cell.hyperlink_rel_id.clone())
        {
            Some(_) => self
                .relationships
                .iter()
                .position(|rel| {
                    self.cells
                        .get(coordinate)
                        .and_then(|cell| cell.hyperlink_rel_id.clone())
                        .as_deref()
                        == rel.id.as_deref()
                })
                .unwrap_or_else(|| self.create_relationship(RelationshipType::Hyperlink)),
            None => self.create_relationship(RelationshipType::Hyperlink),
        };
        let id = self.relationships[relationship_index]
            .id
            .clone()
            .unwrap_or_else(|| format!("rId{}", relationship_index + 1));
        self.relationships[relationship_index].target = Some(target.to_string());
        self.relationships[relationship_index].target_mode = Some("External".to_string());
        self.create_cell(coordinate)?;
        let cell = self.cells.get_mut(coordinate).expect("cell just created");
        cell.hyperlink_rel_id = Some(id.clone());
        cell.set_hyperlink(target);
        Ok(id)
    }

    /// The cells that carry a hyperlink.
    pub fn cells_with_hyperlinks(&self) -> impl Iterator<Item = &Cell> {
        self.cells
            .values()
            .filter(|cell| cell.hyperlink_rel_id.is_some())
    }

    /// The data type recorded for a coordinate, if the cell exists.
    pub fn cell_data_type(&self, coordinate: &str) -> Option<DataType> {
        self.cells.get(coordinate).map(|cell| cell.data_type)
    }

    /// Add a print-title named range definition for this sheet.
    ///
    /// Returns the range string so the caller can register it as a named range.
    pub fn print_title_range(n: u32, rows_or_cols: &str) -> Result<String> {
        if rows_or_cols == "cols" {
            Ok(format!("$A:${}", get_column_letter(n)?))
        } else {
            Ok(format!("$1:${n}"))
        }
    }
}

impl std::fmt::Display for Worksheet {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "<Worksheet \"{}\">", self.title)
    }
}

#[test]
fn rows_come_back_left_to_right() {
    let mut ws = Worksheet::new("S").expect("title");
    ws.set("A1", 1).expect("A1");
    ws.set("C1", 3).expect("C1");
    ws.set("A2", 4).expect("A2");

    let rows = ws.iter_rows();
    // B1 is empty but still occupies its place, so the rows line up with the columns.
    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0].len(), 3);
    assert_eq!(
        rows[0][0].map(|c| c.internal_value().clone()),
        Some(CellValue::Number(1.0))
    );
    assert!(rows[0][1].is_none(), "B1 is empty but present");
    assert_eq!(
        rows[0][2].map(|c| c.internal_value().clone()),
        Some(CellValue::Number(3.0))
    );
    assert_eq!(
        rows[1][0].map(|c| c.internal_value().clone()),
        Some(CellValue::Number(4.0))
    );
}

#[test]
fn columns_come_back_top_to_bottom() {
    let mut ws = Worksheet::new("S").expect("title");
    ws.set("A1", 1).expect("A1");
    ws.set("A3", 3).expect("A3");
    ws.set("B1", 4).expect("B1");

    let cols = ws.iter_cols();
    assert_eq!(cols.len(), 2);
    assert_eq!(cols[0].len(), 3);
    assert_eq!(
        cols[0][0].map(|c| c.internal_value().clone()),
        Some(CellValue::Number(1.0))
    );
    assert!(cols[0][1].is_none(), "A2 is empty but present");
    assert_eq!(
        cols[0][2].map(|c| c.internal_value().clone()),
        Some(CellValue::Number(3.0))
    );
}

#[test]
fn rows_and_columns_are_transposes_of_each_other() {
    let mut ws = Worksheet::new("S").expect("title");
    for (coordinate, value) in [("A1", 1), ("B1", 2), ("A2", 3), ("B2", 4)] {
        ws.set(coordinate, value).expect("set");
    }
    let rows = ws.iter_rows();
    let cols = ws.iter_cols();
    for (r, row) in rows.iter().enumerate() {
        for (c, cell) in row.iter().enumerate() {
            let value = cell.map(|c| c.internal_value().clone());
            assert_eq!(
                value,
                cols[c][r].map(|c| c.internal_value().clone()),
                "row {r} column {c} disagrees with its transpose"
            );
        }
    }
}

#[test]
fn an_empty_sheet_yields_one_empty_row() {
    let ws = Worksheet::new("S").expect("title");
    let rows = ws.iter_rows();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].len(), 1);
    assert!(rows[0][0].is_none());
}

#[test]
fn the_bounds_are_inclusive_at_both_ends() {
    // A1:B5 is five rows of two columns. An exclusive upper bound here would silently
    // drop row 5 and column B, which is the sort of off-by-one that reads as data loss.
    let mut ws = Worksheet::new("S").expect("title");
    ws.set("B5", 1).expect("B5");
    let rows = ws.iter_rows_within(1, 5, 1, 2);
    assert_eq!(rows.len(), 5);
    assert!(rows.iter().all(|row| row.len() == 2));
    assert!(rows[4][1].is_some(), "B5 is inside the bounds");
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::NaiveDate;

    fn sheet() -> Worksheet {
        Worksheet::new("Sheet1").unwrap()
    }

    #[test]
    fn titles_are_validated() {
        let mut ws = sheet();
        assert!(ws.set_title("My Sheet", &["Sheet1".to_string()]).is_ok());
        assert_eq!(ws.title(), "My Sheet");
        for bad in ["a/b", "a\\b", "a*b", "a?b", "a:b", "a[b", "a]b"] {
            assert!(matches!(ws.set_title(bad, &[]), Err(Error::SheetTitle(_))));
        }
        let long = "x".repeat(32);
        assert!(matches!(
            ws.set_title(&long, &[]),
            Err(Error::SheetTitle(_))
        ));
        assert_eq!(
            ws.title(),
            "My Sheet",
            "failed set must not clobber the title"
        );
    }

    #[test]
    fn duplicate_titles_get_a_counter() {
        let existing = vec!["Sheet".to_string(), "Sheet2".to_string()];
        let mut ws = sheet();
        ws.set_title("Sheet", &existing).unwrap();
        assert_eq!(ws.title(), "Sheet3");
    }

    #[test]
    fn empty_title_gets_a_default() {
        let ws = Worksheet::with_title("", 0).unwrap();
        assert_eq!(ws.title(), "Sheet1");
        let ws = Worksheet::with_title("", 1).unwrap();
        assert_eq!(ws.title(), "Sheet2");
    }

    #[test]
    fn cells_are_created_lazily_with_dimensions() {
        let mut ws = sheet();
        assert_eq!(ws.cell_count(), 0);
        ws.cell("B3").unwrap();
        assert_eq!(ws.cell_count(), 1);
        assert!(ws.column_dimensions.contains_key("B"));
        assert!(ws.row_dimensions.contains_key(&3));
    }

    #[test]
    fn values_round_trip() {
        let mut ws = sheet();
        ws.set("A1", "hello").unwrap();
        assert_eq!(ws.cell_value("A1"), Some(CellValue::text("hello")));
        ws.set("A2", 42.0).unwrap();
        assert_eq!(ws.cell_value("A2"), Some(CellValue::Number(42.0)));
        ws.set("A3", true).unwrap();
        assert_eq!(ws.cell_value("A3"), Some(CellValue::Bool(true)));
        ws.set("A4", CellValue::None).unwrap();
        assert_eq!(ws.cell_value("A4"), Some(CellValue::Text(String::new())));
    }

    #[test]
    fn dates_get_a_number_format_and_convert_back() {
        let mut ws = sheet();
        let date = NaiveDate::from_ymd_opt(2010, 1, 18).unwrap();
        ws.set("A1", CellValue::Date(date)).unwrap();
        assert_eq!(ws.number_format("A1"), "yyyy-mm-dd");
        assert_eq!(ws.cell_value("A1"), Some(CellValue::Date(date)));
    }

    #[test]
    fn dimensions_track_data() {
        let mut ws = sheet();
        assert_eq!(ws.highest_row(), 1);
        assert_eq!(ws.highest_column(), 1);
        ws.set("C5", 1.0).unwrap();
        assert_eq!(ws.highest_row(), 5);
        assert_eq!(ws.highest_column(), 3);
        assert_eq!(ws.calculate_dimension().unwrap(), "A1:C5");
    }

    #[test]
    fn ranges_are_read_row_major() {
        let mut ws = sheet();
        ws.set("A1", 1.0).unwrap();
        ws.set("B1", 2.0).unwrap();
        ws.set("A2", 3.0).unwrap();
        let rows = ws.range_values("A1:B2").unwrap();
        assert_eq!(rows.len(), 2);
        assert_eq!(
            rows[0],
            vec![CellValue::Number(1.0), CellValue::Number(2.0)]
        );
        assert_eq!(rows[1][0], CellValue::Number(3.0));
        // The gap in row 2 reads as an empty cell rather than a missing one.
        assert_eq!(rows[1][1], CellValue::None);
        assert_eq!(ws.range_coordinates("A1:B1").unwrap().len(), 2);
    }

    #[test]
    fn writing_rows() {
        let mut ws = sheet();
        ws.write_row(1, 1, &[CellValue::text("a"), CellValue::Number(2.0)])
            .unwrap();
        assert_eq!(ws.cell_value("A1"), Some(CellValue::text("a")));
        assert_eq!(ws.cell_value("B1"), Some(CellValue::Number(2.0)));
    }

    #[test]
    fn merge_and_unmerge() {
        let mut ws = sheet();
        ws.merge_cells("A1:B2").unwrap();
        assert_eq!(ws.merged_cells(), &["A1:B2".to_string()]);
        // Non-anchor cells exist and are marked merged and empty.
        let cell = ws.get_cell("B2").unwrap();
        assert!(cell.merged);
        assert!(cell.internal_value().is_empty());

        ws.unmerge_cells("A1:B2").unwrap();
        assert!(ws.merged_cells().is_empty());
        assert!(!ws.get_cell("B2").unwrap().merged);

        assert!(ws.merge_cells("A1").is_err());
        assert!(ws.unmerge_cells("A9:B9").is_err());
    }

    #[test]
    fn merging_stays_inside_the_range() {
        // `RangeBounds::max_col` is exclusive, the way openpyxl's `get_range_boundaries`
        // makes it. Treating it as inclusive put a blank cell one column past the merge,
        // which inflated the sheet's dimension and wrote a cell the range never covered.
        let mut ws = sheet();
        ws.merge_cells("A6:D6").unwrap();
        for coordinate in ["B6", "C6", "D6"] {
            assert!(
                ws.get_cell(coordinate).is_some_and(|cell| cell.merged),
                "{coordinate}"
            );
        }
        assert!(ws.get_cell("E6").is_none(), "E6 is outside A6:D6");
        assert_eq!(ws.calculate_dimension().unwrap(), "A1:D6");

        ws.unmerge_cells("A6:D6").unwrap();
        assert!(
            ws.get_cell("E6").is_none(),
            "unmerging must not create E6 either"
        );
    }

    #[test]
    fn garbage_collect_drops_empty_cells() {
        let mut ws = sheet();
        ws.set("A1", "x").unwrap();
        ws.set("B1", "").unwrap();
        ws.set("C1", CellValue::None).unwrap();
        ws.get_style_mut("D1").unwrap().font.bold = true;
        ws.set("E1", "").unwrap();
        ws.garbage_collect();
        assert!(ws.get_cell("A1").is_some());
        assert!(ws.get_cell("B1").is_none());
        assert!(ws.get_cell("C1").is_none());
        // A styled empty cell survives.
        assert!(ws.get_cell("D1").is_some());
        assert!(ws.get_cell("E1").is_none());
    }

    #[test]
    fn static_styles_are_copied_on_write() {
        let mut ws = sheet();
        let mut static_style = Style::static_style();
        static_style.font.bold = true;
        let table_entry = static_style.clone();
        ws.set_style("A1", static_style).unwrap();
        assert!(ws.get_style_read_only("A1").unwrap().is_static);
        let style = ws.get_style_mut("A1").unwrap();
        assert!(!style.is_static);
        assert!(style.font.bold);
        style.font.italic = true;
        // The first mutable access replaced the shared entry with a private copy, so the
        // caller's own copy is untouched by the edit.
        assert!(table_entry.is_static);
        assert!(!table_entry.font.italic);
        // The sheet now owns the edited copy.
        assert!(!ws.get_style_read_only("A1").unwrap().is_static);
        assert!(ws.get_style_read_only("A1").unwrap().font.italic);
    }

    #[test]
    fn freeze_panes_normalise() {
        let mut ws = sheet();
        ws.set_freeze_panes("b2");
        assert_eq!(ws.freeze_panes.as_deref(), Some("B2"));
        ws.set_freeze_panes("A1");
        assert_eq!(ws.freeze_panes, None);
        ws.set_freeze_panes("");
        assert_eq!(ws.freeze_panes, None);
    }

    #[test]
    fn printer_settings_validate_orientation() {
        let mut ws = sheet();
        ws.set_printer_settings("9", Worksheet::ORIENTATION_LANDSCAPE)
            .unwrap();
        assert_eq!(ws.page_setup.paper_size.as_deref(), Some("9"));
        assert_eq!(ws.page_setup.orientation.as_deref(), Some("landscape"));
        assert!(ws.set_printer_settings("9", "sideways").is_err());
    }

    #[test]
    fn hyperlinks_create_relationships() {
        let mut ws = sheet();
        let id = ws.set_hyperlink("A1", "http://example.com").unwrap();
        assert_eq!(id, "rId1");
        let cell = ws.get_cell("A1").unwrap();
        assert_eq!(cell.hyperlink(), "http://example.com");
        assert_eq!(cell.hyperlink_rel_id.as_deref(), Some("rId1"));
        assert_eq!(ws.relationships.len(), 1);
        assert_eq!(ws.cells_with_hyperlinks().count(), 1);

        // Setting a second hyperlink on the same cell reuses the relationship.
        ws.set_hyperlink("A1", "http://other.com").unwrap();
        assert_eq!(ws.relationships.len(), 1);
        assert_eq!(ws.get_cell("A1").unwrap().hyperlink(), "http://other.com");
    }

    #[test]
    fn comments_are_counted() {
        let mut ws = sheet();
        ws.set_comment("A1", Some(Comment::new("hi", "me")))
            .unwrap();
        assert_eq!(ws.comment_count(), 1);
        ws.set_comment("B2", Some(Comment::new("yo", "me")))
            .unwrap();
        assert_eq!(ws.comment_count(), 2);
        ws.set_comment("A1", None).unwrap();
        assert_eq!(ws.comment_count(), 1);
    }

    #[test]
    fn anchors_and_point_positions() {
        let mut ws = sheet();
        ws.set("D4", 1.0).unwrap();
        let anchor = ws.cell_anchor("D4").unwrap();
        assert!(anchor.0 > 0 && anchor.1 > 0);
        let (column, row) = ws.point_pos(anchor.0 + 1, anchor.1 + 1).unwrap();
        assert!(!column.is_empty());
        assert!(row > 0);
    }

    #[test]
    fn append_writes_to_the_next_row() {
        let mut ws = sheet();
        ws.set("A1", 1.0).unwrap();
        ws.append(&[(0, CellValue::text("x")), (2, CellValue::Number(3.0))])
            .unwrap();
        // The append lands on the next row rather than overwriting the existing data.
        assert_eq!(ws.cell_value("A1"), Some(CellValue::Number(1.0)));
        assert_eq!(ws.cell_value("A2"), Some(CellValue::text("x")));
        assert_eq!(ws.cell_value("C2"), Some(CellValue::Number(3.0)));
    }

    #[test]
    fn print_title_ranges() {
        assert_eq!(Worksheet::print_title_range(2, "rows").unwrap(), "$1:$2");
        assert_eq!(Worksheet::print_title_range(3, "cols").unwrap(), "$A:$C");
    }

    #[test]
    fn display_formatting() {
        let ws = sheet();
        assert_eq!(ws.to_string(), "<Worksheet \"Sheet1\">");
    }
}
