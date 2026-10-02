//! Excel tables — ListObjects (`openpyxl/worksheet/table.py`).
//!
//! A table is a named rectangle of cells with a header row, optional banding, and per-column
//! names that structured references are built from. `=SUM(Table1[Sales])` means nothing to
//! the workbook engine without the table part that gives those names meaning, so a table is
//! not a formatting convenience bolted onto a range: it is a separate XML part with its own
//! relationship, and Excel will not open a file whose `<tableParts>` points at a missing one.
//!
//! The rule that catches people: **column names must match the header cells exactly.** Excel
//! rewrites the names to match the cells when a file is opened, and a mismatch shows up as a
//! silently renamed column and a broken structured reference. [`Table::initialise_columns`]
//! therefore reads the header row rather than inventing names.
//!
//! ```no_run
//! use ferroxl::worksheet::table::{Table, TableStyleInfo};
//!
//! let mut table = Table::new("Sales", "A1:C10").unwrap();
//! table.set_header_row(false);
//! table.style_info = TableStyleInfo::banded();
//! ```

use std::collections::BTreeMap;
use std::fmt;

use crate::exceptions::{Error, Result};
use crate::worksheet::cell_range::CellRange;
use crate::worksheet::Worksheet;

/// The functions a totals row can compute.
///
/// openpyxl accepts these as strings, and so does this, because the XML stores a string and a
/// closed Rust enum would only make the value harder to read out of a loaded file.
pub const TOTALS_ROW_FUNCTIONS: [&str; 9] = [
    "sum",
    "min",
    "max",
    "average",
    "count",
    "countNums",
    "stdDev",
    "var",
    "custom",
];

/// Whether and how a table is banded.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TableStyleInfo {
    /// The style name, e.g. `TableStyleMedium9`. `None` means no built-in style.
    pub name: Option<String>,
    /// Emphasise the first column.
    pub show_first_column: Option<bool>,
    /// Emphasise the last column.
    pub show_last_column: Option<bool>,
    /// Draw alternating row bands.
    pub show_row_stripes: Option<bool>,
    /// Draw alternating column bands.
    pub show_column_stripes: Option<bool>,
}

impl TableStyleInfo {
    /// No styling at all.
    pub fn new() -> Self {
        TableStyleInfo::default()
    }

    /// A named style with row banding, which is what Excel applies by default.
    pub fn banded() -> Self {
        TableStyleInfo {
            name: Some("TableStyleMedium9".to_string()),
            show_row_stripes: Some(true),
            ..TableStyleInfo::default()
        }
    }

    /// Set the style name.
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Turn row banding on or off.
    pub fn with_row_stripes(mut self, stripes: bool) -> Self {
        self.show_row_stripes = Some(stripes);
        self
    }

    /// Turn column banding on or off.
    pub fn with_column_stripes(mut self, stripes: bool) -> Self {
        self.show_column_stripes = Some(stripes);
        self
    }

    /// Emphasise the first and last columns.
    pub fn with_outer_columns(mut self, emphasis: bool) -> Self {
        self.show_first_column = Some(emphasis);
        self.show_last_column = Some(emphasis);
        self
    }
}

/// A formula belonging to a table column.
///
/// `calculated_column_formula` is the one Excel fills down an entire column; `totals_row_formula`
/// is the one in the totals row. The text is the formula without its leading `=`, as the XML
/// stores it.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TableFormula {
    /// Whether the formula is an array formula.
    pub array: Option<bool>,
    /// The formula text, without a leading `=`.
    pub text: String,
}

impl TableFormula {
    /// A formula from its text, with or without a leading `=`.
    pub fn new(text: &str) -> Self {
        TableFormula {
            array: None,
            text: text.strip_prefix('=').unwrap_or(text).to_string(),
        }
    }

    /// Mark the formula as an array formula.
    pub fn as_array(mut self) -> Self {
        self.array = Some(true);
        self
    }
}

/// One column of a table.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TableColumn {
    /// The column's 1-based index within the table.
    pub id: u32,
    /// The name structured references use.
    ///
    /// Must match the header cell exactly. Excel rewrites the part when it disagrees with the
    /// cell, which turns `=SUM(Table1[Sales])` into a broken reference without saying so.
    pub name: String,
    /// The function this column's totals row cell computes.
    pub totals_row_function: Option<String>,
    /// Literal text for the totals row cell, used instead of a function.
    pub totals_row_label: Option<String>,
    /// The differential style index for the header row.
    pub header_row_dxf_id: Option<u32>,
    /// The differential style index for the data cells.
    pub data_dxf_id: Option<u32>,
    /// The differential style index for the totals row.
    pub totals_row_dxf_id: Option<u32>,
    /// The named cell style for the header row.
    pub header_row_cell_style: Option<String>,
    /// The named cell style for the data cells.
    pub data_cell_style: Option<String>,
    /// The named cell style for the totals row.
    pub totals_row_cell_style: Option<String>,
    /// The formula Excel fills down this column.
    pub calculated_column_formula: Option<TableFormula>,
    /// The formula in the totals row cell.
    pub totals_row_formula: Option<TableFormula>,
}

impl TableColumn {
    /// A column with the given index and name.
    pub fn new(id: u32, name: impl Into<String>) -> Self {
        TableColumn {
            id,
            name: name.into(),
            ..TableColumn::default()
        }
    }

    /// Compute the totals row cell with `function`.
    ///
    /// The name is checked against [`TOTALS_ROW_FUNCTIONS`], because Excel rejects an unknown
    /// one when the file is opened and the error names the file rather than the table.
    pub fn with_totals_function(mut self, function: &str) -> Result<Self> {
        if !TOTALS_ROW_FUNCTIONS.contains(&function) {
            return Err(Error::Value(format!(
                "{function} is not a totals row function; expected one of {}",
                TOTALS_ROW_FUNCTIONS.join(", ")
            )));
        }
        self.totals_row_function = Some(function.to_string());
        Ok(self)
    }

    /// Fill this column down from `formula`.
    pub fn with_calculated_formula(mut self, formula: &str) -> Self {
        self.calculated_column_formula = Some(TableFormula::new(formula));
        self
    }
}

/// An Excel table: a named range with columns, styling and a totals row.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Table {
    /// The table's id, which also numbers its part in the package.
    pub id: u32,
    /// The name shown in Excel's field list. May not contain a space.
    pub display_name: String,
    /// The table's name in formulas. Defaults to the display name.
    pub name: Option<String>,
    /// The range the table covers.
    pub reference: String,
    /// A note shown in Excel's table properties.
    pub comment: Option<String>,
    /// `worksheet`, `xml` or `queryTable`. `None` means a normal table.
    pub table_type: Option<String>,
    /// How many header rows. `0` means the table has no header row.
    pub header_row_count: Option<u32>,
    /// Whether inserting a table row inserts a worksheet row too.
    pub insert_row: Option<bool>,
    /// How many totals rows.
    pub totals_row_count: Option<u32>,
    /// Whether the totals row is shown.
    pub totals_row_shown: Option<bool>,
    /// The differential style index for the header row.
    pub header_row_dxf_id: Option<u32>,
    /// The differential style index for the data cells.
    pub data_dxf_id: Option<u32>,
    /// The differential style index for the totals row.
    pub totals_row_dxf_id: Option<u32>,
    /// The named cell style for the header row.
    pub header_row_cell_style: Option<String>,
    /// The named cell style for the data cells.
    pub data_cell_style: Option<String>,
    /// The named cell style for the totals row.
    pub totals_row_cell_style: Option<String>,
    /// The columns, left to right.
    pub columns: Vec<TableColumn>,
    /// The table's style.
    pub style_info: TableStyleInfo,
}

impl Table {
    /// A table called `display_name` over `reference`, with a header row and no columns yet.
    ///
    /// The columns are left empty because they have to be read from the header cells; call
    /// [`Table::initialise_columns`] with the sheet before writing, or set them by hand.
    ///
    /// A display name containing a space is refused. Excel rejects it when the file is opened,
    /// and the error it gives names the file rather than the table, so catching it here is
    /// worth a special case.
    pub fn new(display_name: &str, reference: &str) -> Result<Self> {
        if display_name.contains(' ') {
            return Err(Error::Value(
                "a table name cannot contain a space".to_string(),
            ));
        }
        if display_name.is_empty() {
            return Err(Error::Value("a table needs a name".to_string()));
        }
        CellRange::parse(reference)?;
        Ok(Table {
            id: 1,
            display_name: display_name.to_string(),
            name: None,
            reference: reference.to_string(),
            comment: None,
            table_type: None,
            header_row_count: Some(1),
            insert_row: None,
            totals_row_count: None,
            totals_row_shown: None,
            header_row_dxf_id: None,
            data_dxf_id: None,
            totals_row_dxf_id: None,
            header_row_cell_style: None,
            data_cell_style: None,
            totals_row_cell_style: None,
            columns: Vec::new(),
            style_info: TableStyleInfo::default(),
        })
    }

    /// The name used in formulas, defaulting to the display name.
    pub fn formula_name(&self) -> &str {
        self.name.as_deref().unwrap_or(&self.display_name)
    }

    /// Set the table's part id.
    pub fn with_id(mut self, id: u32) -> Self {
        self.id = id;
        self
    }

    /// Add a note to the table.
    pub fn with_comment(mut self, comment: impl Into<String>) -> Self {
        self.comment = Some(comment.into());
        self
    }

    /// Set the style.
    pub fn with_style(mut self, style: TableStyleInfo) -> Self {
        self.style_info = style;
        self
    }

    /// Add a column.
    pub fn with_column(mut self, column: TableColumn) -> Self {
        self.columns.push(column);
        self
    }

    /// Add columns.
    pub fn with_columns(mut self, columns: impl IntoIterator<Item = TableColumn>) -> Self {
        self.columns.extend(columns);
        self
    }

    /// Show or hide the header row.
    ///
    /// Turning it off sets `header_row_count` to `0`, which is how the XML spells "no header
    /// row" — not an absent attribute.
    pub fn set_header_row(&mut self, present: bool) -> &mut Self {
        self.header_row_count = if present { Some(1) } else { Some(0) };
        self
    }

    /// Show a totals row.
    pub fn set_totals_row(&mut self, shown: bool) -> &mut Self {
        self.totals_row_shown = Some(shown);
        self.totals_row_count = Some(if shown { 1 } else { 0 });
        self
    }

    /// The column names, in order.
    pub fn column_names(&self) -> Vec<&str> {
        self.columns.iter().map(|c| c.name.as_str()).collect()
    }

    /// Read the column names out of the sheet's header row.
    ///
    /// This is the step that makes structured references work. Naming columns `Column1`,
    /// `Column2` without looking at the cells produces a table whose references Excel silently
    /// rewrites on open.
    ///
    /// A header cell holding a number rather than text is named by its display form, because
    /// Excel shows the formatted value in the field list. A blank cell keeps the placeholder,
    /// since there is nothing better to call it and an empty name is rejected.
    pub fn initialise_columns(&mut self, sheet: &Worksheet) -> Result<()> {
        let bounds = CellRange::parse(&self.reference)?;
        let has_header = self.header_row_count.unwrap_or(1) != 0;
        let header_row = if has_header {
            bounds.min_row
        } else {
            // No header row: the first row is data, so the table has one fewer data row and no
            // name to read. The placeholder names are what openpyxl writes too.
            bounds.min_row
        };

        let mut columns = Vec::new();
        for index in 0..(bounds.max_col - bounds.min_col + 1) {
            let column = bounds.min_col + index;
            let name = if has_header {
                let coordinate = format!(
                    "{}{header_row}",
                    crate::cell::utils::get_column_letter(column)?
                );
                match sheet.cell_value(&coordinate) {
                    Some(crate::cell::cell::CellValue::Text(text)) if !text.is_empty() => text,
                    Some(crate::cell::cell::CellValue::Number(number)) => {
                        crate::xml::functions::safe_string(number)
                    }
                    Some(crate::cell::cell::CellValue::DateTime(date)) => {
                        date.format("%Y-%m-%d").to_string()
                    }
                    _ => format!("Column{column}"),
                }
            } else {
                format!("Column{column}")
            };
            columns.push(TableColumn::new(column, name));
        }
        self.columns = columns;
        Ok(())
    }
}

impl fmt::Display for Table {
    /// `name` and `reference`, the way openpyxl's `TableList.items` reports them.
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{} {}", self.formula_name(), self.reference)
    }
}

/// The tables on a worksheet, keyed by name.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableList {
    tables: BTreeMap<String, Table>,
}

impl TableList {
    /// No tables.
    pub fn new() -> Self {
        TableList::default()
    }

    /// Add a table, keyed by its formula name.
    ///
    /// A second table with the same name replaces the first. Excel does not allow two tables
    /// with one name on a sheet, so the alternative is an error on a load that should succeed.
    pub fn add(&mut self, table: Table) -> Result<()> {
        self.tables.insert(table.formula_name().to_string(), table);
        Ok(())
    }

    /// The table with this name.
    pub fn get(&self, name: &str) -> Option<&Table> {
        self.tables.get(name)
    }

    /// The table with this name, for mutation.
    pub fn get_mut(&mut self, name: &str) -> Option<&mut Table> {
        self.tables.get_mut(name)
    }

    /// The table covering this exact range.
    pub fn get_by_range(&self, reference: &str) -> Option<&Table> {
        self.tables.values().find(|t| t.reference == reference)
    }

    /// Remove the table with this name.
    pub fn remove(&mut self, name: &str) -> Option<Table> {
        self.tables.remove(name)
    }

    /// How many tables.
    pub fn len(&self) -> usize {
        self.tables.len()
    }

    /// Whether there are none.
    pub fn is_empty(&self) -> bool {
        self.tables.is_empty()
    }

    /// The tables, in name order.
    pub fn iter(&self) -> impl Iterator<Item = &Table> {
        self.tables.values()
    }

    /// The names, in order.
    pub fn names(&self) -> Vec<&str> {
        self.tables.keys().map(String::as_str).collect()
    }
}

impl fmt::Display for TableList {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let rendered: Vec<String> = self.tables.values().map(|t| t.to_string()).collect();
        f.write_str(&rendered.join(", "))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::cell::CellValue;

    fn sheet_with_headers() -> Worksheet {
        let mut ws = Worksheet::new("S").expect("title");
        ws.set("A1", CellValue::text("Product")).expect("A1");
        ws.set("B1", CellValue::text("Region")).expect("B1");
        ws.set("C1", CellValue::text("Sales")).expect("C1");
        ws
    }

    #[test]
    fn a_table_needs_a_name_and_a_range() {
        assert!(Table::new("Sales", "A1:C10").is_ok());
        assert!(
            Table::new("Sales Data", "A1:C10").is_err(),
            "a space is refused"
        );
        assert!(Table::new("", "A1:C10").is_err());
        assert!(Table::new("Sales", "nonsense").is_err());
    }

    #[test]
    fn the_formula_name_defaults_to_the_display_name() {
        let table = Table::new("Sales", "A1:C10").expect("a table");
        assert_eq!(table.formula_name(), "Sales");
        let mut renamed = table.clone();
        renamed.name = Some("Q3".to_string());
        assert_eq!(renamed.formula_name(), "Q3");
    }

    #[test]
    fn columns_are_read_from_the_header_row() {
        // This is the step that makes `=SUM(Table1[Sales])` resolve. Naming them Column1,
        // Column2 without looking produces references Excel silently rewrites on open.
        let sheet = sheet_with_headers();
        let mut table = Table::new("Sales", "A1:C3").expect("a table");
        table.initialise_columns(&sheet).expect("columns");
        assert_eq!(table.column_names(), ["Product", "Region", "Sales"]);
        assert_eq!(table.columns[0].id, 1);
        assert_eq!(table.columns[2].id, 3);
    }

    #[test]
    fn a_numeric_header_is_named_by_its_display_form() {
        let mut ws = Worksheet::new("S").expect("title");
        ws.set("A1", CellValue::number(2026.0)).expect("A1");
        ws.set("B1", CellValue::text("Total")).expect("B1");
        let mut table = Table::new("T", "A1:B3").expect("a table");
        table.initialise_columns(&ws).expect("columns");
        assert_eq!(table.column_names(), ["2026", "Total"]);
    }

    #[test]
    fn a_blank_header_keeps_a_placeholder() {
        // An empty column name is rejected by Excel, and there is nothing better to call a
        // header cell that has nothing in it.
        let mut ws = Worksheet::new("S").expect("title");
        ws.set("B1", CellValue::text("Total")).expect("B1");
        let mut table = Table::new("T", "A1:B3").expect("a table");
        table.initialise_columns(&ws).expect("columns");
        assert_eq!(table.column_names(), ["Column1", "Total"]);
    }

    #[test]
    fn a_table_without_a_header_row_keeps_placeholders() {
        let sheet = sheet_with_headers();
        let mut table = Table::new("T", "A1:C3").expect("a table");
        table.set_header_row(false);
        table.initialise_columns(&sheet).expect("columns");
        assert_eq!(table.column_names(), ["Column1", "Column2", "Column3"]);
        // `headerRowCount="0"` is how the XML says "no header row", not an absent attribute.
        assert_eq!(table.header_row_count, Some(0));
    }

    #[test]
    fn a_totals_function_is_checked_against_what_excel_accepts() {
        let column = TableColumn::new(1, "Sales");
        assert!(column.clone().with_totals_function("sum").is_ok());
        assert!(column.clone().with_totals_function("average").is_ok());
        let err = column
            .with_totals_function("median")
            .expect_err("not a totals function")
            .to_string();
        assert!(err.contains("countNums"), "{err}");
    }

    #[test]
    fn a_table_list_keys_by_name_and_replaces_on_a_collision() {
        let mut list = TableList::new();
        list.add(Table::new("Sales", "A1:C3").expect("a table"))
            .expect("added");
        list.add(Table::new("Costs", "E1:G3").expect("a table"))
            .expect("added");
        assert_eq!(list.len(), 2);
        assert_eq!(list.names(), ["Costs", "Sales"]);

        // Excel does not allow two tables with one name, so a reload replaces rather than
        // failing on a file that is valid.
        list.add(Table::new("Sales", "A10:C12").expect("a table"))
            .expect("replaced");
        assert_eq!(list.len(), 2);
        assert_eq!(list.get("Sales").expect("Sales").reference, "A10:C12");
        assert_eq!(
            list.get_by_range("E1:G3").expect("by range").display_name,
            "Costs"
        );
        assert_eq!(list.remove("Sales").expect("removed").reference, "A10:C12");
        assert!(list.get("Sales").is_none());
    }

    #[test]
    fn a_table_renders_as_its_name_and_range() {
        let table = Table::new("Sales", "A1:C3").expect("a table");
        assert_eq!(table.to_string(), "Sales A1:C3");
        let mut list = TableList::new();
        list.add(table).expect("added");
        assert_eq!(list.to_string(), "Sales A1:C3");
    }

    #[test]
    fn a_calculated_column_formula_loses_its_equals_sign() {
        // The XML stores the expression without one, and a leading `=` would be written
        // verbatim into `<tableFormula>`.
        let column = TableColumn::new(1, "Total").with_calculated_formula("=[@Qty]*[@Price]");
        let formula = column.calculated_column_formula.expect("a formula");
        assert_eq!(formula.text, "[@Qty]*[@Price]");
    }

    #[test]
    fn banding_defaults_match_excel() {
        let style = TableStyleInfo::banded();
        assert_eq!(style.name.as_deref(), Some("TableStyleMedium9"));
        assert_eq!(style.show_row_stripes, Some(true));
        assert_eq!(style.show_column_stripes, None);
    }
}
