//! Row and column dimensions (`openpyxl/worksheet/dimensions.rs`).

use crate::cell::utils::{column_index_from_string, get_column_letter};
use crate::exceptions::Result;

/// Information about the display properties of a row or column.
#[derive(Debug, Clone, PartialEq)]
pub struct Dimension {
    /// Row number, or column letters for a column dimension.
    pub index: String,
    /// Whether the row/column is visible.
    pub visible: bool,
    /// Outline (grouping) level.
    pub outline_level: u32,
    /// Whether the row/column is collapsed.
    pub collapsed: bool,
}

/// Information about the display properties of a row.
#[derive(Debug, Clone, PartialEq)]
pub struct RowDimension {
    /// The 1-based row number.
    pub index: u32,
    /// Row height in points; `-1` means "use the default".
    pub height: f64,
    /// Whether the row is visible.
    pub visible: bool,
    /// Outline (grouping) level.
    pub outline_level: u32,
    /// Whether the row is collapsed.
    pub collapsed: bool,
}

impl RowDimension {
    /// Build a row dimension with defaults.
    pub fn new(index: u32) -> Self {
        RowDimension {
            index,
            height: -1.0,
            visible: true,
            outline_level: 0,
            collapsed: false,
        }
    }

    /// Chainable height setter.
    pub fn with_height(mut self, height: f64) -> Self {
        self.height = height;
        self
    }

    /// Chainable visibility setter.
    pub fn with_visible(mut self, visible: bool) -> Self {
        self.visible = visible;
        self
    }

    /// The shared fields, as a [`Dimension`].
    pub fn as_dimension(&self) -> Dimension {
        Dimension {
            index: self.index.to_string(),
            visible: self.visible,
            outline_level: self.outline_level,
            collapsed: self.collapsed,
        }
    }
}

/// Information about the display properties of a column.
#[derive(Debug, Clone, PartialEq)]
pub struct ColumnDimension {
    /// The column letters.
    pub index: String,
    /// Column width in characters; `-1` means "use the default".
    pub width: f64,
    /// Whether the width should be auto-fitted.
    pub auto_size: bool,
    /// Whether the column is visible.
    pub visible: bool,
    /// Outline (grouping) level.
    pub outline_level: u32,
    /// Whether the column is collapsed.
    pub collapsed: bool,
}

impl ColumnDimension {
    /// Build a column dimension with defaults.
    pub fn new(index: &str) -> Self {
        ColumnDimension {
            index: index.to_uppercase(),
            width: -1.0,
            auto_size: false,
            visible: true,
            outline_level: 0,
            collapsed: false,
        }
    }

    /// Build a column dimension from a 1-based column index.
    pub fn from_index(index: u32) -> Result<Self> {
        Ok(ColumnDimension::new(&get_column_letter(index)?))
    }

    /// The 1-based column index.
    pub fn column_index(&self) -> Result<u32> {
        column_index_from_string(&self.index)
    }

    /// Chainable width setter.
    pub fn with_width(mut self, width: f64) -> Self {
        self.width = width;
        self
    }

    /// Chainable auto-size setter.
    pub fn with_auto_size(mut self, auto_size: bool) -> Self {
        self.auto_size = auto_size;
        self
    }

    /// Chainable visibility setter.
    pub fn with_visible(mut self, visible: bool) -> Self {
        self.visible = visible;
        self
    }

    /// The shared fields, as a [`Dimension`].
    pub fn as_dimension(&self) -> Dimension {
        Dimension {
            index: self.index.clone(),
            visible: self.visible,
            outline_level: self.outline_level,
            collapsed: self.collapsed,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn row_defaults() {
        let row = RowDimension::new(1);
        assert_eq!(row.index, 1);
        assert_eq!(row.height, -1.0);
        assert!(row.visible);
        assert_eq!(row.outline_level, 0);
        assert!(!row.collapsed);
    }

    #[test]
    fn column_defaults_and_index_round_trip() {
        let col = ColumnDimension::from_index(27).unwrap();
        assert_eq!(col.index, "AA");
        assert_eq!(col.column_index().unwrap(), 27);
        assert_eq!(col.width, -1.0);
        assert!(!col.auto_size);
    }

    #[test]
    fn builders_apply() {
        let row = RowDimension::new(3).with_height(30.0).with_visible(false);
        assert_eq!(row.height, 30.0);
        assert!(!row.visible);

        let col = ColumnDimension::new("c")
            .with_width(12.5)
            .with_auto_size(true);
        assert_eq!(col.index, "C");
        assert_eq!(col.width, 12.5);
        assert!(col.auto_size);
    }

    #[test]
    fn dimension_view_shares_common_fields() {
        let dim = RowDimension::new(2).with_visible(false).as_dimension();
        assert_eq!(dim.index, "2");
        assert!(!dim.visible);
        assert_eq!(dim.outline_level, 0);
    }
}
