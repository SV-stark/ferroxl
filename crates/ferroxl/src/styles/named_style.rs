//! Named cell styles (`openpyxl/styles/named_styles.py` and `builtins.py`).
//!
//! A named style is a reusable bundle — a font, a fill, a number format — that a cell
//! references by name rather than owning. `cell.style = "Good"` is how banding, totals rows
//! and themed headers are normally built, so without this a spreadsheet that uses them
//! loses them on the next save.
//!
//! Two things make it more than a lookup table:
//!
//! - A cell's `xf` carries an **`xfId`**, an index into `<cellStyleXfs>`. That index is the
//!   link; without it the style is a name in a list that nothing points at.
//! - A **`builtinId`** names one of the 49 styles Excel itself defines. Excel renders a
//!   built-in correctly from the id alone, so [`BUILTIN_STYLES`] is a table of ids and names
//!   rather than the XML blobs openpyxl carries. The blobs exist because openpyxl has to
//!   produce a file other readers can render, which is a different problem from writing one
//!   Excel will accept.
//!
//! ```
//! use ferroxl::styles::named_style::{builtin_style, BUILTIN_STYLES};
//!
//! // "Good" is built in, so applying it needs no definition.
//! assert!(builtin_style("Good").is_some());
//! assert_eq!(BUILTIN_STYLES.len(), 49);
//! ```

use std::fmt;

use crate::exceptions::{Error, Result};
use crate::styles::alignment::Alignment;
use crate::styles::borders::Borders;
use crate::styles::fills::Fill;
use crate::styles::fonts::Font;
use crate::styles::numbers::NumberFormat;
use crate::styles::protection::Protection;
use crate::styles::style::Style;

/// The 49 named styles Excel defines, as `(builtinId, name)`.
///
/// The order and ids match `openpyxl/styles/builtins.py`, so a style named here has the same
/// identity in both libraries. Ids are not contiguous — `0` is followed by `3` — because
/// Excel numbers by internal category, so the gaps are deliberate and must not be closed up.
pub const BUILTIN_STYLES: [(&str, &str); 49] = [
    ("0", "Normal"),
    ("3", "Comma"),
    ("6", "Comma [0]"),
    ("4", "Currency"),
    ("7", "Currency [0]"),
    ("5", "Percent"),
    ("8", "Hyperlink"),
    ("9", "Followed Hyperlink"),
    ("15", "Title"),
    ("16", "Headline 1"),
    ("17", "Headline 2"),
    ("18", "Headline 3"),
    ("19", "Headline 4"),
    ("26", "Good"),
    ("27", "Bad"),
    ("28", "Neutral"),
    ("20", "Input"),
    ("21", "Output"),
    ("22", "Calculation"),
    ("24", "Linked Cell"),
    ("23", "Check Cell"),
    ("11", "Warning Text"),
    ("10", "Note"),
    ("53", "Explanatory Text"),
    ("25", "Total"),
    ("29", "Accent1"),
    ("30", "20 % - Accent1"),
    ("31", "40 % - Accent1"),
    ("32", "60 % - Accent1"),
    ("33", "Accent2"),
    ("34", "20 % - Accent2"),
    ("35", "40 % - Accent2"),
    ("36", "60 % - Accent2"),
    ("37", "Accent3"),
    ("38", "20 % - Accent3"),
    ("39", "40 % - Accent3"),
    ("40", "60 % - Accent3"),
    ("41", "Accent4"),
    ("42", "20 % - Accent4"),
    ("43", "40 % - Accent4"),
    ("44", "60 % - Accent4"),
    ("45", "Accent5"),
    ("46", "20 % - Accent5"),
    ("47", "40 % - Accent5"),
    ("48", "60 % - Accent5"),
    ("49", "Accent6"),
    ("50", "20 % - Accent6"),
    ("51", "40 % - Accent6"),
    ("52", "60 % - Accent6"),
];

/// One built-in style's formatting.
///
/// Generated from openpyxl 3.1.5's `styles/builtins.py` rather than transcribed by hand:
/// 49 entries is well past the point where a typo in one is invisible, and the source is
/// right there. Regenerate rather than editing.
// `size` is an f64, so no `Eq`. Nothing hashes or interns these, so a value comparison is
// enough -- a manual bit comparison would be a lie about what two sizes mean.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct BuiltinStyle {
    /// The `builtinId` Excel uses to identify it.
    pub builtin_id: &'static str,
    /// The name as a cell writes it.
    pub name: &'static str,
    /// The fill's foreground colour, if it has one.
    pub fill_color: Option<&'static str>,
    /// The font colour, as `RRGGBB` or `theme:N`.
    pub font_color: Option<&'static str>,
    /// The font size in points.
    pub size: f64,
    /// Whether the font is bold.
    pub bold: bool,
    /// Whether the font is italic.
    pub italic: bool,
    /// The theme font slot, `major` or `minor`.
    pub scheme: Option<&'static str>,
    /// The font name.
    pub font_name: Option<&'static str>,
    /// The number format code, for the styles that carry one.
    pub number_format: Option<&'static str>,
}

// 49 built-in styles, generated from openpyxl 3.1.5 styles/builtins.py
/// The formatting behind each built-in style, in openpyxl's order.
///
/// Generated from openpyxl 3.1.5's `styles/builtins.py` rather than transcribed by hand:
/// 49 entries is well past the point where a typo in one is invisible, and the source is
/// right there. Regenerate rather than editing.
/// The formatting behind each built-in style, in openpyxl's order.
///
/// Generated from openpyxl 3.1.5's `styles/builtins.py` rather than transcribed by hand:
/// 49 entries is well past the point where a typo in one is invisible, and the source is
/// right there. Regenerate rather than editing.
pub const BUILTIN_DETAILS: [BuiltinStyle; 49] = [
    BuiltinStyle {
        builtin_id: "0",
        name: "Normal",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "3",
        name: "Comma",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: Some("_-* #,##0.00\\ _$_-;\\-* #,##0.00\\ _$_-;_-* \"-\"??\\ _$_-;_-@_-"),
    },
    BuiltinStyle {
        builtin_id: "6",
        name: "Comma [0]",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: Some("_-* #,##0\\ _$_-;\\-* #,##0\\ _$_-;_-* \"-\"\\ _$_-;_-@_-"),
    },
    BuiltinStyle {
        builtin_id: "4",
        name: "Currency",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: Some(
            "_-* #,##0.00\\ \"$\"_-;\\-* #,##0.00\\ \"$\"_-;_-* \"-\"??\\ \"$\"_-;_-@_-",
        ),
    },
    BuiltinStyle {
        builtin_id: "7",
        name: "Currency [0]",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: Some("_-* #,##0\\ \"$\"_-;\\-* #,##0\\ \"$\"_-;_-* \"-\"\\ \"$\"_-;_-@_-"),
    },
    BuiltinStyle {
        builtin_id: "5",
        name: "Percent",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: Some("0%"),
    },
    BuiltinStyle {
        builtin_id: "8",
        name: "Hyperlink",
        fill_color: None,
        font_color: Some("theme:10"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "9",
        name: "Followed Hyperlink",
        fill_color: None,
        font_color: Some("theme:11"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "15",
        name: "Title",
        fill_color: None,
        font_color: Some("theme:3"),
        size: 18.0,
        bold: true,
        italic: false,
        scheme: Some("major"),
        font_name: Some("Cambria"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "16",
        name: "Headline 1",
        fill_color: None,
        font_color: Some("theme:3"),
        size: 15.0,
        bold: true,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "17",
        name: "Headline 2",
        fill_color: None,
        font_color: Some("theme:3"),
        size: 13.0,
        bold: true,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "18",
        name: "Headline 3",
        fill_color: None,
        font_color: Some("theme:3"),
        size: 11.0,
        bold: true,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "19",
        name: "Headline 4",
        fill_color: None,
        font_color: Some("theme:3"),
        size: 11.0,
        bold: true,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "26",
        name: "Good",
        fill_color: Some("FFC6EFCE"),
        font_color: Some("FF006100"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "27",
        name: "Bad",
        fill_color: Some("FFFFC7CE"),
        font_color: Some("FF9C0006"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "28",
        name: "Neutral",
        fill_color: Some("FFFFEB9C"),
        font_color: Some("FF9C6500"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "20",
        name: "Input",
        fill_color: Some("FFFFCC99"),
        font_color: Some("FF3F3F76"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "21",
        name: "Output",
        fill_color: Some("FFF2F2F2"),
        font_color: Some("FF3F3F3F"),
        size: 12.0,
        bold: true,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "22",
        name: "Calculation",
        fill_color: Some("FFF2F2F2"),
        font_color: Some("FFFA7D00"),
        size: 12.0,
        bold: true,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "24",
        name: "Linked Cell",
        fill_color: None,
        font_color: Some("FFFA7D00"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "23",
        name: "Check Cell",
        fill_color: Some("FFA5A5A5"),
        font_color: Some("theme:0"),
        size: 12.0,
        bold: true,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "11",
        name: "Warning Text",
        fill_color: None,
        font_color: Some("FFFF0000"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "10",
        name: "Note",
        fill_color: Some("FFFFFFCC"),
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "53",
        name: "Explanatory Text",
        fill_color: None,
        font_color: Some("FF7F7F7F"),
        size: 12.0,
        bold: false,
        italic: true,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "25",
        name: "Total",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: true,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "29",
        name: "Accent1",
        fill_color: None,
        font_color: Some("theme:0"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "30",
        name: "20 % - Accent1",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "31",
        name: "40 % - Accent1",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "32",
        name: "60 % - Accent1",
        fill_color: None,
        font_color: Some("theme:0"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "33",
        name: "Accent2",
        fill_color: None,
        font_color: Some("theme:0"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "34",
        name: "20 % - Accent2",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "35",
        name: "40 % - Accent2",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "36",
        name: "60 % - Accent2",
        fill_color: None,
        font_color: Some("theme:0"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "37",
        name: "Accent3",
        fill_color: None,
        font_color: Some("theme:0"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "38",
        name: "20 % - Accent3",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "39",
        name: "40 % - Accent3",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "40",
        name: "60 % - Accent3",
        fill_color: None,
        font_color: Some("theme:0"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "41",
        name: "Accent4",
        fill_color: None,
        font_color: Some("theme:0"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "42",
        name: "20 % - Accent4",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "43",
        name: "40 % - Accent4",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "44",
        name: "60 % - Accent4",
        fill_color: None,
        font_color: Some("theme:0"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "45",
        name: "Accent5",
        fill_color: None,
        font_color: Some("theme:0"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "46",
        name: "20 % - Accent5",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "47",
        name: "40 % - Accent5",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "48",
        name: "60 % - Accent5",
        fill_color: None,
        font_color: Some("theme:0"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "49",
        name: "Accent6",
        fill_color: None,
        font_color: Some("theme:0"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "50",
        name: "20 % - Accent6",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "51",
        name: "40 % - Accent6",
        fill_color: None,
        font_color: Some("theme:1"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
    BuiltinStyle {
        builtin_id: "52",
        name: "60 % - Accent6",
        fill_color: None,
        font_color: Some("theme:0"),
        size: 12.0,
        bold: false,
        italic: false,
        scheme: Some("minor"),
        font_name: Some("Calibri"),
        number_format: None,
    },
];

/// Look up a built-in style's `builtinId` by name.
///
/// Case-insensitive, because Excel's own field list matches `good` against `Good` and a
/// caller writing `cell.style = "good"` means the same thing.
pub fn builtin_style(name: &str) -> Option<&'static str> {
    builtin(name).map(|style| style.builtin_id)
}

/// The full record for a built-in style by name.
pub fn builtin(name: &str) -> Option<&'static BuiltinStyle> {
    let wanted = name.trim().to_lowercase();
    BUILTIN_DETAILS
        .iter()
        .find(|style| style.name.to_lowercase() == wanted)
}

/// Every built-in style, in openpyxl's order.
pub fn builtin_styles() -> &'static [BuiltinStyle] {
    &BUILTIN_DETAILS
}

/// The name for a `builtinId`.
pub fn builtin_name(id: &str) -> Option<&'static str> {
    BUILTIN_DETAILS
        .iter()
        .find(|style| style.builtin_id == id)
        .map(|style| style.name)
}

/// A reusable style a cell can reference by name.
#[derive(Debug, Clone, PartialEq)]
pub struct NamedStyle {
    /// The style's name, as a cell writes it.
    pub name: String,
    /// The style the cells referencing it inherit from.
    ///
    /// A named style is a *delta*: `Good` is mostly the default font with a colour change,
    /// and the `xf` a cell points at carries only the difference. So a cell that uses a
    /// named style still needs its own `xf` with the differences applied.
    pub font: Font,
    /// The fill.
    pub fill: Fill,
    /// The border.
    pub border: Borders,
    /// The alignment.
    pub alignment: Alignment,
    /// The number format.
    pub number_format: NumberFormat,
    /// The protection settings.
    pub protection: Protection,
    /// The `builtinId`, for one of Excel's own styles.
    pub builtin_id: Option<String>,
    /// Whether the style is hidden from Excel's field list.
    pub hidden: Option<bool>,
}

impl Default for NamedStyle {
    fn default() -> Self {
        NamedStyle {
            name: String::new(),
            font: Font::new(),
            fill: Fill::new(),
            border: Borders::default(),
            alignment: Alignment::new(),
            number_format: NumberFormat::default(),
            protection: Protection::default(),
            builtin_id: None,
            hidden: None,
        }
    }
}

impl NamedStyle {
    /// A named style with a name and the default formatting.
    pub fn new(name: &str) -> Self {
        NamedStyle {
            name: name.to_string(),
            ..NamedStyle::default()
        }
    }

    /// One of Excel's built-in styles, by name.
    ///
    /// Returns `None` for a name Excel does not define.
    ///
    /// The formatting is filled in as well as the `builtinId`. Excel renders a built-in from
    /// the id alone, so an id-only style looks right in Excel and blank in anything that does
    /// not have Excel's built-in table -- which is most readers, including openpyxl. Carrying
    /// the colours costs a few hundred bytes and makes `cell.style = "Good"` mean the same
    /// thing everywhere.
    pub fn builtin(name: &str) -> Option<Self> {
        let record = builtin(name)?;
        let mut font = Font::new();
        if let Some(font_name) = record.font_name {
            font.name = font_name.to_string();
        }
        font.size = record.size;
        font.bold = record.bold;
        font.italic = record.italic;
        if let Some(colour) = record.font_color {
            font.color = crate::styles::colors::Color::new(colour);
        }
        if let Some(scheme) = record.scheme {
            font.scheme = scheme.to_string();
        }

        let mut fill = Fill::new();
        if let Some(colour) = record.fill_color {
            fill.fill_type = Some(Fill::FILL_SOLID.to_string());
            fill.start_color = crate::styles::colors::Color::new(colour);
        }

        let mut number_format = NumberFormat::default();
        if let Some(code) = record.number_format {
            number_format.set_format_code(code);
        }

        Some(NamedStyle {
            name: record.name.to_string(),
            font,
            fill,
            number_format,
            builtin_id: Some(record.builtin_id.to_string()),
            ..NamedStyle::default()
        })
    }

    /// A named style from the style a cell currently carries.
    pub fn from_style(name: &str, style: &Style) -> Self {
        NamedStyle {
            name: name.to_string(),
            font: style.font.clone(),
            fill: style.fill.clone(),
            border: style.borders.clone(),
            alignment: style.alignment.clone(),
            number_format: style.number_format.clone(),
            protection: style.protection,
            builtin_id: None,
            hidden: None,
        }
    }

    /// The style this describes, as the cell-level shape.
    pub fn to_style(&self) -> Style {
        Style {
            font: self.font.clone(),
            fill: self.fill.clone(),
            borders: self.border.clone(),
            alignment: self.alignment.clone(),
            number_format: self.number_format.clone(),
            protection: self.protection,
            ..Style::default()
        }
    }

    /// The number format code this style applies.
    pub fn number_format_code(&self) -> &str {
        self.number_format.format_code()
    }

    /// Set the font.
    pub fn with_font(mut self, font: Font) -> Self {
        self.font = font;
        self
    }

    /// Set the fill.
    pub fn with_fill(mut self, fill: Fill) -> Self {
        self.fill = fill;
        self
    }

    /// Set the border.
    pub fn with_border(mut self, border: Borders) -> Self {
        self.border = border;
        self
    }

    /// Set the alignment.
    pub fn with_alignment(mut self, alignment: Alignment) -> Self {
        self.alignment = alignment;
        self
    }

    /// Set the number format code.
    pub fn with_number_format(mut self, code: &str) -> Self {
        self.number_format.set_format_code(code);
        self
    }

    /// Hide the style from Excel's field list.
    pub fn hidden(mut self, hidden: bool) -> Self {
        self.hidden = Some(hidden);
        self
    }
}

impl fmt::Display for NamedStyle {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.name)
    }
}

/// The workbook's named styles, in the order they appear in `<cellStyles>`.
///
/// A `Vec` rather than a map, because **order is semantic**: a cell's `xfId` is an index into
/// this list, and Excel resolves `xfId="0"` to `Normal`. A sorted map would put `Good` before
/// `Normal` and every style reference in the workbook would mean the wrong thing. Insertion
/// order also makes a load-then-save round trip byte-stable, which a map would not.
///
/// `Normal` is kept first whatever order styles are added in, for the same reason.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct NamedStyleList {
    styles: Vec<NamedStyle>,
}

impl NamedStyleList {
    /// No named styles.
    pub fn new() -> Self {
        NamedStyleList::default()
    }

    /// Add or replace a style.
    ///
    /// Replaces rather than refusing, which is what editing a style in Excel does: refusing
    /// would make a load-then-edit-then-save round trip fail on a file that was valid.
    pub fn add(&mut self, style: NamedStyle) -> Result<()> {
        if style.name.trim().is_empty() {
            return Err(Error::Value("a named style needs a name".to_string()));
        }
        match self.index_of(&style.name) {
            Some(at) => self.styles[at] = style,
            None => {
                // `Normal` has to be the base every other style inherits from, so it is pinned
                // to index 0 however late it arrives.
                if style.name == NORMAL {
                    self.styles.insert(0, style);
                } else {
                    self.styles.push(style);
                }
            }
        }
        Ok(())
    }

    /// The style with this name.
    pub fn get(&self, name: &str) -> Option<&NamedStyle> {
        self.index_of(name).map(|at| &self.styles[at])
    }

    /// The style with this name, for mutation.
    pub fn get_mut(&mut self, name: &str) -> Option<&mut NamedStyle> {
        let at = self.index_of(name)?;
        Some(&mut self.styles[at])
    }

    /// Remove a style by name.
    ///
    /// Refuses to remove `Normal`, which is the base every other style inherits from and the
    /// target of every `xfId="0"`. Removing it would leave the workbook with a dangling
    /// reference, which Excel reports as a corrupt file.
    pub fn remove(&mut self, name: &str) -> Option<NamedStyle> {
        if name == NORMAL {
            return None;
        }
        let at = self.index_of(name)?;
        Some(self.styles.remove(at))
    }

    /// How many styles.
    pub fn len(&self) -> usize {
        self.styles.len()
    }

    /// Whether there are none.
    pub fn is_empty(&self) -> bool {
        self.styles.is_empty()
    }

    /// The names, in list order.
    pub fn names(&self) -> Vec<&str> {
        self.styles.iter().map(|s| s.name.as_str()).collect()
    }

    /// The styles, in list order.
    pub fn iter(&self) -> impl Iterator<Item = &NamedStyle> {
        self.styles.iter()
    }

    /// The index of a style, which is the `xfId` a cell writes.
    pub fn index_of(&self, name: &str) -> Option<usize> {
        self.styles.iter().position(|style| style.name == name)
    }

    /// The style at an `xfId`.
    pub fn by_index(&self, index: usize) -> Option<&NamedStyle> {
        self.styles.get(index)
    }

    /// Whether a name is one of Excel's built-ins, whether or not it is defined here.
    ///
    /// A built-in needs no definition: Excel renders it from its `builtinId`. So this is what
    /// decides whether applying a name is possible, rather than a lookup in `styles`.
    pub fn is_known(name: &str) -> bool {
        builtin_style(name).is_some()
    }
}

/// The base style every other named style inherits from, and the target of `xfId="0"`.
pub const NORMAL: &str = "Normal";

impl fmt::Display for NamedStyleList {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.names().join(", "))
    }
}

/// A colour lookup failure that names the style, for a clearer message than a bare miss.
pub fn unknown_style(name: &str, available: &[&str]) -> Error {
    let hint = if available.is_empty() {
        "the workbook defines no named styles".to_string()
    } else {
        format!("known: {}", available.join(", "))
    };
    Error::Key(format!("{name} is not a named style; {hint}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_built_in_table_matches_openpyxl() {
        // 49, with `Normal` at 0 and `Good` at 26. Both are load-bearing: 0 is the base every
        // other style inherits from, and 26 is the one a caller is most likely to reach for.
        assert_eq!(BUILTIN_STYLES.len(), 49);
        assert_eq!(BUILTIN_STYLES[0], ("0", "Normal"));
        assert_eq!(builtin_style("Good"), Some("26"));
        assert_eq!(builtin_style("Bad"), Some("27"));
        assert_eq!(builtin_style("Title"), Some("15"));
        assert_eq!(builtin_name("26"), Some("Good"));
    }

    #[test]
    fn built_in_lookup_ignores_case_and_padding() {
        // Excel's own field list matches `good` against `Good`, so a caller writing lowercase
        // means the same thing.
        assert_eq!(builtin_style("good"), Some("26"));
        assert_eq!(builtin_style("  Good  "), Some("26"));
        assert_eq!(builtin_style("not a style"), None);
    }

    #[test]
    fn a_built_in_style_carries_its_formatting_as_well_as_its_id() {
        // Excel renders a built-in from the id alone, so an id-only style looks right in Excel
        // and blank in everything else -- openpyxl included. Carrying the colours costs a few
        // hundred bytes and makes `cell.style = "Good"` mean the same thing everywhere.
        let good = NamedStyle::builtin("Good").expect("Good is built in");
        assert_eq!(good.builtin_id.as_deref(), Some("26"));
        assert_eq!(good.name, "Good");
        assert_eq!(good.fill.start_color.index, "FFC6EFCE");
        assert_eq!(good.font.color.index, "FF006100");
        assert_eq!(good.font.scheme, "minor");

        // A style with no fill of its own still reports the default rather than a guess.
        let normal = NamedStyle::builtin("Normal").expect("Normal is built in");
        assert!(normal.fill.start_color == Fill::new().start_color);
        assert!(NamedStyle::builtin("Nonsense").is_none());
    }

    #[test]
    fn the_generated_table_agrees_with_the_49_the_tool_counts() {
        assert_eq!(BUILTIN_DETAILS.len(), 49);
        // Two facts worth pinning, because they are the ones a transcription would get wrong.
        assert_eq!(BUILTIN_DETAILS[0].name, "Normal");
        assert_eq!(BUILTIN_DETAILS[0].builtin_id, "0");
        let title = builtin("Title").expect("Title is built in");
        assert_eq!(title.font_name, Some("Cambria"));
        assert!(title.bold);
        // The Accent family is border-only: no fill, and the font colour flips between
        // theme 0 and theme 1 by which shade it is.
        let accent = builtin("Accent1").expect("Accent1 is built in");
        assert!(accent.fill_color.is_none());
        assert_eq!(accent.font_color, Some("theme:0"));
    }

    #[test]
    fn a_number_format_survives_the_two_levels_of_unescaping() {
        // The source holds Python-escaped backslashes; the value has single ones. Skipping
        // either step yields a format that still parses and aligns wrongly.
        let comma = builtin("Comma").expect("Comma is built in");
        let code = comma.number_format.expect("Comma carries a format");
        assert!(
            code.contains(r"\ "),
            "single backslash before the space: {code}"
        );
        // Two backslashes in a row would mean the unescaping ran twice.
        assert!(!code.contains("\\\\"), "no doubled backslashes: {code}");
    }

    #[test]
    fn a_named_style_round_trips_through_a_cell_style() {
        let style = NamedStyle::new("Band")
            .with_number_format("0.00%")
            .with_font(Font::new().with_bold(true));
        let as_style = style.to_style();
        assert_eq!(as_style.number_format_code(), "0.00%");
        assert!(as_style.font.bold);
        assert_eq!(
            NamedStyle::from_style("Band", &as_style).number_format_code(),
            "0.00%"
        );
    }

    #[test]
    fn the_list_replaces_rather_than_refuses() {
        // Editing a style in Excel replaces it, so refusing would make a load-edit-save round
        // trip fail on a valid file.
        let mut list = NamedStyleList::new();
        list.add(NamedStyle::new("Total")).expect("added");
        list.add(NamedStyle::new("Total").with_number_format("0.00"))
            .expect("replaced");
        assert_eq!(list.len(), 1);
        assert_eq!(
            list.get("Total")
                .expect("Total")
                .number_format
                .format_code(),
            "0.00"
        );
    }

    #[test]
    fn the_index_is_the_xf_id_a_cell_writes() {
        let mut list = NamedStyleList::new();
        list.add(NamedStyle::new("Normal")).expect("added");
        list.add(NamedStyle::new("Good")).expect("added");
        assert_eq!(list.index_of("Normal"), Some(0));
        assert_eq!(list.index_of("Good"), Some(1));
        assert_eq!(list.index_of("Missing"), None);
    }

    #[test]
    fn normal_stays_first_however_late_it_arrives() {
        // `xfId="0"` resolves to `Normal`, so its position is not a matter of taste: a sorted
        // map would put "Good" first and every style reference would mean the wrong thing.
        let mut list = NamedStyleList::new();
        list.add(NamedStyle::new("Good")).expect("added");
        list.add(NamedStyle::new("Accent1")).expect("added");
        list.add(NamedStyle::new("Normal")).expect("added");
        assert_eq!(list.index_of("Normal"), Some(0));
        assert_eq!(list.names(), ["Normal", "Good", "Accent1"]);
    }

    #[test]
    fn normal_cannot_be_removed() {
        // Removing the base would leave every `xfId="0"` dangling, which Excel reports as a
        // corrupt file rather than as a missing style.
        let mut list = NamedStyleList::new();
        list.add(NamedStyle::new("Normal")).expect("added");
        list.add(NamedStyle::new("Good")).expect("added");
        assert!(list.remove("Normal").is_none());
        assert!(list.remove("Good").is_some());
        assert_eq!(list.len(), 1);
    }

    #[test]
    fn a_style_needs_a_name() {
        let mut list = NamedStyleList::new();
        assert!(list.add(NamedStyle::new("  ")).is_err());
    }

    #[test]
    fn an_unknown_style_says_what_is_known() {
        let err = unknown_style("Nonsense", &["Normal", "Good"]);
        let message = err.to_string();
        assert!(message.contains("Normal"), "{message}");
        assert!(message.contains("Good"), "{message}");
    }

    #[cfg(test)]
    mod round_trip_tests {
        use super::*;
        use crate::cell::cell::CellValue;
        use crate::workbook::Workbook;

        #[test]
        fn a_named_style_survives_a_package_round_trip() {
            let mut workbook = Workbook::new();
            let mut list = NamedStyleList::new();
            list.add(NamedStyle::new("Normal")).expect("normal");
            list.add(
                NamedStyle::new("Band")
                    .with_number_format("0.00%")
                    .with_font(Font::new().with_bold(true)),
            )
            .expect("band");
            workbook.named_styles = list;

            let bytes = workbook.to_bytes().expect("saved");
            let loaded = crate::reader::excel::load_workbook_from_bytes(bytes, Default::default())
                .expect("loaded");

            assert_eq!(loaded.named_styles.len(), 2);
            let band = loaded.named_styles.get("Band").expect("Band survived");
            assert_eq!(band.number_format_code(), "0.00%");
            assert!(band.font.bold);
            // `xfId="0"` has to still be Normal, or every style reference means the wrong thing.
            assert_eq!(loaded.named_styles.index_of("Normal"), Some(0));
            assert_eq!(loaded.named_styles.index_of("Band"), Some(1));
        }

        #[test]
        fn a_built_in_style_can_be_applied_and_reads_back() {
            let mut workbook = Workbook::new();
            let sheet = workbook.active_sheet_mut().expect("sheet");
            sheet.set("A1", CellValue::number(1.0)).expect("cell");
            sheet
                .apply_named_style("A1", "Good")
                .expect("Good is built in");

            let bytes = workbook.to_bytes().expect("saved");
            let loaded = crate::reader::excel::load_workbook_from_bytes(bytes, Default::default())
                .expect("loaded");
            // The cell's formatting survives, which is the part that matters: `xfId` is a
            // reference a reader may ignore, so the formatting has to be on the cell itself.
            assert!(loaded.worksheets[0].get_style("A1").fill != Fill::new());
        }

        #[test]
        fn an_unknown_name_is_refused_rather_than_defaulted() {
            let mut workbook = Workbook::new();
            let sheet = workbook.active_sheet_mut().expect("sheet");
            sheet.set("A1", CellValue::number(1.0)).expect("cell");
            // Applying a default and calling it "Good" would be worse than an error: the file
            // would open and show nothing that asked to be styled.
            assert!(sheet.apply_named_style("A1", "Nonsense").is_err());
            assert_eq!(sheet.get_style("A1"), Style::default());
        }
    }
}
