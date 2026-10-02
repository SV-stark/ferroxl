//! Page layout for a sheet (`openpyxl/worksheet/page.rs`).

/// Information about page layout for this sheet.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct PageSetup {
    /// `portrait` or `landscape`.
    pub orientation: Option<String>,
    /// The paper size code.
    ///
    /// [`PAPERSIZES`](crate::worksheet::Worksheet::PAPERSIZES) maps each of openpyxl's
    /// `PAPERSIZE_*` names to the digit OOXML stores here.
    pub paper_size: Option<String>,
    /// Print scaling percentage.
    pub scale: Option<String>,
    /// Whether "fit to page" is enabled.
    pub fit_to_page: Option<String>,
    /// Number of pages tall to fit.
    pub fit_to_height: Option<String>,
    /// Number of pages wide to fit.
    pub fit_to_width: Option<String>,
    /// The first page number.
    pub first_page_number: Option<String>,
    /// Whether `first_page_number` is honoured.
    pub use_first_page_number: Option<String>,
    /// Whether the content is horizontally centred.
    pub horizontal_centered: Option<String>,
    /// Whether the content is vertically centred.
    pub vertical_centered: Option<String>,
}

/// The attributes written to `<pageSetup>`, in schema order.
///
/// Values are only emitted when set, and the numeric fields are validated, so a bad
/// `fitToHeight` cannot produce invalid XML.
impl PageSetup {
    /// The `pageSetup` attribute order used by the writer.
    pub const VALID_SETUP: [&'static str; 8] = [
        "orientation",
        "paperSize",
        "scale",
        "fitToPage",
        "fitToHeight",
        "fitToWidth",
        "firstPageNumber",
        "useFirstPageNumber",
    ];

    /// The `printOptions` attribute order used by the writer.
    pub const VALID_OPTIONS: [&'static str; 2] = ["horizontalCentered", "verticalCentered"];

    /// Attributes for `<pageSetup>`.
    pub fn setup_attributes(&self) -> Vec<(String, String)> {
        let mut out = Vec::new();
        for name in PageSetup::VALID_SETUP {
            let value = match name {
                "orientation" => self.orientation.clone(),
                "paperSize" => self.paper_size.clone(),
                "scale" => self.scale.clone(),
                "fitToPage" => self.fit_to_page.clone(),
                "fitToHeight" => self.fit_to_height.clone(),
                "fitToWidth" => self.fit_to_width.clone(),
                "firstPageNumber" => self.first_page_number.clone(),
                "useFirstPageNumber" => self.use_first_page_number.clone(),
                _ => None,
            };
            let Some(value) = value else { continue };
            match name {
                "orientation" => out.push((name.to_string(), value)),
                "paperSize" | "scale" => {
                    if let Some(formatted) = format_as_int(&value) {
                        out.push((name.to_string(), formatted));
                    }
                }
                "fitToHeight" | "fitToWidth" => {
                    // Only non-negative values are written; Excel treats a negative as 0.
                    match value.trim().parse::<i64>() {
                        Ok(v) if v >= 0 => out.push((name.to_string(), v.to_string())),
                        _ => {}
                    }
                }
                _ => out.push((name.to_string(), value)),
            }
        }
        out
    }

    /// Attributes for `<printOptions>`, which only emits flags.
    pub fn option_attributes(&self) -> Vec<(String, String)> {
        let mut out = Vec::new();
        for name in PageSetup::VALID_OPTIONS {
            let value = match name {
                "horizontalCentered" => self.horizontal_centered.clone(),
                "verticalCentered" => self.vertical_centered.clone(),
                _ => None,
            };
            if value.is_some() {
                out.push((name.to_string(), "1".to_string()));
            }
        }
        out
    }

    /// Whether "fit to page" is requested, accepting the string forms Excel writes.
    pub fn fit_to_page_enabled(&self) -> bool {
        matches!(self.fit_to_page.as_deref(), Some("1") | Some("true"))
    }
}

/// Python's `'%d' % int(value)` with a fallback to the raw text.
fn format_as_int(value: &str) -> Option<String> {
    value
        .trim()
        .parse::<f64>()
        .ok()
        .map(|v| format!("{}", v.trunc() as i64))
        .or_else(|| Some(value.to_string()))
}

/// Information about page margins for view and print layouts.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct PageMargins {
    /// Left margin in inches.
    pub left: Option<f64>,
    /// Right margin in inches.
    pub right: Option<f64>,
    /// Top margin in inches.
    pub top: Option<f64>,
    /// Bottom margin in inches.
    pub bottom: Option<f64>,
    /// Header margin in inches.
    pub header: Option<f64>,
    /// Footer margin in inches.
    pub footer: Option<f64>,
}

impl PageMargins {
    /// The margin names, in schema order.
    pub const VALID_MARGINS: [&'static str; 6] =
        ["left", "right", "top", "bottom", "header", "footer"];

    /// Build empty margins.
    pub fn new() -> Self {
        PageMargins::default()
    }

    /// Attributes for `<pageMargins>`, formatted to two decimals.
    ///
    /// Zero values are skipped, matching the Python truthiness test.
    pub fn margin_attributes(&self) -> Vec<(String, String)> {
        let mut out = Vec::new();
        for name in PageMargins::VALID_MARGINS {
            let value = match name {
                "left" => self.left,
                "right" => self.right,
                "top" => self.top,
                "bottom" => self.bottom,
                "header" => self.header,
                "footer" => self.footer,
                _ => None,
            };
            // Python tests the raw value for truthiness, so 0.0 is omitted.
            if let Some(value) = value {
                if value != 0.0 {
                    out.push((name.to_string(), format!("{value:.2}")));
                }
            }
        }
        out
    }

    /// Apply the library's default margins.
    pub fn with_defaults(mut self) -> Self {
        self.left = Some(crate::units::DEFAULT_LEFT_MARGIN);
        self.right = Some(crate::units::DEFAULT_LEFT_MARGIN);
        self.top = Some(crate::units::DEFAULT_TOP_MARGIN);
        self.bottom = Some(crate::units::DEFAULT_TOP_MARGIN);
        self.header = Some(crate::units::DEFAULT_HEADER);
        self.footer = Some(crate::units::DEFAULT_HEADER);
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_setup_writes_nothing() {
        let setup = PageSetup::default();
        assert!(setup.setup_attributes().is_empty());
        assert!(setup.option_attributes().is_empty());
    }

    #[test]
    fn setup_attributes_are_ordered_and_numeric_fields_normalised() {
        let setup = PageSetup {
            orientation: Some("landscape".into()),
            paper_size: Some("9.0".into()),
            scale: Some("150.7".into()),
            fit_to_page: Some("1".into()),
            fit_to_height: Some("2".into()),
            fit_to_width: Some("-1".into()),
            first_page_number: Some("3".into()),
            use_first_page_number: Some("1".into()),
            horizontal_centered: Some("1".into()),
            vertical_centered: None,
        };
        let attrs = setup.setup_attributes();
        let names: Vec<&str> = attrs.iter().map(|(k, _)| k.as_str()).collect();
        assert_eq!(
            names,
            vec![
                "orientation",
                "paperSize",
                "scale",
                "fitToPage",
                "fitToHeight",
                "firstPageNumber",
                "useFirstPageNumber"
            ]
        );
        assert_eq!(attrs[1].1, "9");
        assert_eq!(attrs[2].1, "150");
        // fitToWidth is negative and therefore dropped.
        assert!(!names.contains(&"fitToWidth"));

        let options = setup.option_attributes();
        assert_eq!(
            options,
            vec![("horizontalCentered".to_string(), "1".to_string())]
        );
    }

    #[test]
    fn fit_to_page_flag() {
        let setup = PageSetup {
            fit_to_page: Some("1".into()),
            ..PageSetup::default()
        };
        assert!(setup.fit_to_page_enabled());
        let setup = PageSetup {
            fit_to_page: Some("true".into()),
            ..PageSetup::default()
        };
        assert!(setup.fit_to_page_enabled());
        assert!(!PageSetup::default().fit_to_page_enabled());
    }

    #[test]
    fn margins_are_two_decimals_and_skip_zero() {
        let margins = PageMargins {
            left: Some(0.7),
            right: Some(0.0),
            top: Some(0.7874),
            bottom: None,
            header: Some(0.3),
            footer: Some(0.5),
        };
        let attrs = margins.margin_attributes();
        assert_eq!(
            attrs,
            vec![
                ("left".to_string(), "0.70".to_string()),
                ("top".to_string(), "0.79".to_string()),
                ("header".to_string(), "0.30".to_string()),
                ("footer".to_string(), "0.50".to_string()),
            ]
        );
    }

    #[test]
    fn default_margins_helper() {
        let attrs = PageMargins::new().with_defaults().margin_attributes();
        assert_eq!(attrs.len(), 6);
    }
}
