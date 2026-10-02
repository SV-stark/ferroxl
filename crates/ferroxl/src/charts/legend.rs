//! Chart legend (`openpyxl/charts/legend.rs`).

/// The legend of a chart.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Legend {
    /// Legend position: `r`, `l`, `t`, `b` or `tr`.
    pub position: String,
    /// Manual layout, currently unused by the writer.
    pub layout: Option<String>,
}

impl Default for Legend {
    fn default() -> Self {
        Legend {
            position: "r".to_string(),
            layout: None,
        }
    }
}

impl Legend {
    /// A right-hand legend.
    pub fn new() -> Self {
        Legend::default()
    }

    /// Move the legend.
    pub fn with_position(mut self, position: &str) -> Self {
        self.position = position.to_string();
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_position_is_right() {
        assert_eq!(Legend::new().position, "r");
        assert!(Legend::new().layout.is_none());
    }

    #[test]
    fn position_is_configurable() {
        assert_eq!(Legend::new().with_position("b").position, "b");
    }
}
