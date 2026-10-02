//! Named colours (`openpyxl/styles/colors.py`).

/// The legacy indexed colour palette used by Excel 2007-era files.
pub const COLOR_INDEX: [&str; 56] = [
    "FF000000", "FFFFFFFF", "FFFF0000", "FF00FF00", "FF0000FF", "FFFFFF00", "FFFF00FF", "FF00FFFF",
    "FF800000", "FF008000", "FF000080", "FF808000", "FF800080", "FF008080", "FFC0C0C0", "FF808080",
    "FF9999FF", "FF993366", "FFFFFFCC", "FFCCFFFF", "FF660066", "FFFF8080", "FF0066CC", "FFCCCCFF",
    "FF000080", "FFFF00FF", "FFFFFF00", "FF00FFFF", "FF800080", "FF800000", "FF008080", "FF0000FF",
    "FF00CCFF", "FFCCFFFF", "FFCCFFCC", "FFFFFF99", "FF99CCFF", "FFFF99CC", "FFCC99FF", "FFFFCC99",
    "FF3366FF", "FF33CCCC", "FF99CC00", "FFFFCC00", "FFFF9900", "FFFF6600", "FF666699", "FF969696",
    "FF003366", "FF339966", "FF003300", "FF333300", "FF993300", "FF993366", "FF333399", "FF333333",
];

/// A colour, stored as a string that may be an RGB value, an indexed value, or a
/// `theme:<n>[:<tint>]` marker.
///
/// The type is intentionally a newtype over `String` rather than a structured enum:
/// openpyxl round-trips unknown colour encodings verbatim, and rejecting or normalising
/// them would lose data on load/save cycles.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord, Default)]
pub struct Color {
    /// The raw colour token.
    pub index: String,
}

impl Color {
    /// Pure black.
    pub const BLACK: &'static str = "FF000000";
    /// Pure white.
    pub const WHITE: &'static str = "FFFFFFFF";
    /// Pure red.
    pub const RED: &'static str = "FFFF0000";
    /// Dark red.
    pub const DARKRED: &'static str = "FF800000";
    /// Pure blue.
    pub const BLUE: &'static str = "FF0000FF";
    /// Dark blue.
    pub const DARKBLUE: &'static str = "FF000080";
    /// Pure green.
    pub const GREEN: &'static str = "FF00FF00";
    /// Dark green.
    pub const DARKGREEN: &'static str = "FF008000";
    /// Pure yellow.
    pub const YELLOW: &'static str = "FFFFFF00";
    /// Dark yellow.
    pub const DARKYELLOW: &'static str = "FF808000";

    /// Wrap a colour token.
    pub fn new(index: impl Into<String>) -> Self {
        Color {
            index: index.into(),
        }
    }

    /// Build the `theme:<n>:<tint>` marker used by the style reader.
    pub fn theme(theme: &str, tint: Option<&str>) -> Self {
        match tint {
            Some(t) => Color::new(format!("theme:{theme}:{t}")),
            None => Color::new(format!("theme:{theme}:")),
        }
    }

    /// Whether the token references a theme colour.
    pub fn is_theme(&self) -> bool {
        self.index.starts_with("theme:")
    }

    /// The theme index, when [`Color::is_theme`] holds.
    pub fn theme_index(&self) -> Option<&str> {
        self.index
            .strip_prefix("theme:")
            .map(|rest| rest.split(':').next().unwrap_or(rest))
    }

    /// The tint component, when present.
    pub fn theme_tint(&self) -> Option<&str> {
        if !self.is_theme() {
            return None;
        }
        let mut parts = self.index.splitn(3, ':');
        parts.next();
        parts.next();
        match parts.next() {
            Some(tint) if !tint.is_empty() => Some(tint),
            _ => None,
        }
    }

    /// Look up a colour from the legacy indexed palette.
    pub fn from_index(index: usize) -> Option<Self> {
        COLOR_INDEX.get(index).map(|c| Color::new(*c))
    }
}

impl std::fmt::Display for Color {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.index)
    }
}

impl From<&str> for Color {
    fn from(value: &str) -> Self {
        Color::new(value)
    }
}

impl From<String> for Color {
    fn from(value: String) -> Self {
        Color::new(value)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn palette_lookup() {
        assert_eq!(Color::from_index(0).unwrap().index, "FF000000");
        assert_eq!(Color::from_index(55).unwrap().index, "FF333333");
        // openpyxl's palette has 56 entries, not Excel's documented 64, so an index past
        // the end of the table is rejected.
        assert!(Color::from_index(56).is_none());
    }

    #[test]
    fn theme_markers_round_trip() {
        let plain = Color::theme("9", None);
        assert_eq!(plain.index, "theme:9:");
        assert_eq!(plain.theme_index(), Some("9"));
        assert_eq!(plain.theme_tint(), None);
        assert!(plain.is_theme());

        let tinted = Color::theme("4", Some("0.5"));
        assert_eq!(tinted.theme_index(), Some("4"));
        assert_eq!(tinted.theme_tint(), Some("0.5"));
    }

    #[test]
    fn equality_is_by_index() {
        assert_eq!(Color::new("FF000000"), Color::new("FF000000".to_string()));
        assert_ne!(Color::new("FF000000"), Color::new("FFFFFFFF"));
    }
}
