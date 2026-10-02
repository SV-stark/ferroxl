//! Protection options (`openpyxl/styles/protection.py`).

/// The tri-state used for `locked` and `hidden`.
///
/// openpyxl mixes booleans and the string `"inherit"` in these fields, which Rust cannot
/// do; the enum captures all three states faithfully.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ProtectionFlag {
    /// Take the value from the cell style / default (`"inherit"`).
    #[default]
    Inherit,
    /// Explicitly locked.
    Protected,
    /// Explicitly unlocked.
    Unprotected,
}

impl ProtectionFlag {
    /// The XML attribute value, when the flag needs to be written.
    pub fn as_attr(self) -> Option<&'static str> {
        match self {
            ProtectionFlag::Inherit => None,
            ProtectionFlag::Protected => Some("1"),
            ProtectionFlag::Unprotected => Some("0"),
        }
    }

    /// Interpret a Python-style value: `"inherit"`, or a truthy/falsy flag.
    pub fn from_py(value: &str) -> Self {
        match value {
            "inherit" => ProtectionFlag::Inherit,
            v => match v.parse::<i64>() {
                Ok(0) => ProtectionFlag::Unprotected,
                _ => ProtectionFlag::Protected,
            },
        }
    }

    /// Interpret an XML attribute, honouring ElementTree's truthiness rules.
    ///
    /// openpyxl calls `bool(...)` on the attribute, so any non-empty, non-`"0"` string is
    /// true and `None` is false.
    pub fn from_xml(value: Option<&str>) -> Self {
        match value {
            None => ProtectionFlag::Unprotected,
            Some("") => ProtectionFlag::Protected,
            Some("0") | Some("false") | Some("False") => ProtectionFlag::Unprotected,
            Some(_) => ProtectionFlag::Protected,
        }
    }
}

/// Protection options for use in styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Protection {
    /// Whether the cell is locked.
    pub locked: ProtectionFlag,
    /// Whether the cell's formula is hidden.
    pub hidden: ProtectionFlag,
}

impl Protection {
    /// The `"inherit"` marker used by the Python defaults.
    pub const PROTECTION_INHERIT: &'static str = "inherit";

    /// Build the default (inheriting) protection.
    pub fn new() -> Self {
        Protection::default()
    }

    /// Whether anything differs from the default protection.
    pub fn is_default(&self) -> bool {
        self.locked == ProtectionFlag::Inherit && self.hidden == ProtectionFlag::Inherit
    }

    /// Attributes for the `<protection/>` element.
    pub fn attributes(&self) -> Vec<(String, String)> {
        let mut attrs = Vec::new();
        if let Some(v) = self.locked.as_attr() {
            attrs.push(("locked".to_string(), v.to_string()));
        }
        if let Some(v) = self.hidden.as_attr() {
            attrs.push(("hidden".to_string(), v.to_string()));
        }
        attrs
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_inherit() {
        let p = Protection::new();
        assert_eq!(p.locked, ProtectionFlag::Inherit);
        assert_eq!(p.hidden, ProtectionFlag::Inherit);
        assert!(p.is_default());
        assert!(p.attributes().is_empty());
    }

    #[test]
    fn explicit_values_serialise() {
        let p = Protection {
            locked: ProtectionFlag::Protected,
            hidden: ProtectionFlag::Unprotected,
        };
        let attrs = p.attributes();
        assert!(attrs.contains(&("locked".to_string(), "1".to_string())));
        assert!(attrs.contains(&("hidden".to_string(), "0".to_string())));
        assert!(!p.is_default());
    }

    #[test]
    fn xml_truthiness_matches_bool() {
        assert_eq!(ProtectionFlag::from_xml(None), ProtectionFlag::Unprotected);
        assert_eq!(
            ProtectionFlag::from_xml(Some("")),
            ProtectionFlag::Protected
        );
        assert_eq!(
            ProtectionFlag::from_xml(Some("0")),
            ProtectionFlag::Unprotected
        );
        assert_eq!(
            ProtectionFlag::from_xml(Some("1")),
            ProtectionFlag::Protected
        );
        assert_eq!(ProtectionFlag::from_py("inherit"), ProtectionFlag::Inherit);
    }
}
