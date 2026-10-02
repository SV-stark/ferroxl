//! Chart error bars (`openpyxl/charts/error_bar.py`).

use super::reference::Reference;

/// Which side of the data point the error bar extends to.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorBarType {
    /// Positive side only.
    Plus,
    /// Negative side only.
    Minus,
    /// Both sides.
    PlusMinus,
}

impl ErrorBarType {
    /// The XML `errBarType` token.
    pub fn as_str(self) -> &'static str {
        match self {
            ErrorBarType::Plus => "plus",
            ErrorBarType::Minus => "minus",
            ErrorBarType::PlusMinus => "both",
        }
    }
}

/// Error bars attached to a series.
#[derive(Debug, Clone, PartialEq)]
pub struct ErrorBar {
    /// Which side the bars extend to.
    pub bar_type: ErrorBarType,
    /// The reference holding the error magnitudes.
    pub reference: Reference,
}

impl ErrorBar {
    /// Build error bars over the given reference.
    pub fn new(bar_type: ErrorBarType, reference: Reference) -> Self {
        ErrorBar {
            bar_type,
            reference,
        }
    }

    /// The values held by the reference, if they have been resolved.
    pub fn values(&self) -> Option<&[super::reference::CellValue]> {
        self.reference.values()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tokens_match_python_flags() {
        assert_eq!(ErrorBarType::Plus.as_str(), "plus");
        assert_eq!(ErrorBarType::Minus.as_str(), "minus");
        assert_eq!(ErrorBarType::PlusMinus.as_str(), "both");
    }

    #[test]
    fn carries_a_reference() {
        let reference = Reference::new("S", (0, 0), Some((2, 0)), None, None).unwrap();
        let bar = ErrorBar::new(ErrorBarType::PlusMinus, reference);
        assert_eq!(bar.bar_type, ErrorBarType::PlusMinus);
        assert!(bar.values().is_none());
    }
}
