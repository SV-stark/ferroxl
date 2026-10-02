//! Error types mirroring `openpyxl.exceptions`.
//!
//! The Python library raises a distinct exception class per failure mode. Rust has no
//! exception hierarchy for library errors, so every class becomes a variant of
//! [`Error`] and the original class name is preserved for diagnostics.

use std::fmt;

/// Errors raised by the library.
///
/// Each variant corresponds 1:1 to a class in `openpyxl/exceptions.py`.
#[derive(Debug, thiserror::Error)]
pub enum Error {
    /// Error for converting between numeric and A1-style cell references.
    #[error("Invalid cell coordinates ({0})")]
    CellCoordinates(String),

    /// The data submitted cannot be used directly in Excel files.
    #[error("IllegalCharacterError: data contains characters that cannot be stored in Excel: {0}")]
    IllegalCharacter(char),

    /// Error for bad column names in A1-style cell references.
    #[error("ColumnStringIndexException: {0}")]
    ColumnStringIndex(String),

    /// Error for any data type inconsistencies.
    #[error("DataTypeException: {0}")]
    DataType(String),

    /// Error for badly formatted named ranges.
    #[error("NamedRangeException: {0}")]
    NamedRange(String),

    /// Error for bad sheet names.
    #[error("SheetTitleException: {0}")]
    SheetTitle(String),

    /// Error for partially specified cell coordinates.
    #[error("InsufficientCoordinatesException: {0}")]
    InsufficientCoordinates(String),

    /// Error for fileobj opened in non-binary mode.
    #[error("OpenModeError: {0}")]
    OpenMode(String),

    /// Error for trying to open a non-ooxml file.
    #[error("InvalidFileException: {0}")]
    InvalidFile(String),

    /// Error for trying to modify a read-only workbook.
    #[error("ReadOnlyWorkbookException: {0}")]
    ReadOnlyWorkbook(String),

    /// Error when a referenced number format is not in the stylesheet.
    #[error("MissingNumberFormat: {0}")]
    MissingNumberFormat(String),

    /// Error when a dump workbook is used after it has been dumped once.
    #[error("WorkbookAlreadySaved: {0}")]
    WorkbookAlreadySaved(String),

    /// Generic `ValueError`.
    #[error("ValueError: {0}")]
    Value(String),

    /// Generic `TypeError`.
    #[error("TypeError: {0}")]
    Type(String),

    /// Generic `KeyError`.
    #[error("KeyError: {0}")]
    Key(String),

    /// Generic `AttributeError`.
    #[error("AttributeError: {0}")]
    Attribute(String),

    /// Underlying zip archive failure.
    #[error("BadZipFile: {0}")]
    BadZipFile(String),

    /// XML parse failure.
    #[error("XMLParseError: {0}")]
    Xml(String),

    /// Filesystem / IO failure.
    #[error("IOError: {0}")]
    Io(String),

    /// Anything raised by `NotImplementedError`.
    #[error("NotImplementedError: {0}")]
    NotImplemented(String),
}

impl Error {
    /// Convenience constructor for [`Error::Value`].
    pub fn value(msg: impl Into<String>) -> Self {
        Error::Value(msg.into())
    }

    /// Convenience constructor for [`Error::Type`].
    pub fn type_error(msg: impl Into<String>) -> Self {
        Error::Type(msg.into())
    }

    /// Convenience constructor for [`Error::Key`].
    pub fn key(msg: impl Into<String>) -> Self {
        Error::Key(msg.into())
    }

    /// Convenience constructor for [`Error::Attribute`].
    pub fn attribute(msg: impl Into<String>) -> Self {
        Error::Attribute(msg.into())
    }
}

impl From<std::io::Error> for Error {
    fn from(value: std::io::Error) -> Self {
        Error::Io(value.to_string())
    }
}

impl From<quick_xml::Error> for Error {
    fn from(value: quick_xml::Error) -> Self {
        Error::Xml(value.to_string())
    }
}

impl From<quick_xml::events::attributes::AttrError> for Error {
    fn from(value: quick_xml::events::attributes::AttrError) -> Self {
        Error::Xml(value.to_string())
    }
}

/// Result alias used throughout the crate.
pub type Result<T, E = Error> = std::result::Result<T, E>;

/// Escape hatch used by `openpyxl`'s bare `except:` clauses, where any failure is
/// swallowed and replaced by a fallback value.
pub trait OrDefault<T: Default> {
    /// Return `T::default()` when `self` is an error.
    fn or_default(self) -> T;
}

impl<T: Default> OrDefault<T> for Result<T> {
    fn or_default(self) -> T {
        self.unwrap_or_default()
    }
}

/// Formats a value the way Python's `repr()`/`str()` would for the handful of types
/// that reach XML output (`openpyxl.compat.strings.safe_string`).
pub fn safe_string<T: fmt::Display>(value: T) -> String {
    value.to_string()
}
