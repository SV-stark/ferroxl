//! Formula text handling.
//!
//! Three things live here, and they are different jobs. [`Translator`] rewrites a formula so it
//! means the same thing from a different cell, which is what copy-paste and fill do.
//! [`evaluate`] computes a formula's value so a saved workbook carries one rather than a blank.
//! Formula *sharing* - one `<f t="shared">` element serving many cells - is in
//! [`crate::cell::formula`] instead, because it is a property of the stored XML rather than of
//! the expression.

pub mod eval;
mod translator;

pub use eval::{evaluate, supports, Recalculation, Unresolved, ValueSource};
pub use translator::{Translator, TranslatorError};
