//! Formula text handling.
//!
//! Two things live here, and they are different jobs. [`Translator`] rewrites a formula so
//! it means the same thing from a different cell, which is what copy-paste and fill do.
//! Formula *sharing* - one `<f t="shared">` element serving many cells - is in
//! [`crate::cell::formula`] instead, because it is a property of the stored XML rather than
//! of the expression.

mod translator;

pub use translator::{Translator, TranslatorError};
