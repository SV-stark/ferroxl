//! XML helpers: namespace constants, DOM building/parsing and streaming writers.

pub mod constants;
pub mod functions;

pub use constants::*;
pub use functions::{
    conditional_element, escape_attribute, escape_text, fromstring, repr_float, safe_string,
    serialize, Element, QName, XmlWriter,
};
