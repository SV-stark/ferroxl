//! Style objects (`openpyxl/styles/`).
//!
//! Each class derives from `HashableObject` in the original, whose equality and hash are
//! based on the tuple of its `__fields__`. In Rust that maps to `PartialEq`/`Eq`/`Hash`
//! over the same fields, which is what the style writer relies on to deduplicate
//! fonts, fills, borders and number formats.

pub mod alignment;
pub mod borders;
pub mod colors;
pub mod fills;
pub mod fonts;
pub mod numbers;
pub mod protection;
pub mod style;

pub use alignment::Alignment;
pub use borders::{Border, Borders};
pub use colors::{Color, COLOR_INDEX};
pub use fills::Fill;
pub use fonts::Font;
pub use numbers::{is_builtin, is_date_format, NumberFormat};
pub use protection::Protection;
pub use style::{defaults, same_visual_style, Style};
