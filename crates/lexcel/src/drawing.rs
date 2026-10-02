//! Drawing objects (`openpyxl/drawing/drawing.py`).
//!
//! A [`Drawing`] is an anchor plus a size, in pixels, which the writer converts to EMUs. A
//! [`Shape`] lives inside a chart and is positioned in axis units. An [`Image`] wraps image
//! bytes plus its anchor.

use crate::cell::utils::{column_index_from_string, get_column_letter};
use crate::exceptions::{Error, Result};
use crate::styles::colors::Color;
use crate::units::{emu_to_pixels, pixels_to_emu, short_color};

/// A drop shadow preset (`openpyxl.drawing.Shadow`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Shadow {
    /// Whether the shadow is drawn.
    pub visible: bool,
    /// Blur radius.
    pub blur_radius: i64,
    /// Distance from the shape.
    pub distance: i64,
    /// Direction in degrees.
    pub direction: i64,
    /// Alignment preset; see the `SHADOW_*` constants.
    pub alignment: String,
    /// Shadow colour.
    pub color: Color,
    /// Alpha percentage.
    pub alpha: i64,
}

impl Default for Shadow {
    fn default() -> Self {
        Shadow {
            visible: false,
            blur_radius: 6,
            distance: 2,
            direction: 0,
            alignment: Shadow::SHADOW_BOTTOM_RIGHT.to_string(),
            color: Color::new(Color::BLACK),
            alpha: 50,
        }
    }
}

impl Shadow {
    /// Below the shape.
    pub const SHADOW_BOTTOM: &'static str = "b";
    /// Below and to the left.
    pub const SHADOW_BOTTOM_LEFT: &'static str = "bl";
    /// Below and to the right.
    pub const SHADOW_BOTTOM_RIGHT: &'static str = "br";
    /// Centred.
    pub const SHADOW_CENTER: &'static str = "ctr";
    /// To the left.
    pub const SHADOW_LEFT: &'static str = "l";
    /// Above the shape.
    pub const SHADOW_TOP: &'static str = "t";
    /// Above and to the left.
    pub const SHADOW_TOP_LEFT: &'static str = "tl";
    /// Above and to the right.
    pub const SHADOW_TOP_RIGHT: &'static str = "tr";

    /// An invisible shadow with openpyxl's defaults.
    pub fn new() -> Self {
        Shadow::default()
    }
}

/// How a drawing is anchored to the grid.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AnchorType {
    /// Positioned in absolute EMUs.
    #[default]
    Absolute,
    /// Positioned relative to a single cell.
    OneCell,
}

/// A drawing object: a container for shapes or charts.
///
/// Dimensions are given in pixels by the caller and converted to EMUs when written.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Drawing {
    /// Name.
    pub name: String,
    /// Description.
    pub description: String,
    /// Anchor corner coordinates in axis units.
    pub coordinates: ((i64, i64), (i64, i64)),
    /// Left offset in pixels.
    pub left: i64,
    /// Top offset in pixels.
    pub top: i64,
    width: i64,
    height: i64,
    /// Whether setting one dimension scales the other.
    pub resize_proportional: bool,
    /// Rotation in degrees.
    pub rotation: i64,
    /// How the drawing is anchored.
    pub anchor_type: AnchorType,
    /// Zero-based anchor column, for `OneCell` anchoring.
    pub anchor_col: Option<i64>,
    /// Zero-based anchor row, for `OneCell` anchoring.
    pub anchor_row: Option<i64>,
}

impl Default for Drawing {
    fn default() -> Self {
        Drawing {
            name: String::new(),
            description: String::new(),
            coordinates: ((1, 2), (16, 8)),
            left: 0,
            top: 0,
            // The Python defaults come from EMU_to_pixels(200000) and (1828800).
            width: emu_to_pixels(200_000),
            height: emu_to_pixels(1_828_800),
            resize_proportional: false,
            rotation: 0,
            anchor_type: AnchorType::Absolute,
            anchor_col: None,
            anchor_row: None,
        }
    }
}

impl Drawing {
    /// The default size (200000 × 1828800 EMU).
    pub fn new() -> Self {
        Drawing::default()
    }

    /// Width in pixels.
    pub fn width(&self) -> i64 {
        self.width
    }

    /// Set the width, scaling the height when `resize_proportional` is set.
    pub fn set_width(&mut self, width: i64) {
        if self.resize_proportional && width != 0 {
            let ratio = self.height as f64 / self.width as f64;
            self.height = (ratio * width as f64).round() as i64;
        }
        self.width = width;
    }

    /// Height in pixels.
    pub fn height(&self) -> i64 {
        self.height
    }

    /// Set the height, scaling the width when `resize_proportional` is set.
    pub fn set_height(&mut self, height: i64) {
        if self.resize_proportional && height != 0 {
            let ratio = self.width as f64 / self.height as f64;
            self.width = (ratio * height as f64).round() as i64;
        }
        self.height = height;
    }

    /// Resize to fit a box, keeping the aspect ratio when proportional.
    pub fn set_dimension(&mut self, width: i64, height: i64) {
        if !self.resize_proportional || width == 0 || height == 0 {
            return;
        }
        let x_ratio = width as f64 / self.width as f64;
        let y_ratio = height as f64 / self.height as f64;
        if x_ratio * (self.height as f64) < height as f64 {
            self.height = (x_ratio * self.height as f64).ceil() as i64;
            self.width = width;
        } else {
            self.width = (y_ratio * self.width as f64).ceil() as i64;
            self.height = height;
        }
    }

    /// `(x, y, cx, cy)` in EMUs.
    pub fn emu_dimensions(&self) -> (i64, i64, i64, i64) {
        (
            pixels_to_emu(self.left as f64),
            pixels_to_emu(self.top as f64),
            pixels_to_emu(self.width as f64),
            pixels_to_emu(self.height as f64),
        )
    }
}

/// A shape drawn inside a chart.
///
/// Coordinates are given in axis units and converted to percentages of the plot area.
#[derive(Debug, Clone, PartialEq)]
pub struct Shape {
    /// The anchor corner in axis units: bottom-left and top-right.
    pub axis_coordinates: ((f64, f64), (f64, f64)),
    /// The position as a fraction of the plot area.
    pub coordinates: (f64, f64, f64, f64),
    /// Text drawn in the shape.
    pub text: Option<String>,
    /// Theme colour scheme name.
    pub scheme: String,
    /// Preset geometry name; see [`Shape::RECT`].
    pub style: String,
    /// Border colour, stored in short form.
    border_color: String,
    /// Fill colour, stored in short form.
    color: String,
    /// Text colour, stored in short form.
    text_color: String,
    /// Border width.
    border_width: i64,
}

impl Default for Shape {
    fn default() -> Self {
        Shape {
            axis_coordinates: ((0.0, 0.0), (1.0, 1.0)),
            coordinates: (0.0, 0.0, 1.0, 1.0),
            text: None,
            scheme: "accent1".to_string(),
            style: Shape::RECT.to_string(),
            border_color: short_color(Color::BLACK),
            color: short_color(Color::WHITE),
            text_color: short_color(Color::BLACK),
            border_width: 0,
        }
    }
}

impl Shape {
    /// Left margin allowance, in pixels: font width plus the plot-area gutter.
    pub const MARGIN_LEFT: i64 = 6 + 13 + 1;
    /// Bottom margin allowance, in pixels.
    pub const MARGIN_BOTTOM: i64 = 17 + 11;
    /// Assumed font width in pixels, used to size the left margin.
    pub const FONT_WIDTH: i64 = 7;
    /// Assumed font height in pixels, used to size the bottom margin.
    pub const FONT_HEIGHT: i64 = 8;
    /// Rectangle geometry.
    pub const RECT: &'static str = "rect";
    /// Rounded rectangle geometry.
    pub const ROUND_RECT: &'static str = "roundRect";

    /// A rectangle shape covering the whole plot area.
    pub fn new() -> Self {
        Shape::default()
    }

    /// A rectangle shape with text.
    pub fn with_text(text: impl Into<String>) -> Self {
        Shape {
            text: Some(text.into()),
            ..Shape::default()
        }
    }

    /// Border colour in short (`RRGGBB`) form.
    pub fn border_color(&self) -> &str {
        &self.border_color
    }

    /// Set the border colour; the alpha prefix is stripped.
    pub fn set_border_color(&mut self, color: &str) {
        self.border_color = short_color(color);
    }

    /// Fill colour in short form.
    pub fn color(&self) -> &str {
        &self.color
    }

    /// Set the fill colour; the alpha prefix is stripped.
    pub fn set_color(&mut self, color: &str) {
        self.color = short_color(color);
    }

    /// Text colour in short form.
    pub fn text_color(&self) -> &str {
        &self.text_color
    }

    /// Set the text colour; the alpha prefix is stripped.
    pub fn set_text_color(&mut self, color: &str) {
        self.text_color = short_color(color);
    }

    /// Border width.
    pub fn border_width(&self) -> i64 {
        self.border_width
    }

    /// Set the border width.
    pub fn set_border_width(&mut self, width: i64) {
        self.border_width = width;
    }

    /// Set the anchor corner in axis units.
    pub fn set_axis_coordinates(&mut self, coordinates: ((f64, f64), (f64, f64))) {
        self.axis_coordinates = coordinates;
    }

    /// Force coordinates into the 0..=1 range so oversized shapes stay visible.
    ///
    /// A NaN passes through unchanged, which is what `clamp` does and what the writer
    /// needs: a broken anchor should be written as given rather than snapped to a corner.
    pub fn norm_pct(pct: f64) -> f64 {
        pct.clamp(0.0, 1.0)
    }
}

/// Fit `(width, height)` inside `(box_width, box_height)` keeping the aspect ratio.
///
/// Returns the new dimensions.
pub fn bounding_box(bw: i64, bh: i64, w: i64, h: i64) -> (i64, i64) {
    let mut new_width = w;
    let mut new_height = h;
    if bw != 0 && new_width > bw {
        new_width = bw;
        new_height = (new_width as f64 / (w as f64 / h as f64)).round() as i64;
    }
    if bh != 0 && new_height > bh {
        new_height = bh;
        new_width = (new_height as f64 * (w as f64 / h as f64)).round() as i64;
    }
    (new_width, new_height)
}

/// An image anchored to the sheet.
///
/// openpyxl requires PIL to measure and re-encode images; the Rust port keeps the raw bytes
/// and needs only the pixel dimensions, which `png` reads without decoding the image data.
#[derive(Debug, Clone, PartialEq)]
pub struct Image {
    /// The encoded image bytes.
    pub data: Vec<u8>,
    /// The image format, e.g. `png`.
    pub format: String,
    /// The pixel dimensions.
    pub size: (u32, u32),
    /// Whether the aspect ratio is locked.
    pub no_change_aspect: bool,
    /// Whether arrowheads are locked.
    pub no_change_arrowheads: bool,
    /// The anchor and size.
    pub drawing: Drawing,
}

impl Image {
    /// Decode a PNG and build an image from it.
    pub fn from_png(data: Vec<u8>) -> Result<Image> {
        let size = read_png_dimensions(&data)?;
        Ok(Image {
            data,
            format: "png".to_string(),
            size,
            no_change_aspect: true,
            no_change_arrowheads: true,
            drawing: Drawing::new(),
        })
    }

    /// Build an image from bytes whose dimensions are already known.
    pub fn new(data: Vec<u8>, format: &str, size: (u32, u32)) -> Self {
        Image {
            data,
            format: format.to_string(),
            size,
            no_change_aspect: true,
            no_change_arrowheads: true,
            drawing: Drawing::new(),
        }
    }

    /// The extension used for the media part.
    pub fn extension(&self) -> String {
        self.format.to_lowercase()
    }

    /// Resize the drawing to fit a box, preserving the aspect ratio.
    pub fn fit_to(&mut self, size: (Option<i64>, Option<i64>)) {
        let (new_width, new_height) = bounding_box(
            size.0.unwrap_or(0),
            size.1.unwrap_or(0),
            self.size.0 as i64,
            self.size.1 as i64,
        );
        self.drawing.set_width(new_width);
        self.drawing.set_height(new_height);
    }

    /// Anchor absolutely to a cell.
    ///
    /// The drawing's left/top are set from the cell's pixel anchor, and `point_pos` is used
    /// to resolve the cell containing the drawing's bottom-right corner, which is returned.
    pub fn anchor_absolute(
        &mut self,
        cell: (&str, u32),
        cell_left: i64,
        cell_top: i64,
        point_pos: &PointResolver,
    ) -> Result<(String, u32)> {
        self.drawing.anchor_type = AnchorType::Absolute;
        self.drawing.left = cell_left;
        self.drawing.top = cell_top;
        self.drawing.anchor_col = Some(column_index_from_string(cell.0)? as i64 - 1);
        self.drawing.anchor_row = Some(cell.1 as i64 - 1);
        Ok(point_pos(
            self.drawing.top + self.drawing.height(),
            self.drawing.left + self.drawing.width(),
        ))
    }

    /// Anchor to a single cell without moving it.
    pub fn anchor_one_cell(&mut self, column: &str, row: u32) -> Result<()> {
        self.drawing.anchor_type = AnchorType::OneCell;
        self.drawing.anchor_col = Some(column_index_from_string(column)? as i64 - 1);
        self.drawing.anchor_row = Some(row as i64 - 1);
        Ok(())
    }
}

/// Resolves a `(top, left)` pixel position to a `(column letters, row)` pair.
///
/// This is the callback a worksheet provides to [`Image::anchor_absolute`] so that the
/// drawing module does not need to know about worksheets.
pub type PointResolver = dyn Fn(i64, i64) -> (String, u32);

/// Read `(width, height)` from a PNG's IHDR chunk.
///
/// Only the header is parsed, so no image data is decoded and no allocation proportional to
/// the image size occurs.
pub fn read_png_dimensions(data: &[u8]) -> Result<(u32, u32)> {
    const SIGNATURE: [u8; 8] = [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];
    if data.len() < 24 || data[..8] != SIGNATURE {
        return Err(Error::Value(
            "Image data is not a valid PNG file".to_string(),
        ));
    }
    let width = u32::from_be_bytes([data[16], data[17], data[18], data[19]]);
    let height = u32::from_be_bytes([data[20], data[21], data[22], data[23]]);
    if width == 0 || height == 0 {
        return Err(Error::Value("PNG has zero dimensions".to_string()));
    }
    Ok((width, height))
}

/// Convenience: the column letters for a 1-based index, or an empty string when invalid.
pub fn column_letters(index: u32) -> String {
    get_column_letter(index).unwrap_or_default()
}

/// Convenience: resolve a 1-based column index from letters.
pub fn column_index(letters: &str) -> Result<u32> {
    column_index_from_string(letters)
}

/// A shape's colour attributes, for the writer.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeStyle {
    /// Preset geometry.
    pub style: String,
    /// Fill colour.
    pub color: String,
    /// Border colour.
    pub border_color: String,
    /// Border width.
    pub border_width: i64,
    /// Text colour.
    pub text_color: String,
}

impl Shape {
    /// The attributes the shape writer needs.
    pub fn style(&self) -> ShapeStyle {
        ShapeStyle {
            style: self.style.clone(),
            color: self.color.clone(),
            border_color: self.border_color.clone(),
            border_width: self.border_width,
            text_color: self.text_color.clone(),
        }
    }
}

/// Validate that an anchor column exists, for callers building drawings by hand.
pub fn validate_anchor_column(column: &str) -> Result<u32> {
    column_index_from_string(column)
        .map_err(|_| Error::ColumnStringIndex(format!("{column} is not a valid column name")))
}

/// The default column letters used for anchoring, for parity with the Python default `A`.
pub const DEFAULT_ANCHOR_COLUMN: &str = "A";

/// Produce a one-based column index from a letter run without bounds checks.
pub fn loose_column_index(letters: &str) -> u32 {
    letters
        .chars()
        .filter(|c| c.is_ascii_alphabetic())
        .fold(0u32, |acc, c| {
            acc * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1)
        })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn drawing_defaults_match_python() {
        let drawing = Drawing::new();
        assert_eq!(drawing.width(), emu_to_pixels(200_000));
        assert_eq!(drawing.height(), emu_to_pixels(1_828_800));
        assert_eq!(drawing.coordinates, ((1, 2), (16, 8)));
        assert_eq!(drawing.anchor_type, AnchorType::Absolute);
    }

    #[test]
    fn emu_dimensions_convert_pixels() {
        let mut drawing = Drawing::new();
        drawing.left = 10;
        drawing.top = 400;
        drawing.set_width(800);
        drawing.set_height(400);
        let (x, y, w, h) = drawing.emu_dimensions();
        assert_eq!(x, pixels_to_emu(10.0));
        assert_eq!(y, pixels_to_emu(400.0));
        assert_eq!(w, pixels_to_emu(800.0));
        assert_eq!(h, pixels_to_emu(400.0));
    }

    #[test]
    fn proportional_resize_keeps_ratio() {
        let mut drawing = Drawing::new();
        drawing.set_width(1000);
        drawing.set_height(500);
        drawing.resize_proportional = true;
        drawing.set_width(2000);
        assert_eq!(drawing.width(), 2000);
        assert_eq!(drawing.height(), 1000);
        drawing.set_height(500);
        assert_eq!(drawing.width(), 1000);
        assert_eq!(drawing.height(), 500);
    }

    #[test]
    fn set_dimension_only_applies_when_proportional() {
        let mut drawing = Drawing::new();
        drawing.set_width(100);
        drawing.set_height(100);
        drawing.set_dimension(50, 25);
        // Not proportional: dimensions unchanged.
        assert_eq!(drawing.width(), 100);
        assert_eq!(drawing.height(), 100);

        drawing.resize_proportional = true;
        drawing.set_dimension(50, 25);
        assert!(drawing.width() <= 50);
        assert!(drawing.height() <= 25);
    }

    #[test]
    fn colours_are_stored_in_short_form() {
        let mut shape = Shape::new();
        shape.set_color("FFFF0000");
        shape.set_border_color("FF00FF00");
        shape.set_text_color("FF0000FF");
        assert_eq!(shape.color(), "FF0000");
        assert_eq!(shape.border_color(), "00FF00");
        assert_eq!(shape.text_color(), "0000FF");
    }

    #[test]
    fn norm_pct_clamps() {
        assert_eq!(Shape::norm_pct(-0.5), 0.0);
        assert_eq!(Shape::norm_pct(0.25), 0.25);
        assert_eq!(Shape::norm_pct(1.5), 1.0);
    }

    #[test]
    fn bounding_box_preserves_ratio() {
        assert_eq!(bounding_box(100, 100, 200, 100), (100, 50));
        assert_eq!(bounding_box(0, 50, 200, 100), (100, 50));
        assert_eq!(bounding_box(1000, 1000, 200, 100), (200, 100));
    }

    #[test]
    fn shadow_defaults() {
        let shadow = Shadow::new();
        assert!(!shadow.visible);
        assert_eq!(shadow.alignment, "br");
        assert_eq!(shadow.alpha, 50);
    }

    /// A PNG containing only the signature and an `IHDR` chunk.
///
/// The decoder reads the dimensions from the header, so the image data is not needed for
/// the tests.
fn minimal_png(width: u32, height: u32) -> Vec<u8> {
        let mut png = vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];
        png.extend_from_slice(&13u32.to_be_bytes());
        png.extend_from_slice(b"IHDR");
        png.extend_from_slice(&width.to_be_bytes());
        png.extend_from_slice(&height.to_be_bytes());
        // Bit depth, colour type, compression, filter and interlace.
        png.extend_from_slice(&[8, 6, 0, 0, 0]);
        png.extend_from_slice(&[0u8; 4]);
        png
    }

    #[test]
    fn png_dimensions_are_read_from_the_header() {
        let png = minimal_png(64, 32);
        assert_eq!(read_png_dimensions(&png).unwrap(), (64, 32));
        // A zero dimension is rejected rather than producing a degenerate drawing.
        assert!(read_png_dimensions(&minimal_png(0, 32)).is_err());
        assert!(read_png_dimensions(b"not a png").is_err());
    }

    #[test]
    fn image_preserves_bytes_and_fits() {
        let png = minimal_png(200, 100);
        let mut image = Image::from_png(png.clone()).unwrap();
        assert_eq!(image.size, (200, 100));
        assert_eq!(image.extension(), "png");
        assert_eq!(image.data, png);
        image.fit_to((Some(50), None));
        assert_eq!(image.drawing.width(), 50);
        assert_eq!(image.drawing.height(), 25);
    }

    #[test]
    fn one_cell_anchor_records_zero_based_position() {
        let mut image = Image::new(vec![], "png", (10, 10));
        image.anchor_one_cell("C", 4).unwrap();
        assert_eq!(image.drawing.anchor_type, AnchorType::OneCell);
        assert_eq!(image.drawing.anchor_col, Some(2));
        assert_eq!(image.drawing.anchor_row, Some(3));
        assert!(image.anchor_one_cell("!!", 1).is_err());
    }

    #[test]
    fn column_helpers() {
        assert_eq!(column_letters(1), "A");
        assert_eq!(column_index("AA").unwrap(), 27);
        assert_eq!(loose_column_index("AA"), 27);
        assert_eq!(validate_anchor_column("B").unwrap(), 2);
    }
}
