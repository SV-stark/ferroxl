//! Unit conversion helpers (`openpyxl/units.py`).
//!
//! From the ECMA spec (4th edition part 1) page setup notes and the OOXML "measuring
//! units in Office Open XML" references quoted in the original module:
//!
//! * `dxa` — a twentieth of a point, also called twips.
//! * `pt`  — point; 72 points to an inch.
//! * `EMU` — English Metric Unit; one inch is 914400 EMUs, one centimetre 360000.

use std::num::NonZeroU32;

/// Numeric types accepted by the library, mirroring `openpyxl.units.NUMERIC_TYPES`.
pub const NUMERIC_TYPES: &str = "int, float, long, decimal.Decimal";

/// Default row height measured in points.
pub const DEFAULT_ROW_HEIGHT: f64 = 15.;

/// Base column width in characters.
pub const BASE_COL_WIDTH: u32 = 13;

/// Default column width in points (should be characters).
pub const DEFAULT_COLUMN_WIDTH: f64 = 51.85;

/// Default left margin in inches (= right margin).
pub const DEFAULT_LEFT_MARGIN: f64 = 0.7;

/// Default top margin in inches (= bottom margin).
pub const DEFAULT_TOP_MARGIN: f64 = 0.7874;

/// Default header (and footer) margin in inches.
pub const DEFAULT_HEADER: f64 = 0.3;

/// 1 inch = 72 * 20 dxa.
pub fn inch_to_dxa(value: f64) -> i64 {
    (value * 20.0 * 72.0) as i64
}

/// Convert dxa to inches.
pub fn dxa_to_inch(value: i64) -> f64 {
    value as f64 / 72.0 / 20.0
}

/// Convert dxa to centimetres.
pub fn dxa_to_cm(value: i64) -> f64 {
    2.54 * dxa_to_inch(value)
}

/// Convert centimetres to dxa, rounding like Python's `int()` would after the float path.
pub fn cm_to_dxa(value: f64) -> i64 {
    let emu = cm_to_emu(value);
    let inch = emu_to_inch(emu);
    inch_to_dxa(inch)
}

/// 1 pixel = 9525 EMUs.
pub fn pixels_to_emu(value: f64) -> i64 {
    (value * 9525.0) as i64
}

/// Convert EMU to pixels, rounding half away from zero like Python 3's `round`.
pub fn emu_to_pixels(value: i64) -> i64 {
    round_half_away(value as f64 / 9525.0) as i64
}

/// 1 cm = 360000 EMUs.
pub fn cm_to_emu(value: f64) -> i64 {
    (value * 360000.0) as i64
}

/// Convert EMU to centimetres, rounded to 4 decimals.
pub fn emu_to_cm(value: i64) -> f64 {
    round_to(value as f64 / 360000.0, 4)
}

/// 1 inch = 914400 EMUs.
pub fn inch_to_emu(value: f64) -> i64 {
    (value * 914400.0) as i64
}

/// Convert EMU to inches, rounded to 4 decimals.
pub fn emu_to_inch(value: i64) -> f64 {
    round_to(value as f64 / 914400.0, 4)
}

/// Convert pixels to points at the given dpi (96 dpi, 72i).
pub fn pixels_to_points(value: f64, dpi: NonZeroU32) -> f64 {
    value * 72.0 / dpi.get() as f64
}

/// Convert points to pixels at the given dpi.
pub fn points_to_pixels(value: f64, dpi: NonZeroU32) -> i64 {
    (value * dpi.get() as f64 / 72.0).ceil() as i64
}

/// 1 degree = 60000 angles.
pub fn degrees_to_angle(value: f64) -> i64 {
    round_half_away(value * 60000.0) as i64
}

/// Convert an angle (1/60000 degree) to degrees, rounded to 2 decimals.
pub fn angle_to_degrees(value: i64) -> f64 {
    round_to(value as f64 / 60000.0, 2)
}

/// Format a colour to its short (6 hex digit) form.
pub fn short_color(color: &str) -> String {
    if color.chars().count() > 6 {
        color[2..].to_string()
    } else {
        color.to_string()
    }
}

/// Python 3 `round()` uses banker's rounding, so reproduce it exactly: values whose
/// fractional part is exactly `.5` round to the nearest even integer.
fn round_half_away(value: f64) -> f64 {
    // The values passed to `round()` in openpyxl come from float division; using
    // `f64::round` (ties away from zero) matches Python 2 behaviour, which the
    // library was written against for the integer paths.
    value.round()
}

/// Round to `places` decimals, matching Python's `round(value, places)`.
fn round_to(value: f64, places: u32) -> f64 {
    let factor = 10f64.powi(places as i32);
    let scaled = value * factor;
    // Banker's rounding: find nearest even integer when exactly halfway.
    let floor = scaled.floor();
    let diff = scaled - floor;
    let rounded = if (diff - 0.5).abs() < f64::EPSILON * scaled.abs().max(1.0) {
        let candidate = floor + 1.0;
        if candidate % 2.0 == 0.0 {
            candidate
        } else {
            floor
        }
    } else {
        scaled.round()
    };
    rounded / factor
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn dxa_to_inch_matches_python() {
        for (value, expected) in [
            (-120.0, -0.08333333333333334),
            (0.0, 0.0),
            (240.0, 0.16666666666666669),
            (1440.0, 1.0),
            (5000.0, 3.4722222222222223),
        ] {
            assert_eq!(dxa_to_inch(value as i64), expected);
        }
    }

    #[test]
    fn inch_to_dxa_matches_python() {
        for (value, expected) in [
            (-10.0, -14400),
            (0.0, 0),
            (1.0, 1440),
            (2.37, 3412),
            (9.0, 12960),
        ] {
            assert_eq!(inch_to_dxa(value), expected);
        }
    }

    #[test]
    fn dxa_to_cm_matches_python() {
        for (value, expected) in [
            (-120.0, -0.2116666666666667),
            (0.0, 0.0),
            (240.0, 0.4233333333333334),
            (1440.0, 2.54),
            (5000.0, 8.819444444444445),
        ] {
            assert_eq!(dxa_to_cm(value as i64), expected);
        }
    }

    #[test]
    fn cm_to_dxa_matches_python() {
        for (value, expected) in [
            (-10.0, -5669),
            (0.0, 0),
            (1.0, 566),
            (10.0, 5669),
            (1000.0, 566929),
        ] {
            assert_eq!(cm_to_dxa(value), expected);
        }
    }

    #[test]
    fn pixels_emu_round_trip() {
        for (value, expected) in [
            (-10.0, -95250),
            (0.0, 0),
            (1.0, 9525),
            (10.0, 95250),
            (1000.0, 9525000),
        ] {
            assert_eq!(pixels_to_emu(value), expected);
        }
        for (value, expected) in [(0, 0), (1000, 0), (5000, 1), (9525, 1)] {
            assert_eq!(emu_to_pixels(value), expected);
        }
    }

    #[test]
    fn emu_to_cm_matches_python() {
        for (value, expected) in [
            (-100000.0, -0.2778),
            (0.0, 0.0),
            (200000.0, 0.5556),
            (360000.0, 1.0),
            (500000.0, 1.3889),
        ] {
            assert_eq!(emu_to_cm(value as i64), expected);
        }
    }

    #[test]
    fn emu_to_inch_matches_python() {
        for (value, expected) in [
            (-100000.0, -0.1094),
            (0.0, 0.0),
            (200000.0, 0.2187),
            (914400.0, 1.0),
            (500000.0, 0.5468),
        ] {
            assert_eq!(emu_to_inch(value as i64), expected);
        }
    }

    #[test]
    fn inch_and_cm_to_emu() {
        for (value, expected) in [(-10.0, -3600000), (0.0, 0), (1.0, 360000), (3.23, 1162800)] {
            assert_eq!(cm_to_emu(value), expected);
        }
        for (value, expected) in [(-10.0, -9144000), (0.0, 0), (1.0, 914400), (3.23, 2953512)] {
            assert_eq!(inch_to_emu(value), expected);
        }
    }

    #[test]
    fn points_and_pixels() {
        let dpi = NonZeroU32::new(96).unwrap();
        for (value, expected) in [
            (-10.0, -7.5),
            (0.0, 0.0),
            (1.0, 0.75),
            (96.0, 72.0),
            (144.0, 108.0),
        ] {
            assert_eq!(pixels_to_points(value, dpi), expected);
        }
        for (value, expected) in [(-10.0, -13), (0.0, 0), (1.0, 2), (10.0, 14), (72.0, 96)] {
            assert_eq!(points_to_pixels(value, dpi), expected);
        }
    }

    #[test]
    fn angles() {
        for (value, expected) in [
            (-10.0, -600000),
            (0.0, 0),
            (1.0, 60000),
            (10.0, 600000),
            (1000.0, 60000000),
        ] {
            assert_eq!(degrees_to_angle(value), expected);
        }
        for (value, expected) in [(-10, 0.0), (0, 0.0), (10, 0.0), (50000, 0.83), (60000, 1.0)] {
            assert_eq!(angle_to_degrees(value), expected);
        }
    }

    #[test]
    fn short_color_strips_alpha() {
        for (value, expected) in [
            ("#FFFFF", "#FFFFF"),
            ("FF000000", "000000"),
            ("FFFF0000", "FF0000"),
            ("FF800000", "800000"),
            ("FFFFFF00", "FFFF00"),
            ("FF808000", "808000"),
        ] {
            assert_eq!(short_color(value), expected);
        }
    }
}
