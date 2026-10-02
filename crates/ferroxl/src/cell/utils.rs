//! Coordinate utilities (`openpyxl/cell/cell.py`, module-level helpers).

use crate::exceptions::{Error, Result};

/// The highest column index Excel 2007 supports (XFD).
pub const MAX_COLUMN_INDEX: u32 = 18278;

/// Convert a coordinate string like `B12` into a `(column_letters, row)` pair.
///
/// `$` anchors are tolerated on both the column and the row.
pub fn coordinate_from_string(coord_string: &str) -> Result<(String, u32)> {
    let upper = coord_string.to_uppercase();
    let bytes = upper.as_bytes();
    let mut idx = 0usize;
    if idx < bytes.len() && bytes[idx] == b'$' {
        idx += 1;
    }
    let col_start = idx;
    while idx < bytes.len() && bytes[idx].is_ascii_alphabetic() {
        idx += 1;
    }
    if idx == col_start {
        return Err(Error::CellCoordinates(format!(
            "Invalid cell coordinates ({coord_string})"
        )));
    }
    let column = upper[col_start..idx].to_string();
    if idx < bytes.len() && bytes[idx] == b'$' {
        idx += 1;
    }
    let row_part = &upper[idx..];
    if row_part.is_empty() || !row_part.bytes().all(|b| b.is_ascii_digit()) {
        return Err(Error::CellCoordinates(format!(
            "Invalid cell coordinates ({coord_string})"
        )));
    }
    let row: u32 = row_part.parse().map_err(|_| {
        Error::CellCoordinates(format!("Invalid cell coordinates ({coord_string})"))
    })?;
    if row == 0 {
        return Err(Error::CellCoordinates(format!(
            "There is no row 0 ({coord_string})"
        )));
    }
    Ok((column, row))
}

/// Convert a coordinate to an absolute coordinate string (`B12` → `$B$12`).
///
/// Range forms (`B12:D14`) expand to `$B$12:$D$14`.
pub fn absolute_coordinate(coord_string: &str) -> String {
    let parts: Vec<&str> = coord_string.split(':').collect();
    if parts.len() == 2 {
        if let (Ok((c1, r1)), Ok((c2, r2))) = (
            coordinate_from_string(parts[0]),
            coordinate_from_string(parts[1]),
        ) {
            return format!("${c1}${r1}:${c2}${r2}");
        }
    }
    match coordinate_from_string(coord_string) {
        Ok((column, row)) => format!("${column}${row}"),
        Err(_) => coord_string.to_string(),
    }
}

/// Convert a column number into a column letter (`3` → `C`).
///
/// Right shift by 26 to find the letters in reverse order; these numbers are 1-based
/// and become ASCII ordinals by adding 64.
pub fn get_column_letter(col_idx: u32) -> Result<String> {
    if !(1..=MAX_COLUMN_INDEX).contains(&col_idx) {
        return Err(Error::value(format!("Invalid column index {col_idx}")));
    }
    let mut letters = Vec::new();
    let mut idx = col_idx;
    while idx > 0 {
        // Shifting by one turns the 1-based index into the 0-based bijective base-26
        // form, so `Z` (26) yields remainder 25 rather than an exact division.
        let quotient = (idx - 1) / 26;
        let remainder = (idx - 1) % 26;
        idx = quotient;
        letters.push((b'A' + remainder as u8) as char);
    }
    letters.reverse();
    Ok(letters.into_iter().collect())
}

/// Convert a column name into its 1-based index (`A` → `1`).
pub fn column_index_from_string(str_col: &str) -> Result<u32> {
    let upper = str_col.to_uppercase();
    if upper.is_empty() || !upper.chars().all(|c| c.is_ascii_alphabetic()) {
        return Err(Error::ColumnStringIndex(format!(
            "{str_col} is not a valid column name"
        )));
    }
    let mut idx: u32 = 0;
    for ch in upper.chars() {
        idx = idx * 26 + (ch as u32 - 'A' as u32 + 1);
    }
    if !(1..=MAX_COLUMN_INDEX).contains(&idx) {
        return Err(Error::ColumnStringIndex(format!(
            "{str_col} is not a valid column name"
        )));
    }
    Ok(idx)
}

/// Build an A1 coordinate from a 1-based column index and row.
pub fn coordinate_from_index(col_idx: u32, row: u32) -> Result<String> {
    Ok(format!("{}{}", get_column_letter(col_idx)?, row))
}

/// Convert a 1-based column index and row into the `(letters, row)` pair.
pub fn split_coordinate(col_idx: u32, row: u32) -> Result<(String, u32)> {
    Ok((get_column_letter(col_idx)?, row))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_coordinates() {
        assert_eq!(
            coordinate_from_string("B12").unwrap(),
            ("B".to_string(), 12)
        );
        assert_eq!(
            coordinate_from_string("b12").unwrap(),
            ("B".to_string(), 12)
        );
        assert_eq!(
            coordinate_from_string("$B$12").unwrap(),
            ("B".to_string(), 12)
        );
        assert_eq!(
            coordinate_from_string("XFD1048576").unwrap(),
            ("XFD".to_string(), 1048576)
        );
    }

    #[test]
    fn rejects_bad_coordinates() {
        assert!(matches!(
            coordinate_from_string("B"),
            Err(Error::CellCoordinates(_))
        ));
        assert!(matches!(
            coordinate_from_string("12"),
            Err(Error::CellCoordinates(_))
        ));
        assert!(matches!(
            coordinate_from_string("B0"),
            Err(Error::CellCoordinates(_))
        ));
        assert!(matches!(
            coordinate_from_string("B1A"),
            Err(Error::CellCoordinates(_))
        ));
    }

    #[test]
    fn absolute_forms() {
        assert_eq!(absolute_coordinate("B12"), "$B$12");
        assert_eq!(absolute_coordinate("b12"), "$B$12");
        assert_eq!(absolute_coordinate("$B$12"), "$B$12");
        assert_eq!(absolute_coordinate("B12:D14"), "$B$12:$D$14");
        assert_eq!(absolute_coordinate("nonsense"), "nonsense");
    }

    #[test]
    fn column_letter_round_trip() {
        for (idx, expected) in [
            (1u32, "A"),
            (2, "B"),
            (3, "C"),
            (26, "Z"),
            (27, "AA"),
            (52, "AZ"),
            (53, "BA"),
            (702, "ZZ"),
            (703, "AAA"),
            // openpyxl allows indices up to 18278, which is `ZZZ`, not Excel's own
            // 16384-column `XFD` limit.
            (18278, "ZZZ"),
        ] {
            assert_eq!(get_column_letter(idx).unwrap(), expected);
            assert_eq!(column_index_from_string(expected).unwrap(), idx);
        }
        assert!(matches!(get_column_letter(0), Err(Error::Value(_))));
        assert!(matches!(get_column_letter(18279), Err(Error::Value(_))));
        assert!(matches!(
            column_index_from_string("A1"),
            Err(Error::ColumnStringIndex(_))
        ));
    }

    #[test]
    fn coordinate_from_index_matches_column_and_row() {
        assert_eq!(coordinate_from_index(1, 1).unwrap(), "A1");
        assert_eq!(coordinate_from_index(27, 3).unwrap(), "AA3");
        assert_eq!(split_coordinate(27, 3).unwrap(), ("AA".to_string(), 3));
    }
}
