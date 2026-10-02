//! The shared string table (`openpyxl/writer/strings.py`).

use std::collections::HashMap;

use crate::xml::constants::SHEET_MAIN_NS;
use crate::xml::functions::XmlWriter;

/// A mapping from cell text to its index in the shared string table.
pub type StringTable = HashMap<String, usize>;

/// Collect the shared strings for a workbook.
///
/// Strings are collected from cells whose type is `s`, sorted so the table is
/// deterministic, and indexed from zero.
pub fn create_string_table(worksheets: &[crate::worksheet::Worksheet]) -> StringTable {
    let mut strings: Vec<String> = Vec::new();
    for sheet in worksheets {
        for cell in sheet.cells() {
            if cell.data_type != crate::cell::DataType::SharedString {
                continue;
            }
            if let crate::cell::CellValue::Text(text) = cell.internal_value() {
                if !strings.contains(text) {
                    strings.push(text.clone());
                }
            }
        }
    }
    strings.sort();
    strings
        .into_iter()
        .enumerate()
        .map(|(index, text)| (text, index))
        .collect()
}

/// Serialise the shared string table.
///
/// Entries whose text has leading or trailing whitespace get `xml:space="preserve"` so
/// Excel keeps them intact.
pub fn write_string_table(table: &StringTable) -> String {
    let mut doc = XmlWriter::new();
    doc.start_tag(
        "sst",
        [
            ("xmlns", SHEET_MAIN_NS),
            ("uniqueCount", &table.len().to_string()),
        ],
    );
    let mut ordered: Vec<(&String, &usize)> = table.iter().collect();
    ordered.sort_by_key(|(_, index)| **index);
    for (text, _) in ordered {
        doc.start_tag("si", [] as [(&str, &str); 0]);
        if text.trim() != text.as_str() {
            doc.tag("t", [("xml:space", "preserve")], Some(text));
        } else {
            doc.tag("t", [] as [(&str, &str); 0], Some(text));
        }
        doc.end_tag("si");
    }
    doc.end_tag("sst");
    doc.into_string()
}

/// An incremental string-table builder used by the streaming writer.
///
/// Assigns indices in insertion order, so a streaming writer can write each string exactly
/// once as it encounters it.
#[derive(Debug, Clone, Default)]
pub struct StringTableBuilder {
    counter: usize,
    table: HashMap<String, usize>,
}

impl StringTableBuilder {
    /// An empty builder.
    pub fn new() -> Self {
        StringTableBuilder::default()
    }

    /// Return the index for a string, assigning a new one if unseen.
    ///
    /// The key is stripped before lookup, so surrounding whitespace is not significant.
    pub fn add(&mut self, key: &str) -> usize {
        let key = key.trim().to_string();
        if let Some(index) = self.table.get(&key) {
            return *index;
        }
        let index = self.counter;
        self.table.insert(key, index);
        self.counter += 1;
        index
    }

    /// The accumulated table.
    pub fn get_table(&self) -> &StringTable {
        &self.table
    }

    /// Number of distinct strings.
    pub fn len(&self) -> usize {
        self.table.len()
    }

    /// Whether no strings have been added.
    pub fn is_empty(&self) -> bool {
        self.table.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::{CellValue, DataType};
    use crate::worksheet::Worksheet;

    #[test]
    fn collects_and_sorts_strings() {
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        sheet.set("A1", CellValue::text("zebra")).unwrap();
        sheet.set("A2", CellValue::text("apple")).unwrap();
        sheet.set("A3", CellValue::text("zebra")).unwrap();
        let table = create_string_table(std::slice::from_ref(&sheet));
        assert_eq!(table.len(), 2);
        assert_eq!(table["apple"], 0);
        assert_eq!(table["zebra"], 1);
    }

    #[test]
    fn ignores_non_string_cells() {
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        sheet.set("A1", 42.0).unwrap();
        sheet.set("A2", CellValue::text("=SUM(A1:A1)")).unwrap();
        assert!(create_string_table(std::slice::from_ref(&sheet)).is_empty());
    }

    #[test]
    fn writes_sst_with_indices_in_order() {
        let mut table = StringTable::new();
        table.insert("second".to_string(), 1);
        table.insert("first".to_string(), 0);
        let xml = write_string_table(&table);
        assert!(xml.starts_with("<sst"));
        assert!(xml.contains("uniqueCount=\"2\""));
        // first (index 0) must appear before second (index 1).
        let first_pos = xml.find(">first<").expect("first entry");
        let second_pos = xml.find(">second<").expect("second entry");
        assert!(first_pos < second_pos);
    }

    #[test]
    fn preserves_padded_strings() {
        let mut table = StringTable::new();
        table.insert("  padded  ".to_string(), 0);
        table.insert("tight".to_string(), 1);
        let xml = write_string_table(&table);
        assert!(xml.contains("xml:space=\"preserve\""));
        // Only the padded entry carries the attribute.
        assert_eq!(xml.matches("xml:space=\"preserve\"").count(), 1);
    }

    #[test]
    fn builder_assigns_incrementing_indices() {
        let mut builder = StringTableBuilder::new();
        assert_eq!(builder.add("a"), 0);
        assert_eq!(builder.add("b"), 1);
        assert_eq!(builder.add("a"), 0);
        assert_eq!(builder.len(), 2);
        assert!(!builder.is_empty());
        assert_eq!(builder.get_table().len(), 2);
    }

    #[test]
    fn builder_strips_surrounding_whitespace() {
        let mut builder = StringTableBuilder::new();
        assert_eq!(builder.add("a"), 0);
        assert_eq!(builder.add("  a  "), 0);
        assert_eq!(builder.len(), 1);
    }

    #[test]
    fn empty_table_still_produces_valid_xml() {
        let xml = write_string_table(&StringTable::new());
        assert!(xml.contains("uniqueCount=\"0\""));
        assert!(xml.contains("</sst>"));
    }

    #[test]
    fn explicit_types_are_respected() {
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        sheet
            .cell_mut("A1")
            .unwrap()
            .set_explicit_value(CellValue::text("kept"), DataType::SharedString)
            .unwrap();
        let table = create_string_table(std::slice::from_ref(&sheet));
        assert_eq!(table.get("kept"), Some(&0));
    }
}
