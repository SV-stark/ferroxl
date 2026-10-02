//! The shared string table (`openpyxl/reader/strings.py`).

use std::collections::HashMap;

use crate::xml::constants::{SHEET_MAIN_NS, XML_NS};
use crate::xml::functions::{fromstring, Element};

/// Read all shared strings, keyed by their index.
///
/// Rich-text runs are concatenated and their formatting discarded. The literal
/// `x005F_` is stripped, which is how Excel escapes a leading underscore.
pub fn read_string_table(xml_source: &[u8]) -> HashMap<usize, String> {
    let mut table = HashMap::new();
    let Ok(root) = fromstring(xml_source) else {
        return table;
    };
    let tag = format!("{{{SHEET_MAIN_NS}}}si");
    for (index, node) in root.find_all(&tag).into_iter().enumerate() {
        let text = get_string(node).replace("x005F_", "");
        table.insert(index, text);
    }
    table
}

/// Read the contents of a specific `si` node.
pub fn get_string(node: &Element) -> String {
    let rich_tag = format!("{{{SHEET_MAIN_NS}}}r");
    let rich_nodes = node.find_all(&rich_tag);
    if rich_nodes.is_empty() {
        return get_text(node);
    }
    rich_nodes.iter().map(|rich| get_text(rich)).collect()
}

/// Read text, discarding formatting unless it must be preserved.
///
/// Whitespace is stripped unless `xml:space="preserve"` is set.
pub fn get_text(node: &Element) -> String {
    let text_tag = format!("{{{SHEET_MAIN_NS}}}t");
    let Some(text_node) = node.find(&text_tag) else {
        return String::new();
    };
    let text = text_node.text.clone().unwrap_or_default();
    if text_node
        .get(format!("{{{XML_NS}}}space"))
        .map(|space| space != "preserve")
        .unwrap_or(true)
    {
        return text.trim().to_string();
    }
    text
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reads_simple_strings() {
        let xml = br#"<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>hello</t></si><si><t>world</t></si></sst>"#;
        let table = read_string_table(xml);
        assert_eq!(table.get(&0).unwrap(), "hello");
        assert_eq!(table.get(&1).unwrap(), "world");
        assert_eq!(table.len(), 2);
    }

    #[test]
    fn concatenates_rich_text_runs() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t>a</t></r><r><t>b</t></r></si></sst>"#;
        let table = read_string_table(xml);
        assert_eq!(table.get(&0).unwrap(), "ab");
    }

    #[test]
    fn strips_escaped_leading_underscore() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>x005F_abc</t></si></sst>"#;
        let table = read_string_table(xml);
        assert_eq!(table.get(&0).unwrap(), "abc");
    }

    #[test]
    fn whitespace_is_trimmed_unless_preserved() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xml="http://www.w3.org/XML/1998/namespace"><si><t>  padded  </t></si><si><t xml:space="preserve">  kept  </t></si></sst>"#;
        let table = read_string_table(xml);
        assert_eq!(table.get(&0).unwrap(), "padded");
        assert_eq!(table.get(&1).unwrap(), "  kept  ");
    }

    #[test]
    fn empty_si_yields_empty_string() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si/></sst>"#;
        let table = read_string_table(xml);
        assert_eq!(table.get(&0).unwrap(), "");
    }

    #[test]
    fn malformed_input_yields_an_empty_table() {
        assert!(read_string_table(b"not xml").is_empty());
    }
}
