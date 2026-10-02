//! Writing comments (`openpyxl/writer/comments.rs`).
//!
//! Comments live in two parts: `xl/commentsN.xml` holds the text and authors, and a
//! companion VML drawing holds the note shapes that Excel positions on the sheet.

use std::collections::BTreeMap;

use crate::cell::utils::column_index_from_string;
use crate::comments::Comment;
use crate::exceptions::Result;
use crate::worksheet::Worksheet;
use crate::xml::constants::SHEET_MAIN_NS;
use crate::xml::functions::Element;

/// An owned attribute.
///
/// Attribute arrays must be homogeneous, so literals and computed values are both widened
/// to String before being collected.
/// An owned attribute, so a literal and a computed value can share one array.
fn attr(key: impl Into<String>, value: impl Into<String>) -> (String, String) {
    (key.into(), value.into())
}

/// The VML namespace used for note shapes.
pub const VMLNS: &str = "urn:schemas-microsoft-com:vml";
/// The Office VML namespace used for note shapes.
pub const OFFICENS: &str = "urn:schemas-microsoft-com:office:office";
/// The Excel VML namespace used for note shapes.
pub const EXCELNS: &str = "urn:schemas-microsoft-com:office:excel";

/// A comment paired with the cell it is anchored to.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnchoredComment {
    /// The cell coordinate.
    pub coordinate: String,
    /// The 1-based row.
    pub row: u32,
    /// The column letters.
    pub column: String,
    /// The comment itself.
    pub comment: Comment,
}

/// Collect a sheet's comments with their coordinates and assign author ids.
pub fn collect_comments(worksheet: &Worksheet) -> (Vec<AnchoredComment>, Vec<String>) {
    let mut comments: Vec<AnchoredComment> = worksheet
        .cells()
        .filter(|cell| cell.comment.is_some())
        .map(|cell| AnchoredComment {
            coordinate: cell.coordinate(),
            row: cell.row,
            column: cell.column.clone(),
            comment: cell.comment.clone().expect("filtered above"),
        })
        .collect();
    // `cells()` iterates a hash map, so the order is sorted to make the output reproducible
    // and to match the row-major order openpyxl writes.
    comments.sort_by(|a, b| {
        let column = |entry: &AnchoredComment| {
            crate::cell::utils::column_index_from_string(&entry.column).unwrap_or(0)
        };
        a.row.cmp(&b.row).then_with(|| column(a).cmp(&column(b)))
    });
    let mut authors: Vec<String> = Vec::new();
    for entry in &comments {
        if !authors.contains(&entry.comment.author().to_string()) {
            authors.push(entry.comment.author().to_string());
        }
    }
    (comments, authors)
}

/// Serialise `xl/commentsN.xml`.
pub fn write_comments(worksheet: &Worksheet) -> String {
    let (comments, authors) = collect_comments(worksheet);
    let mut root = Element::new(format!("{{{SHEET_MAIN_NS}}}comments"));

    let mut author_list = Element::new(format!("{{{SHEET_MAIN_NS}}}authors"));
    for author in &authors {
        let mut node = Element::new(format!("{{{SHEET_MAIN_NS}}}author"));
        node.set_text(author);
        author_list.append(node);
    }
    root.append(author_list);

    let mut comment_list = Element::new(format!("{{{SHEET_MAIN_NS}}}commentList"));
    for entry in &comments {
        let author_id = authors
            .iter()
            .position(|a| a == entry.comment.author())
            .unwrap_or(0);
        let mut comment_node = Element::with_attributes(
            format!("{{{SHEET_MAIN_NS}}}comment"),
            [
                ("ref", entry.coordinate.clone()),
                ("authorId", author_id.to_string()),
                ("shapeId", "0".to_string()),
            ],
        );
        let mut text_tag = Element::new(format!("{{{SHEET_MAIN_NS}}}text"));
        let mut run_tag = Element::new(format!("{{{SHEET_MAIN_NS}}}r"));
        run_tag.append(Element::new(format!("{{{SHEET_MAIN_NS}}}rPr")));
        let mut value_tag = Element::new(format!("{{{SHEET_MAIN_NS}}}t"));
        value_tag.set_text(entry.comment.text());
        run_tag.append(value_tag);
        text_tag.append(run_tag);
        comment_node.append(text_tag);
        comment_list.append(comment_node);
    }
    root.append(comment_list);
    root.to_pretty_string()
}

/// Serialise the VML drawing that positions the note shapes.
///
/// The element name is literally `xml`, matching the Python implementation; Excel accepts
/// this unprefixed root because no namespace is applied to it.
pub fn write_comments_vml(worksheet: &Worksheet) -> Result<String> {
    let (comments, _) = collect_comments(worksheet);
    let mut root = Element::new("xml");

    let mut shape_layout = Element::with_attributes(
        format!("{{{OFFICENS}}}shapelayout"),
        [attr(format!("{{{VMLNS}}}ext"), "edit".to_string())],
    );
    shape_layout.append(Element::with_attributes(
        format!("{{{OFFICENS}}}idmap"),
        [
            attr(format!("{{{VMLNS}}}ext"), "edit".to_string()),
            attr("data".to_string(), "1".to_string()),
        ],
    ));
    root.append(shape_layout);

    let mut shape_type = Element::with_attributes(
        format!("{{{VMLNS}}}shapetype"),
        [
            attr("id".to_string(), "_x0000_t202".to_string()),
            attr("coordsize".to_string(), "21600,21600".to_string()),
            attr(format!("{{{OFFICENS}}}spt"), "202".to_string()),
            attr("path".to_string(), "m,l,21600r21600,l21600,xe".to_string()),
        ],
    );
    shape_type.append(Element::with_attributes(
        format!("{{{VMLNS}}}stroke"),
        [("joinstyle", "miter")],
    ));
    shape_type.append(Element::with_attributes(
        format!("{{{VMLNS}}}path"),
        [
            attr("gradientshapeok".to_string(), "t".to_string()),
            attr(format!("{{{OFFICENS}}}connecttype"), "rect".to_string()),
        ],
    ));
    root.append(shape_type);

    for (index, entry) in comments.iter().enumerate() {
        let row = entry.row - 1;
        let column = column_index_from_string(&entry.column)? - 1;
        let mut shape = Element::with_attributes(
            format!("{{{VMLNS}}}shape"),
            [
                attr("id".to_string(), format!("_x0000_s{}", index + 1026)),
                attr("type".to_string(), "#_x0000_t202".to_string()),
                attr(
                    "style".to_string(),
                    "position:absolute; margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden".to_string(),
                ),
                attr("fillcolor".to_string(), "#ffffe1".to_string()),
                attr(format!("{{{OFFICENS}}}insetmode"), "auto".to_string()),
            ],
        );
        shape.append(Element::with_attributes(
            format!("{{{VMLNS}}}fill"),
            [("color2", "#ffffe1")],
        ));
        shape.append(Element::with_attributes(
            format!("{{{VMLNS}}}shadow"),
            [("color", "black"), ("obscured", "t")],
        ));
        shape.append(Element::with_attributes(
            format!("{{{VMLNS}}}path"),
            [attr(
                format!("{{{OFFICENS}}}connecttype"),
                "none".to_string(),
            )],
        ));
        let mut textbox = Element::with_attributes(
            format!("{{{VMLNS}}}textbox"),
            [("style", "mso-direction-alt:auto")],
        );
        textbox.append(Element::with_attributes(
            "div",
            [("style", "text-align:left")],
        ));
        shape.append(textbox);

        let mut client_data =
            Element::with_attributes(format!("{{{EXCELNS}}}ClientData"), [("ObjectType", "Note")]);
        client_data.append(Element::new(format!("{{{EXCELNS}}}MoveWithCells")));
        client_data.append(Element::new(format!("{{{EXCELNS}}}SizeWithCells")));
        let mut auto_fill = Element::new(format!("{{{EXCELNS}}}AutoFill"));
        auto_fill.set_text("False");
        client_data.append(auto_fill);
        let mut row_node = Element::new(format!("{{{EXCELNS}}}Row"));
        row_node.set_text(row.to_string());
        client_data.append(row_node);
        let mut column_node = Element::new(format!("{{{EXCELNS}}}Column"));
        column_node.set_text(column.to_string());
        client_data.append(column_node);
        shape.append(client_data);
        root.append(shape);
    }
    Ok(root.to_pretty_string())
}

/// The author-to-id mapping used by the comments part.
pub fn author_ids(worksheet: &Worksheet) -> BTreeMap<String, String> {
    let (_, authors) = collect_comments(worksheet);
    authors
        .into_iter()
        .enumerate()
        .map(|(index, author)| (author, index.to_string()))
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::functions::fromstring;

    fn sheet_with_comments() -> Worksheet {
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        sheet
            .set_comment("A1", Some(Comment::new("first note", "Alice")))
            .unwrap();
        sheet
            .set_comment("C3", Some(Comment::new("second note", "Bob")))
            .unwrap();
        sheet
            .set_comment("B2", Some(Comment::new("shared author", "Alice")))
            .unwrap();
        sheet
    }

    #[test]
    fn collects_comments_and_authors_in_order() {
        let (comments, authors) = collect_comments(&sheet_with_comments());
        assert_eq!(comments.len(), 3);
        assert_eq!(authors, vec!["Alice".to_string(), "Bob".to_string()]);
        assert_eq!(comments[0].coordinate, "A1");
        assert_eq!(comments[0].row, 1);
        assert_eq!(comments[0].column, "A");
    }

    #[test]
    fn comments_part_lists_authors_and_bodies() {
        let xml = write_comments(&sheet_with_comments());
        let root = fromstring(xml.as_bytes()).expect("comments must parse");
        let authors = root
            .find(format!("{{{SHEET_MAIN_NS}}}authors"))
            .unwrap()
            .find_all(format!("{{{SHEET_MAIN_NS}}}author"));
        assert_eq!(authors.len(), 2);
        assert_eq!(authors[0].text.as_deref(), Some("Alice"));

        let comments = root
            .find(format!("{{{SHEET_MAIN_NS}}}commentList"))
            .unwrap()
            .find_all(format!("{{{SHEET_MAIN_NS}}}comment"));
        assert_eq!(comments.len(), 3);
        // Comments are written in row-major order, so A1, B2 then C3.
        assert_eq!(comments[0].get("ref"), Some("A1"));
        assert_eq!(comments[0].get("authorId"), Some("0"));
        assert_eq!(comments[1].get("ref"), Some("B2"));
        assert_eq!(
            comments[1].get("authorId"),
            Some("0"),
            "a repeated author reuses the id"
        );
        assert_eq!(comments[2].get("ref"), Some("C3"));
        assert_eq!(comments[2].get("authorId"), Some("1"));
        assert!(xml.contains("first note"));
        assert!(xml.contains("second note"));
    }

    #[test]
    fn vml_has_one_shape_per_comment() {
        let xml = write_comments_vml(&sheet_with_comments()).unwrap();
        let root = fromstring(xml.as_bytes()).expect("vml must parse");
        let shapes = root.find_all(format!("{{{VMLNS}}}shape"));
        assert_eq!(shapes.len(), 3);
        assert_eq!(shapes[0].get("id"), Some("_x0000_s1026"));
        assert_eq!(shapes[2].get("id"), Some("_x0000_s1028"));
        assert!(xml.contains("urn:schemas-microsoft-com:office:excel"));
    }

    #[test]
    fn vml_positions_shapes_zero_based() {
        let xml = write_comments_vml(&sheet_with_comments()).unwrap();
        let root = fromstring(xml.as_bytes()).unwrap();
        let shapes = root.find_all(format!("{{{VMLNS}}}shape"));
        // Shapes follow the row-major comment order, so B2 is the second shape, and its
        // anchor is zero-based: (row 1, column 1).
        let client_data = shapes[1].find(format!("{{{EXCELNS}}}ClientData")).unwrap();
        assert_eq!(client_data.find_text(format!("{{{EXCELNS}}}Row"), ""), "1");
        assert_eq!(
            client_data.find_text(format!("{{{EXCELNS}}}Column"), ""),
            "1"
        );
    }

    #[test]
    fn no_comments_produces_empty_parts() {
        let sheet = Worksheet::new("Sheet1").unwrap();
        let (comments, authors) = collect_comments(&sheet);
        assert!(comments.is_empty());
        assert!(authors.is_empty());
        let xml = write_comments(&sheet);
        let root = fromstring(xml.as_bytes()).unwrap();
        assert!(root
            .find(format!("{{{SHEET_MAIN_NS}}}commentList"))
            .unwrap()
            .children()
            .is_empty());
    }

    #[test]
    fn author_ids_are_indexed_from_zero() {
        let ids = author_ids(&sheet_with_comments());
        assert_eq!(ids.get("Alice").unwrap(), "0");
        assert_eq!(ids.get("Bob").unwrap(), "1");
    }
}
