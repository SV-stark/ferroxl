//! Writing drawings (`openpyxl/writer/drawings.py`).
//!
//! One `drawingN.xml` per sheet holds the anchors for that sheet's charts and images; chart
//! user shapes go into their own `drawingN.xml` referenced from the chart part.

use crate::drawing::{AnchorType, Drawing, Shape};
use crate::worksheet::Worksheet;
use crate::xml::constants::{
    CHART_DRAWING_NS, CHART_NS, DRAWING_NS, PKG_REL_NS, REL_NS, SHEET_DRAWING_NS,
};
use crate::xml::functions::Element;

/// An owned attribute.
///
/// Attribute arrays must be homogeneous, so literals and computed values are both widened
/// to String before being collected.
/// An owned attribute, so a literal and a computed value can share one array.
fn attr(key: impl Into<String>, value: impl Into<String>) -> (String, String) {
    (key.into(), value.into())
}

/// Serialise `xl/drawings/drawingN.xml` for a sheet.
pub fn write_drawing(worksheet: &Worksheet) -> String {
    let mut root = Element::new(format!("{{{SHEET_DRAWING_NS}}}wsDr"));
    for (index, chart) in worksheet.charts.iter().enumerate() {
        write_chart(&mut root, &chart.drawing, index + 1);
    }
    for (index, image) in worksheet.images.iter().enumerate() {
        write_image(
            &mut root,
            &image.drawing,
            index + 1,
            image.no_change_aspect,
            image.no_change_arrowheads,
        );
    }
    root.to_pretty_string()
}

fn write_chart(root: &mut Element, drawing: &Drawing, index: usize) {
    let (x, y, w, h) = drawing.emu_dimensions();
    let mut anchor = Element::new(format!("{{{SHEET_DRAWING_NS}}}absoluteAnchor"));
    anchor.append(Element::with_attributes(
        format!("{{{SHEET_DRAWING_NS}}}pos"),
        [attr("x", x.to_string()), attr("y", y.to_string())],
    ));
    anchor.append(Element::with_attributes(
        format!("{{{SHEET_DRAWING_NS}}}ext"),
        [attr("cx", w.to_string()), attr("cy", h.to_string())],
    ));

    let mut frame = Element::with_attributes(
        format!("{{{SHEET_DRAWING_NS}}}graphicFrame"),
        [("macro", "")],
    );
    let mut name = Element::new(format!("{{{SHEET_DRAWING_NS}}}nvGraphicFramePr"));
    name.append(Element::with_attributes(
        format!("{{{SHEET_DRAWING_NS}}}cNvPr"),
        [
            ("id", &index.to_string()),
            ("name", &format!("Chart {index}")),
        ],
    ));
    name.append(Element::new(format!(
        "{{{SHEET_DRAWING_NS}}}cNvGraphicFramePr"
    )));
    frame.append(name);

    let mut transform = Element::new(format!("{{{SHEET_DRAWING_NS}}}xfrm"));
    transform.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}off"),
        [("x", "0"), ("y", "0")],
    ));
    transform.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}ext"),
        [("cx", "0"), ("cy", "0")],
    ));
    frame.append(transform);

    let mut graphic = Element::new(format!("{{{DRAWING_NS}}}graphic"));
    let mut data =
        Element::with_attributes(format!("{{{DRAWING_NS}}}graphicData"), [("uri", CHART_NS)]);
    data.append(Element::with_attributes(
        format!("{{{CHART_NS}}}chart"),
        [(format!("{{{REL_NS}}}id").as_str(), &format!("rId{index}"))],
    ));
    graphic.append(data);
    frame.append(graphic);
    anchor.append(frame);
    anchor.append(Element::new(format!("{{{SHEET_DRAWING_NS}}}clientData")));
    root.append(anchor);
}

/// Build an anchor element for a drawing.
///
/// The element is returned so the caller can populate the picture that belongs inside it.
fn write_anchor(drawing: &Drawing) -> Element {
    let (x, y, w, h) = drawing.emu_dimensions();
    let mut anchor = match drawing.anchor_type {
        AnchorType::OneCell => {
            let mut node = Element::new(format!("{{{SHEET_DRAWING_NS}}}oneCellAnchor"));
            let mut from = Element::new(format!("{{{SHEET_DRAWING_NS}}}from"));
            let mut col = Element::new(format!("{{{SHEET_DRAWING_NS}}}col"));
            col.set_text(drawing.anchor_col.unwrap_or(0).to_string());
            from.append(col);
            let mut col_off = Element::new(format!("{{{SHEET_DRAWING_NS}}}colOff"));
            col_off.set_text(x.to_string());
            from.append(col_off);
            let mut row = Element::new(format!("{{{SHEET_DRAWING_NS}}}row"));
            row.set_text(drawing.anchor_row.unwrap_or(0).to_string());
            from.append(row);
            let mut row_off = Element::new(format!("{{{SHEET_DRAWING_NS}}}rowOff"));
            row_off.set_text(y.to_string());
            from.append(row_off);
            node.append(from);
            node
        }
        AnchorType::Absolute => {
            let mut node = Element::new(format!("{{{SHEET_DRAWING_NS}}}absoluteAnchor"));
            node.append(Element::with_attributes(
                format!("{{{SHEET_DRAWING_NS}}}pos"),
                [attr("x", x.to_string()), attr("y", y.to_string())],
            ));
            node
        }
    };
    anchor.append(Element::with_attributes(
        format!("{{{SHEET_DRAWING_NS}}}ext"),
        [attr("cx", w.to_string()), attr("cy", h.to_string())],
    ));
    anchor
}

fn write_image(
    root: &mut Element,
    drawing: &Drawing,
    index: usize,
    no_change_aspect: bool,
    no_change_arrowheads: bool,
) {
    let mut anchor = write_anchor(drawing);

    let mut picture = Element::new(format!("{{{SHEET_DRAWING_NS}}}pic"));
    let mut name = Element::new(format!("{{{SHEET_DRAWING_NS}}}nvPicPr"));
    name.append(Element::with_attributes(
        format!("{{{SHEET_DRAWING_NS}}}cNvPr"),
        [
            ("id", &index.to_string()),
            ("name", &format!("Picture {index}")),
        ],
    ));
    // `cNvPicPr` carries `preferRelativeResize` and nothing else; `noChangeAspect` and
    // `noChangeArrowheads` belong on the `a:picLocks` child. Putting them on both is
    // harmless to Excel but makes openpyxl's `NonVisualPictureProperties.from_tree` raise a
    // TypeError, and a reader that raises drops every image in the drawing.
    let mut locks = Element::new(format!("{{{SHEET_DRAWING_NS}}}cNvPicPr"));
    locks.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}picLocks"),
        [
            ("noChangeAspect", if no_change_aspect { "1" } else { "0" }),
            (
                "noChangeArrowheads",
                if no_change_arrowheads { "1" } else { "0" },
            ),
        ],
    ));
    name.append(locks);
    picture.append(name);

    let mut blip_fill = Element::new(format!("{{{SHEET_DRAWING_NS}}}blipFill"));
    blip_fill.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}blip"),
        [
            attr(format!("{{{REL_NS}}}embed"), format!("rId{index}")),
            attr("cstate".to_string(), "print".to_string()),
        ],
    ));
    blip_fill.append(Element::new(format!("{{{DRAWING_NS}}}srcRect")));
    let mut stretch = Element::new(format!("{{{DRAWING_NS}}}stretch"));
    stretch.append(Element::new(format!("{{{DRAWING_NS}}}fillRect")));
    blip_fill.append(stretch);
    picture.append(blip_fill);

    let mut shape_properties =
        Element::with_attributes(format!("{{{SHEET_DRAWING_NS}}}spPr"), [("bwMode", "auto")]);
    let mut transform = Element::new(format!("{{{DRAWING_NS}}}xfrm"));
    transform.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}off"),
        [("x", "0"), ("y", "0")],
    ));
    transform.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}ext"),
        [("cx", "0"), ("cy", "0")],
    ));
    shape_properties.append(transform);
    let mut geometry =
        Element::with_attributes(format!("{{{DRAWING_NS}}}prstGeom"), [("prst", "rect")]);
    geometry.append(Element::new(format!("{{{DRAWING_NS}}}avLst")));
    shape_properties.append(geometry);
    shape_properties.append(Element::new(format!("{{{DRAWING_NS}}}noFill")));
    let mut line = Element::with_attributes(format!("{{{DRAWING_NS}}}ln"), [("w", "1")]);
    line.append(Element::new(format!("{{{DRAWING_NS}}}noFill")));
    line.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}miter"),
        [("lim", "800000")],
    ));
    line.append(Element::new(format!("{{{DRAWING_NS}}}headEnd")));
    line.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}tailEnd"),
        [("type", "none"), ("w", "med"), ("len", "med")],
    ));
    shape_properties.append(line);
    shape_properties.append(Element::new(format!("{{{DRAWING_NS}}}effectLst")));
    picture.append(shape_properties);
    anchor.append(picture);
    anchor.append(Element::new(format!("{{{SHEET_DRAWING_NS}}}clientData")));
    root.append(anchor);
}

/// Serialise `xl/drawings/_rels/drawingN.xml.rels` for a sheet.
pub fn write_drawing_rels(worksheet: &Worksheet, chart_id: u32, image_id: u32) -> String {
    let mut root = Element::new(format!("{{{PKG_REL_NS}}}Relationships"));
    for (index, _) in worksheet.charts.iter().enumerate() {
        root.append(Element::with_attributes(
            format!("{{{PKG_REL_NS}}}Relationship"),
            [
                ("Id", &format!("rId{}", index + 1)),
                ("Type", &format!("{REL_NS}/chart")),
                (
                    "Target",
                    &format!("../charts/chart{}.xml", chart_id + index as u32),
                ),
            ],
        ));
    }
    let offset = worksheet.charts.len();
    for (index, _) in worksheet.images.iter().enumerate() {
        root.append(Element::with_attributes(
            format!("{{{PKG_REL_NS}}}Relationship"),
            [
                ("Id", &format!("rId{}", offset + index + 1)),
                ("Type", &format!("{REL_NS}/image")),
                (
                    "Target",
                    &format!("../media/image{}.png", image_id + index as u32),
                ),
            ],
        ));
    }
    root.to_pretty_string()
}

/// Serialise the chart user-shapes drawing (`chartUserShapes`).
pub fn write_shapes(shapes: &[Shape], shape_id: usize) -> String {
    let mut root = Element::new(format!("{{{CHART_NS}}}userShapes"));
    for (offset, shape) in shapes.iter().enumerate() {
        let id = shape_id + offset;
        let mut anchor = Element::new(format!("{{{CHART_DRAWING_NS}}}relSizeAnchor"));
        let (x_start, y_start, x_end, y_end) = shape.coordinates;
        let mut from = Element::new(format!("{{{CHART_DRAWING_NS}}}from"));
        let mut x = Element::new(format!("{{{CHART_DRAWING_NS}}}x"));
        x.set_text(format_percent(x_start));
        from.append(x);
        let mut y = Element::new(format!("{{{CHART_DRAWING_NS}}}y"));
        y.set_text(format_percent(y_start));
        from.append(y);
        anchor.append(from);
        let mut to = Element::new(format!("{{{CHART_DRAWING_NS}}}to"));
        let mut x = Element::new(format!("{{{CHART_DRAWING_NS}}}x"));
        x.set_text(format_percent(x_end));
        to.append(x);
        let mut y = Element::new(format!("{{{CHART_DRAWING_NS}}}y"));
        y.set_text(format_percent(y_end));
        to.append(y);
        anchor.append(to);

        let style = shape.style();
        let mut sp = Element::with_attributes(
            format!("{{{CHART_DRAWING_NS}}}sp"),
            [("macro", ""), ("textlink", "")],
        );
        let mut non_visual = Element::new(format!("{{{CHART_DRAWING_NS}}}nvSpPr"));
        non_visual.append(Element::with_attributes(
            format!("{{{CHART_DRAWING_NS}}}cNvPr"),
            [("id", &id.to_string()), ("name", &format!("shape {id}"))],
        ));
        non_visual.append(Element::new(format!("{{{CHART_DRAWING_NS}}}cNvSpPr")));
        sp.append(non_visual);

        let mut shape_properties = Element::new(format!("{{{CHART_DRAWING_NS}}}spPr"));
        let mut transform = Element::new(format!("{{{DRAWING_NS}}}xfrm"));
        transform.append(Element::with_attributes(
            format!("{{{DRAWING_NS}}}off"),
            [("x", "0"), ("y", "0")],
        ));
        transform.append(Element::with_attributes(
            format!("{{{DRAWING_NS}}}ext"),
            [("cx", "0"), ("cy", "0")],
        ));
        shape_properties.append(transform);
        let mut geometry = Element::with_attributes(
            format!("{{{DRAWING_NS}}}prstGeom"),
            [("prst", &style.style)],
        );
        geometry.append(Element::new(format!("{{{DRAWING_NS}}}avLst")));
        shape_properties.append(geometry);
        let mut fill = Element::new(format!("{{{DRAWING_NS}}}solidFill"));
        fill.append(Element::with_attributes(
            format!("{{{DRAWING_NS}}}srgbClr"),
            [("val", &style.color)],
        ));
        shape_properties.append(fill);
        let mut line = Element::with_attributes(
            format!("{{{DRAWING_NS}}}ln"),
            [("w", &style.border_width.to_string())],
        );
        let mut line_fill = Element::new(format!("{{{DRAWING_NS}}}solidFill"));
        line_fill.append(Element::with_attributes(
            format!("{{{DRAWING_NS}}}srgbClr"),
            [("val", &style.border_color)],
        ));
        line.append(line_fill);
        shape_properties.append(line);
        sp.append(shape_properties);

        write_style(&mut sp);
        write_shape_text(&mut sp, shape, &style.text_color);
        anchor.append(sp);
        root.append(anchor);
    }
    root.to_pretty_string()
}

fn write_style(node: &mut Element) {
    let mut style = Element::new(format!("{{{CHART_DRAWING_NS}}}style"));
    let mut line_ref = Element::with_attributes(format!("{{{DRAWING_NS}}}lnRef"), [("idx", "2")]);
    let mut line_color =
        Element::with_attributes(format!("{{{DRAWING_NS}}}schemeClr"), [("val", "accent1")]);
    line_color.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}shade"),
        [("val", "50000")],
    ));
    line_ref.append(line_color);
    style.append(line_ref);

    let mut fill_ref = Element::with_attributes(format!("{{{DRAWING_NS}}}fillRef"), [("idx", "1")]);
    fill_ref.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}schemeClr"),
        [("val", "accent1")],
    ));
    style.append(fill_ref);

    let mut effect_ref =
        Element::with_attributes(format!("{{{DRAWING_NS}}}effectRef"), [("idx", "0")]);
    effect_ref.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}schemeClr"),
        [("val", "accent1")],
    ));
    style.append(effect_ref);

    let mut font_ref =
        Element::with_attributes(format!("{{{DRAWING_NS}}}fontRef"), [("idx", "minor")]);
    font_ref.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}schemeClr"),
        [("val", "lt1")],
    ));
    style.append(font_ref);
    node.append(style);
}

fn write_shape_text(node: &mut Element, shape: &Shape, text_color: &str) {
    let mut body = Element::new(format!("{{{CHART_DRAWING_NS}}}txBody"));
    body.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}bodyPr"),
        [("vertOverflow", "clip")],
    ));
    body.append(Element::new(format!("{{{DRAWING_NS}}}lstStyle")));
    let mut paragraph = Element::new(format!("{{{DRAWING_NS}}}p"));
    match &shape.text {
        Some(text) if !text.is_empty() => {
            let mut run = Element::new(format!("{{{DRAWING_NS}}}r"));
            let mut properties =
                Element::with_attributes(format!("{{{DRAWING_NS}}}rPr"), [("lang", "en-US")]);
            let mut fill = Element::new(format!("{{{DRAWING_NS}}}solidFill"));
            fill.append(Element::with_attributes(
                format!("{{{DRAWING_NS}}}srgbClr"),
                [("val", text_color)],
            ));
            properties.append(fill);
            run.append(properties);
            let mut value = Element::new(format!("{{{DRAWING_NS}}}t"));
            value.set_text(text);
            run.append(value);
            paragraph.append(run);
        }
        _ => {
            paragraph.append(Element::with_attributes(
                format!("{{{DRAWING_NS}}}endParaRPr"),
                [("lang", "en-US")],
            ));
        }
    }
    body.append(paragraph);
    node.append(body);
}

fn format_percent(value: f64) -> String {
    if value == value.trunc() {
        format!("{}", value as i64)
    } else {
        format!("{value}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{BarChart, LineChart};
    use crate::xml::functions::fromstring;

    fn sheet_with_chart() -> Worksheet {
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        sheet.charts.push(BarChart::new().0.base);
        sheet
    }

    #[test]
    fn drawing_contains_one_anchor_per_chart() {
        let xml = write_drawing(&sheet_with_chart());
        let root = fromstring(xml.as_bytes()).expect("drawing must parse");
        let anchors = root.find_all(format!("{{{SHEET_DRAWING_NS}}}absoluteAnchor"));
        assert_eq!(anchors.len(), 1);
        let anchor = &anchors[0];
        assert!(anchor.find(format!("{{{SHEET_DRAWING_NS}}}pos")).is_some());
        assert!(anchor.find(format!("{{{SHEET_DRAWING_NS}}}ext")).is_some());
        assert!(anchor
            .find(format!("{{{SHEET_DRAWING_NS}}}clientData"))
            .is_some());
    }

    #[test]
    fn chart_anchor_references_the_chart_relationship() {
        let xml = write_drawing(&sheet_with_chart());
        // The chart part is referenced through a relationship id attribute.
        assert!(xml.contains("r:id=\"rId1\""), "{xml}");
        assert!(xml.contains("graphicFrame"));
    }

    #[test]
    fn image_anchor_uses_absolute_by_default() {
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        sheet.add_image(crate::drawing::Image::new(vec![], "png", (100, 50)));
        let xml = write_drawing(&sheet);
        let root = fromstring(xml.as_bytes()).unwrap();
        assert_eq!(
            root.find_all(format!("{{{SHEET_DRAWING_NS}}}absoluteAnchor"))
                .len(),
            1
        );
        assert!(xml.contains("blipFill"));
        assert!(xml.contains("noChangeAspect=\"1\""));
    }

    /// openpyxl's `NonVisualPictureProperties` accepts only `preferRelativeResize`, so a
    /// `cNvPicPr` carrying `noChangeAspect` makes its reader raise a `TypeError` -- and a
    /// reader that raises drops every image in the drawing. The attributes belong on the
    /// `a:picLocks` child.
    #[test]
    fn the_picture_lock_attributes_are_on_pic_locks_not_cnvpicpr() {
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        sheet.add_image(crate::drawing::Image::new(vec![], "png", (100, 50)));
        let xml = write_drawing(&sheet);
        let root = fromstring(xml.as_bytes()).unwrap();

        // `find` matches direct children, and `pic` sits inside the anchor, so each level of the
        // path down to `cNvPicPr` is walked explicitly.
        let anchor = &root.find_all(format!("{{{SHEET_DRAWING_NS}}}absoluteAnchor"))[0];
        let picture = anchor
            .find(format!("{{{SHEET_DRAWING_NS}}}pic"))
            .expect("the anchor must contain a picture");
        let non_visual = picture
            .find(format!("{{{SHEET_DRAWING_NS}}}nvPicPr"))
            .expect("nvPicPr");
        let properties = non_visual
            .find(format!("{{{SHEET_DRAWING_NS}}}cNvPicPr"))
            .expect("cNvPicPr");
        assert!(
            properties.attributes.is_empty(),
            "cNvPicPr must carry no attributes openpyxl does not accept, found {:?}",
            properties.attributes
        );
        let locks = properties
            .find(format!("{{{DRAWING_NS}}}picLocks"))
            .expect("the lock attributes must still be written, on picLocks");
        assert_eq!(locks.get("noChangeAspect"), Some("1"));
        assert_eq!(locks.get("noChangeArrowheads"), Some("1"));
    }

    #[test]
    fn one_cell_anchors_record_zero_based_positions() {
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        let mut image = crate::drawing::Image::new(vec![], "png", (100, 50));
        image.anchor_one_cell("C", 4).unwrap();
        sheet.add_image(image);
        let xml = write_drawing(&sheet);
        let root = fromstring(xml.as_bytes()).unwrap();
        let anchor = &root.find_all(format!("{{{SHEET_DRAWING_NS}}}oneCellAnchor"))[0];
        let from = anchor.find(format!("{{{SHEET_DRAWING_NS}}}from")).unwrap();
        assert_eq!(
            from.find_text(format!("{{{SHEET_DRAWING_NS}}}col"), ""),
            "2"
        );
        assert_eq!(
            from.find_text(format!("{{{SHEET_DRAWING_NS}}}row"), ""),
            "3"
        );
    }

    #[test]
    fn drawing_rels_number_charts_then_images() {
        let mut sheet = sheet_with_chart();
        sheet.charts.push(LineChart::new().0.base);
        sheet.add_image(crate::drawing::Image::new(vec![], "png", (10, 10)));
        let xml = write_drawing_rels(&sheet, 1, 1);
        let root = fromstring(xml.as_bytes()).unwrap();
        let relationships = root.find_all(format!("{{{PKG_REL_NS}}}Relationship"));
        assert_eq!(relationships.len(), 3);
        assert_eq!(relationships[0].get("Id"), Some("rId1"));
        assert_eq!(relationships[0].get("Target"), Some("../charts/chart1.xml"));
        assert_eq!(relationships[1].get("Target"), Some("../charts/chart2.xml"));
        // Images continue the chart id sequence rather than restarting.
        assert_eq!(relationships[2].get("Id"), Some("rId3"));
        assert_eq!(relationships[2].get("Target"), Some("../media/image1.png"));
    }

    #[test]
    fn shapes_are_written_with_style_and_text() {
        let shapes = vec![Shape::with_text("annotation")];
        let xml = write_shapes(&shapes, 1);
        let root = fromstring(xml.as_bytes()).expect("shapes must parse");
        let anchors = root.find_all(format!("{{{CHART_DRAWING_NS}}}relSizeAnchor"));
        assert_eq!(anchors.len(), 1);
        assert!(xml.contains("annotation"));
        assert!(xml.contains("cNvPr"));
        assert!(xml.contains("prstGeom"));
        assert!(xml.contains("lnRef"));
        assert!(!xml.contains("endParaRPr"));
    }

    #[test]
    fn shapes_without_text_emit_end_para_properties() {
        let shapes = vec![Shape::new()];
        let xml = write_shapes(&shapes, 1);
        assert!(xml.contains("endParaRPr"));
    }

    #[test]
    fn shape_ids_increment() {
        let shapes = vec![Shape::new(), Shape::new()];
        let xml = write_shapes(&shapes, 5);
        assert!(xml.contains("id=\"5\""));
        assert!(xml.contains("id=\"6\""));
    }

    #[test]
    fn empty_sheet_produces_an_empty_root() {
        let xml = write_drawing(&Worksheet::new("Sheet1").unwrap());
        let root = fromstring(xml.as_bytes()).unwrap();
        assert!(root.children().is_empty());
    }
}
