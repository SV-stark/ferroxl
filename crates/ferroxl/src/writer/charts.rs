//! Writing charts (`openpyxl/writer/charts.py`).
//!
//! Axis scaling is recomputed from the series data immediately before writing, so the
//! `<c:max>`/`<c:min>` values always match what is actually in the chart.

use crate::charts::axis::Axis;
use crate::charts::chart::{BarChart, Chart, GraphChart, LineChart, PieChart, ScatterChart};
use crate::charts::error_bar::ErrorBarType;
use crate::charts::reference::{CellValue, Reference, ReferenceDataType};
use crate::charts::series::Series;
use crate::exceptions::{Error, Result};
use crate::xml::constants::{CHART_NS, DRAWING_NS, PKG_REL_NS, REL_NS};
use crate::xml::functions::{safe_string, Element};

/// A chart that has had its axes scaled and is ready to serialise.
struct PreparedChart {
    chart_type: String,
    grouping: String,
    series: Vec<Series>,
    legend_position: String,
    show_legend: bool,
    lang: String,
    title: String,
    margin_left: f64,
    margin_top: f64,
    width: f64,
    height: f64,
    print_margins: Vec<(String, String)>,
    has_shapes: bool,
    axes: Option<(Axis, Axis)>,
}

fn prepare(chart: &Chart) -> PreparedChart {
    PreparedChart {
        chart_type: chart.chart_type.to_string(),
        grouping: chart.grouping.to_string(),
        series: chart.series.clone(),
        legend_position: chart.legend.position.clone(),
        show_legend: chart.show_legend,
        lang: chart.lang.clone(),
        title: chart.title.clone(),
        margin_left: chart.margin_left(),
        margin_top: chart.margin_top(),
        width: chart.width,
        height: chart.height,
        print_margins: chart.print_margins.clone(),
        has_shapes: !chart.shapes.is_empty(),
        axes: None,
    }
}

fn prepare_graph(chart: &GraphChart) -> PreparedChart {
    let mut prepared = prepare(&chart.base);
    prepared.axes = Some((chart.x_axis.clone(), chart.y_axis.clone()));
    prepared
}

/// Serialise a chart to `xl/charts/chartN.xml`.
///
/// The chart variant is selected from the concrete type: pie charts omit axes, scatter
/// charts use `yVal`/`xVal` series elements and every other graph chart uses `val`.
pub fn write_chart(chart: &Chart) -> Result<String> {
    let prepared = prepare(chart);
    let mut root = Element::new(format!("{{{CHART_NS}}}chartSpace"));
    root.append(Element::with_attributes(
        format!("{{{CHART_NS}}}lang"),
        [("val", &prepared.lang)],
    ));
    write_chart_node(&mut root, &prepared)?;
    write_print_settings(&mut root, &prepared.print_margins);
    if prepared.has_shapes {
        root.append(Element::with_attributes(
            format!("{{{CHART_NS}}}userShapes"),
            [(format!("{{{REL_NS}}}id").as_str(), "rId1")],
        ));
    }
    Ok(root.to_pretty_string())
}

/// Serialise a graph chart, recomputing the axis bounds first.
pub fn write_graph_chart(chart: &GraphChart) -> String {
    let mut scaled = chart.clone();
    if scaled.auto_axis {
        scaled.compute_axes();
    }
    let prepared = prepare_graph(&scaled);
    let mut root = Element::new(format!("{{{CHART_NS}}}chartSpace"));
    root.append(Element::with_attributes(
        format!("{{{CHART_NS}}}lang"),
        [("val", &prepared.lang)],
    ));
    // The axes are already scaled, so the helper must not recompute them again.
    let _ = write_chart_node(&mut root, &prepared);
    write_print_settings(&mut root, &prepared.print_margins);
    if prepared.has_shapes {
        root.append(Element::with_attributes(
            format!("{{{CHART_NS}}}userShapes"),
            [(format!("{{{REL_NS}}}id").as_str(), "rId1")],
        ));
    }
    root.to_pretty_string()
}

/// Serialise whichever chart variant a sheet holds.
pub fn write_any_chart(
    chart: &Chart,
    is_graph: bool,
    axes: Option<&(Axis, Axis)>,
) -> Result<String> {
    let _ = is_graph;
    let mut prepared = prepare(chart);
    prepared.axes = axes.cloned();
    let mut root = Element::new(format!("{{{CHART_NS}}}chartSpace"));
    root.append(Element::with_attributes(
        format!("{{{CHART_NS}}}lang"),
        [("val", &prepared.lang)],
    ));
    write_chart_node(&mut root, &prepared)?;
    write_print_settings(&mut root, &prepared.print_margins);
    if prepared.has_shapes {
        root.append(Element::with_attributes(
            format!("{{{CHART_NS}}}userShapes"),
            [(format!("{{{REL_NS}}}id").as_str(), "rId1")],
        ));
    }
    Ok(root.to_pretty_string())
}

fn write_chart_node(root: &mut Element, prepared: &PreparedChart) -> Result<()> {
    let mut chart_node = Element::new(format!("{{{CHART_NS}}}chart"));
    if !prepared.title.is_empty() {
        write_title(&mut chart_node, &prepared.title, &prepared.lang);
    }

    let mut plot_area = Element::new(format!("{{{CHART_NS}}}plotArea"));
    let mut layout = Element::new(format!("{{{CHART_NS}}}layout"));
    let mut manual = Element::new(format!("{{{CHART_NS}}}manualLayout"));
    manual.append(Element::with_attributes(
        format!("{{{CHART_NS}}}layoutTarget"),
        [("val", "inner")],
    ));
    manual.append(Element::with_attributes(
        format!("{{{CHART_NS}}}xMode"),
        [("val", "edge")],
    ));
    manual.append(Element::with_attributes(
        format!("{{{CHART_NS}}}yMode"),
        [("val", "edge")],
    ));
    manual.append(Element::with_attributes(
        format!("{{{CHART_NS}}}x"),
        [("val", &safe_string(prepared.margin_left))],
    ));
    manual.append(Element::with_attributes(
        format!("{{{CHART_NS}}}y"),
        [("val", &safe_string(prepared.margin_top))],
    ));
    manual.append(Element::with_attributes(
        format!("{{{CHART_NS}}}w"),
        [("val", &safe_string(prepared.width))],
    ));
    manual.append(Element::with_attributes(
        format!("{{{CHART_NS}}}h"),
        [("val", &safe_string(prepared.height))],
    ));
    layout.append(manual);
    plot_area.append(layout);

    let mut subchart = Element::new(format!("{{{CHART_NS}}}{}", prepared.chart_type));
    write_options(&mut subchart, &prepared.chart_type, &prepared.grouping);
    write_series(&mut subchart, prepared)?;
    plot_area.append(subchart);

    if let Some((x_axis, y_axis)) = &prepared.axes {
        plot_area.append(Element::with_attributes(
            format!("{{{CHART_NS}}}axId"),
            [("val", &x_axis.id.to_string())],
        ));
        plot_area.append(Element::with_attributes(
            format!("{{{CHART_NS}}}axId"),
            [("val", &y_axis.id.to_string())],
        ));
        write_axis(
            &mut plot_area,
            x_axis,
            &prepared.lang,
            x_axis.axis_type.clone(),
        );
        write_axis(
            &mut plot_area,
            y_axis,
            &prepared.lang,
            y_axis.axis_type.clone(),
        );
    }
    chart_node.append(plot_area);

    if prepared.show_legend {
        let mut legend = Element::new(format!("{{{CHART_NS}}}legend"));
        legend.append(Element::with_attributes(
            format!("{{{CHART_NS}}}legendPos"),
            [("val", &prepared.legend_position)],
        ));
        legend.append(Element::new(format!("{{{CHART_NS}}}layout")));
        chart_node.append(legend);
    }
    chart_node.append(Element::with_attributes(
        format!("{{{CHART_NS}}}plotVisOnly"),
        [("val", "1")],
    ));
    root.append(chart_node);
    Ok(())
}

fn write_options(subchart: &mut Element, chart_type: &str, grouping: &str) {
    match chart_type {
        "pieChart" => {
            subchart.append(Element::with_attributes(
                format!("{{{CHART_NS}}}varyColors"),
                [("val", "1")],
            ));
        }
        "barChart" => {
            subchart.append(Element::with_attributes(
                format!("{{{CHART_NS}}}barDir"),
                [("val", "col")],
            ));
            subchart.append(Element::with_attributes(
                format!("{{{CHART_NS}}}grouping"),
                [("val", grouping)],
            ));
        }
        "lineChart" => {
            subchart.append(Element::with_attributes(
                format!("{{{CHART_NS}}}grouping"),
                [("val", grouping)],
            ));
        }
        "scatterChart" => {
            subchart.append(Element::with_attributes(
                format!("{{{CHART_NS}}}scatterStyle"),
                [("val", "lineMarker")],
            ));
        }
        _ => {}
    }
}

fn write_title(parent: &mut Element, title: &str, lang: &str) {
    let mut node = Element::new(format!("{{{CHART_NS}}}title"));
    let mut text = Element::new(format!("{{{CHART_NS}}}tx"));
    let mut rich = Element::new(format!("{{{CHART_NS}}}rich"));
    rich.append(Element::new(format!("{{{DRAWING_NS}}}bodyPr")));
    rich.append(Element::new(format!("{{{DRAWING_NS}}}lstStyle")));
    let mut paragraph = Element::new(format!("{{{DRAWING_NS}}}p"));
    let mut properties = Element::new(format!("{{{DRAWING_NS}}}pPr"));
    properties.append(Element::new(format!("{{{DRAWING_NS}}}defRPr")));
    paragraph.append(properties);
    let mut run = Element::new(format!("{{{DRAWING_NS}}}r"));
    run.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}rPr"),
        [("lang", lang)],
    ));
    let mut value = Element::new(format!("{{{DRAWING_NS}}}t"));
    value.set_text(title);
    run.append(value);
    paragraph.append(run);
    rich.append(paragraph);
    text.append(rich);
    node.append(text);
    node.append(Element::new(format!("{{{CHART_NS}}}layout")));
    parent.append(node);
}

fn write_axis(plot_area: &mut Element, axis: &Axis, lang: &str, label: String) {
    let mut node = Element::new(format!("{{{CHART_NS}}}{label}"));
    node.append(Element::with_attributes(
        format!("{{{CHART_NS}}}axId"),
        [("val", &axis.id.to_string())],
    ));
    let mut scaling = Element::new(format!("{{{CHART_NS}}}scaling"));
    scaling.append(Element::with_attributes(
        format!("{{{CHART_NS}}}orientation"),
        [("val", &axis.orientation)],
    ));
    if axis.delete_axis {
        scaling.append(Element::with_attributes(
            format!("{{{CHART_NS}}}delete"),
            [("val", "1")],
        ));
    }
    if axis.axis_type == "valAx" {
        // Excel validates these: a non-finite value would make the chart unreadable.
        scaling.append(Element::with_attributes(
            format!("{{{CHART_NS}}}max"),
            [("val", &axis.max().to_string())],
        ));
        scaling.append(Element::with_attributes(
            format!("{{{CHART_NS}}}min"),
            [("val", &axis.min().to_string())],
        ));
    }
    node.append(scaling);
    node.append(Element::with_attributes(
        format!("{{{CHART_NS}}}axPos"),
        [("val", &axis.position)],
    ));
    if axis.axis_type == "valAx" {
        node.append(Element::new(format!("{{{CHART_NS}}}majorGridlines")));
        node.append(Element::with_attributes(
            format!("{{{CHART_NS}}}numFmt"),
            [("formatCode", "General"), ("sourceLinked", "1")],
        ));
    }
    if !axis.title.is_empty() {
        write_title(&mut node, &axis.title, lang);
    }
    node.append(Element::with_attributes(
        format!("{{{CHART_NS}}}tickLblPos"),
        [("val", &axis.tick_label_position)],
    ));
    node.append(Element::with_attributes(
        format!("{{{CHART_NS}}}crossAx"),
        [("val", &axis.cross.to_string())],
    ));
    node.append(Element::with_attributes(
        format!("{{{CHART_NS}}}crosses"),
        [("val", &axis.crosses)],
    ));
    if axis.auto {
        node.append(Element::with_attributes(
            format!("{{{CHART_NS}}}auto"),
            [("val", "1")],
        ));
    }
    if let Some(align) = &axis.label_align {
        node.append(Element::with_attributes(
            format!("{{{CHART_NS}}}lblAlgn"),
            [("val", align)],
        ));
    }
    if let Some(offset) = axis.label_offset {
        if offset != 0 {
            node.append(Element::with_attributes(
                format!("{{{CHART_NS}}}lblOffset"),
                [("val", &offset.to_string())],
            ));
        }
    }
    if axis.axis_type == "valAx" {
        node.append(Element::with_attributes(
            format!("{{{CHART_NS}}}crossBetween"),
            [("val", &axis.cross_between)],
        ));
        node.append(Element::with_attributes(
            format!("{{{CHART_NS}}}majorUnit"),
            [("val", &axis.unit().to_string())],
        ));
    }
    plot_area.append(node);
}

fn write_series(subchart: &mut Element, prepared: &PreparedChart) -> Result<()> {
    let value_element = if prepared.chart_type == "scatterChart" {
        "yVal"
    } else {
        "val"
    };
    for (index, series) in prepared.series.iter().enumerate() {
        let mut node = Element::new(format!("{{{CHART_NS}}}ser"));
        node.append(Element::with_attributes(
            format!("{{{CHART_NS}}}idx"),
            [("val", &index.to_string())],
        ));
        node.append(Element::with_attributes(
            format!("{{{CHART_NS}}}order"),
            [("val", &index.to_string())],
        ));
        if let Some(title) = &series.title {
            let mut text = Element::new(format!("{{{CHART_NS}}}tx"));
            let mut value = Element::new(format!("{{{CHART_NS}}}v"));
            value.set_text(title);
            text.append(value);
            node.append(text);
        }
        if let Some(color) = series.color() {
            let mut properties = Element::new(format!("{{{CHART_NS}}}spPr"));
            write_series_color(&mut properties, color, &prepared.chart_type);
            node.append(properties);
        }
        if let Some(error_bar) = &series.error_bar {
            write_error_bar(&mut node, error_bar);
        }
        if let Some(labels) = &series.labels {
            let mut categories = Element::new(format!("{{{CHART_NS}}}cat"));
            write_serial(&mut categories, labels, true)?;
            node.append(categories);
        }
        if prepared.chart_type == "scatterChart" {
            if let Some(x_reference) = &series.x_reference {
                let mut x_values = Element::new(format!("{{{CHART_NS}}}xVal"));
                write_serial(&mut x_values, x_reference, true)?;
                node.append(x_values);
            }
        }
        let mut values = Element::new(format!("{{{CHART_NS}}}{value_element}"));
        match &series.reference {
            Some(reference) => write_serial(&mut values, reference, true)?,
            None => write_literal(&mut values, 1.0),
        }
        node.append(values);
        subchart.append(node);
    }
    Ok(())
}

fn write_series_color(node: &mut Element, color: &str, chart_type: &str) {
    // A bar series colours the whole mark, so it gets a solid fill; a line series only
    // colours the stroke.
    if chart_type == "barChart" {
        let mut fill = Element::new(format!("{{{DRAWING_NS}}}solidFill"));
        fill.append(Element::with_attributes(
            format!("{{{DRAWING_NS}}}srgbClr"),
            [("val", color)],
        ));
        node.append(fill);
    }
    let mut line = Element::new(format!("{{{DRAWING_NS}}}ln"));
    let mut fill = Element::new(format!("{{{DRAWING_NS}}}solidFill"));
    fill.append(Element::with_attributes(
        format!("{{{DRAWING_NS}}}srgbClr"),
        [("val", color)],
    ));
    line.append(fill);
    node.append(line);
}

fn write_error_bar(node: &mut Element, error_bar: &crate::charts::error_bar::ErrorBar) {
    let flag = match error_bar.bar_type {
        ErrorBarType::PlusMinus => "both",
        ErrorBarType::Plus => "plus",
        ErrorBarType::Minus => "minus",
    };
    let mut bars = Element::new(format!("{{{CHART_NS}}}errBars"));
    bars.append(Element::with_attributes(
        format!("{{{CHART_NS}}}errBarType"),
        [("val", flag)],
    ));
    bars.append(Element::with_attributes(
        format!("{{{CHART_NS}}}errValType"),
        [("val", "cust")],
    ));
    let mut plus = Element::new(format!("{{{CHART_NS}}}plus"));
    write_serial(&mut plus, &error_bar.reference, true).ok();
    bars.append(plus);
    let mut minus = Element::new(format!("{{{CHART_NS}}}minus"));
    write_serial(&mut minus, &error_bar.reference, true).ok();
    bars.append(minus);
    node.append(bars);
}

/// Write a `numRef`/`strRef` (or the literal fallback) for a reference.
fn write_serial(parent: &mut Element, reference: &Reference, has_values: bool) -> Result<()> {
    let data_type = reference.effective_data_type();
    let mut reference_node = Element::new(format!("{{{CHART_NS}}}{}", data_type.ref_element()));
    let mut formula = Element::new(format!("{{{CHART_NS}}}f"));
    formula.set_text(reference.to_reference_string());
    reference_node.append(formula);
    let mut cache = Element::new(format!("{{{CHART_NS}}}{}", data_type.cache_element()));
    write_cache(&mut cache, reference, data_type, has_values);
    reference_node.append(cache);
    parent.append(reference_node);
    Ok(())
}

fn write_cache(
    cache: &mut Element,
    reference: &Reference,
    data_type: ReferenceDataType,
    has_values: bool,
) {
    if data_type == ReferenceDataType::Numeric {
        let mut format = Element::new(format!("{{{CHART_NS}}}formatCode"));
        format.set_text(reference.effective_number_format());
        cache.append(format);
    }
    let values: &[CellValue] = reference.values().unwrap_or(&[]);
    let count = if has_values { values.len() } else { 0 };
    cache.append(Element::with_attributes(
        format!("{{{CHART_NS}}}ptCount"),
        [("val", &count.to_string())],
    ));
    for (index, value) in values.iter().enumerate() {
        let mut point =
            Element::with_attributes(format!("{{{CHART_NS}}}pt"), [("idx", &index.to_string())]);
        let mut text = Element::new(format!("{{{CHART_NS}}}v"));
        text.set_text(value.to_cache_string());
        point.append(text);
        cache.append(point);
    }
}

fn write_literal(parent: &mut Element, value: f64) {
    let mut literal = Element::new(format!("{{{CHART_NS}}}numLit"));
    let mut count = Element::with_attributes(format!("{{{CHART_NS}}}ptCount"), [("val", "1")]);
    let mut point = Element::with_attributes(format!("{{{CHART_NS}}}pt"), [("idx", "0")]);
    let mut text = Element::new(format!("{{{CHART_NS}}}v"));
    text.set_text(value.to_string());
    point.append(text);
    count.append(point);
    literal.append(count);
    parent.append(literal);
}

fn write_print_settings(root: &mut Element, margins: &[(String, String)]) {
    let mut settings = Element::new(format!("{{{CHART_NS}}}printSettings"));
    settings.append(Element::new(format!("{{{CHART_NS}}}headerFooter")));
    let mut page_margins = Element::new(format!("{{{CHART_NS}}}pageMargins"));
    for (name, value) in margins {
        page_margins.set(name.clone(), value.clone());
    }
    settings.append(page_margins);
    settings.append(Element::new(format!("{{{CHART_NS}}}pageSetup")));
    root.append(settings);
}

/// Serialise the chart's user-shape relationships.
pub fn write_chart_rels(drawing_id: u32) -> String {
    let mut root = Element::new(format!("{{{PKG_REL_NS}}}Relationships"));
    root.append(Element::with_attributes(
        format!("{{{PKG_REL_NS}}}Relationship"),
        [
            ("Id", "rId1"),
            ("Type", &format!("{REL_NS}/chartUserShapes")),
            ("Target", &format!("../drawings/drawing{drawing_id}.xml")),
        ],
    ));
    root.to_pretty_string()
}

/// Whether a chart type is a graph chart (and therefore has axes).
pub fn is_graph_chart(chart: &Chart) -> bool {
    matches!(chart.chart_type, "barChart" | "lineChart" | "scatterChart")
}

/// The axis pair for a graph chart, if the chart has one.
pub fn axes_for(chart: &GraphChart) -> (Axis, Axis) {
    (chart.x_axis.clone(), chart.y_axis.clone())
}

/// Serialise a pie chart.
pub fn write_pie_chart(chart: &PieChart) -> Result<String> {
    write_chart(&chart.0)
}

/// Serialise a bar chart.
pub fn write_bar_chart(chart: &BarChart) -> String {
    write_graph_chart(&chart.0)
}

/// Serialise a line chart.
pub fn write_line_chart(chart: &LineChart) -> String {
    write_graph_chart(&chart.0)
}

/// Serialise a scatter chart.
pub fn write_scatter_chart(chart: &ScatterChart) -> String {
    write_graph_chart(&chart.0)
}

/// Report a chart type the writer cannot serialise.
pub fn unsupported_chart_error(chart_type: &str) -> Error {
    Error::Value(format!("Don't know how to handle {chart_type}"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::series::Series;
    use crate::xml::functions::fromstring;

    fn reference(values: &[f64]) -> Reference {
        let mut reference = Reference::new("Sheet1", (0, 0), Some((2, 0)), None, None).unwrap();
        reference.set_values(
            values.iter().map(|v| CellValue::Number(*v)).collect(),
            ReferenceDataType::Numeric,
        );
        reference
    }

    #[test]
    fn pie_charts_have_no_axes() {
        let mut chart = PieChart::new();
        chart.add_series(Series::new(reference(&[1.0, 2.0, 3.0])));
        let xml = write_chart(&chart.0).unwrap();
        let root = fromstring(xml.as_bytes()).expect("chart must parse");
        let plot_area = root
            .find(format!("{{{CHART_NS}}}chart"))
            .and_then(|c| c.find(format!("{{{CHART_NS}}}plotArea")))
            .expect("plotArea");
        assert!(plot_area.find(format!("{{{CHART_NS}}}pieChart")).is_some());
        assert!(plot_area.find(format!("{{{CHART_NS}}}catAx")).is_none());
        assert!(plot_area.find(format!("{{{CHART_NS}}}valAx")).is_none());
        assert!(xml.contains("varyColors"));
    }

    #[test]
    fn bar_charts_declare_direction_and_grouping() {
        let mut chart = BarChart::new();
        chart.add_series(Series::new(reference(&[1.0, 2.0])));
        let xml = write_bar_chart(&chart);
        assert!(xml.contains("barDir"));
        assert!(xml.contains("grouping"));
        assert!(xml.contains("clustered"));
        assert!(xml.contains("catAx"));
        assert!(xml.contains("valAx"));
    }

    #[test]
    fn scatter_charts_use_yval_and_xval() {
        let mut chart = ScatterChart::new();
        chart.add_series(Series::new(reference(&[1.0, 2.0])).with_xvalues(reference(&[3.0, 4.0])));
        let xml = write_scatter_chart(&chart);
        assert!(xml.contains("yVal"));
        assert!(xml.contains("xVal"));
        assert!(xml.contains("scatterStyle"));
        // Scatter charts treat the x axis as a value axis, so both carry bounds.
        assert_eq!(
            xml.matches("c:min").count(),
            2,
            "both axes scale as value axes"
        );
    }

    #[test]
    fn series_carry_indices_titles_and_values() {
        let mut chart = BarChart::new();
        chart.add_series(Series::new(reference(&[1.0, 2.0])).with_title("Revenue"));
        let xml = write_bar_chart(&chart);
        let root = fromstring(xml.as_bytes()).unwrap();
        let plot_area = root
            .find(format!("{{{CHART_NS}}}chart"))
            .and_then(|c| c.find(format!("{{{CHART_NS}}}plotArea")))
            .unwrap();
        let subchart = plot_area.find(format!("{{{CHART_NS}}}barChart")).unwrap();
        let series = &subchart.find_all(format!("{{{CHART_NS}}}ser"))[0];
        // `c:idx` and `c:order` carry their value in a `val` attribute.
        let idx = series.find(format!("{{{CHART_NS}}}idx")).unwrap();
        assert_eq!(idx.get("val"), Some("0"));
        let title = series
            .find(format!("{{{CHART_NS}}}tx"))
            .and_then(|tx| tx.find(format!("{{{CHART_NS}}}v")))
            .and_then(|v| v.text.clone());
        assert_eq!(title.as_deref(), Some("Revenue"));
        let values = series.find(format!("{{{CHART_NS}}}val")).unwrap();
        let reference_node = values.find(format!("{{{CHART_NS}}}numRef")).unwrap();
        assert!(reference_node
            .find_text(format!("{{{CHART_NS}}}f"), "")
            .starts_with("'Sheet1'!"));
        assert_eq!(
            reference_node
                .find(format!("{{{CHART_NS}}}numCache"))
                .and_then(|c| c.find(format!("{{{CHART_NS}}}ptCount")))
                .and_then(|p| p.get("val"))
                .unwrap_or(""),
            "2"
        );
    }

    #[test]
    fn axis_bounds_are_computed_from_the_data() {
        let mut chart = BarChart::new();
        chart.add_series(Series::new(reference(&[1.0, 5.0, 9.0])));
        let xml = write_bar_chart(&chart);
        let root = fromstring(xml.as_bytes()).unwrap();
        let plot_area = root
            .find(format!("{{{CHART_NS}}}chart"))
            .and_then(|c| c.find(format!("{{{CHART_NS}}}plotArea")))
            .unwrap();
        let value_axis = plot_area.find(format!("{{{CHART_NS}}}valAx")).unwrap();
        let scaling = value_axis.find(format!("{{{CHART_NS}}}scaling")).unwrap();
        let max: f64 = scaling
            .find(format!("{{{CHART_NS}}}max"))
            .and_then(|n| n.get("val"))
            .and_then(|v| v.parse().ok())
            .unwrap();
        let min: f64 = scaling
            .find(format!("{{{CHART_NS}}}min"))
            .and_then(|n| n.get("val"))
            .and_then(|v| v.parse().ok())
            .unwrap();
        assert!(max >= 9.0, "max {max} must cover the data");
        assert!(min <= 1.0, "min {min} must cover the data");
    }

    #[test]
    fn flat_series_do_not_produce_invalid_bounds() {
        // A zero-width range would divide by zero in the Python implementation.
        let mut chart = BarChart::new();
        chart.add_series(Series::new(reference(&[7.0, 7.0, 7.0])));
        let xml = write_bar_chart(&chart);
        assert!(!xml.contains("NaN"), "NaN would make the chart unreadable");
        assert!(!xml.contains("inf"));
    }

    #[test]
    fn legends_are_written_only_when_enabled() {
        let mut chart = BarChart::new();
        chart.add_series(Series::new(reference(&[1.0])));
        assert!(write_bar_chart(&chart).contains("legendPos"));

        let mut hidden = BarChart::new();
        hidden.0.base.show_legend = false;
        hidden.add_series(Series::new(reference(&[1.0])));
        assert!(!write_bar_chart(&hidden).contains("legendPos"));
    }

    #[test]
    fn titles_are_written_when_set() {
        let mut chart = BarChart::new();
        chart.0.base.title = "Quarterly revenue".to_string();
        chart.add_series(Series::new(reference(&[1.0])));
        let xml = write_bar_chart(&chart);
        assert!(xml.contains("Quarterly revenue"));
        assert!(xml.contains("layoutTarget"));
    }

    #[test]
    fn series_colours_are_short_and_applied() {
        let mut chart = BarChart::new();
        chart.add_series(Series::new(reference(&[1.0])).with_color("FF3366FF"));
        let xml = write_bar_chart(&chart);
        assert!(
            xml.contains("val=\"3366FF\""),
            "the alpha prefix must be stripped"
        );
        // A bar series gets both a solid fill and a stroke.
        assert!(xml.contains("solidFill"));
    }

    #[test]
    fn error_bars_use_the_right_flags() {
        let mut chart = BarChart::new();
        chart.add_series(Series::new(reference(&[1.0, 2.0])).with_error_bar(
            crate::charts::error_bar::ErrorBar::new(
                ErrorBarType::PlusMinus,
                reference(&[0.5, 0.5]),
            ),
        ));
        let xml = write_bar_chart(&chart);
        assert!(xml.contains("errBarType"));
        assert!(xml.contains("val=\"both\""));
    }

    #[test]
    fn print_settings_are_always_written() {
        let mut chart = BarChart::new();
        chart.add_series(Series::new(reference(&[1.0])));
        let xml = write_bar_chart(&chart);
        assert!(xml.contains("printSettings"));
        assert!(xml.contains("pageMargins"));
    }

    #[test]
    fn chart_rels_point_at_the_shape_drawing() {
        let xml = write_chart_rels(3);
        let root = fromstring(xml.as_bytes()).unwrap();
        let relationship = &root.find_all(format!("{{{PKG_REL_NS}}}Relationship"))[0];
        assert_eq!(relationship.get("Target"), Some("../drawings/drawing3.xml"));
        assert_eq!(relationship.get("Id"), Some("rId1"));
    }

    #[test]
    fn shapes_add_a_user_shapes_element() {
        let mut chart = BarChart::new();
        chart.add_series(Series::new(reference(&[1.0])));
        chart.0.base.add_shape(crate::drawing::Shape::new());
        let xml = write_bar_chart(&chart);
        assert!(xml.contains("userShapes"));
    }

    #[test]
    fn unresolved_references_have_zero_points() {
        let mut chart = BarChart::new();
        // A reference whose values were never resolved.
        chart.add_series(Series::new(
            Reference::new("Sheet1", (0, 0), None, None, None).unwrap(),
        ));
        let xml = write_bar_chart(&chart);
        let root = fromstring(xml.as_bytes()).unwrap();
        let plot_area = root
            .find(format!("{{{CHART_NS}}}chart"))
            .and_then(|c| c.find(format!("{{{CHART_NS}}}plotArea")))
            .unwrap();
        let subchart = plot_area.find(format!("{{{CHART_NS}}}barChart")).unwrap();
        let series = &subchart.find_all(format!("{{{CHART_NS}}}ser"))[0];
        let values = series.find(format!("{{{CHART_NS}}}val")).unwrap();
        let cache = values
            .find(format!("{{{CHART_NS}}}numRef"))
            .and_then(|r| r.find(format!("{{{CHART_NS}}}numCache")))
            .unwrap();
        assert_eq!(
            cache
                .find(format!("{{{CHART_NS}}}ptCount"))
                .unwrap()
                .get("val"),
            Some("0")
        );
    }

    #[test]
    fn type_helpers() {
        assert!(is_graph_chart(&BarChart::new().0.base));
        assert!(!is_graph_chart(&PieChart::new().0));
        let bar = BarChart::new();
        let (x, y) = axes_for(&bar.0);
        assert_eq!(x.axis_type, "catAx");
        assert_eq!(y.axis_type, "valAx");
        assert!(unsupported_chart_error("bubbleChart")
            .to_string()
            .contains("bubbleChart"));
    }
}
