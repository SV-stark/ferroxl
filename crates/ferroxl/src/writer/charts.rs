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
    options: crate::charts::chart::ChartOptions,
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
        options: chart.options.clone(),
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
    write_view_3d(&mut root, &prepared);
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
    write_view_3d(&mut root, &prepared);
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
    write_view_3d(&mut root, &prepared);
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
    write_options(&mut subchart, prepared);
    write_series(&mut subchart, prepared)?;
    write_options_after_series(&mut subchart, prepared);
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

/// Write the type-specific option elements that precede the series.
///
/// Element order is fixed by the schema, not by taste: `grouping` before `varyColors`,
/// `barDir` first of all, and the trailing elements (`holeSize`, `bubbleScale`) after the
/// series, which this function cannot reach — [`write_options_after_series`] does.
fn write_options(subchart: &mut Element, prepared: &PreparedChart) {
    let options = &prepared.options;
    let flag =
        |name: &str| Element::with_attributes(format!("{{{CHART_NS}}}{name}"), [("val", "1")]);
    let value = |name: &str, val: &str| {
        Element::with_attributes(format!("{{{CHART_NS}}}{name}"), [("val", val)])
    };

    match prepared.chart_type.as_str() {
        // `varyColors` gives every slice a different colour, which is what makes a pie
        // readable at all. Excel writes it for every pie-family chart.
        "pieChart" | "pie3DChart" | "doughnutChart" | "ofPieChart" => {
            subchart.append(flag("varyColors"));
        }
        "barChart" | "bar3DChart" => {
            subchart.append(value("barDir", "col"));
            subchart.append(value("grouping", &prepared.grouping));
        }
        "lineChart" | "line3DChart" | "areaChart" | "area3DChart" => {
            subchart.append(value("grouping", &prepared.grouping));
        }
        "scatterChart" => {
            subchart.append(value("scatterStyle", "lineMarker"));
        }
        "radarChart" => {
            subchart.append(value("radarStyle", options.radar_style));
            subchart.append(flag("varyColors"));
        }
        // A bubble chart varies colours too: the bubbles are the categories, not the
        // series, so a single colour per series would hide them.
        "bubbleChart" => {
            subchart.append(flag("varyColors"));
        }
        // A surface chart is a wireframe unless asked otherwise, and `wireframe` comes
        // before the series.
        "surfaceChart" | "surface3DChart" if options.wireframe => {
            subchart.append(flag("wireframe"));
        }
        // A stock chart has no options before its series: openpyxl's `__elements__` starts
        // with `ser`.
        _ => {}
    }

    // A projected pie names its second plot after the `varyColors` element.
    if prepared.chart_type == "ofPieChart" {
        subchart.append(value("ofPieType", options.of_pie_type));
    }
}

/// The option elements that follow the series.
///
/// Split from [`write_options`] because the schema puts them on the other side of `ser`, and
/// a function that can only append to one element cannot express that.
fn write_options_after_series(subchart: &mut Element, prepared: &PreparedChart) {
    let options = &prepared.options;
    let value = |name: &str, val: &str| {
        Element::with_attributes(format!("{{{CHART_NS}}}{name}"), [("val", val)])
    };
    let flag =
        |name: &str| Element::with_attributes(format!("{{{CHART_NS}}}{name}"), [("val", "1")]);

    match prepared.chart_type.as_str() {
        "bubbleChart" => {
            if options.bubble_3d {
                subchart.append(flag("bubble3D"));
            }
            if let Some(scale) = options.bubble_scale {
                subchart.append(value("bubbleScale", &scale.to_string()));
            }
            if options.show_negative_bubbles {
                subchart.append(flag("showNegBubbles"));
            }
            subchart.append(value("sizeRepresents", options.size_represents));
        }
        "doughnutChart" => {
            if let Some(angle) = options.first_slice_angle {
                subchart.append(value("firstSliceAng", &angle.to_string()));
            }
            // Excel's default hole is 10%; openpyxl's is the same. Writing it explicitly
            // keeps a round trip from shifting the ring's thickness.
            subchart.append(value(
                "holeSize",
                &options.hole_size.unwrap_or(10).to_string(),
            ));
        }
        "pie3DChart" | "ofPieChart" => {
            if let Some(angle) = options.first_slice_angle {
                subchart.append(value("firstSliceAng", &angle.to_string()));
            }
        }
        _ => {}
    }
}

/// Write `<c:view3D>` for a 3-D chart.
///
/// The element is a sibling of `<c:chart>` inside `<c:chartSpace>` and comes *before* it,
/// which is why this is a separate pass rather than part of `write_options`.
fn write_view_3d(root: &mut Element, prepared: &PreparedChart) {
    let Some(view) = &prepared.options.view_3d else {
        return;
    };
    if !is_three_dimensional(&prepared.chart_type) {
        // A 2-D chart has nowhere to put this. Ignoring it is deliberate: the caller may
        // have reused an options struct, and refusing would fail a call that is otherwise
        // perfectly valid.
        return;
    }
    let mut node = Element::new(format!("{{{CHART_NS}}}view3D"));
    node.append(Element::with_attributes(
        format!("{{{CHART_NS}}}rotX"),
        [("val", view.rot_x.to_string())],
    ));
    node.append(Element::with_attributes(
        format!("{{{CHART_NS}}}rotY"),
        [("val", view.rot_y.to_string())],
    ));
    node.append(Element::with_attributes(
        format!("{{{CHART_NS}}}depthPercent"),
        [("val", view.depth_percent.to_string())],
    ));
    if view.right_angle_axes {
        node.append(Element::with_attributes(
            format!("{{{CHART_NS}}}rAngAx"),
            [("val", "1".to_string())],
        ));
    }
    root.append(node);
}

/// Whether a chart tag denotes a 3-D chart.
fn is_three_dimensional(chart_type: &str) -> bool {
    matches!(
        chart_type,
        "area3DChart"
            | "bar3DChart"
            | "line3DChart"
            | "pie3DChart"
            | "surface3DChart"
            | "ofPieChart"
    )
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
    // Scatter and bubble charts name their values `yVal`; everything else says `val`.
    let value_element = if matches!(prepared.chart_type.as_str(), "scatterChart" | "bubbleChart") {
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
        if matches!(prepared.chart_type.as_str(), "scatterChart" | "bubbleChart") {
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
        // A bubble's size is a third series of numbers. It goes in `<c:bubbleSize>`, after
        // the y values; openpyxl calls the same thing `zVal` on the series it builds.
        if prepared.chart_type == "bubbleChart" {
            if let Some(size) = &series.bubble_size {
                let mut size_node = Element::new(format!("{{{CHART_NS}}}bubbleSize"));
                write_serial(&mut size_node, size, true)?;
                node.append(size_node);
            }
        }
        subchart.append(node);
    }
    Ok(())
}

fn write_series_color(node: &mut Element, color: &str, chart_type: &str) {
    // A bar series colours the whole mark, so it gets a solid fill; a line series only
    // colours the stroke.
    if matches!(
        chart_type,
        "barChart" | "bar3DChart" | "areaChart" | "area3DChart" | "surface3DChart"
    ) {
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
    // Everything with an `axId` in its `__elements__`: every chart type except the
    // pie family, which has no axes.
    !matches!(
        chart.chart_type,
        "pieChart" | "pie3DChart" | "doughnutChart" | "ofPieChart"
    )
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
    use crate::charts::chart::{AreaChart, AreaChart3D, BarChart3D, LineChart3D, PieChart3D};
    use crate::charts::{
        BubbleChart, ChartOptions, DoughnutChart, ProjectedPieChart, RadarChart, StockChart,
        SurfaceChart, SurfaceChart3D, View3D,
    };

    /// One numeric series, for the tests that need a `<c:ser>` to order against.
    fn a_series() -> Series {
        let values: Vec<CellValue> = vec![CellValue::Number(1.0), CellValue::Number(2.0)];
        let mut reference = Reference::new("S", (0, 0), None, None, None).unwrap();
        reference.set_values(values, ReferenceDataType::Numeric);
        Series::new(reference)
    }

    #[test]
    fn every_chart_type_writes_its_own_tag() {
        // One test across all of them because the failure mode is the same for each: a
        // mistyped tag writes a file Excel refuses to open, and nothing else would catch it.
        let cases: Vec<(&str, String)> = vec![
            (
                "areaChart",
                write_chart(&AreaChart::new().into_chart()).unwrap(),
            ),
            (
                "area3DChart",
                write_chart(&AreaChart3D::new().into_chart()).unwrap(),
            ),
            (
                "barChart",
                write_chart(&BarChart::new().into_chart()).unwrap(),
            ),
            (
                "bar3DChart",
                write_chart(&BarChart3D::new().into_chart()).unwrap(),
            ),
            (
                "lineChart",
                write_chart(&LineChart::new().into_chart()).unwrap(),
            ),
            (
                "line3DChart",
                write_chart(&LineChart3D::new().into_chart()).unwrap(),
            ),
            (
                "pieChart",
                write_chart(&PieChart::new().into_chart()).unwrap(),
            ),
            (
                "pie3DChart",
                write_chart(&PieChart3D::new().into_chart()).unwrap(),
            ),
            (
                "doughnutChart",
                write_chart(&DoughnutChart::new().into_chart()).unwrap(),
            ),
            (
                "ofPieChart",
                write_chart(&ProjectedPieChart::new().into_chart()).unwrap(),
            ),
            (
                "scatterChart",
                write_chart(&ScatterChart::new().into_chart()).unwrap(),
            ),
            (
                "radarChart",
                write_chart(&RadarChart::new().into_chart()).unwrap(),
            ),
            (
                "bubbleChart",
                write_chart(&BubbleChart::new().into_chart()).unwrap(),
            ),
            (
                "stockChart",
                write_chart(&StockChart::new().into_chart()).unwrap(),
            ),
            (
                "surfaceChart",
                write_chart(&SurfaceChart::new().into_chart()).unwrap(),
            ),
            (
                "surface3DChart",
                write_chart(&SurfaceChart3D::new().into_chart()).unwrap(),
            ),
        ];
        for (tag, xml) in cases {
            // An element serialises three ways: `<c:x/>` when it has no children (a stock
            // chart with no series), `<c:x>` when it has some, and `<c:x attr="v">` when it
            // has attributes. Matching the prefix alone would also match `stockChartExtra`,
            // so the character after the name is checked.
            let needle = format!("<c:{tag}");
            let at = xml
                .find(&needle)
                .unwrap_or_else(|| panic!("{tag} is missing from:\n{xml}"));
            let next = xml[at + needle.len()..].chars().next().unwrap_or('\n');
            assert!(
                matches!(next, '>' | '/' | ' '),
                "{tag} is followed by {next:?}, which means the tag name is wrong:\n{xml}"
            );
            // The root carries the namespace declaration, so it is `<c:chartSpace
            // xmlns:c="...">` and a bare prefix match would be wrong.
            assert!(xml.starts_with("<c:chartSpace xmlns:c="), "{tag}: {xml}");
        }
    }

    #[test]
    fn a_radar_chart_writes_its_style_before_the_series() {
        let mut radar =
            RadarChart::with_options(ChartOptions::default().with_radar_style("marker"));
        radar.add_series(a_series());
        let chart = radar.into_chart();
        let xml = write_chart(&chart).unwrap();
        let radar = xml.find("<c:radarStyle").expect("radarStyle");
        let series = xml.find("<c:ser>").expect("a series");
        assert!(radar < series, "the schema puts radarStyle first:\n{xml}");
        assert!(xml.contains("<c:radarStyle val=\"marker\"/>"), "{xml}");
    }

    #[test]
    fn a_doughnut_hole_size_comes_after_the_series() {
        let mut doughnut = DoughnutChart::with_options(ChartOptions::default().with_hole_size(40));
        doughnut.add_series(a_series());
        let chart = doughnut.into_chart();
        let xml = write_chart(&chart).unwrap();
        let series = xml.find("<c:ser>").expect("a series");
        let hole = xml.find("<c:holeSize").expect("holeSize");
        assert!(hole > series, "the schema puts holeSize after ser:\n{xml}");
        assert!(xml.contains("<c:holeSize val=\"40\"/>"), "{xml}");
    }

    #[test]
    fn a_doughnut_defaults_to_excels_own_hole_size() {
        // Without this a round trip would shift the ring's thickness, because openpyxl
        // writes 10 when the field is unset and Excel would infer its own default.
        let xml = write_chart(&DoughnutChart::new().into_chart()).unwrap();
        assert!(xml.contains("<c:holeSize val=\"10\"/>"), "{xml}");
    }

    #[test]
    fn a_bubble_chart_writes_x_values_and_a_size() {
        let values: Vec<CellValue> = vec![CellValue::Number(1.0), CellValue::Number(2.0)];
        let x_values: Vec<CellValue> = vec![CellValue::Number(10.0), CellValue::Number(20.0)];
        let sizes: Vec<CellValue> = vec![CellValue::Number(5.0), CellValue::Number(8.0)];

        let mut reference = Reference::new("S", (0, 0), None, None, None).unwrap();
        reference.set_values(values, ReferenceDataType::Numeric);
        let mut x = Reference::new("S", (1, 0), None, None, None).unwrap();
        x.set_values(x_values, ReferenceDataType::Numeric);
        let mut size = Reference::new("S", (2, 0), None, None, None).unwrap();
        size.set_values(sizes, ReferenceDataType::Numeric);

        let mut chart = BubbleChart::new();
        chart.add_series(
            Series::new(reference)
                .with_xvalues(x)
                .with_bubble_size(size),
        );
        let xml = write_chart(&chart.into_chart()).unwrap();

        assert!(xml.contains("<c:xVal>"), "a bubble takes x values:\n{xml}");
        // `yVal` rather than `val`, as for a scatter chart.
        assert!(xml.contains("<c:yVal>"), "{xml}");
        assert!(xml.contains("<c:bubbleSize>"), "{xml}");
        let y = xml.find("<c:yVal>").expect("yVal");
        let bubble = xml.find("<c:bubbleSize>").expect("bubbleSize");
        assert!(y < bubble, "bubbleSize follows yVal:\n{xml}");
    }

    #[test]
    fn bubble_options_follow_the_series() {
        let mut bubble = BubbleChart::with_options(
            ChartOptions::default()
                .three_d_bubbles()
                .with_bubble_scale(150)
                .with_negative_bubbles(),
        );
        bubble.add_series(a_series());
        let chart = bubble.into_chart();
        let xml = write_chart(&chart).unwrap();
        let series = xml.find("<c:ser>").expect("a series");
        for element in [
            "bubble3D",
            "bubbleScale",
            "showNegBubbles",
            "sizeRepresents",
        ] {
            let at = xml
                .find(&format!("<c:{element}"))
                .unwrap_or_else(|| panic!("{element} is missing:\n{xml}"));
            assert!(at > series, "{element} must follow ser:\n{xml}");
        }
    }

    #[test]
    fn a_three_d_chart_writes_a_view() {
        let chart = BarChart3D::with_options(
            ChartOptions::default()
                .with_view_3d(View3D::new().with_rot_x(30).with_depth_percent(80)),
        )
        .into_chart();
        let xml = write_chart(&chart).unwrap();
        assert!(xml.contains("<c:view3D>"), "{xml}");
        assert!(xml.contains("<c:rotX val=\"30\"/>"), "{xml}");
        assert!(xml.contains("<c:depthPercent val=\"80\"/>"), "{xml}");
        // The schema puts `view3D` before `chart`.
        let view = xml.find("<c:view3D>").expect("view3D");
        let chart_node = xml.find("<c:chart>").expect("chart");
        assert!(view < chart_node, "view3D comes first:\n{xml}");
    }

    #[test]
    fn a_two_d_chart_ignores_a_view_rather_than_emitting_an_orphan() {
        // A 2-D chart has nowhere to put `view3D`, and a caller may have reused an options
        // struct. Writing it anyway would produce a file Excel rejects.
        let chart = BarChart::with_options(ChartOptions::default().with_view_3d(View3D::new()))
            .into_chart();
        let xml = write_chart(&chart).unwrap();
        assert!(
            !xml.contains("view3D"),
            "a barChart must not carry one:\n{xml}"
        );
    }

    #[test]
    fn a_surface_chart_is_a_wireframe_only_when_asked() {
        let plain = write_chart(&SurfaceChart::new().into_chart()).unwrap();
        assert!(!plain.contains("wireframe"), "{plain}");

        let mesh = SurfaceChart::with_options(ChartOptions::default().wireframe()).into_chart();
        assert!(write_chart(&mesh)
            .unwrap()
            .contains("<c:wireframe val=\"1\"/>"));
    }

    #[test]
    fn the_pie_family_is_the_only_part_without_axes() {
        for chart in [
            PieChart::new().into_chart(),
            PieChart3D::new().into_chart(),
            DoughnutChart::new().into_chart(),
            ProjectedPieChart::new().into_chart(),
        ] {
            assert!(
                !is_graph_chart(&chart),
                "{} should have no axes",
                chart.chart_type
            );
        }
        for chart in [
            BarChart::new().into_chart(),
            AreaChart::new().into_chart(),
            RadarChart::new().into_chart(),
            BubbleChart::new().into_chart(),
            StockChart::new().into_chart(),
            ScatterChart::new().into_chart(),
            SurfaceChart::new().into_chart(),
        ] {
            assert!(
                is_graph_chart(&chart),
                "{} should have axes",
                chart.chart_type
            );
        }
    }

    #[test]
    fn a_projected_pie_names_its_second_plot() {
        let chart =
            ProjectedPieChart::with_options(ChartOptions::default().with_of_pie_type("bar"))
                .into_chart();
        let xml = write_chart(&chart).unwrap();
        assert!(xml.contains("<c:ofPieType val=\"bar\"/>"), "{xml}");
    }
}
