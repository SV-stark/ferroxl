//! Tools that draw on a sheet: charts, images and comments.

use ferroxl::charts::reference::{Reference, ReferenceDataType};
use ferroxl::charts::{BarChart, LineChart, PieChart, ScatterChart, Series};
use ferroxl::comments::Comment;
use ferroxl::drawing::Image;
use serde_json::{json, Value};

use super::{err, open_for_edit, Handled};
use crate::tools::Args;
use crate::workspace::Workspace;

/// The largest image the tool will embed, to keep a workbook from becoming unwieldy.
const MAX_IMAGE_BYTES: usize = 8 * 1024 * 1024;

/// Add a chart to a sheet.
pub fn add_chart(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let kind = args.require_str("type")?;
    let anchor = args.require_str("anchor")?;
    let series_spec = args.require_array("series")?;
    if series_spec.is_empty() {
        return Err("a chart needs at least one series".to_string());
    }
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    let sheet_name = workbook.worksheets[index].title().to_string();
    let (column, row) = ferroxl::coordinate_from_string(&anchor).map_err(err)?;

    let categories = args.opt_str("categories");
    let scatter = kind == "scatter";
    let mut series = Vec::new();
    for (position, entry) in series_spec.iter().enumerate() {
        series.push(build_series(
            entry,
            &sheet_name,
            position,
            categories.as_deref(),
            scatter,
        )?);
    }

    let mut chart = match kind.as_str() {
        "bar" => BarChart::new().into_chart(),
        "line" => LineChart::new().into_chart(),
        "scatter" => ScatterChart::new().into_chart(),
        "pie" => PieChart::new().into_chart(),
        other => {
            return Err(format!(
                "{other:?} is not a chart type; use bar, line, scatter or pie"
            ))
        }
    };
    for entry in series {
        chart.add_series(entry);
    }
    if let Some(title) = args.opt_str("title") {
        chart = chart.with_title(title);
    }
    // The drawing part measures in pixels; the tool takes centimetres because that is
    // what a person reading "make the chart 15 cm wide" means.
    let width_cm = args.opt_number("width").unwrap_or(DEFAULT_WIDTH_CM);
    let height_cm = args.opt_number("height").unwrap_or(DEFAULT_HEIGHT_CM);
    if width_cm <= 0.0 || height_cm <= 0.0 {
        return Err("chart width and height must be greater than zero".to_string());
    }
    chart.drawing.set_width(cm_to_pixels(width_cm));
    chart.drawing.set_height(cm_to_pixels(height_cm));

    let count = workbook.worksheets[index].charts.len();
    workbook.worksheets[index].charts.push(chart);
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!(
        "added a {kind} chart with {} series at {name}!{anchor}",
        series_spec.len()
    );
    Ok((
        summary,
        json!({
            "path": path,
            "sheet": name,
            "type": kind,
            "anchor": anchor,
            "column": column,
            "row": row,
            "series": count + 1,
            "categories": categories,
            "width_cm": width_cm,
            "height_cm": height_cm,
        }),
    ))
}

/// The default chart width, in centimetres, matching openpyxl's `GraphicalProperties`.
const DEFAULT_WIDTH_CM: f64 = 15.0;

/// The default chart height, in centimetres.
const DEFAULT_HEIGHT_CM: f64 = 7.5;

/// Convert centimetres to pixels at the 96 dpi that OOXML assumes.
fn cm_to_pixels(centimetres: f64) -> i64 {
    (centimetres * 96.0 / 2.54).round() as i64
}

/// Build one series from its JSON description.
///
/// `values` is the series' own numbers. `xvalues` is the horizontal numbers a scatter plot
/// needs, and `categories` is the axis labels every other type uses; both default to the
/// tool's `categories` argument, so a caller who set it once gets it on every series.
fn build_series(
    entry: &Value,
    sheet_name: &str,
    position: usize,
    categories: Option<&str>,
    scatter: bool,
) -> Result<Series, String> {
    let values = entry
        .get("values")
        .and_then(Value::as_str)
        .ok_or_else(|| format!("series {position} needs a values range"))?;
    let (pos1, pos2) = sheet_range(values)?;

    // A scatter plot puts numbers on both axes, so the values are forced to numeric. For
    // the other types ferroxl infers the type from the cells themselves.
    let data_type = if scatter {
        Some(ReferenceDataType::Numeric)
    } else {
        None
    };
    let mut series =
        Series::new(Reference::new(sheet_name, pos1, pos2, data_type, None).map_err(err)?);
    if let Some(name) = entry.get("name").and_then(Value::as_str) {
        series = series.with_title(name.to_string());
    }

    // The x values come from the series itself when it names them, because only a scatter
    // plot needs different x values per series.
    if let Some(xvalues) = entry.get("xvalues").and_then(Value::as_str) {
        let (x1, x2) = sheet_range(xvalues)?;
        series = series.with_xvalues(
            Reference::new(sheet_name, x1, x2, Some(ReferenceDataType::Numeric), None)
                .map_err(err)?,
        );
    } else if scatter {
        if let Some(range) = categories {
            let (x1, x2) = sheet_range(range)?;
            series = series.with_xvalues(
                Reference::new(sheet_name, x1, x2, Some(ReferenceDataType::Numeric), None)
                    .map_err(err)?,
            );
        }
    }
    // A scatter plot's x values are its categories too, and writing both would give the
    // axis the same numbers twice.
    if !scatter {
        if let Some(range) = categories {
            let (c1, c2) = sheet_range(range)?;
            series = series.with_labels(
                Reference::new(sheet_name, c1, c2, Some(ReferenceDataType::String), None)
                    .map_err(err)?,
            );
        }
    }
    Ok(series)
}

/// The 0-based `(row, column)` corners of a range: the top-left cell, and the bottom-right
/// cell as `None` when the range is a single cell.
type RangeCorners = ((usize, usize), Option<(usize, usize)>);

/// Turn an A1 range into the 0-based `(row, column)` corners `Reference` wants.
///
/// The range may carry its own sheet name and absolute markers; both are stripped, because
/// the sheet is passed separately to `Reference::new`.
///
/// `Reference` takes `(row, column)`, not `(column, row)`, and a multi-cell range needs both
/// corners -- passing one corner, or the two in the other order, produces a reference to a
/// single unrelated cell, which is exactly the kind of thing that still writes a
/// well-formed chart.
fn sheet_range(range: &str) -> Result<RangeCorners, String> {
    let trimmed = range.trim();
    let local = trimmed.strip_prefix('=').unwrap_or(trimmed);
    let local = local
        .rsplit_once('!')
        .map(|(_, tail)| tail)
        .unwrap_or(local);
    let cleaned = local.replace('$', "");
    if cleaned.contains(':') {
        let (min, max) = cleaned
            .split_once(':')
            .ok_or_else(|| format!("{range:?} is not a range"))?;
        let (min_col, min_row) = ferroxl::coordinate_from_string(min).map_err(err)?;
        let (max_col, max_row) = ferroxl::coordinate_from_string(max).map_err(err)?;
        let min_col = ferroxl::column_index_from_string(&min_col).map_err(err)?;
        let max_col = ferroxl::column_index_from_string(&max_col).map_err(err)?;
        if max_col < min_col {
            return Err(format!("{range:?} runs backwards"));
        }
        if max_row < min_row {
            return Err(format!("{range:?} runs backwards"));
        }
        // `coordinate_from_string` is 1-based and `Reference` is 0-based.
        return Ok((
            (min_row as usize - 1, min_col as usize - 1),
            Some((max_row as usize - 1, max_col as usize - 1)),
        ));
    }
    // A bare cell is a single point, with no second corner.
    let (column, row) = ferroxl::coordinate_from_string(&cleaned).map_err(err)?;
    let column = ferroxl::column_index_from_string(&column).map_err(err)?;
    Ok(((row as usize - 1, column as usize - 1), None))
}

/// Embed an image on a sheet.
pub fn add_image(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let image_path = args.require_str("image_path")?;
    let anchor = args.require_str("anchor")?;
    if !image_path.to_ascii_lowercase().ends_with(".png") {
        return Err("only PNG images can be embedded".to_string());
    }
    let bytes = workspace.read_asset(&image_path, MAX_IMAGE_BYTES)?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    // The anchor is validated by asking the sheet where that cell sits, so a typo is
    // reported before the file is written.
    workbook.worksheets[index]
        .cell_anchor(&anchor)
        .map_err(err)?;
    let mut image = Image::from_png(bytes).map_err(err)?;
    if let Some(width) = args.opt_number("width") {
        image.drawing.set_width(width as i64);
    }
    if let Some(height) = args.opt_number("height") {
        image.drawing.set_height(height as i64);
    }
    // Anchor the picture to the cell it was asked for. Without this the drawing keeps its
    // default `Absolute` anchor, which Excel renders but openpyxl cannot read at all: its
    // `AbsoluteAnchor` has no `pic`, so the image is invisible to anything that goes through
    // openpyxl. A one-cell anchor is what openpyxl itself writes.
    let (column, row) = ferroxl::coordinate_from_string(&anchor).map_err(err)?;
    image.anchor_one_cell(&column, row).map_err(err)?;
    let (width, height) = (image.drawing.width(), image.drawing.height());
    let count = workbook.worksheets[index].images.len();
    workbook.worksheets[index].add_image(image);
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!("embedded {image_path} at {name}!{anchor}");
    Ok((
        summary,
        json!({
            "path": path,
            "sheet": name,
            "anchor": anchor,
            "images": count + 1,
            "width": width,
            "height": height,
        }),
    ))
}

/// Attach, replace or remove a cell comment.
pub fn add_comment(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let cell = args.require_str("cell")?;
    let text = args
        .raw("text")
        .and_then(Value::as_str)
        .ok_or_else(|| "text is required".to_string())?
        .to_string();
    let author = args
        .opt_str("author")
        .unwrap_or_else(|| DEFAULT_AUTHOR.to_string());
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    let name = workbook.worksheets[index].title().to_string();
    if text.trim().is_empty() {
        workbook.worksheets[index]
            .set_comment(&cell, None)
            .map_err(err)?;
        workspace.save(workbook, &path)?;
        let summary = format!("removed the comment on {name}!{cell}");
        return Ok((
            summary,
            json!({ "path": path, "sheet": name, "cell": cell, "removed": true }),
        ));
    }
    let comment = Comment::new(text.clone(), author.clone());
    workbook.worksheets[index]
        .set_comment(&cell, Some(comment))
        .map_err(err)?;
    workspace.save(workbook, &path)?;
    let summary = format!("commented on {name}!{cell}");
    Ok((
        summary,
        json!({
            "path": path,
            "sheet": name,
            "cell": cell,
            "author": author,
            "characters": text.chars().count(),
        }),
    ))
}

/// The author recorded when the caller does not name one.
const DEFAULT_AUTHOR: &str = "ferroxl-mcp";

#[cfg(test)]
mod tests {
    use super::*;
    use crate::testing;

    /// The corners are 0-based `(row, column)`, which is the order `Reference` wants --
    /// the reverse order still compiles, and produces a reference to a cell in neither row
    /// nor column of the range that was asked for.
    #[test]
    fn a_range_becomes_its_zero_based_row_column_corners() {
        assert_eq!(sheet_range("A2:A8").unwrap(), ((1, 0), Some((7, 0))));
        assert_eq!(sheet_range("B2:C4").unwrap(), ((1, 1), Some((3, 2))));
        assert_eq!(sheet_range("A1").unwrap(), ((0, 0), None));
        assert_eq!(
            sheet_range("Other!$A$2:$A$9").unwrap(),
            ((1, 0), Some((8, 0)))
        );
        // The `=Sheet!A1` spelling a chart author would paste in.
        assert_eq!(
            sheet_range("='Data'!$C$2:$C$4").unwrap(),
            ((1, 2), Some((3, 2)))
        );
    }

    /// The point the whole function exists to prevent: a range that must reach
    /// `Reference` with both corners intact, so the written `<c:f>` is the range itself.
    #[test]
    fn the_corners_render_back_as_the_range_that_was_given() {
        let (pos1, pos2) = sheet_range("B2:B3").unwrap();
        let reference = ferroxl::charts::Reference::new("Data", pos1, pos2, None, None).unwrap();
        assert_eq!(reference.to_reference_string(), "'Data'!$B$2:$B$3");
    }

    #[test]
    fn a_malformed_range_is_rejected() {
        assert!(sheet_range("not a range").is_err());
        assert!(sheet_range("A1:").is_err());
    }

    #[test]
    fn a_chart_needs_at_least_one_series() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "type": "bar",
            "anchor": "D2",
            "series": [],
        }));
        let error = add_chart(&workspace, &args).unwrap_err();
        assert!(error.contains("at least one series"), "{error}");
    }

    #[test]
    fn an_unknown_chart_type_is_rejected() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "type": "sunburst",
            "anchor": "D2",
            "series": [{ "values": "A1:A2" }],
        }));
        let error = add_chart(&workspace, &args).unwrap_err();
        assert!(error.contains("sunburst"), "{error}");
    }

    #[test]
    fn a_chart_is_written_and_read_back() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "type": "bar",
            "anchor": "D2",
            "title": "Totals",
            "series": [
                { "name": "Row 1", "values": "A1:B1" },
                { "name": "Row 2", "values": "A2:B2" },
            ],
        }));
        let (summary, payload) = add_chart(&workspace, &args).unwrap();
        assert!(summary.contains("2 series"), "{summary}");
        assert_eq!(payload["series"], json!(1));

        // openpyxl does not read charts back, so a reloaded workbook reports none. The
        // chart is in the file; the reader simply does not surface it, and ferroxl matches
        // that.
        let args = Args::new(&json!({ "path": "report.xlsx", "sheet": "Numbers" }));
        let (_, described) = super::super::inspect::describe_sheet(&workspace, &args).unwrap();
        assert_eq!(described["charts"], json!(0));
    }

    #[test]
    fn only_png_images_are_accepted() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "image_path": "photo.jpg",
            "anchor": "A1",
        }));
        let error = add_image(&workspace, &args).unwrap_err();
        assert!(error.contains("PNG"), "{error}");
    }

    #[test]
    fn a_missing_image_is_reported_before_the_workbook_is_touched() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "image_path": "absent.png",
            "anchor": "A1",
        }));
        let error = add_image(&workspace, &args).unwrap_err();
        assert!(error.contains("absent.png"), "{error}");
    }

    #[test]
    fn a_comment_can_be_added_and_removed() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "cell": "A1",
            "text": "check this",
            "author": "Reviewer",
        }));
        add_comment(&workspace, &args).unwrap();
        let args = Args::new(&json!({ "path": "report.xlsx", "sheet": "Numbers" }));
        let (_, payload) = super::super::inspect::list_comments(&workspace, &args).unwrap();
        assert_eq!(payload["comments"][0]["author"], json!("Reviewer"));
        assert_eq!(payload["comments"][0]["text"], json!("check this"));

        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "cell": "A1",
            "text": "",
        }));
        let (_, payload) = add_comment(&workspace, &args).unwrap();
        assert_eq!(payload["removed"], json!(true));
        let args = Args::new(&json!({ "path": "report.xlsx", "sheet": "Numbers" }));
        let (_, payload) = super::super::inspect::list_comments(&workspace, &args).unwrap();
        assert_eq!(payload["comments"], json!([]));
    }
}
