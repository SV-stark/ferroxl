//! Tools that draw on a sheet: charts, images and comments.

use lexcel::charts::reference::{Reference, ReferenceDataType};
use lexcel::charts::{BarChart, LineChart, PieChart, ScatterChart, Series};
use lexcel::comments::Comment;
use lexcel::drawing::Image;
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
    let (column, row) = lexcel::coordinate_from_string(&anchor).map_err(err)?;

    let mut series = Vec::new();
    for (position, entry) in series_spec.iter().enumerate() {
        series.push(build_series(entry, &sheet_name, position, args)?);
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
fn build_series(
    entry: &Value,
    sheet_name: &str,
    position: usize,
    args: &Args,
) -> Result<Series, String> {
    let values = entry
        .get("values")
        .and_then(Value::as_str)
        .ok_or_else(|| format!("series {position} needs a values range"))?;
    let values_range = sheet_range(values)?;
    // A scatter plot puts numbers on both axes, so the values are forced to numeric. For
    // the other types lexcel infers the type from the cells themselves.
    let data_type = if args.opt_str("type").as_deref() == Some("scatter") {
        Some(ReferenceDataType::Numeric)
    } else {
        None
    };
    let mut series = Series::new(
        Reference::new(sheet_name, values_range, None, data_type, None).map_err(err)?,
    );
    if let Some(name) = entry.get("name").and_then(Value::as_str) {
        series = series.with_title(name.to_string());
    }
    Ok(series)
}

/// Turn an A1 range into the `(first column, point count)` pair `Reference` wants.
///
/// The range may carry its own sheet name and absolute markers; both are stripped, because
/// the sheet is passed separately to `Reference::new`.
fn sheet_range(range: &str) -> Result<(usize, usize), String> {
    let trimmed = range.trim();
    let local = trimmed.strip_prefix('=').unwrap_or(trimmed);
    let local = local.rsplit_once('!').map(|(_, tail)| tail).unwrap_or(local);
    let cleaned = local.replace('$', "");
    if let Some((min, max)) = cleaned.split_once(':') {
        let (start, start_row) = lexcel::coordinate_from_string(min).map_err(err)?;
        let (end, end_row) = lexcel::coordinate_from_string(max).map_err(err)?;
        let start = lexcel::column_index_from_string(&start).map_err(err)?;
        let end = lexcel::column_index_from_string(&end).map_err(err)?;
        if end < start {
            return Err(format!("{range:?} runs backwards"));
        }
        if end_row < start_row {
            return Err(format!("{range:?} runs backwards"));
        }
        return Ok((start as usize, (end_row - start_row + 1) as usize));
    }
    // A bare cell is a single point.
    lexcel::coordinate_from_string(&cleaned).map_err(err)?;
    Ok((1, 1))
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
    workbook.worksheets[index].cell_anchor(&anchor).map_err(err)?;
    let mut image = Image::from_png(bytes).map_err(err)?;
    if let Some(width) = args.opt_number("width") {
        image.drawing.set_width(width as i64);
    }
    if let Some(height) = args.opt_number("height") {
        image.drawing.set_height(height as i64);
    }
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
const DEFAULT_AUTHOR: &str = "lexcel-mcp";

#[cfg(test)]
mod tests {
    use super::*;
    use crate::testing;

    #[test]
    fn a_range_becomes_an_offset_and_a_length() {
        assert_eq!(sheet_range("A2:A8").unwrap(), (1, 7));
        assert_eq!(sheet_range("B2:C4").unwrap(), (2, 3));
        assert_eq!(sheet_range("A1").unwrap(), (1, 1));
        assert_eq!(sheet_range("Other!$A$2:$A$9").unwrap(), (1, 8));
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
        // chart is in the file; the reader simply does not surface it, and lexcel matches
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
