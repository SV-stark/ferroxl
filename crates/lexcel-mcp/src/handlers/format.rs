//! Tools for layout and appearance: dimensions, styling, validation and conditional
//! formatting.

use lexcel::datavalidation::{
    DataValidation, ValidationErrorStyle, ValidationOperator, ValidationType,
};
use lexcel::formatting::rules::{CellIsRule, ColorScaleRule, FormulaRule, Rule};
use lexcel::styles::{Border, Borders, Fill, Font, Style};
use lexcel::worksheet::{ColumnDimension, RowDimension};
use lexcel::Color;
use serde_json::{json, Value};

use super::{check_colour, err, open_for_edit, Handled};
use crate::tools::Args;
use crate::workspace::Workspace;

/// Set the width of one column or a span of columns.
pub fn set_column_width(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let columns = args.require_str("columns")?;
    let width = args.require_number("width")?;
    if width <= 0.0 {
        return Err("width must be greater than zero".to_string());
    }
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    let letters = expand_columns(&columns)?;
    for letter in &letters {
        let mut dimension = workbook.worksheets[index]
            .column_dimensions
            .get(letter)
            .cloned()
            .unwrap_or_else(|| ColumnDimension::new(letter));
        dimension.width = width;
        workbook.worksheets[index]
            .column_dimensions
            .insert(letter.clone(), dimension);
    }
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!("set {columns} on {name} to width {width}");
    Ok((
        summary,
        json!({ "path": path, "sheet": name, "columns": letters, "width": width }),
    ))
}

/// Set the height of one row or a span of rows.
pub fn set_row_height(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let rows = args.require_str("rows")?;
    let height = args.require_number("height")?;
    if height <= 0.0 {
        return Err("height must be greater than zero".to_string());
    }
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    let numbers = expand_rows(&rows)?;
    for number in &numbers {
        let mut dimension = workbook.worksheets[index]
            .row_dimensions
            .get(number)
            .cloned()
            .unwrap_or_else(|| RowDimension::new(*number));
        dimension.height = height;
        workbook.worksheets[index]
            .row_dimensions
            .insert(*number, dimension);
    }
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!("set rows {rows} on {name} to height {height}pt");
    Ok((
        summary,
        json!({ "path": path, "sheet": name, "rows": numbers, "height": height }),
    ))
}

/// Set the printed header and footer.
///
/// Each section is assigned directly rather than through `set_header`, because `&Ltext`
/// and `&L` followed by text are the same thing to Excel but not to openpyxl's own parser,
/// which only recognises a section marker as a standalone `&`-delimited field.
pub fn set_header_footer(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    // The three positions, each with its header and footer counterpart.
    let positions = ["left", "center", "right"];
    let mut applied = 0usize;
    for position in positions {
        for (part, header) in [("header", true), ("footer", false)] {
            let key = format!("{position}_{part}");
            let Some(text) = args.opt_str(&key) else {
                continue;
            };
            applied += 1;
            let section = section_mut(
                &mut workbook.worksheets[index].header_footer,
                position,
                header,
            );
            section.text = if text.is_empty() { None } else { Some(text) };
        }
    }
    if applied == 0 {
        return Err("pass at least one header or footer section".to_string());
    }
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!("set {applied} header/footer section(s) on {name}");
    Ok((
        summary,
        json!({ "path": path, "sheet": name, "sections": applied }),
    ))
}

/// The header or footer section for a position, where `position` is `left`, `center` or
/// `right`.
fn section_mut<'a>(
    header_footer: &'a mut lexcel::worksheet::HeaderFooter,
    position: &str,
    header: bool,
) -> &'a mut lexcel::worksheet::HeaderFooterItem {
    match (position, header) {
        ("left", true) => &mut header_footer.left_header,
        ("center", true) => &mut header_footer.center_header,
        ("right", true) => &mut header_footer.right_header,
        ("left", false) => &mut header_footer.left_footer,
        ("center", false) => &mut header_footer.center_footer,
        _ => &mut header_footer.right_footer,
    }
}

/// Attach a hyperlink to a cell.
pub fn add_hyperlink(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let cell = args.require_str("cell")?;
    let target = args.require_str("target")?;
    let display = args.opt_str("display").unwrap_or_else(|| target.clone());
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    if workbook.worksheets[index].cell_value(&cell).is_none() {
        workbook.worksheets[index]
            .set(&cell, lexcel::CellValue::text(display.clone()))
            .map_err(err)?;
    }
    let relationship = workbook.worksheets[index]
        .set_hyperlink(&cell, &target)
        .map_err(err)?;
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!("linked {name}!{cell} to {target}");
    Ok((
        summary,
        json!({
            "path": path,
            "sheet": name,
            "cell": cell,
            "target": target,
            "relationship": relationship,
        }),
    ))
}

/// Apply a style to a range.
pub fn style_cells(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let range = args.require_str("range")?;
    let requested = args.require_object("style")?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    let coordinates: Vec<String> = workbook.worksheets[index]
        .range_coordinates(&range)
        .map_err(err)?;
    if coordinates.is_empty() {
        return Err(format!("{range} covers no cells"));
    }
    // The style is built once and cloned per cell, so a large range does not repeat the
    // same parsing work thousands of times.
    let mut template = workbook.worksheets[index]
        .get_style_read_only(&coordinates[0])
        .cloned()
        .unwrap_or_else(Style::new);
    apply_style(&mut template, requested)?;
    for coordinate in &coordinates {
        workbook.worksheets[index]
            .set_style(coordinate, template.clone())
            .map_err(err)?;
    }
    let name = workbook.worksheets[index].title().to_string();
    let applied = template.sort_key();
    workspace.save(workbook, &path)?;
    let summary = format!("styled {} cell(s) in {range} on {name}", coordinates.len());
    Ok((
        summary,
        json!({
            "path": path,
            "sheet": name,
            "range": range,
            "cells_styled": coordinates.len(),
            "style_key": applied,
        }),
    ))
}

/// Set a number format on a range.
pub fn set_number_format(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let range = args.require_str("range")?;
    let format = args.require_str("format")?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    let coordinates: Vec<String> = workbook.worksheets[index]
        .range_coordinates(&range)
        .map_err(err)?;
    for coordinate in &coordinates {
        workbook.worksheets[index]
            .set_number_format(coordinate, &format)
            .map_err(err)?;
    }
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!("set {format} on {} cell(s) in {range}", coordinates.len());
    Ok((
        summary,
        json!({
            "path": path,
            "sheet": name,
            "range": range,
            "format": format,
            "cells_styled": coordinates.len(),
        }),
    ))
}

/// Restrict what may be typed into a range.
pub fn add_data_validation(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let range = args.require_str("range")?;
    let kind = args.require_str("type")?;
    let validation_type = ValidationType::from_str(&kind)
        .ok_or_else(|| format!("{kind:?} is not a data-validation type"))?;
    let formula1 = args.require_str("formula1")?;
    let formula2 = args.opt_str("formula2").unwrap_or_default();
    let operator = match args.opt_str("operator") {
        Some(name) => Some(
            ValidationOperator::from_str(&name)
                .ok_or_else(|| format!("{name:?} is not a validation operator"))?,
        ),
        None => None,
    };
    let mut validation = DataValidation::new(
        validation_type,
        operator,
        Some(&formula1),
        Some(&formula2),
        args.opt_bool("allow_blank", true),
    );
    if let Some(message) = args.opt_str("error_message") {
        validation.set_error_message(&message, "Invalid value");
    }
    if let Some(message) = args.opt_str("prompt_message") {
        validation.set_prompt_message(&message, "Enter a value");
    }
    validation.attr_map.insert(
        "errorStyle".to_string(),
        ValidationErrorStyle::Stop.as_str().to_string(),
    );
    validation.ranges = vec![range.clone()];
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    let count = workbook.worksheets[index].data_validations.len();
    workbook.worksheets[index].add_data_validation(validation);
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!("added a {kind} validation on {range} to {name}");
    Ok((
        summary,
        json!({
            "path": path,
            "sheet": name,
            "range": range,
            "type": kind,
            "rules": count + 1,
        }),
    ))
}

/// Add a conditional format to a range.
pub fn add_conditional_format(workspace: &Workspace, args: &Args) -> Handled {
    let path = args.require_str("path")?;
    let range = args.require_str("range")?;
    let kind = args.require_str("kind")?;
    let rule = build_rule(&kind, args)?;
    let (mut workbook, index) = open_for_edit(workspace, args)?;
    let priority = workbook.worksheets[index]
        .conditional_formatting
        .max_priority
        + 1;
    let mut rule = rule;
    rule = rule.with_priority(priority);
    workbook.worksheets[index]
        .conditional_formatting
        .add(&range, rule);
    let name = workbook.worksheets[index].title().to_string();
    workspace.save(workbook, &path)?;
    let summary = format!("added a {kind} rule on {range} to {name}");
    Ok((
        summary,
        json!({
            "path": path,
            "sheet": name,
            "range": range,
            "kind": kind,
            "priority": priority,
        }),
    ))
}

/// Build the rule described by the arguments.
fn build_rule(kind: &str, args: &Args) -> Result<Rule, String> {
    match kind {
        "colorScale" => Ok(colour_scale(args)?.to_rule()),
        "formula" => {
            let formula = args
                .opt_str("formula")
                .ok_or("a formula rule needs a formula")?;
            let builder = FormulaRule::new(Some(&formula), false);
            let mut rule = builder.to_rule();
            if let Some(style) = differential_style(args)? {
                rule = rule.with_dxf(style);
            }
            Ok(rule)
        }
        "cellIs" => {
            let operator = args
                .opt_str("operator")
                .ok_or("a cellIs rule needs an operator")?;
            let formula = args
                .opt_str("formula")
                .ok_or("a cellIs rule needs a formula to compare against")?;
            // A two-sided comparison is written as a single AND, which is how openpyxl
            // expresses it.
            let expression = match args.opt_str("second_formula") {
                Some(second) => format!("AND({formula},{second})"),
                None => formula,
            };
            let mut rule = CellIsRule::new(Some(&operator), Some(&expression), false).to_rule();
            if let Some(style) = differential_style(args)? {
                rule = rule.with_dxf(style);
            }
            Ok(rule)
        }
        other => Err(format!(
            "{other:?} is not a rule kind; use cellIs, formula or colorScale"
        )),
    }
}

/// Build a two- or three-stop colour scale from the supplied colours.
///
/// The midpoint only appears when `mid_color` was given, so the common two-colour case does
/// not carry a stray `percentile` stop.
fn colour_scale(args: &Args) -> Result<ColorScaleRule, String> {
    let start = check_colour(
        "start_color",
        &args
            .opt_str("start_color")
            .ok_or("a colour scale needs start_color")?,
    )?;
    let end = check_colour(
        "end_color",
        &args
            .opt_str("end_color")
            .ok_or("a colour scale needs end_color")?,
    )?;
    let mut scale = ColorScaleRule::new(
        Some("min"),
        None,
        Some(Color::new(start)),
        None,
        None,
        None,
        Some("max"),
        None,
        Some(Color::new(end)),
    );
    if let Some(mid) = args.opt_str("mid_color") {
        scale.mid_type = Some("percentile".to_string());
        scale.mid_value = Some("50".to_string());
        scale.mid_color = Some(Color::new(check_colour("mid_color", &mid)?));
    }
    Ok(scale)
}

/// The differential style a comparison rule applies when it matches.
fn differential_style(args: &Args) -> Result<Option<lexcel::formatting::DxfStyle>, String> {
    let font_colour = args.opt_str("font_color");
    let fill_colour = args.opt_str("fill_color");
    let bold = args.opt_bool("bold", false);
    if font_colour.is_none() && fill_colour.is_none() && !bold {
        return Ok(None);
    }
    let mut font = Font::new();
    let mut has_font = false;
    if let Some(colour) = font_colour {
        font.color = Color::new(check_colour("font_color", &colour)?);
        has_font = true;
    }
    if bold {
        font.bold = true;
        has_font = true;
    }
    let fill = match fill_colour {
        Some(colour) => {
            let mut fill = Fill::new();
            fill.fill_type = Some("solid".to_string());
            fill.start_color = Color::new(check_colour("fill_color", &colour)?);
            Some(fill)
        }
        None => None,
    };
    Ok(Some(lexcel::formatting::DxfStyle {
        font: has_font.then_some(font),
        border: None,
        fill,
    }))
}

/// Apply the keys present in `requested` to `style`, leaving everything else alone.
fn apply_style(style: &mut Style, requested: &Value) -> Result<(), String> {
    let Some(fields) = requested.as_object() else {
        return Err("style must be an object".to_string());
    };
    for (key, value) in fields {
        match key.as_str() {
            "bold" => style.font.bold = value.as_bool().ok_or("bold must be a boolean")?,
            "italic" => style.font.italic = value.as_bool().ok_or("italic must be a boolean")?,
            "underline" => {
                style.font.underline = value
                    .as_str()
                    .ok_or("underline must be a string")?
                    .to_string()
            }
            "font_size" => style.font.size = value.as_f64().ok_or("font_size must be a number")?,
            "font_name" => {
                style.font.name = value
                    .as_str()
                    .ok_or("font_name must be a string")?
                    .to_string()
            }
            "font_color" => {
                style.font.color = Color::new(check_colour(
                    "font_color",
                    value.as_str().ok_or("font_color must be a string")?,
                )?)
            }
            "fill_color" => {
                let colour = check_colour(
                    "fill_color",
                    value.as_str().ok_or("fill_color must be a string")?,
                )?;
                style.fill.fill_type = Some("solid".to_string());
                style.fill.start_color = Color::new(colour);
            }
            "number_format" => style
                .number_format
                .set_format_code(value.as_str().ok_or("number_format must be a string")?),
            "horizontal" => {
                let requested = value.as_str().ok_or("horizontal must be a string")?;
                if !matches!(
                    requested,
                    "general"
                        | "left"
                        | "center"
                        | "right"
                        | "fill"
                        | "justify"
                        | "centerContinuous"
                        | "distributed"
                ) {
                    return Err(format!("{requested:?} is not a horizontal alignment"));
                }
                style.alignment.horizontal = requested.to_string();
            }
            "vertical" => {
                let requested = value.as_str().ok_or("vertical must be a string")?;
                if !matches!(
                    requested,
                    "top" | "center" | "bottom" | "justify" | "distributed"
                ) {
                    return Err(format!("{requested:?} is not a vertical alignment"));
                }
                style.alignment.vertical = requested.to_string();
            }
            "wrap_text" => {
                style.alignment.wrap_text = value.as_bool().ok_or("wrap_text must be a boolean")?
            }
            "indent" => {
                style.alignment.indent = value.as_i64().ok_or("indent must be an integer")?
            }
            "text_rotation" => {
                let rotation = value.as_i64().ok_or("text_rotation must be an integer")?;
                if !(-90..=90).contains(&rotation) && !(91..=180).contains(&rotation) {
                    return Err(format!(
                        "text_rotation must be between -90 and 90, or 91 and 180, not {rotation}"
                    ));
                }
                style.alignment.text_rotation = rotation;
            }
            "border" => {
                let name = value.as_str().ok_or("border must be a string")?;
                // The colour may be given in the same call, which is the common case.
                let colour = sibling_colour(requested, "border_color");
                let side = Border {
                    border_style: Some(name.to_string()),
                    color: Color::new(colour),
                };
                style.borders = Borders {
                    left: side.clone(),
                    right: side.clone(),
                    top: side.clone(),
                    bottom: side.clone(),
                    diagonal: side.clone(),
                    ..Borders::new()
                };
            }
            "border_color" => {
                // Accepted on its own so a caller can set a border style and its colour in
                // two calls; the colour is applied to whatever border style is in place.
                let colour = Color::new(check_colour(
                    "border_color",
                    value.as_str().ok_or("border_color must be a string")?,
                )?);
                for side in [
                    &mut style.borders.left,
                    &mut style.borders.right,
                    &mut style.borders.top,
                    &mut style.borders.bottom,
                ] {
                    side.color = colour.clone();
                }
            }
            other => return Err(format!("{other:?} is not a style property")),
        }
    }
    Ok(())
}

/// A colour declared alongside another key in the same style object.
fn sibling_colour(requested: &Value, name: &str) -> String {
    requested
        .get(name)
        .and_then(Value::as_str)
        .map(|colour| colour.trim_start_matches('#').to_string())
        .unwrap_or_else(|| "FF000000".to_string())
}

/// Expand `A`, `A:C` into the individual column letters.
fn expand_columns(spec: &str) -> Result<Vec<String>, String> {
    let trimmed = spec.trim().to_uppercase();
    match trimmed.split_once(':') {
        Some((from, to)) => {
            let start = lexcel::column_index_from_string(from.trim()).map_err(err)?;
            let end = lexcel::column_index_from_string(to.trim()).map_err(err)?;
            if end < start {
                return Err(format!("{spec:?} runs backwards"));
            }
            Ok((start..=end)
                .map(lexcel::get_column_letter)
                .collect::<Result<Vec<_>, _>>()
                .map_err(err)?)
        }
        None => {
            lexcel::column_index_from_string(&trimmed).map_err(err)?;
            Ok(vec![trimmed])
        }
    }
}

/// Expand `3` or `3:5` into the individual row numbers.
fn expand_rows(spec: &str) -> Result<Vec<u32>, String> {
    let trimmed = spec.trim();
    match trimmed.split_once(':') {
        Some((from, to)) => {
            let start: u32 = from
                .trim()
                .parse()
                .map_err(|_| format!("{from:?} is not a row number"))?;
            let end: u32 = to
                .trim()
                .parse()
                .map_err(|_| format!("{to:?} is not a row number"))?;
            if end < start {
                return Err(format!("{spec:?} runs backwards"));
            }
            Ok((start..=end).collect())
        }
        None => {
            let single: u32 = trimmed
                .parse()
                .map_err(|_| format!("{trimmed:?} is not a row number"))?;
            Ok(vec![single])
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::testing;

    #[test]
    fn a_column_spec_expands_to_letters() {
        assert_eq!(expand_columns("B").unwrap(), vec!["B"]);
        assert_eq!(expand_columns("a:c").unwrap(), vec!["A", "B", "C"]);
        assert!(expand_columns("C:A").is_err());
        assert!(expand_columns("1").is_err());
    }

    #[test]
    fn a_row_spec_expands_to_numbers() {
        assert_eq!(expand_rows("3").unwrap(), vec![3]);
        assert_eq!(expand_rows("3:5").unwrap(), vec![3, 4, 5]);
        assert!(expand_rows("5:3").is_err());
        assert!(expand_rows("x").is_err());
    }

    #[test]
    fn styling_applies_only_the_keys_it_was_given() {
        let mut style = Style::new();
        style.font.italic = true;
        apply_style(&mut style, &json!({ "bold": true })).unwrap();
        assert!(style.font.bold);
        assert!(style.font.italic, "italic should have been left alone");
    }

    #[test]
    fn styling_rejects_a_property_it_does_not_know() {
        let mut style = Style::new();
        let error = apply_style(&mut style, &json!({ "sparkle": true })).unwrap_err();
        assert!(error.contains("sparkle"), "{error}");
    }

    #[test]
    fn alignment_and_rotation_are_validated() {
        let mut style = Style::new();
        assert!(apply_style(&mut style, &json!({ "horizontal": "sideways" })).is_err());
        assert!(apply_style(&mut style, &json!({ "text_rotation": 200 })).is_err());
        assert!(apply_style(&mut style, &json!({ "text_rotation": -45 })).is_ok());
    }

    #[test]
    fn a_border_sets_all_four_sides() {
        let mut style = Style::new();
        apply_style(
            &mut style,
            &json!({ "border": "medium", "border_color": "FF00FF00" }),
        )
        .unwrap();
        for side in [
            &style.borders.left,
            &style.borders.right,
            &style.borders.top,
        ] {
            assert_eq!(side.border_style.as_deref(), Some("medium"));
            assert_eq!(side.color.index, "FF00FF00");
        }
    }

    #[test]
    fn set_number_format_survives_a_round_trip() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "range": "A1:A2",
            "format": "0.00%",
        }));
        let (_, payload) = set_number_format(&workspace, &args).unwrap();
        assert_eq!(payload["cells_styled"], json!(2));

        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "range": "A1:A1",
        }));
        let (_, payload) = super::super::inspect::read_cells(&workspace, &args).unwrap();
        assert_eq!(payload["cells"][0]["number_format"], json!("0.00%"));
    }

    #[test]
    fn widths_and_heights_apply_to_a_span() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "columns": "A:B",
            "width": 20,
        }));
        set_column_width(&workspace, &args).unwrap();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
            "rows": "1:2",
            "height": 30,
        }));
        set_row_height(&workspace, &args).unwrap();
        let args = Args::new(&json!({ "path": "report.xlsx", "sheet": "Numbers" }));
        let (_, payload) = super::super::inspect::describe_sheet(&workspace, &args).unwrap();
        assert_eq!(payload["column_dimensions"], json!(2));
        assert_eq!(payload["row_dimensions"], json!(2));
    }

    #[test]
    fn a_non_positive_size_is_refused() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx", "columns": "A", "width": 0,
        }));
        assert!(set_column_width(&workspace, &args)
            .unwrap_err()
            .contains("greater than zero"));
        let args = Args::new(&json!({
            "path": "report.xlsx", "rows": "1", "height": -1,
        }));
        assert!(set_row_height(&workspace, &args)
            .unwrap_err()
            .contains("greater than zero"));
    }

    #[test]
    fn a_cell_is_rule_needs_an_operator_and_a_formula() {
        let error = build_rule("cellIs", &Args::new(&json!({}))).unwrap_err();
        assert!(error.contains("operator"), "{error}");
        let error = build_rule("cellIs", &Args::new(&json!({ "operator": ">" }))).unwrap_err();
        assert!(error.contains("formula"), "{error}");
    }

    #[test]
    fn an_unknown_rule_kind_is_rejected() {
        let error = build_rule("sparkline", &Args::new(&json!({}))).unwrap_err();
        assert!(error.contains("cellIs"), "{error}");
    }

    #[test]
    fn a_colour_scale_needs_both_ends() {
        let error = build_rule("colorScale", &Args::new(&json!({}))).unwrap_err();
        assert!(error.contains("start_color"), "{error}");
    }

    #[test]
    fn a_three_stop_scale_carries_a_midpoint() {
        let rule = build_rule(
            "colorScale",
            &Args::new(&json!({
                "start_color": "FF00FF00",
                "mid_color": "FFFFFF00",
                "end_color": "FFFF0000",
            })),
        )
        .unwrap();
        let scale = rule.color_scale.expect("a colour scale");
        assert_eq!(scale.cfvo.len(), 3);
        assert_eq!(scale.color.len(), 3);
    }

    #[test]
    fn a_two_stop_scale_has_no_midpoint() {
        let rule = build_rule(
            "colorScale",
            &Args::new(&json!({
                "start_color": "FF00FF00",
                "end_color": "FFFF0000",
            })),
        )
        .unwrap();
        let scale = rule.color_scale.expect("a colour scale");
        assert_eq!(scale.cfvo.len(), 2);
        assert_eq!(scale.color.len(), 2);
    }

    #[test]
    fn a_data_validation_needs_a_known_type() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "rules.xlsx",
            "range": "C1:C5",
            "type": "telepathy",
            "formula1": "1",
        }));
        let error = add_data_validation(&workspace, &args).unwrap_err();
        assert!(error.contains("telepathy"), "{error}");
    }

    #[test]
    fn a_header_needs_at_least_one_section() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({ "path": "report.xlsx", "sheet": "Numbers" }));
        let error = set_header_footer(&workspace, &args).unwrap_err();
        assert!(error.contains("at least one"), "{error}");
    }
}
