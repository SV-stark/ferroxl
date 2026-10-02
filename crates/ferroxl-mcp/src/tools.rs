//! The catalogue of tools this server exposes, and the argument helpers used to read
//! them.
//!
//! Every tool carries a JSON Schema for its arguments. The schemas are the contract with
//! the model, so they are written out in full rather than generated: a missing
//! description here shows up as a tool the agent refuses to call.

use std::sync::OnceLock;

use serde_json::{json, Value};

/// One callable tool.
#[derive(Debug, Clone)]
pub struct ToolSpec {
    /// The name the model sends in `tools/call`.
    pub name: &'static str,
    /// What the tool does, and when to prefer it over its neighbours.
    pub description: &'static str,
    /// The JSON Schema for the arguments object.
    pub input_schema: Value,
}

impl ToolSpec {
    /// Render the spec in the shape `tools/list` expects.
    pub fn to_json(&self) -> Value {
        json!({
            "name": self.name,
            "description": self.description,
            "inputSchema": self.input_schema,
        })
    }
}

/// Look a tool up by name.
pub fn find(name: &str) -> Option<&'static ToolSpec> {
    catalogue().iter().find(|tool| tool.name == name)
}

/// Every tool, in the order they are advertised to the client.
///
/// The catalogue is built once and kept, because `tools/list` is asked for on every
/// connection and the schemas are not cheap to assemble.
pub fn catalogue() -> &'static [ToolSpec] {
    static CATALOGUE: OnceLock<Vec<ToolSpec>> = OnceLock::new();
    CATALOGUE.get_or_init(build_catalogue)
}

fn build_catalogue() -> Vec<ToolSpec> {
    vec![
        // -- Inspection -----------------------------------------------------------------
        ToolSpec {
            name: "list_sheets",
            description: "List the sheets in a workbook with their index, visibility, used \
                          range and cell count. Start here when you do not know what a file \
                          contains.",
            input_schema: object(
                &[
                    ("path", string("Path to the .xlsx or .xlsm file, relative to the workspace root.")),
                ],
                &["path"],
            ),
        },
        ToolSpec {
            name: "describe_sheet",
            description: "Describe one sheet in detail: dimensions, how many cells hold each \
                          kind of value, merged ranges, freeze panes, the autofilter, data \
                          validations, conditional formats, comments, charts and images. Use \
                          this to plan an edit before making it.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                ],
                &["path"],
            ),
        },
        ToolSpec {
            name: "read_cells",
            description: "Read a rectangular range of cells. Returns each cell's coordinate, \
                          value, data type and number format. Use an A1 range such as \
                          \"A1:D20\"; omit `range` to read the sheet's used range.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("range", string("A1 range, for example \"B2:F40\". Defaults to the used range.")),
                    ("include_empty", boolean("Include cells with no value. Defaults to true.")),
                ],
                &["path"],
            ),
        },
        ToolSpec {
            name: "read_formulas",
            description: "List the formula cells in a range together with the formula text. \
                          Formulas are not evaluated: the value returned is what the file \
                          cached when it was last saved by Excel.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("range", string("A1 range to search. Defaults to the used range.")),
                ],
                &["path"],
            ),
        },
        ToolSpec {
            name: "search_values",
            description: "Search a sheet for text. Returns matching cells with their \
                          coordinates and surrounding row context, which is usually enough to \
                          answer a question without reading the whole sheet.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("query", string("Text to look for. Case-insensitive unless case_sensitive is set.")),
                    ("case_sensitive", boolean("Match case exactly. Defaults to false.")),
                    ("match_formulas", boolean("Also search formula text, not just values. Defaults to true.")),
                    ("max_results", integer("Stop after this many matches. Defaults to 100.")),
                ],
                &["path", "query"],
            ),
        },
        ToolSpec {
            name: "summarize_range",
            description: "Compute count, numeric count, sum, mean, min, max and blank count for \
                          a range. Use this instead of reading a large numeric block just to \
                          describe it.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("range", string("A1 range. Defaults to the used range.")),
                ],
                &["path"],
            ),
        },
        ToolSpec {
            name: "list_comments",
            description: "List the cell comments in a sheet, with their author and text.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                ],
                &["path"],
            ),
        },
        ToolSpec {
            name: "list_named_ranges",
            description: "List the workbook's defined names and the ranges or values they point at.",
            input_schema: object(
                &[("path", string("Path to the workbook."))],
                &["path"],
            ),
        },
        ToolSpec {
            name: "export_csv",
            description: "Render a range as RFC 4180 CSV text. Cheaper than reading cells when \
                          you only need the data to paste somewhere else.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("range", string("A1 range. Defaults to the used range.")),
                ],
                &["path"],
            ),
        },
        // -- Structure ------------------------------------------------------------------
        ToolSpec {
            name: "create_workbook",
            description: "Create a new workbook with the given sheet names. Refuses to \
                          overwrite an existing file unless overwrite is true.",
            input_schema: object(
                &[
                    ("path", string("Path for the new .xlsx file.")),
                    ("sheets", strings("Sheet names, in order. Defaults to a single \"Sheet\".")),
                    ("overwrite", boolean("Replace the file if it already exists. Defaults to false.")),
                ],
                &["path"],
            ),
        },
        ToolSpec {
            name: "add_sheet",
            description: "Add a sheet to an existing workbook, optionally at a given position.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("title", string("Name for the new sheet. Must be unique in the workbook.")),
                    ("index", integer("Zero-based position. Appended to the end when omitted.")),
                ],
                &["path", "title"],
            ),
        },
        ToolSpec {
            name: "remove_sheet",
            description: "Delete a sheet. The workbook must keep at least one sheet, so this \
                          fails on a single-sheet workbook.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet to delete.")),
                ],
                &["path", "sheet"],
            ),
        },
        ToolSpec {
            name: "rename_sheet",
            description: "Rename a sheet, keeping its position in the workbook.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Current sheet name.")),
                    ("title", string("New sheet name.")),
                ],
                &["path", "sheet", "title"],
            ),
        },
        ToolSpec {
            name: "merge_cells",
            description: "Merge a range into a single cell, keeping the top-left value.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("range", string("A1 range to merge, for example \"A1:C1\".")),
                ],
                &["path", "range"],
            ),
        },
        ToolSpec {
            name: "unmerge_cells",
            description: "Split a previously merged range back into individual cells.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("range", string("A1 range to unmerge.")),
                ],
                &["path", "range"],
            ),
        },
        ToolSpec {
            name: "freeze_panes",
            description: "Freeze the rows and columns above and to the left of a cell, or unfreeze \
                          by passing no cell.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("cell", string("Cell to freeze at, for example \"B2\". Omit to unfreeze.")),
                ],
                &["path"],
            ),
        },
        ToolSpec {
            name: "set_auto_filter",
            description: "Set or clear a sheet's autofilter range. The range should cover \
                          the header row and every row the filter applies to.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("range", string("A1 range to filter, for example \"A1:F100\". Omit to clear.")),
                ],
                &["path"],
            ),
        },
        ToolSpec {
            name: "add_named_range",
            description: "Define a named range such as `Total` pointing at a sheet range.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("name", string("The name, without spaces or punctuation Excel forbids.")),
                    ("sheet", string("Sheet the range lives on.")),
                    ("range", string("A1 range, for example \"A1:A100\".")),
                    ("scope", string("Sheet name to scope the name to. Workbook-scoped when omitted.")),
                ],
                &["path", "name", "sheet", "range"],
            ),
        },
        // -- Values ---------------------------------------------------------------------
        ToolSpec {
            name: "set_cell",
            description: "Write a single cell. A value beginning with '=' is stored as a formula, \
                          \"YYYY-MM-DD\" as a date, \"YYYY-MM-DDTHH:MM:SS\" as a timestamp, a \
                          trailing '%' as a percentage, and anything else as text.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("cell", string("Cell coordinate, for example \"C7\".")),
                    ("value", value("The value to store.")),
                ],
                &["path", "cell", "value"],
            ),
        },
        ToolSpec {
            name: "write_cells",
            description: "Write a block of values starting at a cell. `rows` is a list of rows, \
                          each a list of values, written left to right.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("start", string("Coordinate of the top-left cell, for example \"A1\".")),
                    ("rows", array("Rows of values, outermost is the row.", object(&[], &[]))),
                ],
                &["path", "start", "rows"],
            ),
        },
        ToolSpec {
            name: "append_row",
            description: "Append one row of values below the last row that has data. Values are \
                          placed in the columns given, so 0 is column A.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("values", strings("Values in column order.")),
                ],
                &["path", "values"],
            ),
        },
        ToolSpec {
            name: "clear_cells",
            description: "Remove the values in a range. Formatting and comments are left alone.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("range", string("A1 range to clear.")),
                ],
                &["path", "range"],
            ),
        },
        // -- Layout ---------------------------------------------------------------------
        ToolSpec {
            name: "set_column_width",
            description: "Set the width of one column or a span of columns, in Excel's character \
                          units (roughly the width of one digit).",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("columns", string("A column, or a span such as \"B:D\".")),
                    ("width", number("Width in character units.")),
                ],
                &["path", "columns", "width"],
            ),
        },
        ToolSpec {
            name: "set_row_height",
            description: "Set the height of one row or a span of rows, in points.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("rows", string("A row number, or a span such as \"3:5\".")),
                    ("height", number("Height in points. 72 points is one inch.")),
                ],
                &["path", "rows", "height"],
            ),
        },
        ToolSpec {
            name: "set_header_footer",
            description: "Set the printed header and footer. Sections are introduced with &L, &C \
                          and &R, and &P, &N and &D insert the page number, page count and date.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("left_header", string("Left header text.")),
                    ("center_header", string("Centre header text.")),
                    ("right_header", string("Right header text.")),
                    ("left_footer", string("Left footer text.")),
                    ("center_footer", string("Centre footer text.")),
                    ("right_footer", string("Right footer text.")),
                ],
                &["path"],
            ),
        },
        ToolSpec {
            name: "add_hyperlink",
            description: "Attach a hyperlink to a cell. A cell with no text of its own is given \
                          the target as its display text.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("cell", string("Cell coordinate.")),
                    ("target", string("URL or file path.")),
                    ("display", string("Text to show. Defaults to the target.")),
                ],
                &["path", "cell", "target"],
            ),
        },
        // -- Formatting -----------------------------------------------------------------
        ToolSpec {
            name: "style_cells",
            description: "Apply a font, fill, border, alignment or number format to a range. Only \
                          the keys you pass are changed, so this composes with earlier calls.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("range", string("A1 range to style.")),
                    ("style", object_style()),
                ],
                &["path", "range", "style"],
            ),
        },
        ToolSpec {
            name: "set_number_format",
            description: "Set the number format of a range, for example \"0.00%\", \"#,##0\" or \
                          \"yyyy-mm-dd\". A cell that looks like a date only becomes a date when \
                          its format says so.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("range", string("A1 range to format.")),
                    ("format", string("A number format code.")),
                ],
                &["path", "range", "format"],
            ),
        },
        ToolSpec {
            name: "add_data_validation",
            description: "Restrict what may be typed into a range: a list, a whole number, a \
                          decimal, a date, a length, or a custom formula.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("range", string("A1 range the rule covers.")),
                    ("type", enum_of(&["list", "whole", "decimal", "date", "time", "textLength", "custom"], "The kind of rule.")),
                    ("formula1", string("The first bound, or a comma-separated list for type \"list\".")),
                    ("formula2", string("The second bound, for ranges such as \"between\".")),
                    ("operator", string("For two-bound rules: between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual.")),
                    ("allow_blank", boolean("Accept an empty cell. Defaults to true.")),
                    ("error_message", string("Message shown when the value is rejected.")),
                    ("prompt_message", string("Message shown when the cell is selected.")),
                ],
                &["path", "range", "type", "formula1"],
            ),
        },
        ToolSpec {
            name: "add_conditional_format",
            description: "Add a conditional format to a range: highlight cells that compare a \
                          certain way against a value, cells matching a formula, or a colour scale.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("range", string("A1 range the rule covers.")),
                    ("kind", enum_of(&["cellIs", "formula", "colorScale"], "The kind of rule.")),
                    ("operator", string("For cellIs: greaterThan, lessThan, between, equal, and the not- variants.")),
                    ("formula", string("The value to compare against, or the formula for kind \"formula\".")),
                    ("second_formula", string("The upper bound when operator is \"between\".")),
                    ("font_color", string("ARGB colour for matching text, for example \"FF9C0006\".")),
                    ("fill_color", string("ARGB colour for the matching background, for example \"FFFFC7CE\".")),
                    ("bold", boolean("Make matching text bold.")),
                    ("start_color", string("Colour at the low end of a colour scale.")),
                    ("mid_color", string("Colour at the midpoint of a colour scale.")),
                    ("end_color", string("Colour at the high end of a colour scale.")),
                ],
                &["path", "range", "kind"],
            ),
        },
        // -- Media and annotations -------------------------------------------------------
        ToolSpec {
            name: "add_chart",
            description: "Add a chart to a sheet. Give the category range and one or more series \
                          as A1 ranges; the axis bounds are computed from the data. openpyxl \
                          does not read charts back, so a reloaded workbook will not list it.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet to draw on. Defaults to the first sheet.")),
                    ("type", enum_of(&["bar", "line", "scatter", "pie"], "The chart type.")),
                    ("anchor", string("Top-left cell of the chart, for example \"E2\".")),
                    ("title", string("Chart title.")),
                    ("categories", string("A1 range holding the category labels, for example \"A2:A8\".")),
                    ("series", array("One entry per series.", series_schema())),
                    ("width", number("Chart width in centimetres. Defaults to 15.")),
                    ("height", number("Chart height in centimetres. Defaults to 7.5.")),
                ],
                &["path", "type", "anchor", "series"],
            ),
        },
        ToolSpec {
            name: "add_image",
            description: "Embed a PNG image on a sheet, anchored at a cell.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet to draw on. Defaults to the first sheet.")),
                    ("image_path", string("Path to a .png file, relative to the workspace root.")),
                    ("anchor", string("Top-left cell for the image, for example \"E2\".")),
                    ("width", number("Width in centimetres. Defaults to the image's own size.")),
                    ("height", number("Height in centimetres. Defaults to the image's own size.")),
                ],
                &["path", "image_path", "anchor"],
            ),
        },
        ToolSpec {
            name: "add_comment",
            description: "Attach a note to a cell. Pass an empty text to remove an existing comment.",
            input_schema: object(
                &[
                    ("path", string("Path to the workbook.")),
                    ("sheet", string("Sheet name. Defaults to the first sheet.")),
                    ("cell", string("Cell coordinate.")),
                    ("text", string("The note. An empty string removes the note.")),
                    ("author", string("Who is annotating. Defaults to \"ferroxl-mcp\".")),
                ],
                &["path", "cell", "text"],
            ),
        },
    ]
}

/// The style object accepted by `style_cells`.
fn object_style() -> Value {
    json!({
        "type": "object",
        "description": "Only the keys present are changed.",
        "properties": {
            "bold": boolean("Bold text."),
            "italic": boolean("Italic text."),
            "underline": string("\"single\", \"double\" or \"none\"."),
            "font_size": number("Font size in points."),
            "font_name": string("Font family, for example \"Calibri\"."),
            "font_color": string("ARGB colour for the text, for example \"FF000000\"."),
            "fill_color": string("ARGB solid background colour, for example \"FFFFFF00\"."),
            "number_format": string("Number format code, for example \"0.00%\"."),
            "horizontal": string("\"left\", \"center\", \"right\" or \"general\"."),
            "vertical": string("\"top\", \"center\", \"bottom\"."),
            "wrap_text": boolean("Wrap long text within the cell."),
            "indent": integer("Left indent, in indent levels."),
            "text_rotation": integer("Text rotation in degrees, from -90 to 90."),
            "border": string("Border style applied to all four edges, for example \"thin\" or \"medium\"."),
            "border_color": string("ARGB colour for the border."),
        },
    })
}

/// One entry of the `series` array accepted by `add_chart`.
fn series_schema() -> Value {
    json!({
        "type": "object",
        "required": ["values"],
        "properties": {
            "name": string("A literal series name, such as \"Revenue\"."),
            "values": string("A1 range holding the series values, for example \"B2:B8\"."),
        },
    })
}

// -- Schema helpers -------------------------------------------------------------------

fn object(properties: &[(&str, Value)], required: &[&str]) -> Value {
    json!({
        "type": "object",
        "properties": properties
            .iter()
            .map(|(name, schema)| (name.to_string(), schema.clone()))
            .collect::<serde_json::Map<String, Value>>(),
        "required": required,
        "additionalProperties": false,
    })
}

fn string(description: &str) -> Value {
    json!({ "type": "string", "description": description })
}

fn number(description: &str) -> Value {
    json!({ "type": "number", "description": description })
}

fn integer(description: &str) -> Value {
    json!({ "type": "integer", "description": description })
}

fn boolean(description: &str) -> Value {
    json!({ "type": "boolean", "description": description })
}

fn array(description: &str, items: Value) -> Value {
    json!({ "type": "array", "description": description, "items": items })
}

fn strings(description: &str) -> Value {
    array(description, json!({ "type": "string" }))
}

fn enum_of(options: &[&str], description: &str) -> Value {
    json!({
        "type": "string",
        "description": description,
        "enum": options,
    })
}

fn value(description: &str) -> Value {
    json!({
        "description": description,
        "type": ["string", "number", "boolean", "null"],
    })
}

// -- Argument reading -----------------------------------------------------------------

/// A typed reader over a tool's `arguments` object.
///
/// The object is owned so a handler can be handed a `&Args` without borrowing the caller's
/// JSON. The error messages name the argument and say what was expected, because a model
/// that gets a clear correction will fix its call rather than guessing again.
#[derive(Debug, Clone)]
pub struct Args {
    value: Value,
}

impl Args {
    /// Wrap an arguments object. Anything that is not an object reads as empty.
    pub fn new(value: &Value) -> Self {
        Args {
            value: if value.is_object() {
                value.clone()
            } else {
                Value::Null
            },
        }
    }

    /// A required string argument.
    pub fn require_str(&self, name: &str) -> Result<String, String> {
        match self.value.get(name) {
            Some(Value::String(text)) if !text.trim().is_empty() => Ok(text.clone()),
            Some(Value::String(_)) => Err(format!("{name} must not be empty")),
            Some(other) => Err(format!("{name} must be a string, got {other}")),
            None => Err(format!("{name} is required")),
        }
    }

    /// An optional string argument, treating an empty string as absent.
    pub fn opt_str(&self, name: &str) -> Option<String> {
        match self.value.get(name) {
            Some(Value::String(text)) if !text.trim().is_empty() => Some(text.clone()),
            _ => None,
        }
    }

    /// A required number argument.
    pub fn require_number(&self, name: &str) -> Result<f64, String> {
        match self.value.get(name) {
            Some(Value::Number(number)) => number
                .as_f64()
                .ok_or_else(|| format!("{name} must be a finite number")),
            Some(other) => Err(format!("{name} must be a number, got {other}")),
            None => Err(format!("{name} is required")),
        }
    }

    /// An optional number argument.
    pub fn opt_number(&self, name: &str) -> Option<f64> {
        self.value.get(name).and_then(Value::as_f64)
    }

    /// An optional integer argument.
    pub fn opt_usize(&self, name: &str) -> Option<usize> {
        self.value
            .get(name)
            .and_then(Value::as_u64)
            .map(|number| number as usize)
    }

    /// An optional boolean argument.
    pub fn opt_bool(&self, name: &str, default: bool) -> bool {
        self.value
            .get(name)
            .and_then(Value::as_bool)
            .unwrap_or(default)
    }

    /// A required array argument.
    pub fn require_array(&self, name: &str) -> Result<&Vec<Value>, String> {
        match self.value.get(name) {
            Some(Value::Array(items)) => Ok(items),
            Some(other) => Err(format!("{name} must be an array, got {other}")),
            None => Err(format!("{name} is required")),
        }
    }

    /// An optional array argument.
    pub fn opt_array(&self, name: &str) -> Option<&Vec<Value>> {
        self.value.get(name).and_then(Value::as_array)
    }

    /// A required object argument.
    pub fn require_object(&self, name: &str) -> Result<&Value, String> {
        match self.value.get(name) {
            Some(value @ Value::Object(_)) => Ok(value),
            Some(other) => Err(format!("{name} must be an object, got {other}")),
            None => Err(format!("{name} is required")),
        }
    }

    /// An argument of any type, including a `null` that the schema allows.
    ///
    /// This is how a value the schema types as "anything" is read, such as a cell value
    /// that may legitimately be `null`.
    pub fn raw(&self, name: &str) -> Option<&Value> {
        self.value.get(name)
    }

    /// Reject any argument the tool's schema does not declare.
    ///
    /// The schemas are written with `additionalProperties: false`, so this makes the
    /// runtime agree with what the model was told. An unrecognised key is nearly always a
    /// typo, and ignoring it silently is how a caller ends up believing it set something it
    /// did not.
    pub fn reject_unknown(&self, spec: &ToolSpec) -> Result<(), String> {
        let Some(declared) = spec
            .input_schema
            .get("properties")
            .and_then(Value::as_object)
        else {
            return Ok(());
        };
        let Some(sent) = self.value.as_object() else {
            return Ok(());
        };
        let unknown: Vec<&str> = sent
            .keys()
            .map(String::as_str)
            .filter(|key| !declared.contains_key(*key))
            .collect();
        if unknown.is_empty() {
            return Ok(());
        }
        let mut allowed: Vec<&str> = declared.keys().map(String::as_str).collect();
        allowed.sort_unstable();
        Err(format!(
            "{} does not take {}; it takes {}",
            spec.name,
            unknown
                .iter()
                .map(|name| format!("{name:?}"))
                .collect::<Vec<_>>()
                .join(", "),
            allowed.join(", ")
        ))
    }

    /// The sheet name, defaulting to the first sheet when the caller did not name one.
    pub fn sheet(&self) -> Option<String> {
        self.opt_str("sheet")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_catalogue_is_non_empty_and_uniquely_named() {
        let catalogue = catalogue();
        assert!(catalogue.len() >= 25, "only {} tools", catalogue.len());
        let mut names: Vec<&str> = catalogue.iter().map(|tool| tool.name).collect();
        names.sort_unstable();
        let before = names.len();
        names.dedup();
        assert_eq!(before, names.len(), "duplicate tool names");
    }

    #[test]
    fn every_tool_has_a_usable_schema() {
        for tool in catalogue() {
            assert!(
                tool.description.len() > 40,
                "{} needs a real description",
                tool.name
            );
            assert_eq!(tool.input_schema["type"], json!("object"));
            let properties = tool.input_schema["properties"]
                .as_object()
                .unwrap_or_else(|| panic!("{} has no properties", tool.name));
            assert!(!properties.is_empty(), "{} has no properties", tool.name);
            for required in tool.input_schema["required"]
                .as_array()
                .unwrap_or(&Vec::new())
            {
                let name = required.as_str().unwrap();
                assert!(
                    properties.contains_key(name),
                    "{} requires {name} but does not declare it",
                    tool.name
                );
            }
        }
    }

    #[test]
    fn tools_serialise_to_the_shape_the_specification_wants() {
        let tool = find("list_sheets").expect("a tool");
        let json = tool.to_json();
        assert_eq!(json["name"], json!("list_sheets"));
        assert!(json["description"].is_string());
        assert_eq!(json["inputSchema"]["type"], json!("object"));
    }

    #[test]
    fn unknown_tool_names_are_not_found() {
        assert!(find("no_such_tool").is_none());
    }

    #[test]
    fn required_arguments_are_reported_by_name() {
        let args = Args::new(&json!({ "cell": "A1" }));
        let error = args.require_str("path").unwrap_err();
        assert_eq!(error, "path is required");
        assert_eq!(args.require_str("cell").unwrap(), "A1");
    }

    #[test]
    fn wrong_argument_types_are_reported_rather_than_coerced() {
        let args = Args::new(&json!({ "cell": 7, "flag": "yes", "list": "not a list" }));
        assert!(args
            .require_str("cell")
            .unwrap_err()
            .contains("must be a string"));
        assert!(args
            .require_array("list")
            .unwrap_err()
            .contains("must be an array"));
        // A string is not read as a number, so an optional number falls back to `None`.
        assert_eq!(args.opt_number("flag"), None);
    }

    #[test]
    fn a_flag_of_the_wrong_type_falls_back_to_its_default() {
        // The schema types this as a boolean, so a string is not a value worth arguing
        // about; the default is used and the model sees the documented behaviour in the
        // tool's own description.
        let args = Args::new(&json!({ "flag": "yes" }));
        assert!(args.opt_bool("flag", true));
        assert!(!args.opt_bool("flag", false));
    }

    #[test]
    fn optional_arguments_fall_back_to_their_defaults() {
        let args = Args::new(&json!({ "blank": "  " }));
        assert_eq!(args.opt_str("blank"), None);
        assert!(args.opt_bool("missing", true));
        assert!(!args.opt_bool("missing", false));
        assert_eq!(args.opt_number("missing"), None);
    }

    #[test]
    fn a_non_object_arguments_value_reads_as_empty() {
        let args = Args::new(&json!("oops"));
        assert!(args.require_str("path").unwrap_err().contains("required"));
        assert!(args.reject_unknown(find("list_sheets").unwrap()).is_ok());
    }

    #[test]
    fn unknown_arguments_are_rejected_against_the_schema() {
        let spec = find("list_sheets").unwrap();
        let args = Args::new(&json!({ "path": "a.xlsx", "shetts": [] }));
        let error = args.reject_unknown(spec).unwrap_err();
        assert!(error.contains("shetts"), "{error}");
        assert!(error.contains("list_sheets"), "{error}");
    }

    #[test]
    fn declared_arguments_are_accepted() {
        let spec = find("set_cell").unwrap();
        let args = Args::new(&json!({
            "path": "a.xlsx", "sheet": "S", "cell": "A1", "value": 1,
        }));
        assert!(args.reject_unknown(spec).is_ok());
    }

    #[test]
    fn the_style_schema_names_every_key_the_handler_reads() {
        let schema = object_style();
        let properties = schema["properties"].as_object().unwrap();
        for key in [
            "bold",
            "italic",
            "underline",
            "font_size",
            "font_name",
            "font_color",
            "fill_color",
            "number_format",
            "horizontal",
            "vertical",
            "wrap_text",
            "indent",
            "text_rotation",
            "border",
            "border_color",
        ] {
            assert!(
                properties.contains_key(key),
                "{key} is read but not declared"
            );
        }
    }
}
