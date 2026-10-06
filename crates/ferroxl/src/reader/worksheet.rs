//! Reading a single worksheet part (`openpyxl/reader/worksheet.py`).

use std::collections::btree_map::Entry;
use std::collections::HashMap;

use crate::cell::cell::{CellValue, DataType, FormulaAttributes};
use crate::cell::utils::{column_index_from_string, get_column_letter};
use crate::comments::Comment;
use crate::datavalidation::DataValidation;
use crate::exceptions::Result;
use crate::formatting::{Cfvo, ColorScale, DataBar, IconSet, Rule, RULE_ATTRIBUTES};
use crate::styles::borders::Borders;
use crate::styles::colors::Color;
use crate::styles::fills::Fill;
use crate::styles::fonts::Font;
use crate::styles::style::Style;
use crate::worksheet::dimensions::{ColumnDimension, RowDimension};
use crate::worksheet::filters::parse_bool_attr;
use crate::worksheet::Worksheet;
use crate::xml::constants::SHEET_MAIN_NS;
use crate::xml::functions::fromstring;

/// The shared state a worksheet parser needs.
pub struct WorksheetParseContext<'a> {
    /// The shared string table.
    pub string_table: &'a HashMap<usize, String>,
    /// The `cellXfs` style table.
    pub style_table: &'a [Style],
    /// The indexed colour palette.
    pub color_index: &'a [String],
    /// Whether cell types should be guessed.
    pub guess_types: bool,
    /// Whether formula cells report their cached value.
    pub data_only: bool,
}

/// Read a worksheet part into a [`Worksheet`].
pub fn read_worksheet(
    xml_source: &[u8],
    title: &str,
    index: usize,
    context: &WorksheetParseContext<'_>,
) -> Result<Worksheet> {
    let mut worksheet = Worksheet::with_title(title, index)?;
    worksheet.context.guess_types = context.guess_types;
    let root = fromstring(xml_source)?;
    let mut parser = WorksheetParser {
        worksheet: &mut worksheet,
        context,
    };
    parser.parse(&root)?;
    if !worksheet.conditional_formatting.parse_rules.is_empty() {
        let rules = std::mem::take(&mut worksheet.conditional_formatting.parse_rules);
        worksheet.conditional_formatting.update(rules);
    }
    worksheet.recount_comments();
    Ok(worksheet)
}

struct WorksheetParser<'a, 'b> {
    worksheet: &'a mut Worksheet,
    context: &'b WorksheetParseContext<'b>,
}

impl WorksheetParser<'_, '_> {
    fn tag(&self, name: &str) -> String {
        format!("{{{SHEET_MAIN_NS}}}{name}")
    }

    /// Walk the tree, dispatching on every tag the reader understands.
    ///
    /// openpyxl's reader dispatches through `iterparse(tag=...)`, which matches at any
    /// depth, so a `<col>` inside `<cols>` or a `<pane>` inside `<sheetView>` is found.
    /// This walk does the same: every element is offered to the dispatcher, and a handler
    /// that owns a subtree stops the walk from descending into it a second time.
    fn parse(&mut self, root: &crate::xml::functions::Element) -> Result<()> {
        self.parse_node(root)
    }

    fn parse_node(&mut self, node: &crate::xml::functions::Element) -> Result<()> {
        let mut descend = true;
        match node.tag.as_str() {
            t if t == self.tag("mergeCells") => {
                self.parse_merge(node);
                descend = false;
            }
            t if t == self.tag("col") => {
                self.parse_column_dimensions(node)?;
                descend = false;
            }
            // A `<row>` outside `<sheetData>` is not produced by Excel, but parse it for
            // its dimensions anyway; rows inside `<sheetData>` are handled there so their
            // cells are read too.
            t if t == self.tag("row") => {
                self.parse_row_dimensions(node)?;
                descend = false;
            }
            t if t == self.tag("sheetData") => {
                self.parse_sheet_data(node)?;
                descend = false;
            }
            t if t == self.tag("dataValidations") => {
                self.parse_data_validations(node)?;
                descend = false;
            }
            t if t == self.tag("printOptions") => {
                self.parse_print_options(node);
                descend = false;
            }
            t if t == self.tag("pageMargins") => {
                self.parse_margins(node);
                descend = false;
            }
            t if t == self.tag("pageSetup") => {
                self.parse_page_setup(node);
                descend = false;
            }
            t if t == self.tag("headerFooter") => {
                self.parse_header_footer(node);
                descend = false;
            }
            t if t == self.tag("conditionalFormatting") => {
                self.parse_conditional_formatting(node);
                descend = false;
            }
            t if t == self.tag("autoFilter") => {
                self.parse_auto_filter(node);
                descend = false;
            }
            t if t == self.tag("sheetProtection") => {
                self.parse_sheet_protection(node);
                descend = false;
            }
            // `<pane>` sits inside `<sheetView>`, so it is dispatched in its own right.
            t if t == self.tag("pane") => {
                self.parse_pane(node);
                descend = false;
            }
            t if t == self.tag("sheetView") => {
                self.parse_sheet_view(node);
                descend = false;
            }
            _ => {}
        }
        if descend {
            for child in node.children() {
                self.parse_node(child)?;
            }
        }
        Ok(())
    }

    fn parse_merge(&mut self, node: &crate::xml::functions::Element) {
        for merge in node.find_all(self.tag("mergeCell")) {
            if let Some(reference) = merge.get("ref") {
                let _ = self.worksheet.merge_cells(reference);
            }
        }
    }

    fn parse_column_dimensions(&mut self, node: &crate::xml::functions::Element) -> Result<()> {
        let min = node
            .get("min")
            .and_then(|v| v.trim().parse::<u32>().ok())
            .unwrap_or(1);
        let max = node
            .get("max")
            .and_then(|v| v.trim().parse::<u32>().ok())
            .unwrap_or(1);
        // Excel writes a single column spanning the whole sheet for defaults; skip it.
        if max == crate::xml::constants::MAX_COLUMN {
            return Ok(());
        }
        for column_id in min..=max {
            let Ok(column) = get_column_letter(column_id) else {
                continue;
            };
            let width = node
                .get("width")
                .and_then(|v| v.trim().parse::<f64>().ok())
                .unwrap_or(-1.0);
            let auto_size = node.get("bestFit") == Some("1");
            let visible = node.get("hidden") != Some("1");
            let outline_level = node
                .get("outlineLevel")
                .and_then(|v| v.trim().parse::<u32>().ok())
                .unwrap_or(0);
            let collapsed = node.get("collapsed") == Some("1");
            let style_index = node
                .get("style")
                .and_then(|v| v.trim().parse::<usize>().ok());
            if let Some(style_index) = style_index {
                if let Some(style) = self.context.style_table.get(style_index) {
                    self.worksheet.set_style(&column, style.clone())?;
                }
            }
            // A `<col>` block may split one span into several, and only the first of them carries
            // the dimensions, so an entry that is already there is left alone.
            if let Entry::Vacant(slot) = self.worksheet.column_dimensions.entry(column.clone()) {
                let mut dimension = ColumnDimension::new(&column);
                dimension.width = width;
                dimension.auto_size = auto_size;
                dimension.visible = visible;
                dimension.outline_level = outline_level;
                dimension.collapsed = collapsed;
                slot.insert(dimension);
            }
        }
        Ok(())
    }

    /// Parse a `<sheetData>` block: each `<row>` carries both dimensions and its cells.
    fn parse_sheet_data(&mut self, node: &crate::xml::functions::Element) -> Result<()> {
        for row in node.find_all(self.tag("row")) {
            self.parse_row_dimensions(row)?;
            for cell in row.find_all(self.tag("c")) {
                self.parse_cell(cell)?;
            }
        }
        Ok(())
    }

    fn parse_row_dimensions(&mut self, node: &crate::xml::functions::Element) -> Result<()> {
        let Some(row_id) = node.get("r").and_then(|v| v.trim().parse::<u32>().ok()) else {
            return Ok(());
        };
        let height = node
            .get("ht")
            .and_then(|v| v.trim().parse::<f64>().ok())
            .unwrap_or(-1.0);
        let entry = self
            .worksheet
            .row_dimensions
            .entry(row_id)
            .or_insert_with(|| RowDimension::new(row_id));
        entry.height = height;
        if let Some(hidden) = node.get("hidden") {
            entry.visible = hidden != "1";
        }
        if let Some(outline) = node
            .get("outlineLevel")
            .and_then(|v| v.trim().parse::<u32>().ok())
        {
            entry.outline_level = outline;
        }
        if let Some(collapsed) = node.get("collapsed") {
            entry.collapsed = collapsed == "1";
        }
        let style_index = node.get("s").and_then(|v| v.trim().parse::<usize>().ok());
        if node.get("customFormat").is_some() {
            if let Some(style_index) = style_index {
                if let Some(style) = self.context.style_table.get(style_index) {
                    self.worksheet
                        .set_style(&row_id.to_string(), style.clone())?;
                }
            }
        }
        Ok(())
    }

    fn parse_cell(&mut self, node: &crate::xml::functions::Element) -> Result<()> {
        let Some(coordinate) = node.get("r") else {
            return Ok(());
        };
        let style_id = node.get("s").and_then(|v| v.trim().parse::<usize>().ok());
        if let Some(style_id) = style_id {
            if let Some(style) = self.context.style_table.get(style_id) {
                self.worksheet.set_style(coordinate, style.clone())?;
            }
        }
        let raw_value = node.find_text(self.tag("v"), "");
        let formula = node.find(self.tag("f"));
        let data_type = node
            .get("t")
            .and_then(DataType::from_str)
            .unwrap_or(DataType::Numeric);

        // An inline string carries its text in `<is><t>` rather than in `<v>`, which is where
        // openpyxl puts a string when the workbook has no shared string table. Nothing else
        // supplies the value, so a cell in this form read back as empty.
        let inline = match data_type {
            DataType::InlineString => node
                .find(self.tag("is"))
                .map(|is| is.find_text(self.tag("t"), ""))
                .unwrap_or_default(),
            _ => String::new(),
        };

        if raw_value.is_empty() && inline.is_empty() && formula.is_none() {
            // An empty styled cell still needs to exist.
            self.worksheet.cell_mut(coordinate)?;
            return Ok(());
        }
        self.worksheet.cell_mut(coordinate)?;

        let mut value = match data_type {
            DataType::SharedString => CellValue::Text(
                self.context
                    .string_table
                    .get(&raw_value.trim().parse::<usize>().unwrap_or(usize::MAX))
                    .cloned()
                    .unwrap_or_else(|| raw_value.clone()),
            ),
            DataType::Bool => CellValue::Bool(raw_value.trim() == "1"),
            DataType::Error => CellValue::Error(raw_value.trim().to_string()),
            // The text is in `<is><t>`, not `<v>`, so it is read from there. A cell with no
            // `<t>` at all is empty rather than absent, which is what keeps it in the sheet.
            DataType::InlineString => CellValue::text(inline),
            // A numeric `<v>` is stored as a number so it round-trips as one.
            _ => match raw_value.trim().parse::<f64>() {
                Ok(number) => CellValue::Number(number),
                // A non-numeric body means the cell is really inline or typed text.
                Err(_) => CellValue::text(raw_value.clone()),
            },
        };

        // A formula cell's `<v>` is its cached result. It has to be taken *before* the
        // reassignment below, which is what turns `value` into the formula the caller asked
        // for -- and it is kept even then, because that result is what `recalculate` produced
        // and dropping it here would make every recalculated workbook lose its values on the
        // way back in.
        let cached_for_formula =
            if formula.is_some() && !self.context.data_only && !raw_value.is_empty() {
                Some(value.clone())
            } else {
                None
            };

        if formula.is_some() && !self.context.data_only {
            let text = formula
                .and_then(|node| node.text.clone())
                .unwrap_or_default();
            value = CellValue::Formula(format!("={text}"));
            if let Some(formula_node) = formula {
                if let Some(formula_type) = formula_node.get("t") {
                    let mut attributes = FormulaAttributes {
                        formula_type: Some(formula_type.to_string()),
                        ..FormulaAttributes::default()
                    };
                    if let Some(si) = formula_node.get("si") {
                        attributes.si = Some(si.to_string());
                    }
                    if let Some(reference) = formula_node.get("ref") {
                        attributes.reference = Some(reference.to_string());
                    }
                    self.worksheet.cell_mut(coordinate)?.formula_attributes = Some(attributes);
                }
            }
        }

        // The context is read before the mutable borrow of the worksheet is taken.
        let context = self.worksheet.context;
        let cell = self.worksheet.cell_mut(coordinate)?;
        if formula.is_some() && !self.context.data_only {
            // The formula is what the user wrote; the cached `<v>` is discarded unless the
            // caller asked for values only.
            let text = formula
                .and_then(|node| node.text.clone())
                .unwrap_or_default();
            cell.set_explicit_value(CellValue::Formula(format!("={text}")), DataType::Formula)?;
        } else if !self.context.guess_types {
            // Trust the stored type, so a numeric `<v>` becomes a number rather than text.
            cell.set_explicit_value(value, data_type)?;
        } else {
            // Type guessing infers from the text, which is what `set_value` does.
            cell.set_value(value, context);
        }
        if let Some(cached) = cached_for_formula {
            self.worksheet.set_cached_value(coordinate, cached);
        }
        Ok(())
    }

    fn parse_print_options(&mut self, node: &crate::xml::functions::Element) {
        if let Some(value) = node.get("horizontalCentered") {
            self.worksheet.page_setup.horizontal_centered = Some(value.to_string());
        }
        if let Some(value) = node.get("verticalCentered") {
            self.worksheet.page_setup.vertical_centered = Some(value.to_string());
        }
    }

    fn parse_margins(&mut self, node: &crate::xml::functions::Element) {
        let set = |setup: &mut crate::worksheet::page::PageMargins, name: &str| {
            if let Some(value) = node.get(name).and_then(|v| v.trim().parse::<f64>().ok()) {
                match name {
                    "left" => setup.left = Some(value),
                    "right" => setup.right = Some(value),
                    "top" => setup.top = Some(value),
                    "bottom" => setup.bottom = Some(value),
                    "header" => setup.header = Some(value),
                    _ => setup.footer = Some(value),
                }
            }
        };
        for name in ["left", "right", "top", "bottom", "header", "footer"] {
            set(&mut self.worksheet.page_margins, name);
        }
    }

    fn parse_page_setup(&mut self, node: &crate::xml::functions::Element) {
        let setup = &mut self.worksheet.page_setup;
        if let Some(value) = node.get("orientation") {
            setup.orientation = Some(value.to_string());
        }
        if let Some(value) = node.get("paperSize") {
            setup.paper_size = Some(value.to_string());
        }
        if let Some(value) = node.get("scale") {
            setup.scale = Some(value.to_string());
        }
        if let Some(value) = node.get("fitToPage") {
            setup.fit_to_page = Some(value.to_string());
        }
        if let Some(value) = node.get("fitToHeight") {
            setup.fit_to_height = Some(value.to_string());
        }
        if let Some(value) = node.get("fitToWidth") {
            setup.fit_to_width = Some(value.to_string());
        }
        if let Some(value) = node.get("firstPageNumber") {
            setup.first_page_number = Some(value.to_string());
        }
        if let Some(value) = node.get("useFirstPageNumber") {
            setup.use_first_page_number = Some(value.to_string());
        }
    }

    fn parse_header_footer(&mut self, node: &crate::xml::functions::Element) {
        if let Some(text) = node
            .find(self.tag("oddHeader"))
            .and_then(|n| n.text.clone())
        {
            self.worksheet.header_footer.set_header(&text);
        }
        if let Some(text) = node
            .find(self.tag("oddFooter"))
            .and_then(|n| n.text.clone())
        {
            self.worksheet.header_footer.set_footer(&text);
        }
    }

    fn parse_sheet_protection(&mut self, node: &crate::xml::functions::Element) {
        self.worksheet.protection.enabled = true;
        if let Some(password) = node.get("password") {
            self.worksheet
                .protection
                // The file already stores the hash, so it must not be hashed again.
                .set_password(password, true);
        }
        let set = |protection: &mut crate::worksheet::protection::SheetProtection, name: &str| {
            if let Some(value) = node.get(name) {
                let flag = parse_bool_attr(Some(value));
                match name {
                    "sheet" => protection.sheet = flag,
                    "objects" => protection.objects = flag,
                    "scenarios" => protection.scenarios = flag,
                    "formatCells" => protection.format_cells = flag,
                    "formatColumns" => protection.format_columns = flag,
                    "formatRows" => protection.format_rows = flag,
                    "insertColumns" => protection.insert_columns = flag,
                    "insertRows" => protection.insert_rows = flag,
                    "insertHyperlinks" => protection.insert_hyperlinks = flag,
                    "deleteColumns" => protection.delete_columns = flag,
                    "deleteRows" => protection.delete_rows = flag,
                    "selectLockedCells" => protection.select_locked_cells = flag,
                    "sort" => protection.sort = flag,
                    "autoFilter" => protection.auto_filter = flag,
                    "pivotTables" => protection.pivot_tables = flag,
                    _ => protection.select_unlocked_cells = flag,
                }
            }
        };
        let protection = &mut self.worksheet.protection;
        for name in [
            "sheet",
            "objects",
            "scenarios",
            "formatCells",
            "formatColumns",
            "formatRows",
            "insertColumns",
            "insertRows",
            "insertHyperlinks",
            "deleteColumns",
            "deleteRows",
            "selectLockedCells",
            "sort",
            "autoFilter",
            "pivotTables",
            "selectUnlockedCells",
        ] {
            set(protection, name);
        }
    }

    fn parse_pane(&mut self, node: &crate::xml::functions::Element) {
        if node.get("state") == Some("frozen") {
            if let Some(top_left) = node.get("topLeftCell") {
                self.worksheet.set_freeze_panes(top_left);
            }
        }
    }

    fn parse_sheet_view(&mut self, node: &crate::xml::functions::Element) {
        // The pane and the selection are children of `<sheetView>`, and the walk stops at
        // this element, so both are reached from this one handler.
        if let Some(pane) = node.find(self.tag("pane")) {
            self.parse_pane(pane);
        }
        // The first selection is the active one, so a frozen sheet reports the cell the
        // cursor was in.
        if let Some(selection) = node.find(self.tag("selection")) {
            if let Some(active) = selection.get("activeCell") {
                self.worksheet.active_cell = active.to_string();
            }
            if let Some(selected) = selection.get("sqref") {
                if !selected.is_empty() {
                    let first = selected.split_whitespace().next().unwrap_or(selected);
                    self.worksheet.selected_cell = first.to_string();
                }
            }
        }
    }

    /// Read every `<dataValidation>` rule in a `<dataValidations>` block.
    fn parse_data_validations(&mut self, node: &crate::xml::functions::Element) -> Result<()> {
        for child in node.find_all(self.tag("dataValidation")) {
            let validation = DataValidation::read(child)?;
            self.worksheet.data_validations.push(validation);
        }
        Ok(())
    }

    fn parse_conditional_formatting(&mut self, node: &crate::xml::functions::Element) {
        let Some(range_string) = node.get("sqref") else {
            return;
        };
        if range_string.is_empty() {
            return;
        }
        let mut rules = Vec::new();
        for rule_node in node.find_all(self.tag("cfRule")) {
            let Some(rule_type) = rule_node.get("type") else {
                continue;
            };
            // Data bars used to be skipped here, on the belief that they need a drawing
            // extension openpyxl does not write. That is true only of the extra properties --
            // gradient fill, border, negative-bar colour, axis. The bar itself is ordinary
            // cfRule content, so skipping it dropped the rule from a loaded workbook entirely.
            let mut rule = Rule::new(rule_type);
            for attribute in RULE_ATTRIBUTES {
                if let Some(value) = rule_node.get(attribute) {
                    if attribute == "priority" {
                        rule.attributes
                            .insert(attribute.to_string(), value.trim().to_string());
                    } else {
                        rule.attributes
                            .insert(attribute.to_string(), value.to_string());
                    }
                }
            }
            for formula in rule_node.find_all(self.tag("formula")) {
                rule.formula.push(formula.text.clone().unwrap_or_default());
            }
            if let Some(color_scale) = rule_node.find(self.tag("colorScale")) {
                let mut scale = ColorScale::default();
                for cfvo in color_scale.find_all(self.tag("cfvo")) {
                    scale.cfvo.push(self.parse_cfvo(cfvo));
                }
                for color in color_scale.find_all(self.tag("color")) {
                    scale.color.push(self.parse_rule_color(color));
                }
                rule.color_scale = Some(scale);
            }
            if let Some(icon_set) = rule_node.find(self.tag("iconSet")) {
                let mut set = IconSet {
                    icon_set: icon_set.get("iconSet").map(|v| v.to_string()),
                    show_value: icon_set.get("showValue").map(|v| v.to_string()),
                    reverse: icon_set.get("reverse").map(|v| v.to_string()),
                    percent: icon_set.get("percent").map(|v| v.to_string()),
                    cfvo: Vec::new(),
                };
                for cfvo in icon_set.find_all(self.tag("cfvo")) {
                    set.cfvo.push(self.parse_cfvo(cfvo));
                }
                rule.icon_set = Some(set);
            }
            // A data bar was not read at all, so a workbook using one loaded with the rule
            // present but the bar gone: no error, no warning, and a cell that had a bar now
            // has a rule that does nothing visible.
            if let Some(data_bar) = rule_node.find(self.tag("dataBar")) {
                let mut bar = DataBar {
                    show_value: data_bar.get("showValue").map(|v| v != "0" && v != "false"),
                    min_length: data_bar
                        .get("minLength")
                        .and_then(|v| v.trim().parse::<u32>().ok()),
                    max_length: data_bar
                        .get("maxLength")
                        .and_then(|v| v.trim().parse::<u32>().ok()),
                    ..DataBar::default()
                };
                for cfvo in data_bar.find_all(self.tag("cfvo")) {
                    bar.cfvo.push(self.parse_cfvo(cfvo));
                }
                if let Some(color) = data_bar.find(self.tag("color")) {
                    bar.color = self.parse_rule_color(color).index;
                }
                rule.data_bar = Some(bar);
            }
            rules.push(rule);
        }
        self.worksheet
            .conditional_formatting
            .parse_rules
            .entry(range_string.to_string())
            .or_default()
            .extend(rules);
    }

    /// Read one `<cfvo/>`, wherever it appears.
    ///
    /// The same element serves a colour scale, an icon set and a data bar, and `gte` is on
    /// all three. It was read as absent, which is not the same as true: the schema defaults it
    /// to true, so a rule written with `gte="0"` came back excluding its boundary value.
    fn parse_cfvo(&self, node: &crate::xml::functions::Element) -> Cfvo {
        Cfvo {
            cfvo_type: node.get("type").map(|v| v.to_string()),
            val: node.get("val").map(|v| v.to_string()),
            gte: node.get("gte").map(|v| v != "0" && v != "false"),
        }
    }

    fn parse_rule_color(&self, node: &crate::xml::functions::Element) -> Color {
        let mut color = Color::new(Color::BLACK);
        if let Some(indexed) = node
            .get("indexed")
            .and_then(|v| v.trim().parse::<usize>().ok())
        {
            if let Some(value) = self.context.color_index.get(indexed) {
                color = Color::new(value.clone());
            }
        }
        if let Some(theme) = node.get("theme") {
            match node.get("tint") {
                Some(tint) => color = Color::new(format!("theme:{theme}:{tint}")),
                None => color = Color::new(format!("theme:{theme}:")),
            }
        } else if let Some(rgb) = node.get("rgb") {
            color = Color::new(rgb);
        }
        color
    }

    fn parse_auto_filter(&mut self, node: &crate::xml::functions::Element) {
        if let Some(reference) = node.get("ref") {
            self.worksheet.auto_filter.set_reference(reference);
        }
        for column in node.find_all(self.tag("filterColumn")) {
            let Some(filters) = column.find(self.tag("filters")) else {
                continue;
            };
            let values: Vec<String> = filters
                .find_all(self.tag("filter"))
                .into_iter()
                .filter_map(|node| node.get("val").map(|v| v.to_string()))
                .collect();
            let blank = parse_bool_attr(filters.get("blank"));
            let col_id = column
                .get("colId")
                .and_then(|v| v.trim().parse::<u32>().ok())
                .unwrap_or(0);
            self.worksheet
                .auto_filter
                .add_filter_column(col_id, values, blank);
        }
        for condition in node.find_all(self.tag("sortCondition")) {
            let reference = condition.get("ref").unwrap_or_default().to_string();
            let descending = parse_bool_attr(condition.get("descending"));
            self.worksheet
                .auto_filter
                .add_sort_condition(&reference, descending);
        }
    }
}

/// Attach a parsed comment to a cell on a worksheet.
pub fn assign_comment(worksheet: &mut Worksheet, coordinate: &str, comment: Comment) -> Result<()> {
    worksheet.set_comment(coordinate, Some(comment))
}

/// Convert a parsed column name to its index, for callers outside the parser.
pub fn resolve_column(letters: &str) -> Result<u32> {
    column_index_from_string(letters)
}

/// A font with all defaults, exposed for callers assembling conditional-format styles.
pub fn default_font() -> Font {
    Font::new()
}

/// A fill with all defaults, exposed for callers assembling conditional-format styles.
pub fn default_fill() -> Fill {
    Fill::new()
}

/// A borders object with all defaults, exposed for callers assembling conditional-format styles.
pub fn default_borders() -> Borders {
    Borders::new()
}
