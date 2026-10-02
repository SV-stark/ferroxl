# Feature parity with openpyxl 3.1.5

ferroxl is a port of [openpyxl](https://github.com/theorchard/openpyxl) 3.1.5. This document
records, module by module, what has been implemented, what has not, and where the Rust
version deliberately behaves differently.

It is written from the source trees, not from memory. `tools/parity.py` walks every module
of the Python original, collects its public names, and reports which have a ferroxl
counterpart:

```console
$ python tools/parity.py path/to/openpyxl/openpyxl
```

Re-run it after any change. The numbers below come from that run.

- [How to read this](#how-to-read-this)
- [Summary](#summary)
- [Implemented](#implemented)
- [Pending](#pending)
- [Different by design](#different-by-design)
- [Packages with no Rust counterpart at all](#packages-with-no-rust-counterpart-at-all)
- [Not ported](#not-ported)
- [How parity is verified](#how-parity-is-verified)

## How to read this

Three labels are used, and they mean different things:

- **Implemented** — the behaviour exists and is covered by a test.
- **Pending** — openpyxl has it, ferroxl does not. Each entry says what the effect is and
  what a caller does instead.
- **Different by design** — ferroxl does it, but not the way Python does. Each entry says
  why, because a difference you did not choose is a bug.

## Summary

The reference is **openpyxl 3.1.5**, the tree at the path recorded in `tools/parity.py`'s
invocation below. This document previously described a 1.9-era surface and understated the
gap by a wide margin; the numbers here are from a run against 3.1.5.

| | |
| --- | --- |
| openpyxl top-level names audited | 994 |
| Modules whose every public name is matched | 32 of 183 |
| Names with no ferroxl counterpart | 679 |

The count fell from 729 to 679 without 50 features being written. `styles/builtins.py`
exposes 51 module-level string literals, and openpyxl names them after their *variables*
(`accent_1_20`) while calling the style something else entirely (`20 % - Accent1`), so no
name-based matcher can ever line them up. ferroxl holds the same 49 as one Rust table,
`BUILTIN_DETAILS`, and `tools/parity.py` now matches on `builtinId` after checking the table
is still 49 entries long -- so the coverage is verified against the source rather than
asserted here. The one name it cannot account for, `pandas_highlight`, is listed below as a
real gap, because it is one.

Unmatched names, by upstream package:

| Package | Unmatched | What it is |
| --- | --- | --- |
| `drawing/` | 126 | Shape geometry and the `spPr` tree; ferroxl has image sizing only |
| `worksheet/` | 117 | Views, filters, OLE, scenarios, print ranges, array formulas |
| `styles/` | 100 | Named styles, `StyleArray`, table styles, dxf extras |
| `chart/` | 81 | Data labels, trendlines, layouts, rich-text titles, the chart reader |
| `pivot/` | 58 | The whole pivot table and pivot cache model |
| `descriptors/` | 49 | The `Serialisable` base and the typed descriptor system |
| `xml/` | 41 | Namespace registration, `iterparse`, tag constants |
| `utils/` | 35 | `FORMULAE`, escaping, `IndexedList`, dataframe bridge, open-ended ranges |
| `workbook/` | 35 | Calculation properties, book views, external links, file sharing, web options |
| `packaging/` | 34 | Manifest, relationship and content-type construction |
| `cell/` | 24 | Rich text, phonetic text, inline fonts |
| `chartsheet/` | 11 | Sheets whose only content is a chart |
| `formatting/` | 7 | Data bars, icon-set rules, `cfvo/@gte`, `Rule/@timePeriod` |
| `comments/` | 5 | Comment shape and sizing details |
| `reader/` | 5 | `SharedStrings`, `ExcelReader`, rich-text reading |
| `formula/` | 3 | The tokenizer and `FORMULAE` — `Translator` **is** ported |

Three further gaps are *method-level* and do not appear in a top-level-name audit:
[`Worksheet::range()` with offsets](#2-worksheet-range-with-offsets-rows-and-columns),
[no read-only loader](#3-there-is-no-read-only-loader), and
[charts not read back](#6-charts-and-images-are-written-but-not-read-back).

The honest summary of where ferroxl stands: **cell values, styles, and the read/write round
trip are solid; everything around the cell is thin.** A workbook with pivot tables, named
styles, gradient fills, rich text or chart-only sheets in it will lose those parts on save.
That is not a rounding error against "feature parity" — it is a materially different library
from openpyxl, and this document exists to say so precisely rather than approximately.

## Implemented

Everything in this section is covered by tests. Names are ferroxl's; the upstream module is
named so a change can be traced back to the Python it mirrors.

### `cell` — `openpyxl/cell`

| Upstream | ferroxl |
| --- | --- |
| `cell.py::Cell` | `cell::Cell` — `coordinate`, `row`, `column`, `value`, `data_type`, `style_id`, `hyperlink` |
| `cell.py::CellValue` | `cell::CellValue`, with `text`, `number`, `formula`, `is_empty`, `as_text`, `as_number`, `is_temporal` |
| `cell.py::DataType` | `cell::DataType` plus `VALID_TYPES` and `ERROR_CODES` |
| `cell.py` type guessing | `Cell::set_value`, `set_explicit_value`, `data_type_for_value`, `cast_numeric`, `cast_percentage`, `cast_time`, `display_value` |
| `cell.py::check_string` | `cell::check_string` |
| `cell.py::CellContext` | `cell::CellContext` — the style table, base date and guess-types flag |
| `formula.py` | `formula::FormulaStore`, `SharedFormula`, `shift_references`, `FormulaStore::expand` |
| `read_only.py::ReadOnlyCell` | `cell::ReadOnlyCell`, `ReadOnlyTables`, `empty_cell` |
| `__init__.py` helpers | `coordinate_from_string`, `absolute_coordinate`, `get_column_letter`, `column_index_from_string`, `coordinate_from_index`, `split_coordinate` |

### `charts` — `openpyxl/charts`

| Upstream | ferroxl |
| --- | --- |
| `chart.py::ChartBase` | `charts::Chart` — `add_series`, `add_shape`, `with_title`, margins, `y_chars` |
| `bar.py`, `line.py`, `pie.py`, `scatter.py` | `BarChart`, `LineChart`, `PieChart`, `ScatterChart`, all newtypes over `GraphChart` or `Chart`, each with `into_chart()` |
| `axis.py::TextAxis`, `NumericAxis`, `CategoryAxis`, `ValueAxis` | `Axis`, `CategoryAxis`, `ValueAxis`, `AxisScale`, `less_than_one` |
| `series.py::Series` | `charts::Series` — `min`, `max`, `len`, `is_empty`, `values_for`, `with_title`, `with_xvalues`, `with_labels`, `with_color`, `with_error_bar` |
| `reference.py::Reference` | `charts::Reference`, `ReferenceDataType`, and the chart cache's own `CellValue` |
| `legend.py::Legend` | `charts::Legend` |
| `error_bar.py::ErrorBars` | `charts::ErrorBar`, `ErrorBarType` |
| `__init__.py` chart factories | `BarChart::new()` and friends |

Charts are **written but not read back** — see [Pending](#6-charts-and-images-are-written-but-not-read-back).

### `comments` — `openpyxl/comments`

`comments::Comment` with `text` and `author`, plus the writer that emits both
`xl/comments1.xml` and the VML shape part Excel needs (`write_comments`,
`write_comments_vml`, `collect_comments`, `author_ids`).

### `datavalidation` — `openpyxl/datavalidation`

`DataValidation` with `ValidationType`, `ValidationOperator`, `ValidationErrorStyle` and
`ImeMode`, plus `add_cell`, `set_error_message`, `set_prompt_message`, `collapse_cell_addresses`
and `attributes`.

### `date_time` — `openpyxl/date_time`

`to_excel`, `from_excel`, `date_to_excel`, `time_to_days`, `time_to_days_datetime`,
`timedelta_to_days`, `micros_to_days`, `days_to_time`, `datetime_to_w3cdtf`,
`w3cdtf_to_datetime`, `SharedDate`, `BaseDate`, `ExcelDateTime`, and the Julian helpers
`gcal2jd`, `jd2gcal`, `jd2gcal_int`.

`jdcal` is no longer on PyPI, so `gcal2jd`/`jd2gcal` are reimplemented here and pinned
against `jdcal`'s own output in the tests.

### `drawing` — `openpyxl/drawing`

`drawing::Image`, `Drawing`, `AnchorType`, `Shape`, `ShapeStyle`, `Shadow`, the
`SHADOW_*` constants, `bounding_box`, `norm_pct`, `read_png_dimensions`,
`validate_anchor_column`, `column_letters` and `column_index`.

Image dimensions come from the PNG header rather than PIL — see
[Different by design](#image-dimensions-come-from-a-png-header).

### `exceptions` — `openpyxl/exceptions`

All twelve openpyxl exception classes are `Error` variants with the same names, plus
`Value`, `Type`, `Key`, `Attribute`, `BadZipFile`, `Xml`, `Io` and `NotImplemented`.
`Result<T>` is the crate's own alias.

### `formatting` — `openpyxl/formatting`

`Rule` with its `RULE_ATTRIBUTES` allow-list, `Cfvo`, `ColorScale`, `IconSet`, `DxfStyle`,
`ConditionalFormatting`, `StyleProperties`, and the builders `CellIsRule`, `FormulaRule`,
`ColorScaleRule`. `OPERATOR_EXPANSION`, `COLOR_SCALE_VALID_TYPES` and `ICON_ATTRIBUTES`
match the upstream tables.

### `namedrange` — `openpyxl/namedrange`

`NamedRange`, `NamedRangeContainingValue`, the `DefinedName` enum, `split_named_range` and
`refers_to_range`.

### `reader` — `openpyxl/reader`

The whole load path: `load_workbook`, `load_workbook_from_bytes`, `LoadOptions` with
`guess_types`, `data_only` and `keep_vba`, plus the `new` / `guessing_types` /
`values_only` / `keeping_vba` builders; `WorkbookSource`, `package_bytes`,
`read_string_table`, `read_style_table`, `read_worksheet`, `read_comments`,
`comments_file_path`, `read_sheets`, `read_rels`, `read_content_types`,
`read_properties_core`, `read_excel_base_date`, `read_workbook_settings`,
`read_named_ranges`, `title_resolver`, `detect_worksheets`.

Three of openpyxl 3.1.5's six `load_workbook` parameters are implemented, and the three
that are not are not interchangeable with the one ferroxl has that upstream dropped:

| Parameter | ferroxl |
| --- | --- |
| `keep_vba` | `keep_vba` |
| `data_only` | `data_only` |
| `read_only` | **absent** — needs the streaming loader, see [Pending](#3-there-is-no-read-only-loader) |
| `keep_links` | **absent** — no external-link parts are read or written |
| `rich_text` | **absent** — no inline-runs model in cells |
| *(removed in 3.0)* `guess_types` | `guess_types` — a 2.x flag with no 3.x counterpart |

The worksheet reader walks the XML tree recursively rather than matching a tag list, which
matches openpyxl's `iterparse(tag=...)` behaviour: a `<col>` inside `<cols>` and a `<pane>`
inside `<sheetView>` are both found.

### `styles` — `openpyxl/styles`

`Style`, `Font`, `Fill`, `Borders`, `Border`, `Alignment`, `Protection`, `ProtectionFlag`,
`Color`, `NumberFormat`, plus `COLOR_INDEX`, `BUILTIN_FORMATS`, the `FORMAT_*` constants,
`is_builtin`, `is_date_format`, `Style::sort_key`, `same_visual_style`, `defaults`,
`static_style` and `copy_style`.

### `units` — `openpyxl/units`

`points_to_pixels`, `pixels_to_points`, `pixels_to_emu`, `emu_to_pixels`, `emu_to_cm`,
`cm_to_emu`, `emu_to_inch`, `inch_to_emu`, `cm_to_dxa`, `dxa_to_cm`, `inch_to_dxa`,
`dxa_to_inch`, `degrees_to_angle`, `angle_to_degrees`, `short_color`, and the
`DEFAULT_*` / `BASE_COL_WIDTH` constants.

### `workbook` — `openpyxl/workbook`

`Workbook` with `create_sheet`, `create_sheet_at`, `add_sheet`, `remove_sheet`, `active`,
`set_active`, `active_sheet`, `active_sheet_mut`, `get_sheet_by_name`, `get_index`,
`get_sheet_names`, `cell_context`, `excel_base_date`, `add_named_range`, `add_named_value`,
`get_named_ranges`, `save`, `to_bytes`, plus `DocumentProperties` and `DocumentSecurity`.

### `worksheet` — `openpyxl/worksheet`

`Worksheet` with cells, `row_dimensions`, `column_dimensions`, `merged_cells`, `charts`,
`images`, `relationships`, `data_validations`, `conditional_formatting`, `page_setup`,
`page_margins`, `header_footer`, `protection`, `auto_filter`, `freeze_panes`, `sheet_state`,
`show_gridlines` and `print_gridlines`, plus `append`, `append_rows`,
`calculate_dimension`, `highest_row`, `highest_column`, `range_values` and
`range_coordinates`.

`RowDimension`, `ColumnDimension`, `Dimension`, `AutoFilter`, `FilterColumn`,
`SortCondition`, `HeaderFooter`, `HeaderFooterItem`, `PageSetup`, `PageMargins`,
`SheetProtection`, `Relationship`, `RelationshipType`, `hash_password`, `RangeBounds`,
`SheetDimensions`, `get_range_boundaries`, and the `BREAK_*`, `SHEETSTATE_*`,
`ORIENTATION_*` and `PAPERSIZES` tables.

### `writer` — `openpyxl/writer`

`ExcelWriter`, `save_workbook`, `save_virtual_workbook`, `save_workbook_to`,
`write_worksheet`, `write_worksheet_rels`, `write_workbook`, `write_workbook_rels`,
`write_content_types`, `write_properties_core`, `write_properties_app`, `write_root_rels`,
`write_style_table`, `build_style_tables`, `StyleTables`, `StyleId`, `write_string_table`,
`create_string_table`, `StringTableBuilder`, `write_theme`, `write_comments`,
`write_comments_vml`, `write_drawing`, `write_drawing_rels`, `write_shapes`, `write_chart`,
`write_chart_rels`, `STATIC_CONTENT_TYPES`, `THEME_XML`, `FIRST_CUSTOM_FORMAT_ID`.

### `xml` — `openpyxl/xml`

`Element` with `find`, `find_all`, `find_text`, `iter_tag`, `children`, `insert`, `append`;
`XmlWriter` with `start_tag`, `end_tag`, `tag`, `raw`; `fromstring` with real namespace
resolution; `serialize`, `to_pretty_string`, `conditional_element`, `safe_string`,
`repr_float`, `escape_text`, `escape_attribute`, `REGISTERED_PREFIXES`, and all the
namespace and archive-path constants from `xml/constants.rs`.

### The `ferroxl-mcp` server

Not part of openpyxl. 33 tools over JSON-RPC 2.0 on stdio, grouped as reading (9),
structure (9), values (4), layout and appearance (8) and drawing (3). See the
[README](README.md#the-mcp-server).

## Pending

Six areas where openpyxl has something ferroxl does not, ordered by how likely they are to
matter.

### 1. ~~The streaming writer (writer/dump_worksheet.py)~~ — shipped in 0.1.2

`DumpWorksheet` and `save_dump` are ported. `write_worksheet` is split into a head, the
rows and a tail so both writers share the serialisation, and the two are tested against
each other because a streaming writer that emitted different XML would be worse than none.

### 2. ~~Worksheet range with offsets, rows and columns~~ — shipped in 0.1.4

`worksheet::cell_range::CellRange` is openpyxl's `worksheet.cell_range.CellRange` as a value:
`intersection`, `union`, `issubset`, `issuperset`, `isdisjoint`, `contains`, `shift`,
`expand`, `shrink`, `size`, `top`/`bottom`/`left`/`right`, `rows`, `cols` and `cells`. Bounds
are inclusive at both ends, which is the opposite of `RangeBounds` and the easiest
off-by-one in the area.

`MultiCellRange` is the `sqref` collection — several disjoint rectangles, as conditional
formatting applies to. It keeps overlapping ranges separately rather than merging them: a
caller asking "which rules apply to B2" wants to know they asked twice.

`Worksheet::range(range, row_offset, column_offset)` and `Worksheet::range_cells` are
`ws["A1:C3"]` and friends. The offsets expand a bare coordinate into a rectangle and shift a
range that already has a colon; the argument's shape decides which, because deciding from the
offsets would make `ws["A1:B2", 1, 1]` two plausible things.

### 3. There is no read-only loader

openpyxl's `load_workbook(..., read_only=True)` hands back a `ReadOnlyWorksheet`, which
parses a sheet's XML lazily and yields `ReadOnlyCell`s one at a time without building a
`Worksheet`. That is how openpyxl reads a hundred-megabyte sheet.

> The 2.x spelling of this flag was `use_iterators`; 3.x renamed it to `read_only` and
> `use_iterators` no longer exists. The gap is the same one either name.

ferroxl has the value type — `cell::ReadOnlyCell`, with `coordinate`, `internal_value`,
`number_format`, `is_date`, `value` and `datetime`, plus the `ReadOnlyTables` it resolves
against — and both are public and constructible. Nothing in `load_workbook` produces them:
there is no `read_only` option and no iterator over a worksheet's cells.

`ReadOnlyCell` is also missing the style accessors openpyxl's version has: `style_array`,
`has_style`, `font`, `fill`, `border`, `alignment`, `protection`. It exposes values and
number formats, not formatting.

**Effect.** Reading a workbook materialises every sheet. For the sheets an agent typically
opens this does not matter; for a very large one it would.

**What to do instead.** Nothing today. `ReadOnlyCell` is exercised by tests but is not
reachable from a loaded workbook.

### 4. The stored dimension element is not read

openpyxl's read-only path reads `<dimension ref="A1:D10">` and trusts it, falling back to
the cells when the element is absent. ferroxl computes the dimension from the cells it read
(`calculate_dimension`) and stores nothing.

**Effect.** For a workbook whose stored `<dimension>` disagrees with its cells, the two
report different extents. Excel and openpyxl agree in practice because Excel writes the
dimension correctly, so the divergence only shows on a hand-edited or corrupt file.
`highest_row` and `highest_column` come from the dimension tables rather than the stored
element, so they move with the cells.

### 5. Zip central directory repair

**Resolved.** This entry was wrong in both directions, and measuring it is what showed that.

The gap was described as openpyxl scanning for the end-of-central-directory signature to
tolerate junk appended after it. ferroxl never needed that: the zip crate locates the record by
scanning backwards, so a 4 kB tail of zeroes, a stray `<html>404</html>`, and even a *prepended*
UTF-8 BOM all load. There was never a defect here, and the entry described a bug that did not
exist.

What was actually missing was the neighbouring case: an end-of-central-directory record cut off
by an interrupted download. The central directory is written before that record, so such a file
is reconstructible, and `reader::archive` now walks the directory and rebuilds the missing 22
bytes. `crates/ferroxl/tests/archive_check.rs` covers it end to end.

**Effect.** A workbook whose trailing record was truncated now loads, cells and all. A file
truncated *into* the central directory is still refused, because no record can describe it, and
a wrong one would produce a file that opens and shows the wrong sheets.

### 6. Charts and images are written but not read back

openpyxl 3.1.5 writes chart and drawing parts but its reader does not parse them, so a
reloaded workbook reports no charts and no images. ferroxl matches this rather than being
half-compatible in a different direction.

**Effect.** `worksheet.charts` and `worksheet.images` are empty after a load, even though
the parts are in the file. Anything that depends on reading a chart back — inspecting an
existing chart, preserving one across an edit — does not work, in ferroxl or in openpyxl 3.1.5.

This one is a *parity* gap rather than a *capability* gap: ferroxl behaves exactly as the
reference does.

## Different by design

Places where ferroxl does the job but not Python's way. Each is also documented at the call
site, so a reader of the code does not have to come here first.

### Equality and hashing (styles/hashable.py)

openpyxl's `HashableObject` compares a style by the fields that matter and falls back to
identity for the rest, so a number format hashes by its code and two visually identical
styles collapse to one style-table entry.

Rust's `HashMap` requires `Hash` and `Eq`, so ferroxl derives them but implements them using
the comparison Python uses, with `same_visual_style` and `Style::sort_key` doing the work
`HashableObject` does. The `HashableObject` *type* does not exist.

The visible consequence is the one you would want: a style table does not accumulate
near-duplicate entries.

### Exceptions are an enum, not a class hierarchy

Twelve Python classes you could name in an `except` clause become twelve `Error` variants.
Catching by type in Rust means matching the enum.

Every variant's `Display` string starts with the Python class name, so text that used to say
`SheetTitleException: ...` still does.

### Tri-state attributes and ProtectionFlag

OOXML attributes like `<protection locked>` are genuinely three-valued: present-and-true,
present-and-false, and absent. Python spells the third case `None`; Rust's `bool` has no
such value. ferroxl uses `ProtectionFlag::{Inherit, Locked, Unlocked}` for those attributes
and a plain `bool` where the specification really does mean two states.

### CategoryAxis and ValueAxis are constructors

In Python they are classes that differ only in their class-level attribute defaults. In
ferroxl they are unit structs whose `new()` returns an `Axis` configured the way the
corresponding Python class would configure itself — which is why `new` does not return
`Self`, and why that carries an explicit `#[allow(clippy::new_ret_no_self)]`.

### Image dimensions come from a PNG header

openpyxl opens an image with PIL to read its size, which makes PIL a hard dependency for
anyone embedding a picture. ferroxl parses the PNG header directly, so PNG is the only
supported format and PIL is not a dependency at all. The bytes are stored verbatim, so the
image itself is untouched.

### Names that were renamed

Renaming is not a behaviour change, but it will trip up someone porting code by search and
replace:

| openpyxl | ferroxl | Why |
| --- | --- | --- |
| `pixels_to_EMU`, `EMU_to_pixels`, `cm_to_EMU`, … | `pixels_to_emu`, `emu_to_pixels`, `cm_to_emu`, … | Rust naming convention |
| `PAPERSIZE_LETTER` … `PAPERSIZE_A5`, eleven constants | `PAPERSIZES`, a table of `(name, code)` pairs | Eleven constants are a table |
| `worksheet.add_chart(chart, anchor)` | `worksheet.charts.push(chart)` — the anchor is a field on the chart | The `Vec` is public |
| `worksheet.add_image(image, anchor)` | `worksheet.images.push(image)` | Same |
| `worksheet.add_rel(...)` | `worksheet.relationships.push(...)` | Same |
| `ws.max_row`, `ws.max_column` | `ws.highest_row()`, `ws.highest_column()` | Matches openpyxl 3.1.5's `get_highest_row` |
| `cell.value` | `cell.internal_value()`, or `worksheet.cell_value(coord)` | A method, so the type cast runs |
| `StyleWriter`, `ChartWriter`, `CommentWriter`, `DrawingWriter`, `ShapeWriter` | `write_style_table`, `write_chart`, `write_comments`, `write_drawing`, `write_shapes` | See below |
| `load_workbook(..., read_only=True)` | `LoadOptions::guessing_types()` / `values_only()` / `keeping_vba()` | Builders instead of keyword arguments |
| `ExcelWriter(workbook).save(path)` | `Workbook::save(path)` | See below |

### Writer classes became functions

openpyxl's writers are classes with `write()` and `close()` methods, which lets them stream
through an open file handle. ferroxl's writers build a `String` and hand it to the archive,
so there is no handle to own and no state worth capturing — `write_worksheet(...)`,
`write_workbook(...)`, `write_chart(...)` and the rest are free functions.

`ExcelWriter` does survive, because it holds the workbook, the string table and the style
tables; only its `write()`/`close()` pair collapsed into `to_bytes()` and
`save_workbook_to(...)`. This is also why the streaming writer in
[Pending](#1-the-streaming-writer-writerdump_worksheetpy) is a real gap rather than a
refactor: the function-shaped writers have nowhere to put a stream.

### Style tables are sorted before they are written

openpyxl keeps the style table in insertion order. ferroxl sorts by `Style::sort_key` so two
runs that build the same set of styles produce byte-identical output. Excel does not care
about the order; a diff does.

## Packages with no Rust counterpart at all

These are whole upstream packages with no module to point at. They are the largest part of
the gap and the reason the headline number above is as bad as it is.

### `pivot/` — 58 classes, ~3,700 lines

The pivot table and pivot cache model: `TableDefinition`, `CacheDefinition`, `CacheField`,
`SharedItems`, `PivotField`, `DataField`, `PageField`, `FieldGroup`, `RecordList`,
`CacheSource`, `WorksheetSource`, and 40 more. ferroxl contains the string
`PivotStyleLight16` as a default style name and a `sheetProtection/@pivotTables` flag,
which is the entirety of its pivot surface.

**Effect.** A pivot table cannot be created, inspected or edited. Worse, because the parts
are not preserved, **saving a loaded workbook silently destroys every pivot table and pivot
cache it contained.** This is the most damaging single gap in this document: it is data
loss on a round trip, not a missing feature.

### `chartsheet/` — 11 classes, ~760 lines

`Chartsheet` and its views, properties, protection and relation types: a workbook sheet
whose entire content is one chart, written to `xl/chartsheets/sheetN.xml`. ferroxl's
`Workbook` has `worksheets` and no `chartsheets`; the reader never looks in
`xl/chartsheets/`.

**Effect.** A chart-only sheet cannot be created, and any in a loaded file is dropped on
save. Same round-trip loss as pivot tables.

### `descriptors/` — 49 names

openpyxl's metaprogramming layer: `Serialisable` with `to_tree`/`from_tree`, the typed
descriptors (`Integer`, `Float`, `Bool`, `String`, `Set`, `NoneSet`, `MinMax`, `DateTime`),
the `Sequence` family, the `Nested` family, and `excel.py`'s `HexBinary`, `TextPoint`,
`Percentage`, `Extension`, `ExtensionList`, `Guid`, `Base64Binary`.

There is no trait, macro or derive in ferroxl that mirrors this. Every Rust type is a
hand-written struct with a hand-written attribute list and a hand-written parse function.

**Effect.** The port cannot express a new OOXML construct without hand-writing both halves
of it, so anything not explicitly modelled for the ~20 supported types is dropped without
error. In particular `extLst` is not read or written anywhere.

This is a different *kind* of gap from the others. It is not missing features; it is the
mechanism that would make adding features cheap. Porting it is a rewrite of the reader and
writer, not an addition to them, which is why it is not attempted incrementally.

### `packaging/` — 34 names, ~1,800 lines

`Manifest`, `Relationship`, `Override`, `FileSharing`, `save`, `get_dependents`. ferroxl
builds the manifest, content types and relationships itself in `writer/`, and
`worksheet/relationship.rs` covers the relationship type. The *constructors* differ; the
behaviour largely does not.

**Effect.** Small. This is the closest of the six to a naming-and-shape difference, and
PARITY.md should not have implied otherwise by omitting it.

### `chart/` — 82 unmatched names of 93

**All sixteen chart types are ported** as of 0.1.4: `BarChart`, `LineChart`, `PieChart`,
`ScatterChart`, `AreaChart`, `BubbleChart`, `RadarChart`, `StockChart`, `SurfaceChart`,
`DoughnutChart`, `ProjectedPieChart`, and the 3-D variants `AreaChart3D`, `BarChart3D`,
`LineChart3D`, `PieChart3D`, `SurfaceChart3D`. Type-specific options (`radarStyle`,
`holeSize`, `bubble3D`, `bubbleScale`, `showNegBubbles`, `sizeRepresents`, `firstSliceAng`,
`wireframe`, `ofPieType`) and `View3D` are ported too, so the remaining gap is decoration
rather than chart selection.

Still absent: `DataLabel`, `Trendline`, `UpDownBars`, `Marker` as a type (a series carries a
marker *name*, which is most of what it is for), `Layout`/`ManualLayout`, `ChartSpace` as a
class, `PlotArea`, `DataTable`, `Title`/`Text`/`RichText` as classes, `GraphicalProperties`,
`BandFormat`, and the chart reader.

**Effect.** A user can build any chart type Excel offers, with axes, a legend, solid series
colours and one error-bar type. They cannot add a data label, a trendline, a title's rich
text or a manual layout, and cannot read a chart back.

### `cell/rich_text.py` and `cell/text.py`

`CellRichText`, `TextBlock`, `Text`, `RichText`, `InlineFont`, `PhoneticText`,
`PhoneticProperties`. ferroxl's string-table reader concatenates inline runs and discards
their formatting, which is what openpyxl does when `rich_text=False`.

**Effect.** Formatted text inside a cell cannot be written or read, which also means
openpyxl 3.x's `load_workbook(rich_text=True)` has no counterpart. See the `load_workbook`
table in [`reader`](#reader--openpyxlreader).

### `worksheet/cell_range.py` — ported in 0.1.4

`CellRange` and `MultiCellRange` are now `worksheet::cell_range::CellRange` and
`MultiCellRange`. See [item 2](#2-worksheet-range-with-offsets-rows-and-columns--shipped-in-014).

What is still absent is `openpyxl/worksheet/print_settings.py`, which builds `PrintArea`
and `PrintTitles` on top of the same two classes. That is now the only reason
`print_settings.py` is listed here.

### `worksheet/table.py` — ported in 0.1.5

`Table`, `TableColumn`, `TableStyleInfo`, `TableFormula` and `TableList` are now
`worksheet::table::*`, written to `xl/tables/tableN.xml` and read back through the sheet's
relationships. `Worksheet::add_table` reads the column names out of the header cells, which
is the step that makes `=SUM(Table1[Sales])` resolve: a name that disagrees with its header
cell is silently rewritten by Excel, and a structured reference built on it breaks without
saying so.

Still absent from that module: `XMLColumnProps` (XML-mapped tables), `TablePartList` as a
type, and the query-table and xml table types.

### `worksheet/` remainder

| Module | What a user cannot do |
| --- | --- |
| `views.py` | Model `SheetView`, `Pane`, `Selection` — freeze panes are a pair of fields, not an object |
| `filters.py` | `CustomFilter`, `Top10`, `DynamicFilter`, `DateGroupItem`, `ColorFilter`, `IconFilter`, `Filters`, `SortState`. Only `AutoFilter`, `FilterColumn` and `SortCondition` are ported |
| `errors.py` | `IgnoredError` / `IgnoredErrors` / `ExtensionList` |
| `ole.py` | Embedded OLE objects |
| `smart_tag.py` | Cell and document smart tags |
| `scenario.py` | What-if scenarios and their input cells |
| `print_settings.py` | `PrintArea`, `PrintTitles`, `ColRange`, `RowRange` as parseable values |
| `pagebreak.py` | `Break`/`RowBreak`/`ColBreak` with `min`/`max`/`man`/`pt` — ferroxl has `Vec<u32>` and no span data |
| `properties.py` | `WorksheetProperties`, `Outline`, `PageSetupProperties`; `<sheetPr>` is emitted from hard-coded literals |
| `cell_watch.py`, `controls.py`, `custom.py` | Cell watches, form controls, custom sheet properties |
| `formula.py` | `ArrayFormula` and `DataTableFormula` |
| `hyperlink.py`, `merge.py`, `ole.py` | The XML element types; ferroxl stores the same information as cell fields |

### `workbook/` remainder

| Module | What a user cannot do |
| --- | --- |
| `properties.py::CalcProperties` | Set calculation mode, `fullCalcOnLoad`, `forceFullCalc`, `iterate`. `<calcPr>` is written from fixed literals — 13 fields unreachable |
| `properties.py::WorkbookProperties` | 17 of 19 fields unreachable |
| `views.py::CustomWorkbookView` | Per-user custom workbook views |
| `protection.py::FileSharing` | `readOnlyRecommended`, `reservationPassword` |
| `external_link/` | Read or write any `externalLink` part — 8 classes plus `read_external_link`. This is why `keep_links` has no counterpart |
| `web.py`, `smart_tags.py`, `function_group.py` | Web publishing, smart tags, function groups |
| `defined_name.py` | `workbook/defined_name.py` is a *different* type from the `namedrange.py` one ferroxl ports. `comment`, `description`, `help`, `statusBar`, `hidden`, `function` and nine more are unreachable, as is `RESERVED` / `_xlnm.` handling |

### Shipped in 0.1.6

- **`GradientFill`, `Stop`, `StopList`** — `Fill::linear_gradient`, `Fill::path_gradient`,
  `GradientStop`, `spread_stops`. The reader had been discarding `<gradientFill>` entirely.
- **`DataBar`** and the `Rule::data_bar` / `Rule::icon_set` factories.
- **`cfvo/@gte`, `iconSet/@percent`, `Rule/@timePeriod`**, all previously absent from the
  attribute lists.
- **`Font.charset`, `family`, `scheme`, `outline`, `shadow`, `condense`, `extend`** and
  **`Alignment.relativeIndent`, `justifyLastLine`, `readingOrder`.**
- **`quotePrefix` and `pivotButton`** on `Style`.
- **`CalcProperties`** — all thirteen `<calcPr>` fields.

### `styles/` remainder

| Module | What a user cannot do |
| --- | --- |
| `builtins.py::pandas_highlight` | The one gap left in this module: openpyxl's non-built-in `Pandas` style, used when round-tripping a DataFrame. All 49 built-ins are covered by `BUILTIN_DETAILS` |
| `cell_style.py` | `StyleArray` as a class, and `<cellStyleXfs>` is not read at all. `quotePrefix`, `pivotButton` and `applyNumberFormat` **are** now read and written (0.1.6); `xfId` is not |
| `proxy.py::StyleProxy` | The read-only style proxy that makes `cell.font` non-assignable |
| `table.py` | Custom table styles |
| `differential.py` | `dxf` cannot change number format, alignment or protection — only font, fill and border |
| `colors.py` | `RgbColor`; `<colors>`/`<indexedColors>` is read but never written |
| `numbers.py` | Built-in format ids 48 (`##0.0E+0`) and 49 (`@`) are missing, so `is_builtin("@")` is false |

### `utils/` remainder

`FORMULAE` (≈370 built-in function names), `escape`/`unescape`, `IndexedList`,
`BoundDictionary`, `dataframe_to_rows`, and the open-ended range forms `"A:A"`, `"1:5"`,
`"A1:"` — `get_range_boundaries` rejects all three even though openpyxl accepts them.
`cols_from_range`, `coordinate_to_tuple`, `range_to_tuple` and `quote_sheetname` have no
reusable form. `cast_numeric`/`cast_percentage`/`cast_time` exist as `Cell` methods rather
than free functions.

## Not ported

Python-only infrastructure with no Rust counterpart, or a counterpart that is worse.

| Upstream | Why not |
| --- | --- |
| `compat/` — `functools`, `itertools`, `numbers`, `odict`, `singleton`, `strings` | Shims for Python 2, the `OrderedDict` recipe and `Decimal`. Rust has these in the standard library or as `f64`. |
| `xml/namespace.py::register_namespace` | `lxml`'s global namespace registry. ferroxl has fixed constants and resolves prefixes locally. |
| `xml/functions.py::iterparse`, `safe_iterparse`, `safe_iterator`, `get_document_content` | Streaming `lxml` iterators. `fromstring` reads a whole part. Tied to the streaming-writer gap. |
| `xml/functions.py::pretty_indent` | `to_pretty_string` covers it. |
| `reader/style.py::SharedStylesParser`, `reader/worksheet.py::fast_parse`, `reader/comments.py::get_comments_file` | Internal parser scaffolding. ferroxl exposes `read_style_table`, `read_worksheet`, `read_comments` and `comments_file_path`; the private `WorksheetParser` and `WorksheetParseContext` do the rest. |
| `namedrange.py::NAMED_RANGE_RE`, `SPLIT_NAMED_RANGE_RE` | Module-private regexes; ferroxl parses the grammar by hand, which is what the tests exercise. |
| `cell/cell.py::COORD_RE`, `ABSOLUTE_RE`, `ILLEGAL_CHARACTERS_RE`, `TIME_REGEX`, `TIME_TYPES`, `KNOWN_TYPES` | Compiled once and kept private. The behaviour is in `coordinate_from_string`, `absolute_coordinate`, `get_range_boundaries`, `check_string` and the `cast_*` methods. |
| `datavalidation.py::default_attr_map`, `styles/__init__.py::DEFAULTS`, `writer/workbook.py::static_content_types_config`, `writer/theme.py::theme_xml` | Module-private data. The content types and the theme are built by the writer, from `STATIC_CONTENT_TYPES` and `THEME_XML`. |
| `worksheet/worksheet.py::SheetView` | An empty `pass` class upstream. The fields are on `Worksheet`. |
| `worksheet/iter_worksheet.py::IterableWorksheet`, `ROW_TAG`, `CELL_TAG`, `VALUE_TAG`, `FORMULA_TAG`, `DIMENSION_TAG` | The streaming worksheet class and its module-private tag constants. ferroxl's `ReadOnlyCell` and `ReadOnlyTables` give the same access, and the reader matches tag names at run time. |
| `charts/series.py::Serie` | A backwards-compatibility alias for `Series`. |
| `worksheet/worksheet.py::flatten` | A Python 2 leftover; it takes one argument and returns it. |
| `benchmarks/`, `sample/`, `tests/` | Python's own. ferroxl has its own tests, 481 of them. |

## How parity is verified

Parity is checked three ways, because "it compiles" is not evidence.

**Values are pinned from Python, not from this implementation.** Where a number is
observable from openpyxl, the test asserts the number Python produces:

- the password hasher — `""` → `CE4B`, `"password"` → `83AF`, `"a" * 50` → `20FFFFF4E38`,
  `"a" * 100` → `83FFFFFFFFFFFFFFFFF4E6E`
- the date serials — 1900-02-28 → 59, 1900-03-01 → 61, 2009-12-20 → 40167,
  2010-01-18 → 40196, and the two calendars' Julian days `2415018.5` and `2416480.5`
- the Julian day round trip, against `jdcal`'s own output
- the `is_date_format` rule, including `[hh]:mm:ss` being *not* a date format
- the chart axis padding and rounding
- `COLOR_INDEX` having 56 entries rather than the 64 the documentation mentions

**Files are opened with real openpyxl.** Workbooks ferroxl writes are loaded with openpyxl
3.x, and the values, fonts, fills, borders, number formats, merges, conditional formatting,
data validations, comments, defined names, column widths, freeze panes, header and footer,
autofilters and chart series are read back and compared. `describe_sheet`'s `freeze_panes`
and `set_header_footer` bugs were found this way, not by unit tests.

**The public surface is audited.** `tools/parity.py` compares openpyxl's public names with
ferroxl's and prints what is missing. This document is its output, and the
[Pending](#pending) list is what it still reports.

The suite is 532 tests - 412 in the library, 116 in the MCP server, 4 doctests - and
`cargo build`, `cargo clippy -- -D warnings`, `cargo fmt --check` and
`RUSTDOCFLAGS=-D warnings cargo doc` are all clean.

### What this audit can and cannot see

`tools/parity.py` matches names, not behaviour. It lowercases and drops underscores, so
`iter_rows` and `iterRows` both count as matched against a Python `iter_rows`. It therefore
cannot see a function that exists but behaves differently, a struct that is missing half its
fields, or an XML element that is read but silently dropped. Every field-level loss listed
above — `Font.family`, `cfvo/@gte`, `GradientFill`, `<colors>`, `<cellStyleXfs>` — is
invisible to it and was found by reading both trees, not by running the tool.

The converse also holds: a name can match while the feature behind it is absent. The tool
counts those as matched, so the module figures in [Summary](#summary) are an upper bound on
what works rather than a measurement of it.
