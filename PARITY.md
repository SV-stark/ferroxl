# Feature parity with openpyxl 1.9.0

lexcel is a port of [openpyxl](https://github.com/theorchard/openpyxl) 1.9.0. This document
records, module by module, what has been implemented, what has not, and where the Rust
version deliberately behaves differently.

It is written from the source trees, not from memory. `tools/parity.py` walks every module
of the Python original, collects its public names, and reports which have a lexcel
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
- [Not ported](#not-ported)
- [How parity is verified](#how-parity-is-verified)

## How to read this

Three labels are used, and they mean different things:

- **Implemented** — the behaviour exists and is covered by a test.
- **Pending** — openpyxl has it, lexcel does not. Each entry says what the effect is and
  what a caller does instead.
- **Different by design** — lexcel does it, but not the way Python does. Each entry says
  why, because a difference you did not choose is a bug.

## Summary

| | |
| --- | --- |
| Python modules audited | 70 |
| Modules whose every public name is matched | 47 |
| openpyxl public names | 272 |
| Names with no lexcel counterpart | 77 |

Of those 77 unmatched names:

| Bucket | Names | What it is |
| --- | --- | --- |
| Genuinely pending | 13 | [The streaming writer](#1-the-streaming-writer-writerdump_worksheetpy), [zip repair](#5-zip-central-directory-repair), and the reader types behind them |
| Renamed or reshaped | 37 | A name changed, or a Python container became a Rust type — see [Different by design](#different-by-design) |
| Python-only infrastructure | 27 | `lxml` iterators, namespace registration, the `compat` shims, and regexes and tag constants that openpyxl keeps to itself — see [Not ported](#not-ported) |

Two further gaps are *method-level* and so do not appear in a top-level-name audit:
[`Worksheet.range()` with offsets, `rows` and `columns`](#2-worksheet-range-with-offsets-rows-and-columns),
and [charts not being read back](#6-charts-and-images-are-written-but-not-read-back).

## Implemented

Everything in this section is covered by tests. Names are lexcel's; the upstream module is
named so a change can be traced back to the Python it mirrors.

### `cell` — `openpyxl/cell`

| Upstream | lexcel |
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

| Upstream | lexcel |
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
`guess_types`, `data_only` and `keep_vba` — matching openpyxl 1.9's four parameters, plus
the `new` / `guessing_types` / `values_only` / `keeping_vba` builders; `WorkbookSource`,
`package_bytes`, `read_string_table`, `read_style_table`, `read_worksheet`,
`read_comments`, `comments_file_path`, `read_sheets`, `read_rels`, `read_content_types`,
`read_properties_core`, `read_excel_base_date`, `read_workbook_settings`,
`read_named_ranges`, `title_resolver`, `detect_worksheets`.

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

### The `lexcel-mcp` server

Not part of openpyxl. 33 tools over JSON-RPC 2.0 on stdio, grouped as reading (9),
structure (9), values (4), layout and appearance (8) and drawing (3). See the
[README](README.md#the-mcp-server).

## Pending

Six areas where openpyxl has something lexcel does not, ordered by how likely they are to
matter.

### 1. The streaming writer (writer/dump_worksheet.py)

openpyxl has two writers. The default one builds a tree in memory; `dump_worksheet.py`
provides the second, `lxml`-based path used by `ExcelWriter` for large sheets, which
streams rows out as they are produced and never holds a whole sheet in memory. It carries
`DumpWorksheet`, `ExcelDumpWriter`, `StyleDumpWriter`, `save_dump`,
`create_temporary_file`, and the `STYLES` / `DATETIME_STYLE` / `BOUNDING_BOX_PLACEHOLDER` /
`DESCRIPTORS_CACHE_SIZE` constants that go with it.

**Effect.** lexcel writes a workbook correctly, but peak memory scales with the size of the
sheet being written. For the sheets an agent typically produces this does not matter; for a
hundred-thousand-row export it would.

**What to do instead.** Nothing today — `Workbook::save` and `save_workbook_to` are the
only paths.

### 2. Worksheet range with offsets, rows and columns

`Worksheet.range(range_string, row=0, column=0)` returns a two-dimensional grid of
**`Cell` objects**, which the caller can then mutate. `rows` and `columns` are convenience
properties built on it.

lexcel has `range_values(range) -> Vec<Vec<CellValue>>` and
`range_coordinates(range) -> Vec<String>`. There is no equivalent that hands back a
rectangle of `Cell`s, no row/column offset arguments, and no `rows` or `columns`.

**Effect.** `for row in ws.rows: row[0].value = x` has no direct spelling. The same edit is
`ws.set(coord, value)` per coordinate, or `ws.set_cell_value(coord, value)`. Read-only use
is covered by `range_values`.

**What to do instead.** Use `range_values` to read and `set` / `set_cell_value` to write.

### 3. The use_iterators flag has no loader

openpyxl's `load_workbook(..., use_iterators=True)` hands back an `IterableWorksheet`, which
parses a sheet's XML lazily and yields `ReadOnlyCell`s one at a time without building a
`Worksheet`. That is how openpyxl reads a hundred-megabyte sheet.

lexcel has the value type — `cell::ReadOnlyCell`, with `coordinate`, `internal_value`,
`number_format`, `is_date`, `value` and `datetime`, plus the `ReadOnlyTables` it resolves
against — and both are public and constructible. Nothing in `load_workbook` produces them:
there is no `use_iterators` option and no iterator over a worksheet's cells.

**Effect.** Reading a workbook materialises every sheet. For the sheets an agent typically
opens this does not matter; for a very large one it would.

**What to do instead.** Nothing today. `ReadOnlyCell` is exercised by tests but is not yet
reachable from a loaded workbook.

### 4. The stored dimension element is not read

openpyxl's read-only path reads `<dimension ref="A1:D10">` and trusts it, falling back to
the cells when the element is absent. lexcel computes the dimension from the cells it read
(`calculate_dimension`) and stores nothing.

**Effect.** For a workbook whose stored `<dimension>` disagrees with its cells, the two
report different extents. Excel and openpyxl agree in practice because Excel writes the
dimension correctly, so the divergence only shows on a hand-edited or corrupt file.
`highest_row` and `highest_column` come from the dimension tables rather than the stored
element, so they move with the cells.

### 5. Zip central directory repair

When a workbook fails to open, openpyxl searches for the end-of-central-directory
signature and truncates whatever follows it, then retries. This recovers files truncated in
transit or with junk appended.

**Effect.** A slightly damaged file that openpyxl would open, lexcel rejects with
`Error::BadZipFile`.

### 6. Charts and images are written but not read back

openpyxl 1.9 writes chart and drawing parts but its reader does not parse them, so a
reloaded workbook reports no charts and no images. lexcel matches this rather than being
half-compatible in a different direction.

**Effect.** `worksheet.charts` and `worksheet.images` are empty after a load, even though
the parts are in the file. Anything that depends on reading a chart back — inspecting an
existing chart, preserving one across an edit — does not work, in lexcel or in openpyxl 1.9.

This one is a *parity* gap rather than a *capability* gap: lexcel behaves exactly as the
reference does.

## Different by design

Places where lexcel does the job but not Python's way. Each is also documented at the call
site, so a reader of the code does not have to come here first.

### Equality and hashing (styles/hashable.py)

openpyxl's `HashableObject` compares a style by the fields that matter and falls back to
identity for the rest, so a number format hashes by its code and two visually identical
styles collapse to one style-table entry.

Rust's `HashMap` requires `Hash` and `Eq`, so lexcel derives them but implements them using
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
such value. lexcel uses `ProtectionFlag::{Inherit, Locked, Unlocked}` for those attributes
and a plain `bool` where the specification really does mean two states.

### CategoryAxis and ValueAxis are constructors

In Python they are classes that differ only in their class-level attribute defaults. In
lexcel they are unit structs whose `new()` returns an `Axis` configured the way the
corresponding Python class would configure itself — which is why `new` does not return
`Self`, and why that carries an explicit `#[allow(clippy::new_ret_no_self)]`.

### Image dimensions come from a PNG header

openpyxl opens an image with PIL to read its size, which makes PIL a hard dependency for
anyone embedding a picture. lexcel parses the PNG header directly, so PNG is the only
supported format and PIL is not a dependency at all. The bytes are stored verbatim, so the
image itself is untouched.

### Names that were renamed

Renaming is not a behaviour change, but it will trip up someone porting code by search and
replace:

| openpyxl | lexcel | Why |
| --- | --- | --- |
| `pixels_to_EMU`, `EMU_to_pixels`, `cm_to_EMU`, … | `pixels_to_emu`, `emu_to_pixels`, `cm_to_emu`, … | Rust naming convention |
| `PAPERSIZE_LETTER` … `PAPERSIZE_A5`, eleven constants | `PAPERSIZES`, a table of `(name, code)` pairs | Eleven constants are a table |
| `worksheet.add_chart(chart, anchor)` | `worksheet.charts.push(chart)` — the anchor is a field on the chart | The `Vec` is public |
| `worksheet.add_image(image, anchor)` | `worksheet.images.push(image)` | Same |
| `worksheet.add_rel(...)` | `worksheet.relationships.push(...)` | Same |
| `ws.max_row`, `ws.max_column` | `ws.highest_row()`, `ws.highest_column()` | Matches openpyxl 1.9's `get_highest_row` |
| `cell.value` | `cell.internal_value()`, or `worksheet.cell_value(coord)` | A method, so the type cast runs |
| `StyleWriter`, `ChartWriter`, `CommentWriter`, `DrawingWriter`, `ShapeWriter` | `write_style_table`, `write_chart`, `write_comments`, `write_drawing`, `write_shapes` | See below |
| `load_workbook(..., read_only=True)` | `LoadOptions::guessing_types()` / `values_only()` / `keeping_vba()` | Builders instead of keyword arguments |
| `ExcelWriter(workbook).save(path)` | `Workbook::save(path)` | See below |

### Writer classes became functions

openpyxl's writers are classes with `write()` and `close()` methods, which lets them stream
through an open file handle. lexcel's writers build a `String` and hand it to the archive,
so there is no handle to own and no state worth capturing — `write_worksheet(...)`,
`write_workbook(...)`, `write_chart(...)` and the rest are free functions.

`ExcelWriter` does survive, because it holds the workbook, the string table and the style
tables; only its `write()`/`close()` pair collapsed into `to_bytes()` and
`save_workbook_to(...)`. This is also why the streaming writer in
[Pending](#1-the-streaming-writer-writerdump_worksheetpy) is a real gap rather than a
refactor: the function-shaped writers have nowhere to put a stream.

### Style tables are sorted before they are written

openpyxl keeps the style table in insertion order. lexcel sorts by `Style::sort_key` so two
runs that build the same set of styles produce byte-identical output. Excel does not care
about the order; a diff does.

## Not ported

Python-only infrastructure with no Rust counterpart, or a counterpart that is worse.

| Upstream | Why not |
| --- | --- |
| `compat/` — `functools`, `itertools`, `numbers`, `odict`, `singleton`, `strings` | Shims for Python 2, the `OrderedDict` recipe and `Decimal`. Rust has these in the standard library or as `f64`. |
| `xml/namespace.py::register_namespace` | `lxml`'s global namespace registry. lexcel has fixed constants and resolves prefixes locally. |
| `xml/functions.py::iterparse`, `safe_iterparse`, `safe_iterator`, `get_document_content` | Streaming `lxml` iterators. `fromstring` reads a whole part. Tied to the streaming-writer gap. |
| `xml/functions.py::pretty_indent` | `to_pretty_string` covers it. |
| `reader/style.py::SharedStylesParser`, `reader/worksheet.py::fast_parse`, `reader/comments.py::get_comments_file` | Internal parser scaffolding. lexcel exposes `read_style_table`, `read_worksheet`, `read_comments` and `comments_file_path`; the private `WorksheetParser` and `WorksheetParseContext` do the rest. |
| `namedrange.py::NAMED_RANGE_RE`, `SPLIT_NAMED_RANGE_RE` | Module-private regexes; lexcel parses the grammar by hand, which is what the tests exercise. |
| `cell/cell.py::COORD_RE`, `ABSOLUTE_RE`, `ILLEGAL_CHARACTERS_RE`, `TIME_REGEX`, `TIME_TYPES`, `KNOWN_TYPES` | Compiled once and kept private. The behaviour is in `coordinate_from_string`, `absolute_coordinate`, `get_range_boundaries`, `check_string` and the `cast_*` methods. |
| `datavalidation.py::default_attr_map`, `styles/__init__.py::DEFAULTS`, `writer/workbook.py::static_content_types_config`, `writer/theme.py::theme_xml` | Module-private data. The content types and the theme are built by the writer, from `STATIC_CONTENT_TYPES` and `THEME_XML`. |
| `worksheet/worksheet.py::SheetView` | An empty `pass` class upstream. The fields are on `Worksheet`. |
| `worksheet/iter_worksheet.py::IterableWorksheet`, `ROW_TAG`, `CELL_TAG`, `VALUE_TAG`, `FORMULA_TAG`, `DIMENSION_TAG` | The streaming worksheet class and its module-private tag constants. lexcel's `ReadOnlyCell` and `ReadOnlyTables` give the same access, and the reader matches tag names at run time. |
| `charts/series.py::Serie` | A backwards-compatibility alias for `Series`. |
| `worksheet/worksheet.py::flatten` | A Python 2 leftover; it takes one argument and returns it. |
| `benchmarks/`, `sample/`, `tests/` | Python's own. lexcel has its own tests, 481 of them. |

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

**Files are opened with real openpyxl.** Workbooks lexcel writes are loaded with openpyxl
3.x, and the values, fonts, fills, borders, number formats, merges, conditional formatting,
data validations, comments, defined names, column widths, freeze panes, header and footer,
autofilters and chart series are read back and compared. `describe_sheet`'s `freeze_panes`
and `set_header_footer` bugs were found this way, not by unit tests.

**The public surface is audited.** `tools/parity.py` compares openpyxl's public names with
lexcel's and prints what is missing. This document is its output, and the
[Pending](#pending) list is what it still reports.

The suite is 481 tests — 369 in the library, 110 in the MCP server, two doctests — and
`cargo build`, `cargo clippy` and `cargo fmt --check` are all clean.
