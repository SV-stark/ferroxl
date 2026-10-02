# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.6] - 2026-10-02

### Added

- **`reader::archive`** rebuilds a zip whose end-of-central-directory record was cut off, which
  is what an interrupted download leaves behind. The central directory is written before that
  record, so the archive is usually recoverable: the directory is walked for its entry count,
  size and offset, and the missing 22 bytes are synthesised. Every cut point inside the record
  is covered, down to the two bytes `PK` left by a 20-byte cut.
- **Relationship targets are resolved rather than assumed.** `resolve_part` walks `.` and `..`
  and honours a leading slash, so a generator that writes `Target="/xl/worksheets/sheet1.xml"`
  or `Target="../xl/worksheets/sheet1.xml"` is read instead of losing the sheet.

- **Named styles.** `NamedStyle`, `NamedStyleList` and `Workbook::named_styles` cover
  `openpyxl.styles.named_styles`. `<cellStyles>` and `<cellStyleXfs>` are read and written,
  so a workbook's own styles survive a round trip instead of every cell arriving as `Normal`.
- **All 49 of Excel's built-in styles**, in `BUILTIN_DETAILS`, generated from openpyxl 3.1.5's
  `styles/builtins.py` rather than transcribed. Each carries its fill, font colour, size,
  weight and number format as well as its `builtinId`: Excel renders a built-in from the id
  alone, so an id-only style looks right in Excel and blank in everything else, openpyxl
  included.
- **`Workbook::apply_named_style`** applies *and registers*, so the name reaches Excel's style
  gallery and reads back as `cell.style`. `Worksheet::apply_named_style` applies only, and a
  worksheet has no way to reach the workbook's style list.
- **`Workbook::add_named_style`**, and `Workbook::named_style_names`.

### Fixed

- **A sheet could vanish with no error at all.** `detect_worksheets` built the part name by
  prefixing `xl/` onto the relationship target, so a generator that writes the target as
  absolute (`/xl/worksheets/sheet1.xml`) or with a hop back out of `xl/`
  (`../xl/worksheets/sheet1.xml`) produced `xl//xl/worksheets/sheet1.xml`. That matched no
  content type, the sheet was dropped, and the workbook opened empty -- no exception, no
  warning. `resolve_part` now resolves the target properly, and the path it returns is the
  final name, so the two callers that used to prefix `xl/` again no longer do.
- **`cellXfs` wrote `xfId="0"` unconditionally**, so every cell claimed to derive from `Normal`
  and openpyxl reported `cell.style` as `'Normal'` even for a cell carrying `Good`'s colours.
  A cell whose formatting matches a named style now points at it; one the user has altered
  since no longer does, because claiming the style would make Excel restyle a cell on an edit
  the user never asked for.

The silent-loss batch: everything here was read as absent or written from a literal, so a
workbook using it loaded wrong and saved wrong without saying so.

### Fixed

- **Gradient fills were discarded on read.** The style reader looked only for `<patternFill>`,
  so a workbook using one loaded with every cell falling back to a plain fill. Nothing failed
  and nothing warned, and it looked correct — a missing gradient is just a background colour.
  `Fill` now carries its stops, and `GradientStop` implements `Eq`/`Hash` on the position's
  bits, because `Fill` needs both to de-duplicate the stylesheet.
- **Data bars could be neither written nor read.** Three separate places — the reader, the
  writer, and `is_data_bar`'s documentation — said "openpyxl skips these". The belief came from
  a half-truth: Excel writes data bars' *extra* properties (gradient fill, border,
  negative-bar colour, axis) through an `x14` extension, but the bar itself is ordinary
  `cfRule` content. So the rule was dropped from every workbook containing one.
- **`cfvo/@gte` was read as absent**, which is not the same as true. The schema defaults it to
  true, so a rule written `gte="0"` came back including the boundary value it was meant to
  exclude. Now read, and written only when false.
- **`iconSet/@percent` was missing** from the attribute list, so an icon set built on
  percentages came back with its thresholds reinterpreted against the row count — still
  lighting icons, just the wrong ones.
- **`Rule/@timePeriod` was missing entirely.** It is the attribute that makes a rule relative
  to today rather than to the data, so "yesterday" and "last week" rules could not be
  expressed at all.
- **`Font.family` was written as a hard-coded 2.** Excel's default font happens to be 2, so it
  looked deliberate and was wrong for every file using another family index — and the reader
  could not see the difference, so it could never be recovered.
- **`Font.scheme` was dropped.** Of the font attributes this changes meaning rather than only
  appearance: a themed font names a theme slot, so dropping it pins the font to its resolved
  name and defeats the point of theming.
- **`<strike>` was never written**, and every underline style except `single` collapsed to a
  bare `<u/>` — which is what `doubleAccounting` became.
- **`quotePrefix` and `pivotButton` were not read.** A quote prefix is a style property, not a
  cell value: it makes Excel display a leading apostrophe rather than treat it as an escape,
  so dropping it turns a shown `'007` into a number on the next save.
- **`applyNumberFormat` was not written.** Excel reads the `numFmtId` regardless, which is why
  nothing looked wrong.
- **`CellRange` collapses a one-cell range to `A1`**, matching openpyxl's `coord`. Rendering it
  as `A1:A1` would have lengthened every single-cell reference on a round trip.

### Added

- **`Alignment.relative_indent`, `justify_last_line`, `reading_order`**. `relativeIndent` was
  read as absent, which turned a hanging indent into a plain one — still an indent, so
  nothing looked wrong.
- **`Font.charset`, `family`, `scheme`, `outline`, `shadow`, `condense`, `extend`.**
- **`Fill::linear_gradient`, `path_gradient`, `GradientStop`, `spread_stops`.**
- **`DataBar`** and `Rule::data_bar` / `Rule::icon_set`, openpyxl's `DataBarRule` and
  `IconSetRule`.
- **`CalcProperties`** — all thirteen of openpyxl's `<calcPr>` fields, with
  `Workbook::set_calculation_properties`. `fullCalcOnLoad` is load-bearing rather than
  cosmetic: ferroxl writes formulas with no cached result, so it is what makes Excel calculate
  them on open instead of showing blanks, and it was not something a caller could change.
- **Five MCP tools**: `add_data_bar`, `add_icon_set`, `add_table`, `describe_table`,
  `set_gradient_fill`. 36 to 41.

### Verified

Gradient fills, `Font.family`, `charset`, `scheme`, `outline`, `doubleAccounting` underline and
`quotePrefix` are each confirmed by loading a ferroxl-written file with **openpyxl 3.1.5** and
reading the values back. A Rust round trip proves the crate agrees with itself; only the other
library proves the file agrees with the format.

## [0.1.5] — 2026-10-02

Excel tables (ListObjects).

### Added

- **`worksheet::table`** — `Table`, `TableColumn`, `TableStyleInfo`, `TableFormula` and
  `TableList`, the port of `openpyxl/worksheet/table.py`. Written to `xl/tables/tableN.xml`
  and read back through the sheet's relationships.
- **`Worksheet::add_table`** and **`Worksheet::table`**. The columns are read from the header
  cells rather than invented, which is the step that makes structured references resolve: a
  column name that disagrees with its header cell is rewritten by Excel on open, and
  `=SUM(Table1[Sales])` breaks without saying so.
- The four pieces a table needs in the package — the part, the relationship, the content-type
  Override and `<tableParts>` — with their ids derived from the sheet so the sheet tail and
  the package writer cannot disagree about which part is which.

### Notes

A table is verified against openpyxl 3.1.5 itself rather than only against ferroxl: the
round-trip test saves a real package, and a cross-check loads it with Python and confirms the
ref, display name, header count, column names and style all come back. None of the four
package pieces is visible from any one of them, so testing each in isolation would pass while
Excel still refused the file.

A table name containing a space is refused at construction. Excel rejects it when the file is
opened, and the error it gives names the file rather than the table.

Still absent: `XMLColumnProps` for XML-mapped tables, and the query-table and xml table types.

## [0.1.4] — 2026-10-02

All sixteen chart types, and cell ranges as values.

### Added

- **`worksheet::cell_range::CellRange`** — openpyxl's `worksheet.cell_range.CellRange` as a value rather
  than a pair of corners: `intersection`, `union`, `issubset`, `issuperset`, `isdisjoint`,
  `contains`, `shift`, `expand`, `shrink`, `size`, `top`/`bottom`/`left`/`right`, `rows`,
  `cols` and `cells`. Bounds are inclusive at both ends, the opposite of `RangeBounds` and
  the easiest off-by-one in the area to get wrong.
- **`MultiCellRange`** — the `sqref` collection, for conditional formatting that applies to
  several disjoint rectangles. Overlapping ranges are kept separately rather than merged, so
  asking "which rules apply to B2" reports both rather than quietly deduplicating.
- **`Worksheet::range(range, row_offset, column_offset)`** and **`range_cells`** — `ws["A1:C3"]`
  and friends. A bare coordinate with offsets expands to a rectangle; a range with a colon
  shifts. The argument's shape decides which, because deciding from the offsets would make
  `ws["A1:B2", 1, 1]` two plausible things at once.
- **Twelve chart types**: `AreaChart`, `BubbleChart`, `RadarChart`, `StockChart`,
  `SurfaceChart`, `DoughnutChart`, `ProjectedPieChart`, and every 3-D variant
  (`AreaChart3D`, `BarChart3D`, `LineChart3D`, `PieChart3D`, `SurfaceChart3D`). ferroxl had
  four of thirteen, all 2-D.
- **`ChartOptions`** and **`View3D`** — the per-type options (`radarStyle`, `holeSize`,
  `bubble3D`, `bubbleScale`, `showNegBubbles`, `sizeRepresents`, `firstSliceAng`, `wireframe`,
  `ofPieType`) and the 3-D view.
- **`Series::with_bubble_size`** and `SeriesAttr::BubbleSizes` — a bubble carries a third
  column of numbers.

### Fixed

- **`is_graph_chart`** was a three-way allowlist covering only the chart types that existed
  when it was written, so the twelve new ones were classified as having no axes. It is now
  the rule it should have been: everything except the pie family has an `axId`.
- **A one-cell `CellRange` renders as `A1`**, matching openpyxl's `coord`. Rendering it as
  `A1:A1` would have lengthened every single-cell reference in the workbook on a round trip.

### Changed

- **PARITY.md**: chart types and `worksheet/cell_range.py` move out of the absent list. The
  remaining chart gap is decoration — data labels, trendlines, up/down bars, manual layouts,
  rich-text titles — and the chart reader, rather than chart selection.

## [0.1.3] — 2026-10-02

The first release measured against openpyxl 3.1.5 rather than the 1.9-era surface the
project originally claimed, and the accounting corrected accordingly.

### Added

- **`formula::Translator`** — openpyxl's `formula.translate.Translator`. Translates a formula from
  the cell it was written for to the cell it is going to, which is the operation behind
  copy-paste, fill-right and fill-down. It takes an origin and a destination rather than
  raw deltas, and it translates whole-row (`3:4`) and whole-column (`A:BC`) references,
  which `shift_references` steps over because it looks for `$A$1`-shaped tokens — leaving
  `=SUM(3:4)` unchanged across a copy that should have made it `=SUM(13:14)`.
- **`TranslatorError`**, raised when a translation would push a relative reference off the
  grid. Excel reports that as `#REF!`; clamping would keep the formula loadable while
  quietly meaning something else.
- **`Worksheet::iter_rows`** and **`iter_cols`** — row-major and column-major walks over the
  used range, with `iter_rows_within` / `iter_cols_within` for explicit inclusive bounds.
  `ws.iter_rows()` is what most openpyxl code calls first, and ferroxl had only an
  unordered `cells()`.

### Fixed

- An anchored reference in a whole-row or whole-column range was being shifted anyway, so
  `=SUM($A:$B)` came out as `=SUM($C:$D)`. The dollar sign is the entire mechanism, so the
  check has to happen before the arithmetic rather than after.

### Changed

- **`PARITY.md` is now measured against openpyxl 3.1.5.** Every claim that the project was a
  port of 1.9.0 was wrong: the reference tree is 3.1.5, and the previous document's
  "272 public names, 77 unmatched" was an artefact of an older checkout. Against the real
  tree it is **994 top-level names, 741 unmatched, and 31 of 183 modules fully matched**.
- Six whole upstream packages that `PARITY.md` never mentioned are now documented: `pivot`
  (58 classes), `chartsheet` (11), `descriptors` (49 names), `packaging` (34),
  `cell/rich_text.py` and `worksheet/cell_range.py`.
- `load_workbook` is now compared against 3.1.5's six parameters. Three are implemented;
  three are not (`read_only`, `keep_links`, `rich_text`). ferroxl's `guess_types` has no
  upstream counterpart — openpyxl removed it in 3.0 — so it is kept as a useful extra rather
  than claimed as parity.
- `tools/parity.py` documents 3.1.5, and `PARITY.md` states what the audit cannot see: it
  matches names, so a function that behaves wrongly and an XML element that is silently
  dropped both read as matched.

### Not ported

Named styles, gradient fills, rich text, pivot tables, chart-only sheets, Excel tables,
the `descriptors` layer, and nine of thirteen chart types are documented as absent rather
than implied present. Two of them — pivot tables and chart sheets — are also *lossy on a
round trip*: saving a loaded workbook discards those parts.

## [0.1.2] — 2026-10-02

Cell dependency tracing, and a streaming writer that does not need the whole sheet in
memory.

### Added

#### The `ferroxl` library

- **`worksheet::dependency`** - `Worksheet::trace_precedents` returns every cell that feeds
  a cell, transitively, in the order a recalculation would visit them.
  `Worksheet::trace_dependents` returns the formulas that would go stale if a cell
  changed. `Worksheet::circular_references` reports each cycle once as a closed path.
  `Worksheet::dependency_graph` exposes the whole `cell -> cells it reads` map, and
  `parse_references` pulls the references out of a formula on its own.
- **`DumpWorksheet`** and `save_dump`, the streaming writer. Rows are serialised as they
  are produced rather than collected into a `String` first, so peak memory is set by the
  largest row instead of the largest sheet.

### Changed

- `write_worksheet` is split into a head, the rows and a tail, which both writers share. A
  streaming writer that quietly emitted different XML would be worse than no streaming
  writer, so the two are tested against each other at the worksheet level and across every
  entry in the package.

### Fixed

- The package comparison in the streaming writer's tests compares parts rather than raw
  zip bytes. `start_file` and `writestr!` do not choose the same compression, so a byte
  comparison was asserting an accident of the writer rather than a property of the format.

### Notes on the tracer

Two answers are refused rather than guessed. A reference through a defined name is
reported by `References::named` instead of being resolved, because a name may cover any
range in the workbook and a guess would produce a graph that looks authoritative and is
not. A cross-sheet reference keeps its sheet title in `References::cross_sheet` rather
than being expanded, because the range behind one may be far larger than the formula that
mentions it.

The graph is built from formula text, so a reference that some other tool wrote as an
unresolved name is invisible until that name is resolved. `References::named` is where
that shows up, which is the reason it exists.

## Added to the MCP server

- `trace_precedents` - what a cell is actually built from.
- `trace_dependents` - the blast radius of changing a cell.
- `check_circular_references` - every cycle in a sheet as a closed path. Excel refuses to
  calculate a workbook with one, so this is worth running before trusting a file you did
  not create.

Summaries name at most a dozen cells and then count the rest. A sheet with a few thousand
dependents would fill the context window with a list that cannot change the model's next
decision, and the count is the part that can.

## See also

`ROADMAP.md` records eight further proposals that were considered for this release, what
each would cost, and why only one of them shipped. Three of them - formula evaluation,
non-destructive editing and rayon parallelism - each change the build-time dependency set
or the round-trip model, and are scheduled against the major versions.

## [0.1.0] — 2026-10-02

The first release. ferroxl is a Rust port of openpyxl 3.1.5: the same modules, the same
classes, the same behaviour, including the parts that are surprising.

### Added

#### The `ferroxl` library

- **`cell`** — `Cell` and `CellValue`, the data-type enum, coordinate parsing and
  formatting, read-only cells, and formula handling including shared formulas.
- **`charts`** — `Chart`, `BarChart`, `LineChart`, `ScatterChart` and `PieChart`, with
  series, number and category references, axis scaling and titles, and data labels.
- **`comments`** — `Comment` and the VML shapes Excel uses to carry them.
- **`datavalidation`** — `DataValidation` with its type and operator enums, and the
  collapsing of cell addresses into ranges.
- **`date_time`** — the 1900 and 1904 date systems, serial conversion, ISO-8601 parsing,
  and the Julian day arithmetic openpyxl gets from `jdcal`.
- **`drawing`** — `Image`, `Drawing`, `Shape`, the anchor types and EMU conversion.
- **`exceptions`** — `Error` and its variants, mirroring openpyxl's exception hierarchy.
- **`formatting`** — conditional formatting, colour scales, icon sets, `CellIsRule` and
  `FormulaRule`.
- **`namedrange`** — defined names, their destinations, and the range-string grammar.
- **`reader`** — the whole load path: shared strings, the style table, the workbook, its
  worksheets and its comments.
- **`styles`** — fonts, fills, borders, alignment, number formats, protection and the
  indexed colour palette.
- **`units`** — the unit constants Excel works in, from `points_to_pixels` to `emu_to_cm`.
- **`workbook`** — `Workbook`, document properties, workbook security and sheet
  management.
- **`worksheet`** — `Worksheet`, cell and row and column dimensions, sheet views, freeze
  panes, protection, header and footer, iteration and table definitions.
- **`writer`** — the whole save path, with a streaming XML writer that emits the same
  bytes openpyxl's `XMLGenerator` does.
- **`xml`** — an ElementTree-shaped element tree with namespace resolution, a streaming
  writer, and the OOXML namespace constants.

#### The `ferroxl-mcp` server

A JSON-RPC 2.0 server over stdio, with 33 tools:

- *Reading* — `list_sheets`, `describe_sheet`, `read_cells`, `read_formulas`,
  `search_values`, `summarize_range`, `list_comments`, `list_named_ranges`, `export_csv`.
- *Structure* — `create_workbook`, `add_sheet`, `remove_sheet`, `rename_sheet`,
  `merge_cells`, `unmerge_cells`, `freeze_panes`, `set_auto_filter`, `add_named_range`.
- *Values* — `set_cell`, `write_cells`, `append_row`, `clear_cells`.
- *Layout and appearance* — `set_column_width`, `set_row_height`, `set_header_footer`,
  `add_hyperlink`, `style_cells`, `set_number_format`, `add_data_validation`,
  `add_conditional_format`.
- *Drawing* — `add_chart`, `add_image`, `add_comment`.

The workspace root comes from `--root`, then `$FERROXL_ROOT`, then the working directory,
and a path that escapes it is refused.

### Notes on behaviour

- **The 1900 leap-year bug is reproduced, not fixed.** Serial 60 is the phantom
  1900-02-29. `to_excel` skips it and `from_excel` does not, so the two disagree by one
  below it, exactly as openpyxl's do.
- **Charts and images are written but not read back.** openpyxl 3.1.5 writes the parts and
  its reader does not parse them; ferroxl matches that rather than being half-compatible
  in a different direction.
- **`COLOR_INDEX` has 56 entries**, not the 64 the Excel documentation mentions, because
  matching the upstream table is what callers depend on.
- **Self-closing tags are never written.** `<x/>` becomes `<x></x>`, as openpyxl's
  generator does.
- **Duplicate sheet titles get a counter appended** rather than being refused, again
  matching openpyxl.

### Deviations from openpyxl

Each of these is documented at the call site as well as in the README.

- Styles and number formats derive `PartialEq`, `Eq` and `Hash` using the same fields
  Python compares, because Rust's `HashMap` requires them.
- Tri-state attributes such as `<protection locked>` use a `ProtectionFlag` enum with
  `Inherit`, `Locked` and `Unlocked` in place of Python's `None`.
- `CategoryAxis` and `ValueAxis` are constructors returning the single `Axis` type, since
  in Python they differ only in class-level attribute defaults.
- `jdcal` is replaced by a local `gcal2jd`/`jd2gcal`, checked against `jdcal`'s own
  output in the tests.
- Images are read for their dimensions by parsing the PNG header, so PNG is the only
  supported format and PIL is not a dependency. The image bytes are stored verbatim.

### Automation

- `.github/workflows/ci.yml` -- on every push and pull request: `cargo fmt --check`,
  `cargo clippy -- -D warnings`, `cargo doc` with warnings denied, `cargo nextest run` on
  Linux, Windows and macOS, doctests as a separate step, and an MSRV job that reads
  `rust-version` from `Cargo.toml` so the two cannot drift apart.
- `.github/workflows/release.yml` -- on a `v*` tag: builds `ferroxl-mcp` for five targets,
  packages each with its licence and README, assembles and verifies `SHA256SUMS`, and
  publishes a draft release only once every archive is present.

### Verification

- 481 tests: 369 unit tests and 110 in the MCP server, plus two doctests.
- Values that came from Python are pinned rather than recomputed: the password hashes,
  the date serials, the Julian day numbers, the chart axis arithmetic and the
  `is_date_format` rule.
- Files written by ferroxl are opened with openpyxl 3.x and the values, styles, merges,
  validations, comments, defined names and freeze panes compared.

[0.1.0]: https://github.com/SV-stark/ferroxl/releases/tag/v0.1.0