# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.10] - 2026-10-06

Four bugs, and they share a failure mode this project cares most about: each accepted the work,
reported success, and produced a file that silently meant something else. Two of them destroyed
data -- one made a workbook unopenable, the other emptied every cell holding an ampersand.

### Fixed

- **A chart or an image was destroyed by the next write to the workbook, and left the file
  unreadable.** `add_chart` and `add_image` both reported success, and the next tool call --
  `set_cell`, `write_cells`, even `style_cells` -- deleted `xl/drawings/drawingN.xml` while the
  sheet's `<drawing>` element and its relationship were preserved verbatim. The package was left
  with a relationship naming a part that was no longer in it, so openpyxl raised
  `KeyError: "There is no item named 'xl/drawings/drawing1.xml' in the archive"` and Excel would
  have reported the file as needing repair. Both tools were effectively single-shot: they worked
  once, and the next call bricked the workbook. Editing a real file that already had images in
  it took a single `set_cell`.

  The cause was `is_writer_owned` claiming `xl/drawings/drawingN.xml`, `xl/charts/chartN.xml`
  and `xl/media/*` as parts the writer produces. That holds only for a sheet whose model still
  has drawings, and ferroxl does not read charts or images back -- so for a loaded workbook the
  writer produced none at all, while the preserved relationship still pointed at them. Drawings,
  charts and media are now preserved instead, and `write_parts` renders a preserved part's own
  `.rels`, which is what says which chart or image a drawing holds. Without that the part
  survives as a set of anchors pointing at nothing.

  Part names are now allocated past whatever is preserved rather than always from 1, which
  fixed a second fault at the same time: a chart on one sheet and an image on another both
  claimed `drawing1.xml`, so the second overwrote the first and the chart vanished with nothing
  reporting it. Two sheets now get `drawing1.xml` and `drawing2.xml`.

  Found by building a workbook from scratch through the server -- a case the earlier suite never
  reached, because every one of its drawing cases made its drawing the only call on a fresh copy.
  See `tools/mcp_real_files.py`, which now writes to the workbook *after* adding a drawing, and
  `tools/bug_demo.py`.

  This is preservation, not modelling: a chart still cannot be read back and adjusted, which
  `PARITY.md` records under Pending. That entry also had the direction of the gap wrong -- it
  claimed openpyxl cannot read charts back either, and openpyxl 3.1.5 round-trips both intact.

- **Every `&`, `<`, `>`, `"` and `'` was silently deleted from every element's text.** The XML
  reader reported an entity reference as its own `GeneralRef` event rather than inline, and the
  parser's catch-all arm dropped it, so `Tom & Jerry` came back as `Tom  Jerry`. This was not
  confined to headers: it applied to cell text, defined names, comments, table columns, every
  part of every workbook, and no error was raised anywhere. A workbook whose cells contained an
  ampersand lost that text on the first save, and `read_cells` reported the cells as empty.

  The five entities and `&#NN;` / `&#xNN;` are now resolved, and an entity XML does not define is
  an error rather than a silent omission -- which is what turned this into data loss in the first
  place.

- **A cell holding an inline string read back as empty.** `<c t="inlineStr">` keeps its text in
  `<is><t>`, which is where openpyxl puts a string when the workbook has no shared string table
  -- the default for a file it has just created. `DataType::InlineString` was modelled but neither
  the reader nor the writer handled it: the reader looked only in `<v>`, and the writer put the
  text *there* under an `inlineStr` type, which is a file no reader can read a value from. Every
  string cell in a workbook openpyxl had written came back blank after a ferroxl save.

- **A printed header or footer survived one save and was gone by the second.** `&L`, `&C` and `&R`
  introduce a section, and openpyxl's parser matches them glued to their text -- `&Lleft` is a left
  section carrying `left`. Comparing whole `&`-delimited fields found no `L` field, so the section
  was dropped, and since `section_bounds` was the only route from a file's header to the model,
  nothing carried it into the next save. This is the defect the entity fix above first showed
  itself through, and fixing one without the other leaves a header that still disappears.

## [0.1.9] - 2026-10-06

Three library bugs, all silent: each accepted a call, reported success, and left a file that
said something else. None was a missing name, so `tools/parity.py` -- which matches names --
could not see any of them. Two of the three are in the writer, and both lose what they were
given.

### Fixed

- **A row height set on a row that holds no cells was discarded on save.** `set_row_height`
  reported the rows it set and the height it set, and neither reached the file. `<row>` is the
  only element a row height can live in, and the writer built its `<row>` elements from the
  cells alone, so a row present only in `row_dimensions` was never written - taking its height,
  its hidden flag, its outline level and its style with it. On a sheet whose fourth row is
  empty, `set_row_height(rows="2:4")` set rows 2 and 3 and dropped 4; `set_row_height(rows="9")`
  on an empty row did nothing at all. openpyxl keeps such a row, and so must this. Rows are now
  merged from both sources and written ascending, and a row dimension with nothing set still
  gets no element of its own.
  Found by calling every tool against real workbooks instead of fixtures the server wrote
  itself, checking each mutation by reading the result back with openpyxl rather than by asking
  the server what it thought it had done - see `tools/mcp_real_files.py`.

- **`<col>` elements were written in the wrong order.** Column dimensions are keyed by letters,
  and `"Z"` sorts before `"AA"`, so a sheet given widths on `Z:AB` wrote its `<col>` list as
  `1, 27, 28, 2, 3, 4, 26`. Every width was correct and openpyxl reads all of them back; only
  the order departed from what Excel and openpyxl both write. Entries are now ordered by column
  number. The unit test covering this set a single column, so it could not have seen it.

- **A workbook part not named `xl/workbook.xml` was rejected outright.** The reader assumed the
  conventional name rather than following the package relationships, so a valid OOXML package
  whose workbook was called something else failed to load at all — "the archive has no
  xl/workbook.xml part". openpyxl ships a fixture for precisely this case,
  `tests/data/reader/nonstandard_workbook_name.xlsx`, whose workbook is `xl/workbook10.xml`, and
  openpyxl reads it without complaint. The workbook part is now resolved through the
  `officeDocument` relationship in `_rels/.rels`, with its own `.rels` found beside it rather
  than at a fixed path. Naming is still tried first, so the common case is unaffected and a
  relationship pointing at a part the archive does not contain cannot shadow a real one.

  Found by extending `tools/upstream/corpus.py` from `genuine/` to the whole of `tests/data`,
  which took coverage from five workbooks to all twelve — six of them `.xlsm` and the rest from
  openpyxl's reader fixtures, none of which had been checked before.

- **The parity CI job could never have passed.** Two faults, both in the harness plumbing and
  neither in the library. `find_binary()` looked for the server one level *above* the
  workspace instead of in cargo's own `target/`, so it resolved only on a machine exporting
  `CARGO_TARGET_DIR`; and the step that fetches openpyxl cloned `github.com/theorchard/openpyxl`,
  which does not exist. openpyxl 3.x is hosted on Heptapod, whose anonymous git-over-https is
  refused, and the GitHub mirror is stuck at 1.9.0 with no 3.x tags — so the pin is now an
  archive at tag 3.1.5, verified byte-identical to a checkout. The step also asserts the
  fixture it depends on is present before running, because an empty corpus would have made
  `corpus.py` report a clean run on nothing.
- `find_binary()` now warns when the binary is older than the sources under `crates/`. A
  harness that locates a stale build reports the previous build's behaviour and calls it a
  pass, which is the failure mode this whole change set exists to prevent.
- The openpyxl links in `README.md` and `PARITY.md` pointed at a repository that 404s.

The three fixes above are in the library and are what this release ships. The two harness
faults below are in `tools/`, `.github/`, the repository README and `PARITY.md`, none of which
are part of either package.

## [0.1.8] — 2026-10-05

Six bugs, all of the same shape: a name matched, every unit test passed, and the feature did
something other than what it claimed. None was a missing name, so `tools/parity.py` -- which
matches names -- could not see any of them. Five produced a well-formed file that silently meant
something else; one destroyed the file.

### Fixed

- **The 1904 date system was ignored on read.** `Cell::display_value` converted every serial
  against the 1900 epoch, hardcoded, because it was not passed the workbook's date system. A
  real Mac Excel workbook carrying `date1904="true"` therefore read every date **1462 days
  early** -- four years and one day. `PARITY.md` claimed 1904 workbooks "round-trip exactly".
  The write direction was already correct, so the damage was confined to reading: a date read
  from a 1904 workbook and shown to a user was wrong and nothing said so.
  Found by running openpyxl's own `tests/data/genuine/mac_date.xlsx` through both
  implementations and comparing, in `tools/upstream/corpus.py`.
  `display_value` now takes the workbook's base date; the old one-argument form is kept as the
  1900 default for a caller that genuinely has no workbook in hand.

- **`merge_cells` accepted a range whose corners ran backwards, and wrote a file nothing can
  open.** `merge_cells("C3:A1")` recorded `<mergeCell ref="C3:A1"/>`, which openpyxl rejects
  while *loading* — `1 must be greater than 3` — so the call reported success and left a workbook
  that no longer opened, with the failure surfacing on some later unrelated read. This is the
  worst shape the other four take: it destroys the file rather than quietly losing a feature.
  Backwards ranges are now refused at the call site, as openpyxl's own `merge_cells` refuses
  them, and both the column and the row corner are checked.
- **A conditional format's differential style was written nowhere.** `Rule::dxf` is set when the
  rule is built and belongs in `styles.xml`'s `<dxfs>` once the package is assembled, but
  `ConditionalFormatting::collect_dxf_styles` — which performs that move — was defined and never
  called. Every rule carrying a font colour, fill or bold was written with an empty `<dxfs>`, so
  it matched and highlighted nothing: a rule that appears to do nothing. It now runs before the
  style tables are built, because both halves depend on it — the worksheet needs the assigned
  `dxfId` and `styles.xml` needs the collected styles — and it appends to a loaded workbook's
  existing list rather than replacing it, so a loaded differential style keeps its index.
- **An embedded image was invisible to openpyxl.** `<xdr:cNvPicPr>` was written carrying
  `noChangeAspect` and `noChangeArrowheads`, which belong on its `<a:picLocks>` child.
  openpyxl's `NonVisualPictureProperties` accepts neither, so its reader raises a `TypeError` —
  and a reader that raises drops every image in the drawing. Excel renders it regardless, which
  is what let this stand.
- **`add_image` ignored its `anchor` argument.** The anchor was validated against the sheet and
  then discarded, leaving the drawing on its default `Absolute` anchor. openpyxl's
  `AbsoluteAnchor` has no `pic` attribute at all, so the image was unreadable there even once the
  `cNvPicPr` attributes were fixed. Images are now anchored to the cell they were asked for,
  which is the one-cell anchor openpyxl itself writes.
- **`add_chart` pointed every series at one unrelated cell.** The handler's range parser returned
  `(first column, point count)`, but `Reference` takes 0-based `(row, column)` *corners* — so a
  series over `B2:B3` was written as `'Data'!$C$3`, and `pos2` was never passed at all. The
  parser now returns both corners in the order `Reference` wants, and the categories argument is
  applied instead of being accepted and dropped.

### Added

- **`tools/mcp_client.py`**, a line-delimited JSON-RPC client that keeps one server process alive,
  so a test can make a sequence of calls against one workbook the way an agent would. Everything
  is protocol-level: no Rust is imported and nothing is stubbed.
- **`tools/mcp_parity.py`**, which drives all 41 tools and then opens what they wrote with real
  openpyxl — 127 checks over values, formulas, styles, merges, validations, conditional formats,
  tables, comments, charts, images, named ranges and defined names. Every test gets its own
  workbook, so the order of the file cannot change a result.
- **`tools/mcp_required.py`**, which calls every tool with only the arguments its schema declares
  required. A model reads `tools/list`, sends those, and expects the call to work; if the handler
  then complains about an argument nothing in the schema marked required, the schema is lying to
  the only reader it has. All 41 pass, with five documented exceptions where JSON Schema cannot
  express the requirement (an argument needed only for one value of another, or at least one of
  six) and the error message names what is missing.
- **`tools/mcp_limits.py`**, which asks what the happy-path suite does not: what happens on the
  *second* call to a mutating tool, on an empty range, on a mistyped coordinate, and whether the
  workbook still opens afterwards. 13 checks, and it prints the gaps it does not cover rather than
  implying they work. This is where the backwards-merge bug surfaced — the tool reported success
  and the damage was only visible in the file.
- **`tools/upstream/`**, which takes openpyxl's own test suite and its fixture corpus -- 161 test
  files, about 1,700 test functions -- as far as they can be taken. They are Python calling a
  Python API, so they cannot be *run* against a Rust library, and pretending otherwise would be
  the easiest way to make this look covered when it is not. What is possible:
  - **`corpus.py`** reads every real workbook in `tests/data/genuine/` with both implementations
    and compares them cell by cell, with openpyxl as the oracle. This found the 1904 bug.
  - **`expectations.py`** cross-checks the numbers openpyxl's tests pin against this project's
    tests: 152 of 158 numeric expectations appear in both. The six that do not are reported.
  - **`manifest.py`** maps all 161 test files onto ferroxl modules, so "we have parity" is a
    checklist that can be argued with rather than a claim.
  - **`survey.py`** reports what is in the wider fixture corpus and what ferroxl can read of it.
  - A **CI job** now runs all of it, so none of the six can regress silently. `tools/requirements.txt`
    pins the reference exactly: a version range would let a new openpyxl release turn a green
    build red for reasons unconnected to the commit.

## [0.1.7] — 2026-10-02

Everything since 0.1.6. Four features, and two of them fix failures that were silent: a
worksheet that disappeared from a workbook with no error, and a `<pivotCaches>` element dropped
on the way out. Both were invisible to `tools/parity.py`, which matches names and cannot see
either.

### Added

- **Pass-through preservation, so a round trip stops being destructive.** `Workbook::preserved`
  holds every part, content type, relationship and `<workbook>`/`<worksheet>` child the writer
  does not produce, and the writer writes them back. Copying the bytes is only the easy half: a
  part nobody can reach is inert, so the *relationship* is preserved too, its id remapped when
  the writer has already used it, and the `r:id` in the referencing element rewritten to match.
  Without that a pivot table survives on disk with nothing pointing at it -- a file that opens,
  shows the right cells, and has quietly lost the table.
  What still does not survive is recorded in the module docs rather than left to be discovered:
  unknown *attributes*, content inside `<sheetData>`, and byte-identical zip entries.
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

- **`Workbook::recalculate()`**, so a saved workbook carries values instead of blanks. Opt-in,
  and never silent: a formula it cannot evaluate gets no cached value at all and is named in
  `Recalculation::unresolved` with the reason. `calcPr/@fullCalcOnLoad` is set so Excel
  recomputes on open, which is what stops a mistake here from persisting.
- **`formula::eval`**: a tokenizer and recursive-descent parser for Excel's precedence,
  including the `-2^2 = 4` case and `%` as a postfix operator. 31 functions -- the arithmetic
  and comparison operators, `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `PRODUCT`, the
  rounding family, `IF`, `IFERROR`, `AND`/`OR`/`NOT`, and the common text functions. `supports`
  answers whether a function is handled, and a test asserts the list matches the dispatcher, so
  a name cannot be advertised without being implemented.
- **`Worksheet::cached_value` / `set_cached_value`**, the `<v>` half of a formula cell, read on
  load and written on save with the right `t` attribute (`str`, `b`, or `e` -- not the
  shared-string `s` a literal uses, because a computed string is not in the string table).


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

## [0.1.6] — 2026-10-02

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

[Unreleased]: https://github.com/SV-stark/ferroxl/compare/v0.1.10...HEAD
[0.1.10]: https://github.com/SV-stark/ferroxl/releases/tag/v0.1.10
[0.1.9]: https://github.com/SV-stark/ferroxl/releases/tag/v0.1.9
[0.1.8]: https://github.com/SV-stark/ferroxl/releases/tag/v0.1.8
[0.1.7]: https://github.com/SV-stark/ferroxl/releases/tag/v0.1.7
[0.1.6]: https://github.com/SV-stark/ferroxl/releases/tag/v0.1.6
[0.1.5]: https://github.com/SV-stark/ferroxl/releases/tag/v0.1.5
[0.1.4]: https://github.com/SV-stark/ferroxl/releases/tag/v0.1.4
[0.1.3]: https://github.com/SV-stark/ferroxl/releases/tag/v0.1.3
[0.1.2]: https://github.com/SV-stark/ferroxl/releases/tag/v0.1.2
[0.1.0]: https://github.com/SV-stark/ferroxl/releases/tag/v0.1.0