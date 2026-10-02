# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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