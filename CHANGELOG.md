# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] — 2026-10-02

The first release. ferroxl is a Rust port of openpyxl 1.9.0: the same modules, the same
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
- **Charts and images are written but not read back.** openpyxl 1.9 writes the parts and
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