# ferroxl

A Rust library for reading and writing Excel 2007 `.xlsx`/`.xlsm` files, plus a
[Model Context Protocol](https://modelcontextprotocol.io) server that exposes it to AI
agents.

ferroxl is a feature-parity port of [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl)
3.1.5. It follows the Python package's module layout, class names and semantics, so a
change can be traced back to the Python it mirrors. Where Python's behaviour cannot be
reproduced in Rust — hash-based equality, tri-state booleans, PIL-backed images — the
deviation is documented at the call site and the closest faithful behaviour is implemented
instead.

Parity is real but partial, and `PARITY.md` is the honest accounting: what is ported, what
differs, and what is not there at all. The largest absences are pivot tables, chart-only
sheets, and rich text. Named styles are done: all 49 of Excel's built-ins plus the workbook's
own, with `cell.style` reading back a name again.

Loading and re-saving no longer deletes what it does not model. Parts ferroxl has no API
for — pivot tables and their caches, slicers, query tables, threaded comments, ActiveX
controls, `customXml` — are carried through with the content types and relationships that
make them reachable, so an enterprise template survives being opened and saved. It does not
make those features editable; `PARITY.md` lists exactly what that does and does not cover.

Two capabilities are not ports at all, because openpyxl has no equivalent. **Dependency
tracing** answers what feeds a cell, what would go stale if one changed, and where the
cycles are. **Formula evaluation** (`recalculate()`) fills in cached values for the
formulas it can evaluate and reports the ones it cannot, rather than leaving every formula
blank or guessing at a number.

`PARITY.md` is measured by `tools/parity.py` against a real checkout of the Python, not
written from memory. It also says what the audit *cannot* see: the tool matches names, so a
function that behaves wrongly and an XML element that is silently dropped both read as
matched. Every field-level loss in it was found by reading both trees.

```
crates/
  ferroxl/       the library
  ferroxl-mcp/   the MCP server
```

## Contents

- [Installation](#installation)
- [Quick start](#quick-start)
- [Using it from Rust](#using-it-from-rust)
- [Using it from an AI agent](#using-it-from-an-ai-agent)
- [Reading a workbook](#reading-a-workbook)
- [Writing a workbook](#writing-a-workbook)
- [Formula evaluation](#formula-evaluation)
- [What is covered](#what-is-covered)
- [Feature parity](#feature-parity)
- [Dates and the two calendars](#dates-and-the-two-calendars)
- [Differences from openpyxl](#differences-from-openpyxl)
- [The MCP server](#the-mcp-server)
- [Development](#development)
- [Continuous integration and releases](#continuous-integration-and-releases)
- [License](#license)

## Installation

```toml
[dependencies]
ferroxl = "0.1.10"
```

Both crates are on crates.io — the [library](https://crates.io/crates/ferroxl) and the
[MCP server](https://crates.io/crates/ferroxl-mcp) — and are published in that order, because
the server depends on the library by version. See [docs/PUBLISHING.md](docs/PUBLISHING.md)
for how that works.

To work on it instead, from a checkout:

```console
$ git clone https://github.com/SV-stark/ferroxl
$ cd ferroxl
$ cargo build --workspace
$ cargo run --example build_and_read   # writes orders.xlsx and reads it back
```

To use a checkout from another project on the same machine, point at it:

```toml
[dependencies]
ferroxl = { path = "../ferroxl/crates/ferroxl" }
```

The library has no unsafe code and six dependencies: `chrono` for date arithmetic,
`quick-xml` for XML, `regex` for the few patterns that need one, `zip` for the package
container, `png` for reading an image's dimensions, and `thiserror` for the error enum.
The MCP server adds only `serde`, `serde_json`, `chrono` and `thiserror`.

## Quick start

```console
$ git clone https://github.com/SV-stark/ferroxl
$ cd ferroxl
$ cargo test --workspace        # 762 tests
$ cargo run --example build_and_read
wrote orders.xlsx
sheets: ["Sheet1", "Orders"]
dimension: A1:D6
merged: ["A6:D6"]
frozen at: Some("A2")
  Text("Item") | Text("Qty") | Text("Price") | Text("Total")
  Text("Bolt") | Number(10.0) | Number(1.5) | Formula("=B2*C2")
  Text("Nut") | Number(25.0) | Number(0.75) | Formula("=B3*C3")
A1 bold: true
```

## Using it from Rust

`crates/ferroxl/examples/build_and_read.rs` is a complete round trip — create a sheet, set
cells, style a header, merge a range, freeze the panes, save, load. Read it; it is about a
hundred lines and every line earns its place.

Two things catch people out, and both are openpyxl's behaviour rather than a Rust quirk:

- **`create_sheet` does not make the new sheet active.** `Workbook::new()` already made
  `Sheet1`, so `active_sheet_mut()` after a `create_sheet` is still `Sheet1` — a mistake
  that produces a file where the data is on the wrong sheet. Take the index
  `create_sheet` returns, as [Writing a workbook](#writing-a-workbook) does.

- **`merge_cells` blanks every cell but the top-left one**, so merge before writing the
  value or the value disappears. It also takes one range, and `set_number_format` takes one
  coordinate, not a range.

## Using it from an AI agent

`ferroxl-mcp` is the same library behind a Model Context Protocol server, so an agent can
read and edit workbooks instead of guessing at them.

```console
$ cargo run -p ferroxl-mcp -- --help
ferroxl-mcp 0.1.9 — a Model Context Protocol server for Excel workbooks

USAGE:
    ferroxl-mcp [--root <directory>]

OPTIONS:
    --root <directory>  Bound every path the tools may touch. Defaults to
                        $FERROXL_ROOT, then the working directory.
    -h, --help          Print this help and exit.
    -V, --version       Print the version and exit.

The server speaks JSON-RPC 2.0 over stdio, one message per line.
```

Point an MCP client at it. For Claude Code:

```console
$ claude mcp add ferroxl -- cargo run -p ferroxl-mcp -- --root ./spreadsheets
```

Or in a client that reads `mcpServers` from a config file:

```json
{
  "mcpServers": {
    "ferroxl": {
      "command": "ferroxl-mcp",
      "args": ["--root", "/path/to/spreadsheets"]
    }
  }
}
```

The agent then gets 41 tools — `list_sheets`, `read_cells`, `write_cells`, `summarize_range`,
`trace_precedents`, `add_chart`, `style_cells`, and so on. Two behaviours are worth knowing:

- A tool that ran and failed returns `isError: true` with the message in `content`, so the
  model can read what went wrong and fix its call. Only an unknown tool is a JSON-RPC error.
- A misspelled argument is **rejected**, not ignored. A tool that silently drops an argument
  is worse than one that refuses.

See [The MCP server](#the-mcp-server) for the full tool list and how values are mapped.

## Reading a workbook

```rust
use ferroxl::{load_workbook, LoadOptions};

let workbook = load_workbook("report.xlsx", LoadOptions::default())?;

for name in workbook.get_sheet_names() {
    println!("{name}");
}

let sheet = workbook.active_sheet()?;
for row in sheet.range_values("A1:D10")? {
    for value in row {
        println!("{value:?}");
    }
}

// One cell at a time, when a coordinate is all you have.
println!("{:?}", sheet.cell_value("B2"));
# Ok::<(), ferroxl::Error>(())
```

`load_workbook` accepts anything `AsRef<Path>`, and `load_workbook_from_bytes` takes an
already-read buffer, which is what a web service wants. `LoadOptions` carries the same
switches as the Python call:

| Field | Builder | Meaning |
| --- | --- | --- |
| `guess_types` | `guessing_types()` | Infer a cell's type from its text rather than trusting the stored one, so `"50%"` loads as `0.5` |
| `data_only` | `values_only()` | Return the value Excel last cached instead of the formula |
| `keep_vba` | `keeping_vba()` | Keep the original package bytes so the VBA project survives a save |

Three of openpyxl 3.1.5's `load_workbook` switches are implemented. Three are not, and the
differences run in both directions:

- **`guess_types` has no counterpart upstream.** openpyxl removed it in 3.0; it was a 2.x
  flag. ferroxl keeps it because the behaviour is useful and cheap, but it is not part of
  the surface being ported, so a script written against openpyxl 3 will not find it.
- **`read_only`, `keep_links` and `rich_text` are not implemented.** `read_only` would
  need a streaming loader — see [Pending](PARITY.md#1-there-is-no-read-only-loader).
  `keep_links` needs external-link parts, which are not read or written at all.
  `rich_text` needs the inline-runs model in cells, which does not exist; a rich-text cell
  is concatenated with its formatting discarded, which is what openpyxl does too when
  `rich_text=False`.

A value read back from a file is always reconstructed from the serial, so a date cell
reports a `DateTime` — the same as openpyxl, which also loses the distinction between
`date` and `datetime` on the way through a file. A value written in memory keeps its own
type, so `ws.set("A1", CellValue::Date(..))` reads back as a `Date`.

## Formula evaluation

A saved workbook carries no computed values unless you ask for them, which is what openpyxl
does. `Workbook::recalculate` fills them in:

```rust
# use ferroxl::{CellValue, Workbook};
let mut workbook = Workbook::new();
let sheet = workbook.active_sheet_mut().unwrap();
sheet.set("A1", CellValue::number(2.0)).unwrap();
sheet.set("A2", CellValue::number(3.0)).unwrap();
sheet.set("A3", CellValue::formula("=SUM(A1:A2)")).unwrap();

let report = workbook.recalculate();
assert_eq!(report.computed_count(), 1);
assert_eq!(workbook.active_sheet().unwrap().cached_value("A3"), Some(&CellValue::Number(5.0)));
```

It covers the operators, the aggregates, `IF`/`IFERROR`, `AND`/`OR`/`NOT` and the common
text functions — 31 in all, and `formula::supports("XLOOKUP")` answers whether a name is
handled without calling it and getting a wrong answer. A formula it cannot evaluate gets
**no** cached value and appears in `Recalculation::unresolved` with the reason.
`calcPr/@fullCalcOnLoad` is set, so Excel recomputes on open and a value the engine got
wrong cannot survive a human opening the file.

Formulas are evaluated in one pass over reading order, so a formula reading another
formula's cell sees it as blank. `Worksheet::trace_precedents` gives the order to do it
properly.

## Writing a workbook

```rust
use ferroxl::{CellValue, Style, Workbook};

let mut workbook = Workbook::new();

// `Workbook::new()` already made `Sheet1`, and `create_sheet` appends without
// making the new sheet active, so take the index it returns.
let summary = workbook.create_sheet(Some("Summary"))?;

let sheet = &mut workbook.worksheets[summary];
sheet.set("A1", CellValue::text("Item"))?;
sheet.set("B1", CellValue::text("Revenue"))?;
sheet.set("A2", CellValue::text("Widget"))?;
sheet.set("B2", CellValue::Number(1200.0))?;
sheet.set("C2", CellValue::Formula("=B2*1.2".to_string()))?;

let mut header = Style::new();
header.font.bold = true;
header.fill.fill_type = Some("solid".to_string());
header.fill.start_color = ferroxl::Color::new("FFDDDDDD".to_string());
for cell in ["A1", "B1", "C1"] {
    sheet.set_style(cell, header.clone())?;
}
sheet.set_freeze_panes("A2");

workbook.save("summary.xlsx")?;
# Ok::<(), ferroxl::Error>(())
```

`Workbook::save` writes the package to a path and consumes the workbook, which is what
makes "edit then save" a single linear sequence. `Workbook::to_bytes` returns the package
instead, and `ferroxl::writer::save_workbook_to` streams it to any `Write`.

## What is covered

| Module | Upstream | What it holds |
| --- | --- | --- |
| `cell` | `openpyxl/cell` | `Cell`, `CellValue`, data types, coordinates, read-only cells, formulas, shared formulas |
| `charts` | `openpyxl/charts` | all sixteen chart types, series, references, axes, titles, error bars, 3-D views |
| `comments` | `openpyxl/comments` | `Comment` and the VML shapes that carry them |
| `datavalidation` | `openpyxl/datavalidation` | `DataValidation` and its enums, cell-address collapsing |
| `date_time` | `openpyxl/date_time` | the two date calendars, serial conversion, ISO-8601 parsing, Julian days |
| `drawing` | `openpyxl/drawing` | `Image`, `Drawing`, `Shape`, anchors, EMU conversions |
| `exceptions` | `openpyxl/exceptions` | `Error` and its variants, openpyxl's exception hierarchy |
| `formatting` | `openpyxl/formatting` | conditional formats, colour scales, data bars, icon sets, `CellIsRule`, `FormulaRule` |
| `formula` | `openpyxl/formula` | `Translator` for copy-paste and fill, and the evaluator behind `recalculate()` |
| `namedrange` | `openpyxl/namedrange` | defined names, their destinations, and the range-string grammar |
| `reader` | `openpyxl/reader` | the whole load path: strings, styles, workbook, worksheets, comments, zip repair |
| `styles` | `openpyxl/styles` | fonts, gradient fills, borders, alignment, number formats, protection, named styles, the 49 built-ins, the indexed palette |
| `units` | `openpyxl/units` | the unit constants Excel uses (`points_to_pixels`, `emu_to_cm`, …) |
| `workbook` | `openpyxl/workbook` | `Workbook`, document properties, security, sheet management, calculation properties |
| `worksheet` | `openpyxl/worksheet` | `Worksheet`, dimensions, views, panes, protection, header/footer, tables, cell ranges, dependency tracing |
| `writer` | `openpyxl/writer` | the whole save path, the streaming writer, and the pass-through of unmodelled parts |
| `xml` | `openpyxl/xml` | an ElementTree-shaped element tree, a streaming writer, and the namespace constants |

Two capabilities have no upstream module because openpyxl has no equivalent: cell dependency
tracing (`worksheet::dependency`) and pass-through preservation (`reader::preserved`,
`workbook::preserved`, `writer::preserved`).

## Feature parity

[PARITY.md](PARITY.md) is the module-by-module account of what is implemented, what is
pending, and where the Rust version differs on purpose. It is generated from an audit of
the two source trees, so it can be re-run rather than believed:

```console
$ python tools/parity.py path/to/openpyxl/openpyxl
```

Of openpyxl 3.1.5's 994 top-level public names, 678 have no counterpart, and 32 of its 183
modules are matched name-for-name in full. The rest are concentrated in six packages that
have no Rust module to point at: `drawing/` (126), `worksheet/` (117), `chart/` (81),
`pivot/` (58), `descriptors/` (48) and `xml/` (41). What remains genuinely pending is
short: a read-only loader, reading the stored `<dimension>`, and reading charts back.
`PARITY.md` has the accounting per package.

Two caveats on that number, both of which cut against it. The audit matches *names*, so a
function that behaves wrongly reads as matched, and a name can match while the feature
behind it is absent — the figures are an upper bound on what works, not a measurement of
it. And the gap is not evenly spread: cell values, styles and the read/write round trip are
solid, while everything around the cell is thin.

## Dates and the two calendars

Excel's 1900 date system believes 1900 was a leap year, so it has a day — the phantom
1900-02-29 — that never existed. ferroxl reproduces this exactly, including the
consequences:

- Serial 1 is 1900-01-01 and serial 60 is the phantom day.
- `to_excel` skips the phantom day, so 1900-02-28 is serial 59 and 1900-03-01 is serial 61.
- `from_excel` does not skip it, so serial 60 reads back as 1900-02-28 and serial 1 reads
  as 1899-12-30. The two functions disagree by one below the phantom day and agree above
  it. openpyxl behaves the same way; the tests pin both halves of the asymmetry.
- A 1904 workbook (the classic Mac calendar) has no phantom day, and its serials round-trip
  against the 1904 epoch. Reading one used to convert against 1900 instead — every date came
  back 1462 days early — so `display_value` now takes the workbook's base date.
  `tools/upstream/corpus.py` checks this against openpyxl's own `mac_date.xlsx`.

A cell's serial only becomes a date when its number format says so, so
`ws.set_number_format("A1", "yyyy-mm-dd")` is what turns `40196` into a date.

## Differences from openpyxl

These are the places where a literal port is impossible or would be a bad idea. Each is
also documented at the call site.

**Equality and hashing.** Python compares styles and number formats by the fields that
matter — a number format's code, not its index — and lets everything else fall back to
identity. Rust's `HashMap` needs `Hash` and `Eq`, so ferroxl derives them but implements
them by the same comparison Python uses: `NumberFormat` hashes its code alone, and
`Style` hashes its visual fields. The consequence is that two styles that look identical
dedupe to one entry in the style table, which is what Excel does too.

**Tri-state booleans.** OOXML attributes like `<protection locked>` are genuinely
three-valued: present-and-true, present-and-false, and absent. Python models this with
`None`; Rust's `bool` cannot. ferroxl uses a `ProtectionFlag` enum with `Inherit`, `Locked`
and `Unlocked` for those attributes, and a plain `bool` where the specification really does
mean two states.

**`CategoryAxis` and `ValueAxis`.** In Python these are classes that differ only in their
class-level attribute defaults. In Rust they are constructors returning the single `Axis`
type, so `CategoryAxis::new()` and `ValueAxis::new()` produce an `Axis` configured the way
the corresponding Python class would configure itself.

**Images.** openpyxl opens an image with PIL to read its size, which makes PIL a hard
dependency for anyone embedding a picture. ferroxl parses the PNG header directly and
supports PNG only. The bytes are stored verbatim, so the image itself is untouched.

**Julian days.** openpyxl depends on `jdcal`, which is unmaintained and not on PyPI any
more. ferroxl reimplements `gcal2jd` and `jd2gcal` and pins them against `jdcal`'s own
output in the tests.

**The colour palette.** openpyxl's `COLOR_INDEX` has 56 entries, not the 64 Excel
documents. ferroxl ships the same 56 and rejects an index past the end, because matching
the upstream table is more useful than matching the documentation.

**Charts and images are written but not read back.** openpyxl 3.1.5 writes chart and drawing
parts but its reader does not parse them, so a reloaded workbook reports no charts. ferroxl
matches that, rather than being half-compatible in a different way.

**Self-closing tags.** Python's `XMLGenerator` never writes `<x/>`; it writes `<x></x>`.
ferroxl's streaming writer does the same, so the bytes match.

## The MCP server

`ferroxl-mcp` speaks JSON-RPC 2.0 over stdio, one message per line, which is what a locally
launched MCP server uses.

```console
$ cargo run -p ferroxl-mcp -- --root ./spreadsheets
```

`--root` bounds every path a tool may touch. Without it the server uses `$FERROXL_ROOT`, and
without that the working directory. This is a convenience that keeps an agent inside a
directory you chose, not a security sandbox: a caller that can launch this process can
already do whatever it likes. The root is created if it does not exist.

The client configuration looks like this:

```json
{
  "mcpServers": {
    "ferroxl": {
      "command": "ferroxl-mcp",
      "args": ["--root", "/path/to/spreadsheets"]
    }
  }
}
```

### The tools

Forty-one tools, grouped by what they are for. Every argument is documented in the schema
the server advertises at `tools/list`, and the grouping below is the one the server itself
uses.

**Inspection** — `list_sheets`, `describe_sheet`, `read_cells`, `read_formulas`,
`trace_precedents`, `trace_dependents`, `check_circular_references`, `add_data_bar`,
`add_icon_set`, `add_table`, `describe_table`, `set_gradient_fill`, `search_values`,
`summarize_range`, `list_comments`, `list_named_ranges`, `export_csv`.

The three `trace_*` tools are the ones openpyxl cannot answer at all: what feeds a cell,
what would go stale if a cell changed, and every cycle in a sheet as a closed path. Excel
refuses to calculate a workbook with a cycle, so `check_circular_references` is worth
running before trusting a file you did not create. Summaries name at most a dozen cells and
then give the count, because three thousand dependents cannot change a model's next decision
but will consume the context window needed to make one.

**Structure** — `create_workbook`, `add_sheet`, `remove_sheet`, `rename_sheet`,
`merge_cells`, `unmerge_cells`, `freeze_panes`, `set_auto_filter`, `add_named_range`.

**Values** — `set_cell`, `write_cells`, `append_row`, `clear_cells`.

**Layout** — `set_column_width`, `set_row_height`, `set_header_footer`, `add_hyperlink`.

**Formatting** — `style_cells`, `set_number_format`, `add_data_validation`,
`add_conditional_format`.

**Media and annotations** — `add_chart`, `add_image`, `add_comment`.

The server does not evaluate formulas: a cell written with a formula has no cached result,
exactly as openpyxl writes it, so `read_cells` returns the formula rather than a number.
`trace_precedents` and `trace_dependents` answer the questions about the graph that a
calculated value would otherwise have been needed for.

### How values are interpreted

A tool that writes a value takes JSON, and the mapping is deliberately narrow so a model
does not have to learn a wrapper object:

| JSON | Stored as |
| --- | --- |
| a string beginning `=` | a formula |
| `"2010-01-18"` | a date |
| `"2010-01-18T14:15:20"` | a timestamp |
| `"14:15:20"` | a time of day |
| `"50%"` | the number `0.5` |
| a number | a number |
| `true` / `false` | a boolean |
| `null` | an empty cell |
| anything else | text |

Values come back with their type spelled out, so a formula is never confused with the text
it evaluates to:

```json
{ "cell": "B2", "value": { "type": "formula", "value": "=SUM(B2:B9)" } }
```

### Errors

A tool that ran and failed returns `isError: true` with the message in `content`, so the
model can read what went wrong and correct its call. Only an unknown method or tool is a
JSON-RPC error (`-32601`). A misspelled argument is rejected rather than ignored, because
a tool that silently drops an argument is worse than one that refuses.

## Development

```console
$ cargo build --workspace                  # build
$ cargo test --workspace                   # 762 tests
$ cargo nextest run --workspace            # the same tests, in parallel; this is what CI runs
$ cargo clippy --workspace --all-targets -- -D warnings
$ cargo fmt --all --check
$ cargo doc --workspace --no-deps          # API documentation
$ python tools/check_readme.py             # the README examples are the doctests
$ python tools/parity.py path/to/openpyxl/openpyxl   # audit the public surface
```

The examples in this README are the library's own doctests, checked by
`tools/check_readme.py`, so they cannot drift away from the API.

The library is `#![deny(missing_docs)]`, so every public item carries a doc comment.
Tests live next to the code they cover, and the values they assert were taken from the
Python original rather than from this implementation — the password hashes, the date
serials, the Julian day numbers, the chart axis arithmetic and the `is_date_format` rule
were all checked against openpyxl or `jdcal` and the result pinned in a test.

The workspace is verified against real openpyxl: files this library writes are opened with
openpyxl 3.x and the values, styles, merges, validations, comments, names and freeze panes
are compared.

Two harnesses cover the MCP server, because its output is what an agent's work actually
becomes and openpyxl is how anyone else will read it:

```console
$ python tools/mcp_parity.py    # drive all 41 tools, then read each file back with openpyxl
$ python tools/mcp_required.py  # call each tool with only its schema's required arguments
$ python tools/mcp_limits.py    # second calls, empty ranges, bad input, does the file still open?
```

`tools/upstream/` goes further, using openpyxl's own test suite and its fixture corpus — 161
test files, about 1,700 test functions, and a set of real workbooks from Excel, LibreOffice and
Mac Excel:

```console
$ curl --fail --location --output openpyxl.tar.gz \
      https://foss.heptapod.net/openpyxl/openpyxl/-/archive/3.1.5/openpyxl-3.1.5.tar.gz
$ mkdir -p ../openpyxl && tar -xzf openpyxl.tar.gz --strip-components=1 -C ../openpyxl
$ python tools/upstream/run.py ../openpyxl
```

An archive, not a clone, and not by preference: openpyxl 3.x is hosted on Heptapod, which
refuses anonymous git-over-https, and the GitHub mirror stops at 1.9.0 with no 3.x tags. See
[tools/upstream/README.md](tools/upstream/README.md) for the details.

`corpus.py` reads all twelve workbooks openpyxl ships — from Excel, LibreOffice, Mac Excel and
its own reader fixtures — with both implementations and compares them cell by cell, with
openpyxl as the oracle. `manifest.py` maps all 161 test files onto ferroxl modules,
so the parity claim is a checklist rather than an assertion. See
[tools/upstream/README.md](tools/upstream/README.md) for what is and is not covered.

Between them these found seven bugs that `tools/parity.py` cannot see, because none of them is
a missing name: a conditional format's differential style was never written, an embedded image
was unreadable to openpyxl, `add_image` discarded its `anchor`, `add_chart` pointed every series
at one unrelated cell, `merge_cells` accepted a backwards range and left a workbook that no
longer opened, the 1904 date system was ignored on read — a real Mac Excel workbook came back
four years and one day early — and a workbook whose part was not named `xl/workbook.xml` was
rejected outright. All seven passed every unit test, and all seven run in CI now.

## Continuous integration and releases

Three workflows under `.github/workflows`.

**`ci.yml`** runs on every push to `main` and every pull request, in three jobs:

- **checks** -- `cargo fmt --check`, `cargo clippy -- -D warnings`, `cargo doc` with
  `RUSTDOCFLAGS: -D warnings`, and `tools/check_readme.py`, which fails if a README example
  has drifted from the doctest it mirrors.
- **test** -- `cargo nextest run --workspace --no-fail-fast` on Linux, Windows and macOS,
  then `cargo test --doc` separately. nextest does not run doctests, so without that second
  step the README examples would quietly stop being verified.
- **msrv** -- `cargo check` on whatever `rust-version` says in `Cargo.toml`, read from the
  manifest so the job cannot pass on a toolchain older than the one the crate promises.

The dependency cache is shared across the matrix, so the first job to finish warms it for
the rest.

**`release.yml`** runs when a tag of the form `v0.1.9` is pushed. It builds `ferroxl-mcp`
for five targets:

| Target | Archive |
| --- | --- |
| `x86_64-unknown-linux-gnu` | `tar.gz` |
| `aarch64-unknown-linux-gnu` | `tar.gz` |
| `x86_64-pc-windows-msvc` | `zip` |
| `x86_64-apple-darwin` | `tar.gz` |
| `aarch64-apple-darwin` | `tar.gz` |

Each archive carries the binary with `LICENSE` and `README.md` beside it -- a
redistributable binary without its licence is a licence violation. A build that cannot
print `--version` fails rather than shipping. The tag is checked against the version in
`Cargo.toml`, so a filename cannot lie about what is inside it. The per-target `.sha256`
files are combined into one `SHA256SUMS` and verified before anything is published, and the
release is drafted first and only published once the file count matches the matrix.

**`publish.yml`** runs when a GitHub release is *published*, not when the tag is pushed, so
a failed binary build cannot put a crate on crates.io. It publishes `ferroxl` first, waits
for the registry index to list it, then publishes `ferroxl-mcp` — the server depends on the
library by version, so publishing both at once fails in a way that reads like a version
number is wrong rather than a race. See [docs/PUBLISHING.md](docs/PUBLISHING.md).

```console
$ git tag v0.1.9 && git push origin v0.1.9
```

## License

MIT, the same as openpyxl. See [LICENSE](LICENSE).
