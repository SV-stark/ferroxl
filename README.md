# lexcel

A Rust library for reading and writing Excel 2007 `.xlsx`/`.xlsm` files, plus a
[Model Context Protocol](https://modelcontextprotocol.io) server that exposes it to AI
agents.

lexcel is a feature-parity port of [openpyxl](https://github.com/theorchard/openpyxl).
It follows the Python package's module layout, class names and semantics, so a change can
be traced back to the Python it mirrors. Where Python's behaviour cannot be reproduced in
Rust — hash-based equality, tri-state booleans, PIL-backed images — the deviation is
documented at the call site and the closest faithful behaviour is implemented instead.

```
crates/
  lexcel/       the library
  lexcel-mcp/   the MCP server
```

## Contents

- [Installation](#installation)
- [Quick start](#quick-start)
- [Using it from Rust](#using-it-from-rust)
- [Using it from an AI agent](#using-it-from-an-ai-agent)
- [Reading a workbook](#reading-a-workbook)
- [Writing a workbook](#writing-a-workbook)
- [What is covered](#what-is-covered)
- [Feature parity](#feature-parity)
- [Dates and the two calendars](#dates-and-the-two-calendars)
- [Differences from openpyxl](#differences-from-openpyxl)
- [The MCP server](#the-mcp-server)
- [Development](#development)
- [License](#license)

## Installation

lexcel is not on crates.io yet, so depend on the git tag:

```toml
[dependencies]
lexcel = { git = "https://github.com/SV-stark/lexcel", tag = "v1.9.0" }
```

Or work from a checkout, which is what you want if you want to change it:

```console
$ git clone https://github.com/SV-stark/lexcel
$ cd lexcel
$ cargo build --workspace
$ cargo run --example build_and_read   # writes orders.xlsx and reads it back
```

To use it from another project on the same machine, point at the checkout:

```toml
[dependencies]
lexcel = { path = "../lexcel/crates/lexcel" }
```

The library has no unsafe code and six dependencies: `chrono` for date arithmetic,
`quick-xml` for XML, `regex` for the few patterns that need one, `zip` for the package
container, `png` for reading an image's dimensions, and `thiserror` for the error enum.
The MCP server adds only `serde`, `serde_json` and `chrono`.

## Quick start

```console
$ git clone https://github.com/SV-stark/lexcel
$ cd lexcel
$ cargo test --workspace        # 481 tests
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

`crates/lexcel/examples/build_and_read.rs` is a complete round trip — create a sheet, set
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

`lexcel-mcp` is the same library behind a Model Context Protocol server, so an agent can
read and edit workbooks instead of guessing at them.

```console
$ cargo run -p lexcel-mcp -- --root ./spreadsheets
lexcel-mcp 1.9.0 — a Model Context Protocol server for Excel workbooks

USAGE:
    lexcel-mcp [--root <directory>]
```

Point an MCP client at it. For Claude Code:

```console
$ claude mcp add lexcel -- cargo run -p lexcel-mcp -- --root ./spreadsheets
```

Or in a client that reads `mcpServers` from a config file:

```json
{
  "mcpServers": {
    "lexcel": {
      "command": "lexcel-mcp",
      "args": ["--root", "/path/to/spreadsheets"]
    }
  }
}
```

The agent then gets 33 tools — `list_sheets`, `read_cells`, `write_cells`, `summarize_range`,
`add_chart`, `style_cells`, and so on. Two behaviours are worth knowing:

- A tool that ran and failed returns `isError: true` with the message in `content`, so the
  model can read what went wrong and fix its call. Only an unknown tool is a JSON-RPC error.
- A misspelled argument is **rejected**, not ignored. A tool that silently drops an argument
  is worse than one that refuses.

See [The MCP server](#the-mcp-server) for the full tool list and how values are mapped.

## Reading a workbook

```rust
use lexcel::{load_workbook, LoadOptions};

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
# Ok::<(), lexcel::Error>(())
```

`load_workbook` accepts anything `AsRef<Path>`, and `load_workbook_from_bytes` takes an
already-read buffer, which is what a web service wants. `LoadOptions` carries the same
switches as the Python call:

| Field | Builder | Meaning |
| --- | --- | --- |
| `guess_types` | `guessing_types()` | Infer a cell's type from its text rather than trusting the stored one, so `"50%"` loads as `0.5` |
| `data_only` | `values_only()` | Return the value Excel last cached instead of the formula |
| `keep_vba` | `keeping_vba()` | Keep the original package bytes so the VBA project survives a save |

Those three are the whole of openpyxl 1.9's `load_workbook` switches. Rich text is
concatenated with its formatting discarded, which is also what openpyxl 1.9 does.

A value read back from a file is always reconstructed from the serial, so a date cell
reports a `DateTime` — the same as openpyxl, which also loses the distinction between
`date` and `datetime` on the way through a file. A value written in memory keeps its own
type, so `ws.set("A1", CellValue::Date(..))` reads back as a `Date`.

## Writing a workbook

```rust
use lexcel::{CellValue, Style, Workbook};

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
header.fill.start_color = lexcel::Color::new("FFDDDDDD".to_string());
for cell in ["A1", "B1", "C1"] {
    sheet.set_style(cell, header.clone())?;
}
sheet.set_freeze_panes("A2");

workbook.save("summary.xlsx")?;
# Ok::<(), lexcel::Error>(())
```

`Workbook::save` writes the package to a path and consumes the workbook, which is what
makes "edit then save" a single linear sequence. `Workbook::to_bytes` returns the package
instead, and `lexcel::writer::save_workbook_to` streams it to any `Write`.

## What is covered

| Module | Upstream | What it holds |
| --- | --- | --- |
| `cell` | `openpyxl/cell` | `Cell`, `CellValue`, data types, coordinates, read-only cells, formulas, shared formulas |
| `charts` | `openpyxl/charts` | `Chart`, bar/line/scatter/pie, series, references, axes, titles, data labels |
| `comments` | `openpyxl/comments` | `Comment` and the VML shapes that carry them |
| `datavalidation` | `openpyxl/datavalidation` | `DataValidation` and its enums, cell-address collapsing |
| `date_time` | `openpyxl/date_time` | the two date calendars, serial conversion, ISO-8601 parsing |
| `drawing` | `openpyxl/drawing` | `Image`, `Drawing`, `Shape`, anchors, EMU conversions |
| `exceptions` | `openpyxl/exceptions` | `Error` and its variants, openpyxl's exception hierarchy |
| `formatting` | `openpyxl/formatting` | conditional formats, colour scales, icon sets, `CellIsRule`, `FormulaRule` |
| `namedrange` | `openpyxl/namedrange` | defined names, their destinations, and the range-string grammar |
| `reader` | `openpyxl/reader` | the whole load path: strings, styles, workbook, worksheets, comments |
| `styles` | `openpyxl/styles` | fonts, fills, borders, alignment, number formats, protection, the indexed palette |
| `units` | `openpyxl/units` | the unit constants Excel uses (`points_to_pixels`, `emu_to_cm`, …) |
| `workbook` | `openpyxl/workbook` | `Workbook`, document properties, security, sheet management |
| `worksheet` | `openpyxl/worksheet` | `Worksheet`, dimensions, views, panes, protection, header/footer, iteration |
| `writer` | `openpyxl/writer` | the whole save path, including the streaming XML writer |
| `xml` | `openpyxl/xml` | an ElementTree-shaped element tree, a streaming writer, and the namespace constants |

## Feature parity

[PARITY.md](PARITY.md) is the module-by-module account of what is implemented, what is
pending, and where the Rust version differs on purpose. It is generated from an audit of
the two source trees, so it can be re-run rather than believed:

```console
$ python tools/parity.py path/to/openpyxl/openpyxl
```

Of openpyxl's 272 public names, 195 have a direct counterpart. The 77 that do not are
accounted for: 13 are genuinely pending, 37 are a name or a container that had to change,
and 27 are `lxml` and Python infrastructure with no Rust equivalent. The pending list is
the streaming writer, `Worksheet.range()` with offsets, the `use_iterators` loader, reading
the stored `<dimension>`, zip repair, and reading charts back.

## Dates and the two calendars

Excel's 1900 date system believes 1900 was a leap year, so it has a day — the phantom
1900-02-29 — that never existed. lexcel reproduces this exactly, including the
consequences:

- Serial 1 is 1900-01-01 and serial 60 is the phantom day.
- `to_excel` skips the phantom day, so 1900-02-28 is serial 59 and 1900-03-01 is serial 61.
- `from_excel` does not skip it, so serial 60 reads back as 1900-02-28 and serial 1 reads
  as 1899-12-30. The two functions disagree by one below the phantom day and agree above
  it. openpyxl behaves the same way; the tests pin both halves of the asymmetry.
- A 1904 workbook (the classic Mac calendar) has no phantom day and round-trips exactly.

A cell's serial only becomes a date when its number format says so, so
`ws.set_number_format("A1", "yyyy-mm-dd")` is what turns `40196` into a date.

## Differences from openpyxl

These are the places where a literal port is impossible or would be a bad idea. Each is
also documented at the call site.

**Equality and hashing.** Python compares styles and number formats by the fields that
matter — a number format's code, not its index — and lets everything else fall back to
identity. Rust's `HashMap` needs `Hash` and `Eq`, so lexcel derives them but implements
them by the same comparison Python uses: `NumberFormat` hashes its code alone, and
`Style` hashes its visual fields. The consequence is that two styles that look identical
dedupe to one entry in the style table, which is what Excel does too.

**Tri-state booleans.** OOXML attributes like `<protection locked>` are genuinely
three-valued: present-and-true, present-and-false, and absent. Python models this with
`None`; Rust's `bool` cannot. lexcel uses a `ProtectionFlag` enum with `Inherit`, `Locked`
and `Unlocked` for those attributes, and a plain `bool` where the specification really does
mean two states.

**`CategoryAxis` and `ValueAxis`.** In Python these are classes that differ only in their
class-level attribute defaults. In Rust they are constructors returning the single `Axis`
type, so `CategoryAxis::new()` and `ValueAxis::new()` produce an `Axis` configured the way
the corresponding Python class would configure itself.

**Images.** openpyxl opens an image with PIL to read its size, which makes PIL a hard
dependency for anyone embedding a picture. lexcel parses the PNG header directly and
supports PNG only. The bytes are stored verbatim, so the image itself is untouched.

**Julian days.** openpyxl depends on `jdcal`, which is unmaintained and not on PyPI any
more. lexcel reimplements `gcal2jd` and `jd2gcal` and pins them against `jdcal`'s own
output in the tests.

**The colour palette.** openpyxl's `COLOR_INDEX` has 56 entries, not the 64 Excel
documents. lexcel ships the same 56 and rejects an index past the end, because matching
the upstream table is more useful than matching the documentation.

**Charts and images are written but not read back.** openpyxl 1.9 writes chart and drawing
parts but its reader does not parse them, so a reloaded workbook reports no charts. lexcel
matches that, rather than being half-compatible in a different way.

**Self-closing tags.** Python's `XMLGenerator` never writes `<x/>`; it writes `<x></x>`.
lexcel's streaming writer does the same, so the bytes match.

## The MCP server

`lexcel-mcp` speaks JSON-RPC 2.0 over stdio, one message per line, which is what a locally
launched MCP server uses.

```console
$ cargo run -p lexcel-mcp -- --root ./spreadsheets
```

`--root` bounds every path a tool may touch. Without it the server uses `$LEXCEL_ROOT`, and
without that the working directory. This is a convenience that keeps an agent inside a
directory you chose, not a security sandbox: a caller that can launch this process can
already do whatever it likes. The root is created if it does not exist.

The client configuration looks like this:

```json
{
  "mcpServers": {
    "lexcel": {
      "command": "lexcel-mcp",
      "args": ["--root", "/path/to/spreadsheets"]
    }
  }
}
```

### The tools

Thirty-three tools, grouped by what they are for. Every argument is documented in the
schema the server advertises at `tools/list`.

**Reading** — `list_sheets`, `describe_sheet`, `read_cells`, `read_formulas`,
`search_values`, `summarize_range`, `list_comments`, `list_named_ranges`, `export_csv`.

**Structure** — `create_workbook`, `add_sheet`, `remove_sheet`, `rename_sheet`,
`merge_cells`, `unmerge_cells`, `freeze_panes`, `set_auto_filter`, `add_named_range`.

**Values** — `set_cell`, `write_cells`, `append_row`, `clear_cells`.

**Layout and appearance** — `set_column_width`, `set_row_height`, `set_header_footer`,
`add_hyperlink`, `style_cells`, `set_number_format`, `add_data_validation`,
`add_conditional_format`.

**Drawing** — `add_chart`, `add_image`, `add_comment`.

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
$ cargo test --workspace                   # 480 tests
$ cargo doc --workspace                    # API documentation
$ python tools/check_readme.py             # the README examples are the doctests
$ python tools/parity.py ../openpyxl        # audit the public surface against openpyxl
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

## License

MIT, the same as openpyxl. See [LICENSE](LICENSE).
