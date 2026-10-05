# ferroxl

A Rust library for reading and writing Excel 2007 `.xlsx` and `.xlsm` files.

ferroxl is a feature-parity port of [openpyxl](https://github.com/theorchard/openpyxl) 3.1.5.
It follows the Python package's module layout, class names and semantics, so a change can
be traced back to the Python it mirrors. Where Python's behaviour cannot be reproduced in
Rust — hash-based equality, tri-state booleans, PIL-backed images — the deviation is
documented at the call site and the closest faithful behaviour is implemented instead.

This is the library crate. The [MCP server](https://crates.io/crates/ferroxl-mcp) that
exposes it to AI agents is published separately as `ferroxl-mcp`.

## Installation

```toml
[dependencies]
ferroxl = "0.1.8"
```

Rust 1.88 or newer, which is what `rust-version` in the manifest declares and what CI
checks against.

## Reading

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
# Ok::<(), ferroxl::Error>(())
```

`load_workbook_from_bytes` takes an already-read buffer instead of a path, which is what a
web service wants. `LoadOptions` carries `guessing_types()`, `values_only()` and
`keeping_vba()` — inferring a cell's type from its text, returning Excel's last cached value
instead of the formula, and keeping the original package bytes so a VBA project survives a
save.

To walk the used range row by row rather than asking for a rectangle, `iter_rows()` and
`iter_cols()` come back as whole rows and columns, so the outer and inner iterators always
line up:

```rust
use ferroxl::{load_workbook, LoadOptions};

let workbook = load_workbook("report.xlsx", LoadOptions::default())?;
let sheet = workbook.active_sheet()?;

for row in sheet.iter_rows() {
    for cell in row {
        if let Some(cell) = cell {
            println!("{} = {:?}", cell.coordinate(), cell.internal_value());
        }
    }
}
# Ok::<(), ferroxl::Error>(())
```

## Writing

```rust
use ferroxl::{CellValue, Workbook};

let mut workbook = Workbook::new();
let summary = workbook.create_sheet(Some("Summary"))?;
let sheet = &mut workbook.worksheets[summary];

sheet.set("A1", CellValue::text("Q3 revenue"))?;
sheet.set("B1", 150_000.0)?;
sheet.set("A2", CellValue::text("Total"))?;
sheet.set("B2", CellValue::formula("SUM(B1:B1)"))?;

workbook.save("report.xlsx")?;
# Ok::<(), ferroxl::Error>(())
```

`Workbook::new()` already makes a sheet called `Sheet1`, and `create_sheet` appends without
making the new sheet active — openpyxl behaves the same way — so take the index
`create_sheet` returns rather than reaching for `active_sheet_mut`, which would still be
`Sheet1`. `save` consumes the workbook, which is what makes edit-then-save a single linear
sequence; `to_bytes` returns the package instead, and `save_workbook_to` streams it to any
`Write`.

Formula cells are written without a cached result by default, which is what openpyxl does.
Excel calculates them on open; a reader that does not calculate will report the formula
rather than a value. Use `recalculate()` or `trace_precedents` and `trace_dependents` to
reason about the graph instead of assuming the numbers are there.

## Formula evaluation

`recalculate()` evaluates every formula in the workbook and records the values, so a reader
that is not Excel sees numbers rather than blanks:

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

It covers the operators, the aggregates, `IF`/`IFERROR`, `AND`/`OR`/`NOT` and the common text
functions — 31 in all. A formula it cannot evaluate gets **no** cached value and is listed
in `Recalculation::unresolved` with the reason, and `formula::supports` answers whether a
name is handled without calling it and getting a wrong answer. `calcPr/@fullCalcOnLoad` is
set, so Excel recomputes on open and a value the engine got wrong cannot survive a human
opening the file.

Formulas are evaluated in one pass over reading order, so a formula reading another
formula's cell sees it as blank. `trace_precedents` gives the order to do it properly.

## Dependency tracing

The one thing openpyxl cannot answer is what a cell is connected to. ferroxl can:

```rust
# use ferroxl::{load_workbook, LoadOptions};
let workbook = load_workbook("model.xlsx", LoadOptions::default())?;
let sheet = workbook.active_sheet()?;

// What feeds D10, in the order a recalculation would visit them.
sheet.trace_precedents("D10")?;

// What breaks if B2 changes.
sheet.trace_dependents("B2")?;

// Excel refuses to calculate a workbook with one of these.
for cycle in sheet.circular_references()? {
    println!("{}", cycle.join(" -> "));
}
# Ok::<(), ferroxl::Error>(())
```

References through a defined name are reported rather than resolved, and a cross-sheet
reference keeps its sheet title rather than being expanded. A guess at either would produce
a graph that looks authoritative and is not.

## Tables, named styles and charts

- **Tables (ListObjects)** — `Worksheet::add_table` writes `xl/tables/tableN.xml` and the
  sheet tail that points at it, reading the column names out of the header cells so a
  structured reference like `=SUM(Table1[Sales])` resolves.
- **Named styles** — all 49 of Excel's built-ins as `BUILTIN_DETAILS`, generated from
  openpyxl's `styles/builtins.py`, plus `Workbook::named_styles` for a workbook's own.
  `apply_named_style` applies and registers, so the name reaches Excel's style gallery and
  reads back as `cell.style`.
- **Charts** — all sixteen types, with series, axes, titles, error bars and 3-D views.
  Written, not read back: openpyxl 3.1.5's own reader does not parse chart parts either.

## Round trips are not destructive

`Workbook::preserved` holds every part, content type, relationship and `<workbook>` /
`<worksheet>` child that the writer does not produce, and the writer writes them back — so a
pivot table, slicer, query table, threaded comment, ActiveX control or `customXml` part
survives being opened and saved. Copying the bytes is the easy half: a part nothing points at
is inert, so the relationship is preserved too, its id remapped where the writer has already
used one, and the `r:id` in the referencing element rewritten to match.

Carrying a part through is not editing it. `PARITY.md` records what still does not survive:
unknown attributes on `<worksheet>`, content inside `<sheetData>`, and byte-identical zip
entries.

## Streaming

`save_dump` writes rows as they are produced rather than building the sheet part as a
`String` first, so peak memory is set by the largest row rather than the largest sheet:

```rust
# use ferroxl::Workbook;
let workbook = Workbook::new();
ferroxl::save_dump(workbook, &mut std::io::stdout())?;
# Ok::<(), ferroxl::Error>(())
```

`crates/ferroxl/examples/build_and_read.rs` is a full round trip — create a sheet, set cells,
style a header, merge a range, freeze panes, save, load — in about a hundred lines:

```console
$ cargo run --example build_and_read
```

## Feature parity

`PARITY.md` in the repository tracks what is ported, what differs and why, and what is not
ported at all. `ROADMAP.md` records proposals that were considered and declined, with the
reasoning, so they do not have to be argued again.

Of openpyxl 3.1.5's 994 top-level public names, 678 have no counterpart and 32 of its 183
modules match in full. Not implemented: a read-only (`read_only`) loader, reading the stored
`<dimension>` element, reading charts back, external-link parts, rich text, pivot tables,
chart-only sheets, and the `descriptors` metaprogramming layer. The audit matches names
rather than behaviour, so those figures are an upper bound on what works rather than a
measurement of it — `PARITY.md` says so where it matters.

## License

MIT, the same as openpyxl.