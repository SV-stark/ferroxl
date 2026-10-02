# ferroxl

A Rust library for reading and writing Excel 2007 `.xlsx` and `.xlsm` files.

ferroxl is a feature-parity port of [openpyxl](https://github.com/theorchard/openpyxl) 1.9.
It follows the Python package's module layout, class names and semantics, so a change can
be traced back to the Python it mirrors. Where Python's behaviour cannot be reproduced in
Rust — hash-based equality, tri-state booleans, PIL-backed images — the deviation is
documented at the call site and the closest faithful behaviour is implemented instead.

This is the library crate. The [MCP server](https://crates.io/crates/ferroxl-mcp) that
exposes it to AI agents is published separately as `ferroxl-mcp`.

## Installation

```toml
[dependencies]
ferroxl = "0.1.2"
```

## Reading

```rust
use ferroxl::{Workbook, CellValue};

let workbook = Workbook::open("report.xlsx")?;
let sheet = &workbook.worksheets[0];

for row in sheet.rows(1, 3) {
    for cell in row {
        println!("{} = {:?}", cell.coordinate(), cell.value());
    }
}
```

## Writing

```rust
use ferroxl::{CellValue, Workbook};

let mut workbook = Workbook::new();
let sheet = workbook.active_sheet_mut()?;
sheet.set("A1", CellValue::text("Q3 revenue"))?;
sheet.set("B1", 150_000.0)?;
sheet.set("A2", CellValue::text("Total"))?;
sheet.set("B2", CellValue::formula("SUM(B1:B1)"))?;
workbook.save("report.xlsx")?;
```

Formula cells are written without a cached result, which is what openpyxl does. Excel
calculates them on open; a reader that does not calculate will report the formula rather
than a value, so use `trace_precedents` and `trace_dependents` to reason about the graph
instead of assuming the numbers are there.

## Dependency tracing

The one thing openpyxl cannot answer is what a cell is connected to. ferroxl can:

```rust
// What feeds D10, in the order a recalculation would visit them.
sheet.trace_precedents("D10")?;

// What breaks if B2 changes.
sheet.trace_dependents("B2")?;

// Excel refuses to calculate a workbook with one of these.
for cycle in sheet.circular_references()? {
    println!("{}", cycle.join(" -> "));
}
```

References through a defined name are reported rather than resolved, and a cross-sheet
reference keeps its sheet title rather than being expanded. A guess at either would produce
a graph that looks authoritative and is not.

## Streaming

`save_dump` writes rows as they are produced rather than building the sheet part as a
`String` first, so peak memory is set by the largest row rather than the largest sheet:

```rust
ferroxl::save_dump(workbook, &mut std::io::stdout())?;
```

## Feature parity

`PARITY.md` in the repository tracks what is ported, what differs and why, and what is not
ported at all. `ROADMAP.md` records proposals that were considered and rejected, with the
reasoning, so they do not have to be argued again.

Not implemented: formula evaluation (`recalculate()`), reading charts and images back, and
openpyxl's streaming zip repair.

## Licence

MIT.