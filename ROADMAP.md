# Roadmap

This file records proposals that were considered, what each would actually cost, and what
was decided. It is a record of judgement calls, not a wish list — several items here were
argued against, and the arguments are kept so they do not have to be had again.

Status key: **shipped**, **next**, **later**, **declined**, **partial**.

---

## 1. Formula evaluation (`workbook.recalculate()`)

**Status: done, as a deliberate subset.** Implemented in 0.1.6. This entry is kept because
the reasoning below still governs what is *missing*, which is most of the interesting
functions.

What shipped is the part that can be made correct: the operators, `SUM`, `AVERAGE`, `MIN`,
`MAX`, `COUNT`, `COUNTA`, `PRODUCT`, `ROUND`/`ROUNDUP`/`ROUNDDOWN`, `ABS`, `INT`, `SIGN`,
`SQRT`, `MOD`, `POWER`, `IF`, `IFERROR`, `AND`/`OR`/`NOT`, and the common text functions.
`VLOOKUP`, `XLOOKUP`, `INDEX` and `MATCH` are still absent, for the reasons given below, and a
formula naming one comes back *unresolved* rather than with a plausible number.

### What it would take

An expression parser, a value lattice matching Excel's (numbers, text, booleans, errors,
blanks, and the coercion rules between them), a function library, and a recalculation
order. The recalculation order is now available: [`Worksheet::trace_precedents`][dep] gives
the topological order for a cell, and `Worksheet::circular_references` finds the nodes that
have none. That is roughly a fifth of the work, and it was built for a different reason.

### Why it is not next

A wrong cached value is worse than a missing one. A missing `<v>` is visible — a reader
knows to recalculate. A wrong one is silently believed, and the error surfaces days later
in a report nobody can trace back to this library. So a half-finished engine is not
"50% of the feature"; it is a new way to produce wrong spreadsheets.

The other cost is surface area. "Top 50–100 functions" is 50–100 implementations each
needing its own edge cases: `VLOOKUP` with an approximate match against unsorted data,
`SUMIF` with wildcards, `INDEX` returning a reference rather than a value, date arithmetic
across the 1900 leap-year bug, `IFERROR` swallowing a type error. Each is a day, and each
is a place to be subtly wrong.

### What a defensible version looks like

Read-only, opt-in, and never silent. A `Recalculation` that returns the values it computed
and the errors it could not, and writes `<v>` only where it succeeded. A shipped function
list in the docs, so a caller can check `supports("XLOOKUP")` instead of discovering the
gap by getting a wrong number. Deprecations excluded. Start with arithmetic, comparison,
`IF`, `AND`/`OR`/`NOT`, and the aggregates; leave `SUMIFS` out until it is right.

### Why it is still worth doing

It is the only item on this list that openpyxl's user base asks for by name, and the only
one where being wrong is invisible. Both facts point the same direction: it needs
engineering, not a weekend.

---

## 2. Non-destructive editing

**Status: partial.** The most valuable item here, and the one that most changes the
architecture.

### What exists

`keep_vba` already copies the `vbaProject.bin` family of parts through verbatim
(`reader/excel.rs` retains the raw archive, `writer/excel.rs` copies the parts out). That
is the whole-archive case for the one part people most often lose.

### What is missing, and why it is big

Everything else. The current model is parse-into-structs, write-structs-out. Bit-for-bit
preservation requires *not parsing* parts the library does not model — pivot caches, slicer
state, PowerQuery connections, `calcChain`, threaded comments, drawing XML. So this is a
second code path alongside the current one rather than a modification of it, which is good
news: it is additive and can ship behind a feature flag without putting existing behaviour
at risk.

The genuinely hard part is what the proposal asks for beyond passthrough: rewriting only
the modified row inside `sheet1.xml` and passing through everything else. That needs
byte-offset splicing into `<sheetData>`, with correct handling of the `r`, `spans` and
dimension attributes that span rows. It is doable and it is delicate, and a mistake there
corrupts the file in ways that are hard to attribute.

### Recommended shape

`Workbook::open_preserving(path)` returning a document that holds unmodelled parts as
`Vec<u8>` and modelled parts as structs, with `save()` writing each from whichever
representation it has. Sheets become row-level: parse lazily, and on save re-emit only
rows that changed, keeping the original bytes for the rest.

---

## 3. Parallel parsing and writing with rayon

**Status: later.** Real, but much smaller than advertised.

### What is actually parallelisable

Sheets are independent files in the zip, so parsing and writing them across cores is
available today. That is the honest win.

### What is not

Parsing *one* sheet in parallel. `quick-xml` is a sequential pull parser with state that
spans the whole document, and so is the reader built on it. Threading a single
`sheetData` is not a rayon job; it is a different parser. Shared strings are already built
once into a single `Vec<String>` and shared by reference, so the "concurrent hash map"
part of the proposal solves a problem the design already avoids — parallel building would
need a two-pass index-then-fill to be safe.

### On the 10x–50x figure

Not achievable. Amdahl's law bounds a ten-sheet workbook by the shared work —
`workbook.xml`, relationships, `styles.xml`, the string table — which is a fixed cost no
amount of sheet-level fan-out removes. Realistic expectation on a 10-large-sheet workbook
is **2–4x**, and on a single-sheet workbook it is **1x**. Worth having, worth a feature
gate and a benchmark before believing any figure in a changelog.

---

## 4. Polars and Arrow integration

**Status: declined, as specified. A different API is worth considering.**

### Why the stated version does not work

```rust
sheet.to_dataframe("A1:F50000")?;
```

An Excel column is not a typed column. It can hold numbers, text, dates, errors, formulas
and genuine blanks in the same column, and in real spreadsheets it usually does. Arrow has
no cell type, so this signature must either take a schema — silently nulling or coercing
whatever disagrees with it — or infer one from the first row and be wrong about the rest.
The result is a DataFrame that lost exactly the information the spreadsheet was carrying,
which is the failure mode this library exists to avoid.

### Why "read columnar, not per-cell heap objects" is not the win it looks like

That is already true. `iter_rows` borrows cell data; nothing is heap-allocated per cell
except the `CellValue` enum itself, which is a tag plus a borrowed or inline value. The
cost openpyxl pays is a *Python* object per cell, and crossing into Arrow removes the
Python, not an allocation that exists.

### What would be worth building

An explicitly typed read, where the caller says what the column is and ferroxl errors on
anything that does not fit:

```rust
let frame = sheet.read_typed::<f64>("A1:F50000")?;   // fails loudly on a text cell
```

and the same in reverse. If a caller has a schema, this is good. If they do not, they want
`read_cells`.

Also worth noting: `polars` is a large optional dependency on a crate whose entire
justification is parity with a library that has no such dependency.

---

## 5. Workbook diffing

**Status: next.** The best value per unit of work on this list, and it builds directly on
what shipped in 0.1.2.

### Why it is tractable

A diff is a structural comparison over data the library already holds. No new parser, no
new dependencies. The only interesting decision is what counts as a change: a cell whose
style index differs is changed only if the *resolved* style differs, because reordering
`styles.xml` renumbers every index in the workbook without changing a single visible
format.

### Shape

```rust
let diff = Workbook::diff(&before, &after);
diff.sheets_added, diff.sheets_removed,
diff.cells_changed, diff.cells_moved, diff.formulas_changed, diff.styles_changed
```

Each `cells_changed` entry carrying `before`, `after` and `number_format` — because
"the value changed" and "the value is the same but is now shown as a percentage" are
different facts for an auditor.

This is also the natural companion to dependency tracing: given a diff and a graph, you
can say *which formulas the change invalidated*, which is the actual question behind "show
me what this pull request does to the model".

---

## 6. Dependency tracing

**Status: shipped in 0.1.2.**

[`Worksheet::trace_precedents`][dep], `trace_dependents`, `circular_references`,
`dependency_graph` and `parse_references`, exposed to agents as `trace_precedents`,
`trace_dependents` and `check_circular_references`.

Two answers are refused rather than guessed. References through a defined name are
reported by `References::named` instead of resolved, because a name may cover any range in
the workbook. Cross-sheet references keep their sheet title rather than being expanded,
because the range behind one may be far larger than the formula that mentions it.

The one limitation to know about: the graph is built from formula *text*, so a reference
rewritten by a tool that emitted an unresolved name is invisible until that name is
resolved. `References::named` is where that shows up, which is why it exists.

---

## 7. Lint and de-bloat

**Status: later, and split — because two halves of it are not the same kind of tool.**

### Orphan styles — safe, do this

Styles in `styles.xml` that no cell references can be dropped, and reporting them is
unambiguously useful. This is a reachability walk from every cell's style index.

### Prune phantom rows and columns — do not do this automatically

The proposal describes this as cleanup. It is not. "Someone formatted row 1,000,000 and
Excel wrote a million empty cells" is also "someone formatted row 1,000,000" — the
formatting is the intent, and deleting it destroys a decision someone made. A de-bloater
that silently strips it is worse than the bloat.

So: report it, and let the caller prune. `worksheet.phantom_extent()` returning the styled
region beyond the last cell with a value, plus the byte cost, is honest and actionable. An
explicit `prune_styled_empty_rows()` is fine as a method — the user asked.

### Broken references — tractable now

`#REF!`, `#NAME?` and friends in cached values are a scan. Detecting references that
*point* at deleted sheets or out-of-bounds ranges reuses the parser that shipped with
dependency tracing.

---

## 8. MCP superpowers

**Status: partial — dependency tracing shipped; the rest is next.**

### Done

The three tracing tools. Summaries cap at a dozen cells and then give a count, because a
list of three thousand dependents cannot change a model's next decision but will consume
the context window that would let it make one.

### Next, and worth doing

- **`compress_to_markdown(range)`** — trivial, and it is the highest-leverage tool in the
  server. A hundred rows of cells is the difference between a model reading a table and a
  model reading a JSON object graph.
- **`get_column_by_header(name)`** — also small.
- **`sql_query`** — needs an expression parser over the grid with `SELECT`, `WHERE`,
  `GROUP BY`, `HAVING`, `ORDER BY` and a handful of aggregates. This is most of the
  machinery in item 1 with the function library left out, so item 1 makes it cheaper and
  the two should be sequenced together.

### Fuzzy header matching — argue against

Fuzzy matching on spreadsheet headers fails in a way that is hard to debug: two columns
both called "Total", and a typo-tolerant match silently picks the wrong one. Exact match
first; on failure, return *the candidates* and let the model choose:

```
no column named "revenew"; did you mean Revenue, or Revenue (Q3)?
```

That is one extra round trip instead of a wrong column, and it is the difference between
a tool an agent trusts and one it has to double-check every time.

---

## Ordering

1. **0.1.2** — dependency tracing. Shipped.
2. **0.1.3** — `compress_to_markdown`, `get_column_by_header`, workbook diffing. All small,
   all independent, all useful.
3. **0.2.0** — the expression parser, shared by `sql_query` and `recalculate`. It is the
   only item that has to be right rather than merely present, so it gets a version bump
   and its own review.
4. **0.3.0** — non-destructive editing behind a feature flag. Additive, so it can ship
   without putting 0.2 behaviour at risk.

Items 3, 4 and 5 (formula evaluation, non-destructive editing, rayon) each break the
build-time dependency set or the round-trip model. That is what the major versions are for.

[dep]: https://docs.rs/ferroxl/latest/ferroxl/worksheet/struct.Worksheet.html#method.trace_precedents