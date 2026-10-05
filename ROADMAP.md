# Roadmap

This file records proposals that were considered, what each would actually cost, and what
was decided. It is a record of judgement calls, not a wish list — several items here were
argued against, and the arguments are kept so they do not have to be had again.

Status key: **shipped**, **next**, **later**, **declined**, **partial**.

---

## 1. Formula evaluation (`workbook.recalculate()`)

**Status: shipped in 0.1.7, as a deliberate subset.** This entry is kept because the
reasoning below still governs what is *missing*, which is most of the interesting functions.

What shipped is the part that can be made correct: the operators, `SUM`, `AVERAGE`, `MIN`,
`MAX`, `COUNT`, `COUNTA`, `PRODUCT`, `ROUND`/`ROUNDUP`/`ROUNDDOWN`, `ABS`, `INT`, `SIGN`,
`SQRT`, `MOD`, `POWER`, `IF`, `IFERROR`, `AND`/`OR`/`NOT`, and the common text functions.
`VLOOKUP`, `XLOOKUP`, `INDEX` and `MATCH` are still absent, for the reasons given below, and a
formula naming one comes back *unresolved* rather than with a plausible number.

### What it would take

An expression parser, a value lattice matching Excel's (numbers, text, booleans, errors,
blanks, and the coercion rules between them), a function library, and a recalculation
order. The recalculation order was available before the engine: `Worksheet::trace_precedents`
gives the topological order for a cell, and `Worksheet::circular_references` finds the nodes
that have none. That was roughly a fifth of the work, and it was built for a different reason.

The shipped engine evaluates in one pass over reading order, so a formula reading another
formula's cell sees it as blank. `trace_precedents` is the order to do it properly, which is
the next step rather than a redesign.

### Why it was not shipped whole

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

Read-only, opt-in, and never silent — which is what shipped. A `Recalculation` that returns
the values it computed and the errors it could not, and writes `<v>` only where it succeeded.
A shipped function list in the docs, so a caller can check `supports("XLOOKUP")` instead of
discovering the gap by getting a wrong number; there is a test asserting that list matches the
dispatcher, so a name cannot be advertised without being implemented. Deprecations excluded.
Arithmetic, comparison, `IF`, `AND`/`OR`/`NOT` and the aggregates first; `SUMIFS` stays out
until it is right.

### What is still worth doing

It is the only item on this list that openpyxl's user base asks for by name, and the only
one where being wrong is invisible. Both facts still point the same direction: the remaining
functions need engineering, not a weekend.

---

## 2. Non-destructive editing

**Status: partial — the unmodelled-part half shipped in 0.1.7.** The most valuable item
here, and the one that most changes the architecture.

### What exists

Two things, and they are different in kind.

`keep_vba` copies the `vbaProject.bin` family of parts through verbatim (`reader/excel.rs`
retains the raw archive, `writer/excel.rs` copies the parts out). That is the whole-archive
case for the one part people most often lose.

Since 0.1.7, `Workbook::preserved` holds *every* part, content type, relationship and
`<workbook>`/`<worksheet>` child the writer does not produce, and the writer writes them
back — pivot caches, slicers, query tables, connections, threaded comments, ActiveX
controls, `customXml`. Copying the bytes was the easy half; reachability is the part that is
easy to get wrong, because a preserved part nothing points at is inert. The relationship
travels too, its id remapped where the writer has already used one, and the `r:id` in the
referencing element rewritten to match.

Known limits, recorded rather than left to be found: unknown *attributes* on `<worksheet>`
and `<sheetPr>` (children carry over, attributes do not), content inside `<sheetData>`,
which is rebuilt from the cell model, and byte-identical zip entries — the bytes are
re-compressed, so an entry is content-identical rather than byte-identical.

### What is still missing, and why it is big

**Editing.** A preserved pivot table or slicer can be carried through but not changed, and
that is the gap that matters. It needs the same second code path this entry originally
proposed: modelling the OOXML constructs ferroxl currently passes through.

The genuinely hard part is what the proposal asks for beyond passthrough: rewriting only
the modified row inside `sheet1.xml` and passing through everything else. That needs
byte-offset splicing into `<sheetData>`, with correct handling of the `r`, `spans` and
dimension attributes that span rows. It is doable and it is delicate, and a mistake there
corrupts the file in ways that are hard to attribute.

Bit-for-bit preservation would also require *not parsing* parts the library does not model,
which the current parse-into-structs model cannot express. That part is still open.

### Recommended shape

`Workbook::open_preserving(path)` returning a document that holds unmodelled parts as
`Vec<u8>` and modelled parts as structs, with `save()` writing each from whichever
representation it has — which is what `preserved` now is, minus the open/save split. Sheets
become row-level: parse lazily, and on save re-emit only rows that changed, keeping the
original bytes for the rest.

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
2. **0.1.3** — the `formula::Translator` half of formula evaluation. Shipped.
3. **0.1.7** — the evaluator and pass-through preservation. Shipped, both as deliberate
   subsets: 31 functions and the unmodelled parts.
4. **Next** — `compress_to_markdown`, `get_column_by_header`, workbook diffing. All small,
   all independent, all useful, and none of them has to be right rather than merely present.
5. **0.2.0** — multi-pass recalculation in dependency order, and the lookups
   (`VLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`) that most often want a cached value. These have
   to be right rather than merely present, so they get a version bump and their own review.
6. **0.3.0** — editing the constructs that are currently passed through, and rayon.

What is left on 4 and 5 does not change the dependency set or the round-trip model, which is
why they can ship as patch releases. Non-destructive *editing* (item 2) and rayon (item 3)
both do, and that is what the major versions are for.

[dep]: https://docs.rs/ferroxl/latest/ferroxl/worksheet/struct.Worksheet.html#method.trace_precedents