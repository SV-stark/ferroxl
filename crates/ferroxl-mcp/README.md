# ferroxl-mcp

A [Model Context Protocol](https://modelcontextprotocol.io) server that lets an AI agent
read and edit Excel `.xlsx` and `.xlsm` files.

It is backed by [ferroxl](https://crates.io/crates/ferroxl), a feature-parity port of
[openpyxl](https://github.com/theorchard/openpyxl) 3.1.5.

## Installation

```sh
cargo install ferroxl-mcp
```

Prebuilt archives for Linux, macOS and Windows are attached to each
[GitHub release](https://github.com/SV-stark/ferroxl/releases).

## Running

The server speaks MCP over stdio, so an agent launches it directly:

```sh
ferroxl-mcp
```

`--root` bounds every path the tools may touch. Without it, `$FERROXL_ROOT` is used, and
without that the working directory. Paths in tool arguments are resolved inside that root,
and a path that escapes it is refused rather than followed. The root is created if it does
not exist.

This is a convenience that keeps an agent inside a directory you chose, not a security
sandbox: a caller that can launch this process can already do whatever it likes.

```console
$ ferroxl-mcp --help
ferroxl-mcp 0.1.10 - a Model Context Protocol server for Excel workbooks

USAGE:
    ferroxl-mcp [--root <directory>]

OPTIONS:
    --root <directory>  Bound every path the tools may touch. Defaults to
                        $FERROXL_ROOT, then the working directory.
    -h, --help          Print this help and exit.
    -V, --version       Print the version and exit.

The server speaks JSON-RPC 2.0 over stdio, one message per line.
```

## Configuring an agent

```json
{
  "mcpServers": {
    "ferroxl": {
      "command": "ferroxl-mcp",
      "args": ["--root", "/home/agent/workbooks"]
    }
  }
}
```

## Tools

41 tools, in six groups. To see the arguments each one takes, list them from your agent, or
read `crates/ferroxl-mcp/src/tools.rs`.

**Inspection** — `list_sheets`, `describe_sheet`, `read_cells`, `read_formulas`,
`trace_precedents`, `trace_dependents`, `check_circular_references`, `add_data_bar`,
`add_icon_set`, `add_table`, `describe_table`, `set_gradient_fill`, `search_values`,
`summarize_range`, `list_comments`, `list_named_ranges`, `export_csv`.

**Structure** — `create_workbook`, `add_sheet`, `remove_sheet`, `rename_sheet`,
`merge_cells`, `unmerge_cells`, `freeze_panes`, `set_auto_filter`, `add_named_range`.

**Values** — `set_cell`, `write_cells`, `append_row`, `clear_cells`. Formulas go through
the same tools as values; there is no separate formula tool.

**Layout** — `set_column_width`, `set_row_height`, `set_header_footer`, `add_hyperlink`.

**Formatting** — `style_cells`, `set_number_format`, `add_data_validation`,
`add_conditional_format`.

**Media and annotations** — `add_chart`, `add_image`, `add_comment`.

Every tool returns a one-line summary as well as structured JSON. The summary is what a
model reads first, so it is written to be the useful part: the count, the names, the
coordinates. Where a list could be arbitrarily long — a cell with thousands of dependents —
the summary names the first dozen and gives the count, because the count is what changes
the next decision and the list would only consume the context window needed to make one.

The three tracing tools answer what openpyxl cannot at all: which cells feed a cell, which
formulas would go stale if one changed, and every circular reference as a closed path.
Excel refuses to calculate a workbook with a cycle, so `check_circular_references` is worth
running before trusting a file you did not create.

## Values

A tool that writes a value takes JSON, with a deliberately narrow mapping so a model does
not have to learn a wrapper object:

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

## What it will not do

The server does not evaluate formulas. A cell written with a formula has no cached result,
exactly as openpyxl writes it, so `read_cells` returns the formula rather than a number.
`trace_precedents` and `trace_dependents` answer the questions about the graph that a
calculated value would otherwise have been needed for.

The underlying library can evaluate formulas — `Workbook::recalculate()` — but the server
does not expose it, because a model asking for a number that came from a partial evaluator
is worse off than one reading the formula.

## Errors

A tool that ran and failed returns `isError: true` with the message in `content`, so the
model can read what went wrong and correct its call. Only an unknown method or tool is a
JSON-RPC error (`-32601`). A misspelled argument is rejected rather than ignored, because a
tool that silently drops an argument is worse than one that refuses.

## License

MIT.