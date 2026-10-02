# ferroxl-mcp

A [Model Context Protocol](https://modelcontextprotocol.io) server that lets an AI agent
read and edit Excel `.xlsx` and `.xlsm` files.

It is backed by [ferroxl](https://crates.io/crates/ferroxl), a feature-parity port of
[openpyxl](https://github.com/theorchard/openpyxl) 1.9.

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

`FERROXL_ROOT` sets the directory it will read and write files in. Without it, the current
directory is used. Paths in tool arguments are resolved inside that root, and a path that
escapes it is refused rather than followed.

## Configuring an agent

```json
{
  "mcpServers": {
    "ferroxl": {
      "command": "ferroxl-mcp",
      "env": { "FERROXL_ROOT": "/home/agent/workbooks" }
    }
  }
}
```

## Tools

36 tools. To see the arguments each one takes, list them from your agent, or read
`crates/ferroxl-mcp/src/tools.rs`.

**Inspection** — `list_sheets`, `describe_sheet`, `read_cells`, `read_formulas`,
`search_values`, `summarize_range`, `list_comments`, `list_named_ranges`, `export_csv`.

**Dependency tracing** — `trace_precedents`, `trace_dependents`,
`check_circular_references`.

**Structure** — `create_workbook`, `add_sheet`, `remove_sheet`, `rename_sheet`,
`merge_cells`, `unmerge_cells`, `freeze_panes`, `set_auto_filter`, `add_named_range`.

**Cell contents** — `set_cell`, `write_cells`, `append_row`, `clear_cells`. Formulas go
through the same tools as values; there is no separate formula tool.

**Layout and formatting** — `set_column_width`, `set_row_height`, `set_header_footer`,
`add_hyperlink`, `style_cells`, `set_number_format`, `add_data_validation`,
`add_conditional_format`, `add_comment`.

**Media** — `add_chart`, `add_image`.

Every tool returns a one-line summary as well as structured JSON. The summary is what a
model reads first, so it is written to be the useful part: the count, the names, the
coordinates. Where a list could be arbitrarily long — a cell with thousands of dependents —
the summary names the first dozen and gives the count, because the count is what changes
the next decision and the list would only consume the context window needed to make one.

## What it will not do

The server will not evaluate formulas. A cell written with a formula has no cached result,
exactly as openpyxl writes it, so `read_cells` returns the formula rather than a number.
`trace_precedents` and `trace_dependents` answer questions about the graph that a
calculated value would otherwise have been needed to answer.

## Licence

MIT.