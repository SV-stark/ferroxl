"""Call every tool with only the arguments its schema declares required.

This is the model-facing version of the schema check. A model reads `tools/list`, sends the
arguments marked required, and expects the call to work. If the handler then complains about
an argument nothing in the schema marked required, the schema is lying to the only reader it
has, and the failure is not something the model could have anticipated.

Each tool gets its own workbook so that the order of this loop cannot matter -- otherwise a
tool that removes a sheet makes the next tool's call fail for an unrelated reason.

    python tools/mcp_required.py
"""

import base64
import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from mcp_client import Server  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT / "target" / "mcp_required_probe"

# A 1x1 PNG, so add_image has a real file to read without an asset on disk.
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmM"
    "IQAAAABJRU5ErkJggg=="
)

# Per-tool probe values, where an argument means something different to each tool. A chart
# wants series as objects, a validation wants a validation type, and so on.
VALUES = {
    "add_chart": {
        "series": [{"name": "Qty", "values": "B2:B3"}],
        "categories": "A2:A3",
    },
    "add_data_validation": {"type": "list"},
    "add_conditional_format": {"kind": "cellIs", "operator": "greaterThan"},
    "write_cells": {"start": "A1", "rows": [["X", 1]]},
    "append_row": {"values": ["Screw", 100]},
    "merge_cells": {"range": "A5:B5"},
    "set_cell": {"value": 42},
    "add_table": {"name": "Table1", "range": "A1:B3"},
    "describe_table": {"name": "Table1"},
    "set_gradient_fill": {"start_color": "FF1F3864", "end_color": "FF4F81BD"},
    "add_named_range": {"name": "Prices", "range": "B2:B3"},
    "add_sheet": {"title": "Extra"},
    "create_workbook": {"sheets": ["Data"], "overwrite": True},
    "add_image": {"image_path": "pixel.png", "anchor": "D2"},
}

SHARED = {
    "sheet": "Data",
    "cell": "B2",
    "range": "B2:B3",
    "columns": "B",
    "rows": "2",
    "width": 20.0,
    "height": 20.0,
    "title": "Sheet Title",
    "name": "Name",
    "target": "https://example.com",
    "format": "0.00",
    "kind": "cellIs",
    "type": "bar",
    "anchor": "E2",
    "query": "Bolt",
    "value": 42,
    "text": "A comment",
    "formula1": '"Yes,No"',
    "formula": "5",
    "style": {"bold": True},
    "start_color": "FF1F3864",
    "end_color": "FF4F81BD",
    "second_formula": "10",
    "left_header": "Left",
    "type_list": "list",
}


def probe_value(tool: str, argument: str):
    if tool in VALUES and argument in VALUES[tool]:
        return VALUES[tool][argument]
    return SHARED.get(argument)


# A tool whose only-required-arguments call is *expected* to fail, with the reason. These
# are guards and readers rather than schema defects: each refuses a call that is
# meaningless in the state the probe leaves the workbook in, and each says why in a message
# a model can act on. An entry here is a claim about intent -- if the message changes, the
# entry has to change with it.
EXPECTED = {
    "add_conditional_format": (
        "a cellIs rule needs an operator",
        "the operator is required by `kind`, which JSON Schema cannot express; the message "
        "names the missing argument",
    ),
    "set_header_footer": (
        "pass at least one header or footer section",
        "at least one of six optional sections is needed, which `required` cannot express; "
        "the message says so",
    ),
    "unmerge_cells": (
        "not known as merged",
        "unmerging a range that is not merged is an error, mirroring openpyxl",
    ),
    "remove_sheet": (
        "a workbook must keep at least one sheet",
        "refusing to empty a workbook is a guard, and openpyxl refuses it too",
    ),
    "describe_table": (
        "is not a table on this sheet",
        "a reader cannot conjure the table it reads; the probe workbook has none",
    ),
}


def seed(server, workbook: str) -> None:
    """A `Data` sheet with a header row and two data rows."""
    server.ok("create_workbook", path=workbook, sheets=["Data"], overwrite=True)
    server.ok(
        "write_cells",
        path=workbook,
        sheet="Data",
        start="A1",
        rows=[["Item", "Qty"], ["Bolt", 10], ["Nut", 25]],
    )


def message(result) -> str:
    if "rpcError" in result:
        error = result["rpcError"]
        return f"rpcError {error.get('code')}: {error.get('message')}"
    if result.get("isError"):
        return " ".join(part.get("text", "") for part in result.get("content", []))
    return ""


def main() -> int:
    shutil.rmtree(WORK, ignore_errors=True)
    WORK.mkdir(parents=True)
    (WORK / "pixel.png").write_bytes(PNG)
    problems = []

    with Server(WORK) as server:
        server.initialize()
        tools = {t["name"]: t for t in server.list_tools()}

        for name, tool in sorted(tools.items()):
            schema = tool.get("inputSchema") or {}
            declared = list(schema.get("required") or [])
            properties = set((schema.get("properties") or {}).keys())

            workbook = f"{name}.xlsx"
            # Seed the very file the call targets, so a reader finds the sheet it expects.
            # `create_workbook` is the exception: it refuses to clobber an existing file, so
            # it gets a throwaway seed and its own fresh target.
            seed(server, f"{name}.seed.xlsx" if name == "create_workbook" else workbook)
            arguments = {"path": workbook}
            for argument in declared:
                if argument == "path":
                    continue
                arguments[argument] = probe_value(name, argument)

            result = server.call(name, **arguments)
            complaint = message(result)
            if complaint:
                expected = EXPECTED.get(name)
                if expected and expected[0] in complaint:
                    print(f"note {name}: {expected[1]}")
                else:
                    problems.append(
                        (name, f"required arguments were not enough: {complaint}")
                    )

            # Every other advertised argument, sent on top of the required set, to find one
            # the server rejects despite advertising it.
            for argument in sorted(properties - set(declared)):
                extra = probe_value(name, argument)
                if extra is None:
                    continue
                result = server.call(name, **arguments, **{argument: extra})
                complaint = message(result)
                if "unrecognised argument" in complaint:
                    problems.append(
                        (name, f"advertises {argument!r} but the server rejects it")
                    )

    print(f"{len(tools)} tools called with their declared-required arguments\n")
    if not problems:
        print("every tool accepts the call its schema describes")
        return 0
    for name, detail in problems:
        print(f"{name}: {detail}")
    print(f"\n{len(problems)} problems")
    return 1


if __name__ == "__main__":
    sys.exit(main())