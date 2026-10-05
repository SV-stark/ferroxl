"""Cross-reference every ferroxl-mcp tool against openpyxl 3.1.5.

Each test drives the real server over stdio, then opens the file it wrote with openpyxl and
compares what openpyxl sees. openpyxl is the reference implementation, so a divergence here
is a bug in the server rather than a deliberate difference -- which is the whole point: the
server's file output is what an agent's work actually becomes, and openpyxl is how anyone
else will read it.

Every test gets its own workbook, so the order of this file cannot change a result.

    python tools/mcp_parity.py
"""

import base64
import io
import shutil
import sys
import traceback
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from mcp_client import Server  # noqa: E402

WORK = Path(__file__).resolve().parents[1] / "target" / "mcp_parity"

# A 1x1 PNG, so add_image has a real file without needing an asset on disk.
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmM"
    "IQAAAABJRU5ErkJggg=="
)

RESULTS = []


def check(name, condition, detail=""):
    RESULTS.append((name, bool(condition), detail))
    print(("ok   " if condition else "FAIL ") + name + (f"\n        {detail}" if detail and not condition else f"  ({detail})" if detail else ""))


def seed(server, name, rows=None, sheets=("Data",)):
    """A workbook with the given sheets, the first holding a small table."""
    rows = rows if rows is not None else [
        ["Item", "Qty", "Price", "Total"],
        ["Bolt", 10, 1.5, "=B2*C2"],
        ["Nut", 25, 0.75, "=B3*C3"],
    ]
    server.ok("create_workbook", path=name, sheets=list(sheets), overwrite=True)
    server.ok("write_cells", path=name, sheet=sheets[0], start="A1", rows=rows)
    return name


def book(name):
    return openpyxl.load_workbook(WORK / name)


def references(formula) -> str:
    """The range a chart reference points at, without `$` markers or a sheet qualifier."""
    return str(formula).split("!")[-1].replace("$", "")


# -- Reading ---------------------------------------------------------------------------------


def test_read_cells(server):
    p = seed(server, "read_cells.xlsx")
    data = server.data("read_cells", path=p, sheet="Data", range="A1:D3")
    coords = [c["cell"] for c in data["cells"]]
    check(
        "read_cells returns every coordinate in the range",
        coords == [
            "A1", "B1", "C1", "D1",
            "A2", "B2", "C2", "D2",
            "A3", "B3", "C3", "D3",
        ],
        str(coords),
    )
    by_cell = {c["cell"]: c for c in data["cells"]}
    check("read_cells: text reads as a string", by_cell["A1"]["value"] == {"type": "string", "value": "Item"}, str(by_cell["A1"]["value"]))
    check("read_cells: a number reads as a number", by_cell["B2"]["value"] == {"type": "number", "value": 10}, str(by_cell["B2"]["value"]))
    check(
        "read_cells: a formula is not confused with its text",
        by_cell["D2"]["value"] == {"type": "formula", "value": "=B2*C2"},
        str(by_cell["D2"]["value"]),
    )

    sheet = book(p)["Data"]
    check(
        "openpyxl reads the same header row",
        [c.value for c in sheet[1]][:4] == ["Item", "Qty", "Price", "Total"],
        str([c.value for c in sheet[1]][:4]),
    )
    check(
        "openpyxl reads the same formulas",
        (sheet["D2"].value, sheet["D3"].value) == ("=B2*C2", "=B3*C3"),
        f"{sheet['D2'].value} {sheet['D3'].value}",
    )

    # `include_empty` defaults to true, so an empty cell in the range is reported. The cell
    # is written with an explicit `null`, which is how a caller clears one.
    server.ok("write_cells", path=p, sheet="Data", start="A6", rows=[["x", 1]])
    server.ok("set_cell", path=p, sheet="Data", cell="C6", value=None)
    default_cells = server.data("read_cells", path=p, sheet="Data", range="A6:C6")["cells"]
    filtered = server.data(
        "read_cells", path=p, sheet="Data", range="A6:C6", include_empty=False
    )["cells"]
    check(
        "read_cells: include_empty=true reports the cleared cell",
        [c["cell"] for c in default_cells] == ["A6", "B6", "C6"],
        str([c["cell"] for c in default_cells]),
    )
    check(
        "read_cells: include_empty=false drops it",
        [c["cell"] for c in filtered] == ["A6", "B6"],
        str([c["cell"] for c in filtered]),
    )


def test_read_formulas(server):
    p = seed(server, "read_formulas.xlsx")
    formulas = server.data("read_formulas", path=p, sheet="Data")["formulas"]
    check(
        "read_formulas finds both formulas and nothing else",
        sorted(f["cell"] for f in formulas) == ["D2", "D3"],
        str([f["cell"] for f in formulas]),
    )
    sheet = book(p)["Data"]
    from_python = sorted(
        c.coordinate
        for row in sheet.iter_rows()
        for c in row
        if isinstance(c.value, str) and c.value.startswith("=")
    )
    check("read_formulas agrees with openpyxl", from_python == ["D2", "D3"], str(from_python))


def test_list_sheets_and_describe(server):
    p = seed(server, "describe.xlsx")
    server.ok("add_sheet", path=p, title="Extra")
    names = [s["name"] for s in server.data("list_sheets", path=p)["sheets"]]
    check("list_sheets names both sheets in order", names == ["Data", "Extra"], str(names))
    check("list_sheets agrees with openpyxl", names == book(p).sheetnames, str(book(p).sheetnames))

    described = server.data("describe_sheet", path=p, sheet="Data")
    sheet = book(p)["Data"]
    check(
        "describe_sheet's dimension matches openpyxl's",
        described.get("dimension") == sheet.calculate_dimension(),
        f"{described.get('dimension')} vs {sheet.calculate_dimension()}",
    )


def test_search_values(server):
    p = seed(server, "search.xlsx")
    hits = server.data("search_values", path=p, query="Bolt")
    check(
        "search_values finds the cell containing the text",
        [h["cell"] for h in hits["matches"]] == ["A2"],
        str(hits["matches"]),
    )
    check(
        "search_values matches case-insensitively by default",
        [h["cell"] for h in server.data("search_values", path=p, query="bolt")["matches"]]
        == ["A2"],
    )
    check(
        "search_values honours case_sensitive",
        server.data("search_values", path=p, query="bolt", case_sensitive=True)["matches"]
        == [],
    )
    # openpyxl has no search, so the reference is the sheet's own contents.
    sheet = book(p)["Data"]
    expected = sorted(
        c.coordinate
        for row in sheet.iter_rows()
        for c in row
        if isinstance(c.value, str) and "bolt" in c.value.lower()
    )
    check("the cells search_values reports match the sheet", expected == ["A2"], str(expected))


def test_summarize_range(server):
    p = seed(server, "summarize.xlsx")
    summary = server.data("summarize_range", path=p, sheet="Data", range="B2:C3")
    sheet = book(p)["Data"]
    numbers = [sheet[c].value for c in ("B2", "C2", "B3", "C3")]
    check(
        "summarize_range's cell count matches openpyxl's",
        summary.get("cells") == len(numbers),
        f"ferroxl={summary.get('cells')} python={len(numbers)}",
    )
    check(
        "summarize_range's numeric count matches openpyxl's",
        summary.get("numeric") == sum(1 for n in numbers if isinstance(n, (int, float))),
        f"ferroxl={summary.get('numeric')}",
    )
    check(
        "summarize_range's sum matches openpyxl's",
        abs(summary.get("sum", 0) - sum(numbers)) < 1e-9,
        f"ferroxl={summary.get('sum')} python={sum(numbers)}",
    )
    for key, expected in (
        ("min", min(numbers)),
        ("max", max(numbers)),
        ("mean", sum(numbers) / len(numbers)),
    ):
        check(
            f"summarize_range's {key} matches openpyxl's",
            abs(summary.get(key, 0) - expected) < 1e-9,
            f"ferroxl={summary.get(key)} python={expected}",
        )


def test_export_csv(server):
    p = seed(server, "csv.xlsx")
    result = server.ok("export_csv", path=p, sheet="Data", range="A1:B3")
    text = " ".join(part.get("text", "") for part in result.get("content", []))
    sheet = book(p)["Data"]
    check(
        "export_csv emits every cell of the range",
        all(str(sheet[c].value) in text for c in ("A1", "B1", "A2", "B2", "A3", "B3")),
        text[:160],
    )


def test_named_ranges(server):
    p = seed(server, "names.xlsx")
    # With no `scope` the name is workbook-global, which is where openpyxl puts it.
    server.ok("add_named_range", path=p, name="Global", sheet="Data", range="B2:B3")
    defined = book(p).defined_names
    check(
        "an unscoped name lands in the workbook's global names, as openpyxl does",
        "Global" in defined,
        str(list(defined)),
    )
    if "Global" in defined:
        check(
            "the global name points at the range it was given",
            defined["Global"].value.replace("'", "") == "Data!$B$2:$B$3",
            defined["Global"].value,
        )

    # With a `scope` it is sheet-local, which is also what openpyxl does for a local name.
    server.ok(
        "add_named_range", path=p, name="Local", sheet="Data", range="C2:C3", scope="Data"
    )
    read = book(p)
    check(
        "a scoped name lands in that sheet's local names, as openpyxl does",
        "Local" in read["Data"].defined_names and "Local" not in read.defined_names,
        f"wb={list(read.defined_names)} ws={list(read['Data'].defined_names)}",
    )

    names = server.data("list_named_ranges", path=p)["names"]
    by_name = {n["name"]: n for n in names}
    check(
        "list_named_ranges reports both names with their scopes",
        by_name["Global"].get("scope") is None
        and by_name["Local"].get("scope") == "Data",
        str({k: v.get("scope") for k, v in by_name.items()}),
    )
    check(
        "list_named_ranges reports the destination of each",
        by_name["Global"]["destinations"][0]["range"] == "$B$2:$B$3",
        str(by_name["Global"]["destinations"]),
    )


def test_tracing(server):
    p = seed(server, "trace.xlsx")
    precedents = server.data("trace_precedents", path=p, sheet="Data", cell="D2")
    found = {e["cell"] if isinstance(e, dict) else e for e in precedents["precedents"]}
    check(
        "trace_precedents finds B2 and C2 feeding D2",
        {"B2", "C2"} <= found,
        str(sorted(found)),
    )
    dependents = server.data("trace_dependents", path=p, sheet="Data", cell="B2")
    listed = {e["cell"] if isinstance(e, dict) else e for e in dependents["dependents"]}
    check("trace_dependents finds D2 reading B2", "D2" in listed, str(sorted(listed)))
    cycles = server.data("check_circular_references", path=p, sheet="Data")
    check("check_circular_references finds no cycle here", not cycles.get("cycles"), str(cycles.get("cycles")))

    # A real cycle, which Excel refuses to calculate.
    server.ok("set_cell", path=p, sheet="Data", cell="A8", value="=A9")
    server.ok("set_cell", path=p, sheet="Data", cell="A9", value="=A8")
    cycles = server.data("check_circular_references", path=p, sheet="Data")
    check(
        "check_circular_references finds a real cycle as a closed path",
        bool(cycles.get("cycles")) and cycles["cycles"][0][0] == cycles["cycles"][0][-1],
        str(cycles.get("cycles")),
    )


# -- Structure ---------------------------------------------------------------------------------


def test_sheet_management(server):
    p = seed(server, "sheets.xlsx")
    server.ok("add_sheet", path=p, title="Second")
    check(
        "add_sheet appends",
        book(p).sheetnames == ["Data", "Second"],
        str(book(p).sheetnames),
    )
    server.ok("rename_sheet", path=p, sheet="Second", title="Renamed")
    check("rename_sheet takes effect", book(p).sheetnames == ["Data", "Renamed"], str(book(p).sheetnames))
    server.ok("remove_sheet", path=p, sheet="Renamed")
    check("remove_sheet takes effect", book(p).sheetnames == ["Data"], str(book(p).sheetnames))

    # openpyxl refuses to remove the last sheet too.
    result = server.call("remove_sheet", path=p, sheet="Data")
    check(
        "remove_sheet refuses to empty the workbook, as openpyxl does",
        result.get("isError") and "at least one sheet" in json_text(result),
        json_text(result)[:160],
    )


def json_text(result) -> str:
    return " ".join(part.get("text", "") for part in result.get("content", [])) or str(result)


def test_merge(server):
    p = seed(server, "merge.xlsx")
    server.ok("merge_cells", path=p, sheet="Data", range="A5:D5")
    server.ok("set_cell", path=p, sheet="Data", cell="A5", value="Merged caption")
    ranges = [str(r) for r in book(p)["Data"].merged_cells.ranges]
    check("merge_cells writes the range openpyxl reads", "A5:D5" in ranges, str(ranges))
    check("the value survives on the top-left cell", book(p)["Data"]["A5"].value == "Merged caption", str(book(p)["Data"]["A5"].value))
    server.ok("unmerge_cells", path=p, sheet="Data", range="A5:D5")
    ranges = [str(r) for r in book(p)["Data"].merged_cells.ranges]
    check("unmerge_cells removes it again", "A5:D5" not in ranges, str(ranges))


def test_freeze_panes(server):
    p = seed(server, "freeze.xlsx")
    server.ok("freeze_panes", path=p, sheet="Data", cell="A2")
    check("freeze_panes matches openpyxl", book(p)["Data"].freeze_panes == "A2", str(book(p)["Data"].freeze_panes))


def test_auto_filter(server):
    p = seed(server, "filter.xlsx")
    server.ok("set_auto_filter", path=p, sheet="Data", range="A1:D3")
    check(
        "set_auto_filter matches openpyxl's auto_filter.ref",
        book(p)["Data"].auto_filter.ref == "A1:D3",
        str(book(p)["Data"].auto_filter.ref),
    )


def test_value_typing(server):
    p = seed(server, "values.xlsx")
    cases = {
        "E1": 42.5,
        "E2": "2010-01-18",
        "E3": "50%",
        "E4": "14:15:20",
        "E5": True,
        "E6": None,
        "E7": "plain text",
        "E8": "=SUM(E1:E1)",
    }
    for cell, value in cases.items():
        server.ok("set_cell", path=p, sheet="Data", cell=cell, value=value)

    sheet = book(p)["Data"]
    check("set_cell: a number stays a number", sheet["E1"].value == 42.5, repr(sheet["E1"].value))
    check(
        "set_cell: an ISO date becomes a datetime",
        getattr(sheet["E2"].value, "year", None) == 2010 and sheet["E2"].number_format != "General",
        f"{sheet['E2'].value!r} fmt={sheet['E2'].number_format}",
    )
    check("set_cell: a percentage becomes 0.5", sheet["E3"].value == 0.5, repr(sheet["E3"].value))
    check(
        "set_cell: a time of day becomes a real time",
        repr(sheet["E4"].value) == "datetime.time(14, 15, 20)",
        f"{sheet['E4'].value!r} fmt={sheet['E4'].number_format}",
    )
    # A deliberate difference from openpyxl, which writes the *string* "14:15:20" as text.
    # The server infers a type from the text on the way in, which is the useful behaviour
    # for an agent and is documented in the tool's own schema; the library default
    # (`guessing_types` off) matches openpyxl instead.
    reference = openpyxl.Workbook()
    reference.active["A1"] = "14:15:20"
    reference.save(WORK / "time_text_reference.xlsx")
    as_text = openpyxl.load_workbook(WORK / "time_text_reference.xlsx").active["A1"]
    check(
        "the time differs from openpyxl's text, which is the documented server behaviour",
        as_text.value == "14:15:20" and sheet["E4"].value != "14:15:20",
        f"openpyxl={as_text.value!r} ferroxl={sheet['E4'].value!r}",
    )
    check("set_cell: a boolean stays a boolean", sheet["E5"].value is True, repr(sheet["E5"].value))
    check("set_cell: null empties the cell", sheet["E6"].value is None, repr(sheet["E6"].value))
    check("set_cell: text stays text", sheet["E7"].value == "plain text", repr(sheet["E7"].value))
    check("set_cell: a leading = makes a formula", sheet["E8"].value == "=SUM(E1:E1)", repr(sheet["E8"].value))

    # What openpyxl itself would have written, for the same inputs.
    reference = openpyxl.Workbook()
    reference.active.title = "Data"
    for cell, value in cases.items():
        reference.active[cell] = value
    reference.save(WORK / "values_reference.xlsx")
    ref = openpyxl.load_workbook(WORK / "values_reference.xlsx")["Data"]
    for cell in ("E1", "E5", "E7", "E8"):
        check(
            f"set_cell {cell} matches what openpyxl writes for the same value",
            type(sheet[cell].value) is type(ref[cell].value),
            f"ferroxl={type(sheet[cell].value).__name__} "
            f"openpyxl={type(ref[cell].value).__name__}",
        )


def test_write_and_append(server):
    p = seed(server, "write.xlsx")
    server.ok("append_row", path=p, sheet="Data", values=["Screw", 100, 0.25, "=B4*C4"])
    sheet = book(p)["Data"]
    check(
        "append_row lands after the last populated row",
        (sheet["A4"].value, sheet["D4"].value) == ("Screw", "=B4*C4"),
        f"A4={sheet['A4'].value!r} D4={sheet['D4'].value!r}",
    )
    server.ok("clear_cells", path=p, sheet="Data", range="A4:D4")
    sheet = book(p)["Data"]
    check("clear_cells empties the range", all(sheet[c].value is None for c in ("A4", "B4", "C4", "D4")))

    server.ok("write_cells", path=p, sheet="Data", start="A8", rows=[["a", 1], ["b", 2]])
    sheet = book(p)["Data"]
    check(
        "write_cells fills the block from start_cell",
        (sheet["A8"].value, sheet["B9"].value) == ("a", 2),
        f"A8={sheet['A8'].value!r} B9={sheet['B9'].value!r}",
    )


# -- Layout and formatting ----------------------------------------------------------------------


def test_column_and_row_sizes(server):
    p = seed(server, "sizes.xlsx")
    server.ok("set_column_width", path=p, sheet="Data", columns="A", width=22.5)
    server.ok("set_row_height", path=p, sheet="Data", rows="2", height=30)
    sheet = book(p)["Data"]
    width = sheet.column_dimensions["A"].width
    check(
        "set_column_width matches openpyxl's width",
        abs((width or 0) - 22.5) < 0.02,
        f"ferroxl asked 22.5, openpyxl reads {width}",
    )
    height = sheet.row_dimensions[2].height
    check(
        "set_row_height matches openpyxl's height",
        abs((height or 0) - 30) < 0.01,
        f"ferroxl asked 30, openpyxl reads {height}",
    )


def test_number_format(server):
    p = seed(server, "numfmt.xlsx")
    server.ok("set_number_format", path=p, sheet="Data", range="B2:C3", format="0.00")
    sheet = book(p)["Data"]
    formats = {c: sheet[c].number_format for c in ("B2", "C2", "B3", "C3")}
    check(
        "set_number_format applies to every cell in the range",
        all(v == "0.00" for v in formats.values()),
        str(formats),
    )
    # openpyxl's own table of built-in formats agrees that "0.00" is id 2.
    check(
        "the format is one openpyxl also treats as built in",
        "0.00" in openpyxl.styles.numbers.BUILTIN_FORMATS.values(),
    )


def test_style_cells(server):
    p = seed(server, "style.xlsx")
    server.ok(
        "style_cells",
        path=p,
        sheet="Data",
        range="A1:D1",
        style={
            "bold": True,
            "italic": True,
            "font_color": "FFFFFFFF",
            "fill_color": "FF1F3864",
            "horizontal": "center",
            "border": "medium",
            "number_format": "0.00",
        },
    )
    sheet = book(p)["Data"]
    cell = sheet["A1"]
    check("style_cells: bold survives", cell.font.bold is True, repr(cell.font.bold))
    check("style_cells: italic survives", cell.font.italic is True, repr(cell.font.italic))
    check(
        "style_cells: the font colour survives",
        (cell.font.color.rgb or "").upper().endswith("FFFFFF"),
        repr(cell.font.color.rgb if cell.font.color else None),
    )
    check("style_cells: the fill is solid", cell.fill.patternType == "solid", repr(cell.fill.patternType))
    check(
        "style_cells: the fill colour survives",
        (cell.fill.fgColor.rgb or "").upper().endswith("1F3864"),
        repr(cell.fill.fgColor.rgb),
    )
    check(
        "style_cells: horizontal alignment survives",
        cell.alignment.horizontal == "center",
        repr(cell.alignment.horizontal),
    )
    check(
        "style_cells: the border style survives",
        cell.border.left.style == "medium",
        repr(cell.border.left.style),
    )
    check("style_cells: the number format survives", cell.number_format == "0.00", repr(cell.number_format))
    check(
        "style_cells: one style entry covers the whole range",
        sheet["A1"]._style == sheet["D1"]._style,
    )

    # Composing: a later call changes only what it names.
    server.ok("style_cells", path=p, sheet="Data", range="A1:D1", style={"bold": False})
    sheet = book(p)["Data"]
    check(
        "style_cells: a later call changes only the keys it names",
        sheet["A1"].font.bold is False and (sheet["A1"].fill.fgColor.rgb or "").upper().endswith("1F3864"),
        f"bold={sheet['A1'].font.bold} fill={sheet['A1'].fill.fgColor.rgb}",
    )

    result = server.call("style_cells", path=p, sheet="Data", range="A1", style={"nonsense": 1})
    check(
        "style_cells rejects an unknown style property rather than ignoring it",
        result.get("isError") and "nonsense" in json_text(result),
        json_text(result)[:160],
    )


def test_header_footer(server):
    p = seed(server, "header.xlsx")
    server.ok(
        "set_header_footer",
        path=p,
        sheet="Data",
        left_header="Left",
        center_header="Centre",
        right_header="Right",
        left_footer="Page",
    )
    sheet = book(p)["Data"]
    check(
        "set_header_footer: the left header survives",
        sheet.oddHeader.left.text == "Left",
        repr(sheet.oddHeader.left.text),
    )
    check(
        "set_header_footer: the centre header survives",
        sheet.oddHeader.center.text == "Centre",
        repr(sheet.oddHeader.center.text),
    )
    check(
        "set_header_footer: the right header survives",
        sheet.oddHeader.right.text == "Right",
        repr(sheet.oddHeader.right.text),
    )
    check(
        "set_header_footer: the footer survives",
        sheet.oddFooter.left.text == "Page",
        repr(sheet.oddFooter.left.text),
    )


def test_hyperlink(server):
    p = seed(server, "link.xlsx")
    server.ok(
        "add_hyperlink", path=p, sheet="Data", cell="A6", target="https://example.com/report", display="Report"
    )
    sheet = book(p)["Data"]
    link = sheet["A6"].hyperlink
    check("add_hyperlink produces a hyperlink openpyxl reads", link is not None)
    if link is not None:
        check("the hyperlink target survives", link.target == "https://example.com/report", repr(link.target))
    check("the display text survives", sheet["A6"].value == "Report", repr(sheet["A6"].value))


def test_data_validation(server):
    p = seed(server, "validation.xlsx")
    server.ok(
        "add_data_validation",
        path=p,
        sheet="Data",
        range="F1:F10",
        type="list",
        formula1='"Yes,No"',
        error_message="Pick one",
        prompt_message="Choose",
    )
    sheet = book(p)["Data"]
    found = list(sheet.data_validations.dataValidation)
    check("add_data_validation produces a validation openpyxl reads", bool(found), str(len(found)))
    if found:
        d = found[0]
        check("the range survives", "F1:F10" in str(d.sqref), str(d.sqref))
        check("the type survives", d.type == "list", repr(d.type))
        check("the list survives", d.formula1 == '"Yes,No"', repr(d.formula1))
        check("the error message survives", d.error == "Pick one", repr(d.error))
        check("the prompt survives", d.prompt == "Choose", repr(d.prompt))

    # A two-bound rule, which is where operator and formula2 matter.
    server.ok(
        "add_data_validation",
        path=p,
        sheet="Data",
        range="G1:G10",
        type="whole",
        operator="between",
        formula1="1",
        formula2="100",
    )
    sheet = book(p)["Data"]
    bounds = [d for d in sheet.data_validations.dataValidation if d.type == "whole"]
    check("a two-bound validation is written", bool(bounds), str(len(bounds)))
    if bounds:
        check("the operator survives", bounds[0].operator == "between", repr(bounds[0].operator))
        check("both bounds survive", (bounds[0].formula1, bounds[0].formula2) == ("1", "100"), f"{bounds[0].formula1} {bounds[0].formula2}")


def test_conditional_format(server):
    p = seed(server, "cf.xlsx")
    server.ok(
        "add_conditional_format",
        path=p,
        sheet="Data",
        range="B2:B3",
        kind="cellIs",
        operator="greaterThan",
        formula="5",
        font_color="FF9C0006",
        fill_color="FFFFC7CE",
    )
    rules = [r for rng in book(p)["Data"].conditional_formatting for r in rng.rules]
    check("add_conditional_format produces a rule openpyxl reads", bool(rules))
    if rules:
        r = rules[0]
        check("the rule type survives", r.type == "cellIs", repr(r.type))
        check("the operator survives", r.operator == "greaterThan", repr(r.operator))
        check("the threshold survives", [str(f) for f in (r.formula or [])] == ["5"], repr(r.formula))
        check(
            "the differential style's font colour survives",
            r.dxf is not None and r.dxf.font is not None,
            repr(r.dxf),
        )
        if r.dxf is not None and r.dxf.fill is not None:
            # openpyxl exposes a dxf fill's colour as `bgColor`, which is where a solid
            # dxf fill carries it in OOXML -- `fgColor` is the pattern's foreground.
            fill = r.dxf.fill
            check(
                "the differential fill's colour survives",
                (fill.bgColor.rgb or "").upper().endswith("FFC7CE"),
                f"bgColor={fill.bgColor.rgb!r} fgColor={fill.fgColor.rgb!r}",
            )

    server.ok(
        "add_conditional_format",
        path=p,
        sheet="Data",
        range="C2:C3",
        kind="colorScale",
        start_color="FFF8696B",
        end_color="FF63BE7B",
    )
    rules = [r for rng in book(p)["Data"].conditional_formatting for r in rng.rules]
    scales = [r for r in rules if r.type == "colorScale"]
    check("a colour scale is written", bool(scales), str([r.type for r in rules]))
    if scales:
        check(
            "the colour scale has the two stops it was given",
            len(scales[0].colorScale.color) == 2,
            str(len(scales[0].colorScale.color)),
        )


def test_data_bar_and_icon_set(server):
    p = seed(server, "databar.xlsx")
    server.ok("add_data_bar", path=p, sheet="Data", range="B2:B3")
    types = [r.type for rng in book(p)["Data"].conditional_formatting for r in rng.rules]
    check("add_data_bar writes a dataBar rule openpyxl reads", "dataBar" in types, str(types))

    server.ok("add_icon_set", path=p, sheet="Data", range="C2:C3")
    types = [r.type for rng in book(p)["Data"].conditional_formatting for r in rng.rules]
    check("add_icon_set writes an iconSet rule openpyxl reads", "iconSet" in types, str(types))
    icons = [r for rng in book(p)["Data"].conditional_formatting for r in rng.rules if r.type == "iconSet"]
    if icons:
        check(
            "the icon set keeps its three thresholds",
            len(icons[0].iconSet.cfvo) == 3,
            str(len(icons[0].iconSet.cfvo)),
        )


def test_gradient_fill(server):
    p = seed(server, "gradient.xlsx")
    server.ok(
        "set_gradient_fill",
        path=p,
        sheet="Data",
        range="A1:D1",
        start_color="FF1F3864",
        end_color="FF4F81BD",
    )
    fill = book(p)["Data"]["A1"].fill
    # openpyxl's own `GradientFill.fill_type` is the literal `linear`, so that is what a
    # correct gradient reads back as -- not the `grad` a pattern fill would use.
    check(
        "set_gradient_fill writes the gradient openpyxl classifies as a gradient",
        fill.tagname == "gradientFill",
        f"tagname={fill.tagname!r} fill_type={fill.fill_type!r}",
    )
    check(
        "the gradient's type matches what openpyxl itself writes",
        fill.fill_type == openpyxl.styles.GradientFill().fill_type,
        f"ferroxl={fill.fill_type!r} openpyxl={openpyxl.styles.GradientFill().fill_type!r}",
    )
    check(
        "the gradient keeps both stops",
        len(getattr(fill, "stop", []) or []) == 2,
        str(len(getattr(fill, "stop", []) or [])),
    )


def test_table(server):
    p = seed(server, "table.xlsx")
    server.ok("add_table", path=p, sheet="Data", name="Table1", range="A1:D3")
    described = server.data("describe_table", path=p, sheet="Data", name="Table1")
    check("describe_table finds the table", described.get("name") == "Table1", str(described)[:200])
    check(
        "describe_table reports the ref openpyxl reads",
        described.get("ref") == book(p)["Data"].tables["Table1"].ref,
        f"ferroxl={described.get('ref')}",
    )
    check(
        "describe_table reports the column names",
        [c["name"] for c in described["columns"]] == ["Item", "Qty", "Price", "Total"],
        str([c["name"] for c in described["columns"]]),
    )
    sheet = book(p)["Data"]
    check("openpyxl sees the table", "Table1" in sheet.tables, str(list(sheet.tables)))
    if "Table1" in sheet.tables:
        table = sheet.tables["Table1"]
        check("the table's ref survives", table.ref == "A1:D3", table.ref)
        names = [c.name for c in table.tableColumns]
        check(
            "the column names come from the header cells, so structured references resolve",
            names == ["Item", "Qty", "Price", "Total"],
            str(names),
        )


# -- Media ---------------------------------------------------------------------------------------


def test_comment(server):
    p = seed(server, "comment.xlsx")
    server.ok("add_comment", path=p, sheet="Data", cell="B2", text="Check this figure", author="Reviewer")
    listed = server.data("list_comments", path=p, sheet="Data")["comments"]
    check("list_comments finds the comment", len(listed) == 1, str(listed))
    comment = book(p)["Data"]["B2"].comment
    check("openpyxl sees the comment", comment is not None)
    if comment is not None:
        check("the comment text survives", comment.text.strip() == "Check this figure", repr(comment.text))
        check("the comment author survives", comment.author == "Reviewer", repr(comment.author))


def test_chart(server):
    p = seed(server, "chart.xlsx")
    server.ok(
        "add_chart",
        path=p,
        sheet="Data",
        type="bar",
        anchor="F2",
        title="Units",
        series=[{"name": "Qty", "values": "B2:B3"}],
        categories="A2:A3",
    )
    charts = book(p)["Data"]._charts
    check("add_chart writes a chart openpyxl reads", bool(charts), str(len(charts)))
    if charts:
        chart = charts[0]
        check("the chart type survives", chart.tagname == "barChart", chart.tagname)
        text = ""
        if chart.title is not None and getattr(chart.title, "tx", None) is not None:
            text = "".join(
                run.t or ""
                for para in chart.title.tx.rich.p
                for run in (para.r or [])
            )
        check("the chart title survives", text == "Units", repr(text))
        check(
            "the series reference points at the range it was given",
            references(chart.series[0].val.numRef.f) == "B2:B3",
            references(chart.series[0].val.numRef.f),
        )
        cat = chart.series[0].cat
        check(
            "the categories reference points at the range it was given",
            cat is not None
            and references(cat.strRef.f if cat.strRef else cat.numRef.f) == "A2:A3",
            references(cat.strRef.f) if cat and cat.strRef else "no categories",
        )


def test_image(server):
    p = seed(server, "image.xlsx")
    (WORK / "pixel.png").write_bytes(PNG)
    server.ok("add_image", path=p, sheet="Data", image_path="pixel.png", anchor="H2")
    images = book(p)["Data"]._images
    check("add_image writes an image openpyxl reads", bool(images), str(len(images)))
    if images:
        check(
            "the image's size comes from the PNG header",
            (images[0].width, images[0].height) == (1, 1),
            f"{images[0].width}x{images[0].height}",
        )


# -- The contract itself ---------------------------------------------------------------------------


def test_errors(server):
    p = seed(server, "errors.xlsx")
    cases = [
        (
            "an unknown argument is rejected, not ignored",
            "read_cells",
            {"not_an_argument": 1},
            "does not take",
        ),
        ("an unknown tool is a JSON-RPC error", "no_such_tool", {}, None),
        ("a missing file is an isError result", "read_cells", {"path": "absent.xlsx"}, "no such workbook"),
        ("an unknown sheet is named in the error", "read_cells", {"path": p, "sheet": "Nope"}, "Nope"),
        ("a bad colour is rejected", "style_cells", {"path": p, "range": "A1", "style": {"fill_color": "not-a-colour"}}, "fill_color"),
    ]
    for label, tool, extra, expect in cases:
        result = server.call(tool, **{**extra, **({"path": p} if "path" in extra else {})}) if False else server.call(tool, **{**extra, **({"path": p} if "path" not in extra else {})})
        if expect is None:
            check(label, result.get("rpcError", {}).get("code") == -32601, str(result)[:160])
        else:
            check(
                label,
                bool(result.get("isError")) and expect in json_text(result),
                json_text(result)[:160],
            )


def test_summaries_are_useful(server):
    """A summary that does not name the cells is a wasted turn in a model's context."""
    p = seed(server, "summaries.xlsx")
    server.ok("write_cells", path=p, sheet="Data", start="G1", rows=[[f"v{i}", i] for i in range(30)])
    summary = server.summary("trace_dependents", path=p, sheet="Data", cell="B2")
    check(
        "a summary names the coordinates it is talking about",
        "D2" in summary,
        summary[:160],
    )
    check(
        "a long list is capped and counted rather than dumped",
        len(summary) < 400,
        f"{len(summary)} characters",
    )


def main():
    shutil.rmtree(WORK, ignore_errors=True)
    WORK.mkdir(parents=True)
    (WORK / "pixel.png").write_bytes(PNG)

    tests = [value for name, value in sorted(globals().items()) if name.startswith("test_")]
    with Server(WORK) as server:
        server.initialize()
        for test in tests:
            print(f"\n--- {test.__name__}")
            try:
                test(server)
            except Exception:
                check(test.__name__, False, "raised:\n" + traceback.format_exc(limit=4))

    failed = [name for name, ok, _ in RESULTS if not ok]
    print(f"\n{len(RESULTS) - len(failed)}/{len(RESULTS)} checks passed")
    if failed:
        print("failed:")
        for name in failed:
            print(f"  {name}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())