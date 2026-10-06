"""Call every ferroxl-mcp tool against real Excel workbooks and verify with openpyxl.

The other harnesses in tools/ all build their fixtures from the server's own
`create_workbook`, so a tool can agree with a fixture the server itself wrote. This one
does the opposite: every case starts from a workbook that real software produced, and
every mutation is checked by reading the result back with openpyxl -- the reference
implementation -- rather than by asking the server what it thinks it did.

Each case gets its own copy of its fixture, so no case can be affected by the order of
the loop.

    python tools/mcp_real_files.py
"""

import os
import re
import shutil
import sys
import warnings
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from mcp_client import Server  # noqa: E402

import openpyxl  # noqa: E402

warnings.simplefilter("ignore", UserWarning)

ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT / "target" / "mcp_real_files"

# Real workbooks, each the best honest example of the feature it is used for, taken from the
# openpyxl test corpus, so they are files real software produced.
#
# Where the openpyxl fixtures live. The upstream checkout is a sibling of this repository,
# which is where CI puts it; override with FERROXL_OPENPYXL to point somewhere else.
OPENPYXL = Path(os.environ.get("FERROXL_OPENPYXL") or ROOT.parent / "openpyxl")

SOURCES = {
    "sample": OPENPYXL / "openpyxl" / "reader" / "tests" / "data" / "sample.xlsx",
    "styles": OPENPYXL / "openpyxl" / "reader" / "tests" / "data" / "complex-styles.xlsx",
    "condfmt": OPENPYXL / "openpyxl" / "formatting" / "tests" / "data" / "conditional-formatting.xlsx",
    "comments": OPENPYXL / "openpyxl" / "comments" / "tests" / "data" / "comments.xlsx",
    "images": OPENPYXL / "openpyxl" / "reader" / "tests" / "data" / "sample_with_images.xlsx",
    "table": OPENPYXL / "openpyxl" / "reader" / "tests" / "data" / "print_area_table_defined_name.xlsx",
    "vba": OPENPYXL / "openpyxl" / "tests" / "data" / "reader" / "vba-test.xlsm",
    "bigfoot": OPENPYXL / "openpyxl" / "tests" / "data" / "reader" / "bigfoot.xlsx",
}

# The one workbook that is generated rather than checked in.
#
# It reproduces what `cargo run --example build_and_read` writes, because 90 of the 103 calls
# run against it. Two reasons not to use the example's own output: `*.xlsx` is gitignored, so
# on a fresh checkout -- and in CI, which never runs the examples -- the file simply is not
# there; and if the harness used it when present and generated it when absent, a local run and
# a CI run would be testing different files, which is the worst property a test harness can
# have. Generating it unconditionally removes both.
#
# It is written with openpyxl, never with the server. That is the whole premise of this file:
# nothing under test is allowed to have built the workbook being read back.
GENERATED = {"orders": "write_orders_workbook"}


def write_orders_workbook(path):
    """Write the `orders` fixture: two sheets, formulas, a merge, frozen panes, real styles."""
    from openpyxl.styles import Font, PatternFill

    book = openpyxl.Workbook()
    book.active.title = "Sheet1"
    sheet = book.create_sheet("Orders")

    for row in (["Item", "Qty", "Price", "Total"],
                ["Bolt", 10, 1.5, "=B2*C2"],
                ["Nut", 25, 0.75, "=B3*C3"]):
        sheet.append(row)

    # A bold white-on-navy header, so `style_cells` and `read_cells` have something real to
    # find. openpyxl writes the colour as `FFFFFFFF`, which is what the example writes too.
    for letter in "ABCD":
        sheet[f"{letter}1"].font = Font(bold=True, color="FFFFFFFF")
        sheet[f"{letter}1"].fill = PatternFill("solid", fgColor="FF1F3864")

    # `merge_cells` blanks everything but the top-left cell, so merge before writing.
    sheet.merge_cells("A6:D6")
    sheet["A6"] = "Two line items"
    sheet.freeze_panes = "A2"
    for coordinate in ("B2", "C2", "B3", "C3"):
        sheet[coordinate].number_format = "0.00"

    book.save(path)
    return path

# A 1x1 PNG so add_image has a real file to embed.
PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d494844520000000100000001080600000"
    "01f15c4890000000d4944415478da63f8cfc0f01f0005000101ff9d"
    "b2a80000000049454e44ae426082"
)


class Case:
    """One tool call, the real workbook it runs against, and how to check it landed.

    `expect` receives the name of the file the call wrote and returns True. `check`
    receives the tool's `structuredContent`. `refuses` names the text a call that is
    *meant* to fail has to say -- a guard or an argument mistake is a pass only when the
    message tells a model what to fix.
    """

    def __init__(self, tool, fixture, args, expect=None, check=None, setup=None,
                 note=None, refuses=None, extension=".xlsx", verify_as=None):
        self.tool = tool
        self.fixture = fixture
        self.args = args
        self.expect = expect
        self.check = check
        self.setup = setup or []
        self.note = note
        self.refuses = refuses
        self.extension = extension
        self.verify_as = verify_as


# -- readback helpers ------------------------------------------------------------------
# Every one of these asks openpyxl, never the server.

def fixture_path(name):
    """Where a fixture's pristine copy lives, generating it the first time if it is ours."""
    if name in SOURCES:
        return SOURCES[name]
    generated = WORK / f"{name}.xlsx"
    if not generated.is_file():
        globals()[GENERATED[name]](generated)
    return generated


def load(name):
    return openpyxl.load_workbook(WORK / name)


def sheets(name):
    return load(name).sheetnames


def merged(name, sheet):
    return sorted(str(r) for r in load(name)[sheet].merged_cells.ranges)


def rule_types(name, sheet):
    return [rule.type
            for cf in load(name)[sheet].conditional_formatting
            for rule in cf.rules]


def table_columns(name, sheet, table):
    return [c.name for c in load(name)[sheet].tables[table].tableColumns]


def zip_names(name):
    with zipfile.ZipFile(WORK / name) as zf:
        return zf.namelist()


def sheet_part(name, title):
    """The part name of one sheet's XML, which is not sheet1.xml by sheet order."""
    with zipfile.ZipFile(WORK / name) as zf:
        book = zf.read("xl/workbook.xml").decode()
        rels = zf.read("xl/_rels/workbook.xml.rels").decode()
    targets = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid = re.search(rf'<(?:\w+:)?sheet name="{re.escape(title)}"[^>]*r:id="([^"]+)"',
                    book).group(1)
    return "xl/" + targets[rid].lstrip("/")


def sheet_xml(name, title):
    with zipfile.ZipFile(WORK / name) as zf:
        return zf.read(sheet_part(name, title)).decode()


def col_order(name, title="Orders"):
    return [int(m) for m in re.findall(r'<col min="(\d+)"', sheet_xml(name, title))]


def no_dangling(name):
    """Whether every relationship in the package names a part that is actually there.

    This is the general form of the drawing defect, and the check that would have caught it:
    a chart or an image whose part was deleted left the sheet's relationship pointing at
    nothing, which is a file no reader can open. Checking relationships directly says so,
    rather than leaving it to whoever opens the file next.
    """
    names = set(zip_names(name))
    with zipfile.ZipFile(WORK / name) as zf:
        for entry in zf.namelist():
            if not entry.endswith(".rels"):
                continue
            owner_dir = ""
            if not entry.startswith("_rels/"):
                owner = entry.replace("/_rels/", "/").removesuffix(".rels")
                owner_dir = owner.rsplit("/", 1)[0] if "/" in owner else ""
            text = zf.read(entry).decode()
            for target in re.findall(r'Target="([^"]+)"', text):
                if target.startswith(("http", "mailto:", "file:")):
                    continue
                if target.startswith("/"):
                    resolved = target.lstrip("/")
                else:
                    segments = owner_dir.split("/") if owner_dir else []
                    for piece in target.split("/"):
                        if piece == "..":
                            segments.pop()
                        elif piece != ".":
                            segments.append(piece)
                    resolved = "/".join(segments)
                if resolved not in names:
                    return False
    return True


def row_heights(name, title="Orders"):
    """Every row element's height, straight out of the written XML.

    A row with no height at all is absent from this map, which is the point: the writer
    omits `<row>` elements for rows that hold no cells, so a height set on an empty row
    leaves no trace here and openpyxl cannot see one either.
    """
    out = {}
    for tag in re.findall(r"<row [^>]*/?>", sheet_xml(name, title)):
        r = re.search(r'r="(\d+)"', tag)
        ht = re.search(r'ht="([\d.]+)"', tag)
        if r:
            out[int(r.group(1))] = float(ht.group(1)) if ht else None
    return out


def cells(name, sheet, ref):
    return load(name)[sheet][ref].value


def stored_value(name, title, ref):
    """The value as written, before openpyxl reinterprets it through its number format.

    A numeric cell carrying a date format reads back as a datetime, which says nothing
    about what is in the file; this reads the `<v>` element itself.
    """
    cell = re.search(rf'<c r="{re.escape(ref)}"[^>]*>(.*?)</c>',
                     sheet_xml(name, title), re.S)
    if not cell:
        return None
    value = re.search(r"<v>(.*?)</v>", cell.group(1), re.S)
    return value.group(1) if value else None


# -- the matrix ------------------------------------------------------------------------

CASES = [
    # -- Inspection ---------------------------------------------------------------------
    Case("list_sheets", "orders", {},
         expect=lambda f: sheets(f) == ["Sheet1", "Orders"],
         check=lambda d: [s["name"] for s in d["sheets"]] == ["Sheet1", "Orders"]),
    Case("list_sheets", "vba", {}, extension=".xlsm",
         expect=lambda f: sheets(f) == ["Scratch"],
         note="reads a macro-enabled workbook"),
    Case("list_sheets", "bigfoot", {},
         expect=lambda f: len(sheets(f)) == 1024,
         note="a real 410 KB workbook with 1024 sheets"),

    Case("describe_sheet", "orders", {"sheet": "Orders"},
         expect=lambda f: load(f)["Orders"].max_row == 6,
         note="a real sheet with a merge, freeze panes and formulas"),
    Case("describe_sheet", "condfmt", {},
         expect=lambda f: len(list(load(f)["Sheet1"].conditional_formatting)) == 30,
         note="30 real conditional formats counted"),
    Case("describe_sheet", "images", {},
         expect=lambda f: len(load(f)["Sheet1"]._images) == 3,
         note="3 real embedded images counted"),
    Case("describe_sheet", "comments", {"sheet": "Sheet1"},
         expect=lambda f: sum(1 for row in load(f)["Sheet1"].iter_rows()
                              for c in row if c.comment) == 6,
         note="6 real comments counted"),

    Case("read_cells", "orders", {"sheet": "Orders", "range": "A1:D3"},
         expect=lambda f: cells(f, "Orders", "D2") == "=B2*C2",
         note="real formulas and number formats read back"),
    Case("read_cells", "styles", {"sheet": "Sheet1", "range": "A1:I26",
                                  "include_empty": True},
         expect=lambda f: cells(f, "Sheet1", "A3") == "Arial Font, Bold 12",
         note="a 26x9 block of real styles"),
    Case("read_cells", "sample", {},
         expect=lambda f: cells(f, "OXM", "A1").startswith("Oxford Industries"),
         note="no range: the used range of a real 150-row sheet"),

    Case("read_formulas", "orders", {"sheet": "Orders"},
         expect=lambda f: cells(f, "Orders", "D3") == "=B3*C3",
         check=lambda d: {x["cell"] for x in d["formulas"]} == {"D2", "D3"},
         note="both real formulas listed"),

    Case("trace_precedents", "orders", {"sheet": "Orders", "cell": "D2"},
         expect=lambda f: cells(f, "Orders", "B2") == 10,
         check=lambda d: d["precedents"] == ["B2", "C2"],
         note="D2 = B2*C2, both real precedents in evaluation order"),
    Case("trace_dependents", "orders", {"sheet": "Orders", "cell": "B2"},
         expect=lambda f: cells(f, "Orders", "D2") == "=B2*C2",
         check=lambda d: d["dependents"] == ["D2"],
         note="B2 feeds D2"),

    Case("check_circular_references", "orders", {"sheet": "Orders"},
         expect=lambda f: True,
         check=lambda d: d["cycles"] == [],
         note="a clean real workbook reports no cycles"),
    Case("check_circular_references", "orders", {"sheet": "Orders"},
         setup=[("set_cell", {"sheet": "Orders", "cell": "A9", "value": "=B9"}),
                ("set_cell", {"sheet": "Orders", "cell": "B9", "value": "=A9"})],
         expect=lambda f: True,
         check=lambda d: any(c[0] == "A9" and c[-1] == "A9" for c in d["cycles"]),
         note="a cycle written into a real file is found and closed"),

    Case("add_data_bar", "orders", {"sheet": "Orders", "range": "B2:B4",
                                    "color": "638EC6"},
         expect=lambda f: rule_types(f, "Orders") == ["dataBar"],
         note="a data bar on the real Qty column"),
    Case("add_icon_set", "orders", {"sheet": "Orders", "range": "B2:B4",
                                    "style": "3TrafficLights1"},
         expect=lambda f: rule_types(f, "Orders") == ["iconSet"]),

    Case("add_table", "orders", {"sheet": "Orders", "range": "A1:D4",
                                 "name": "OrderTable"},
         expect=lambda f: table_columns(f, "Orders", "OrderTable")
         == ["Item", "Qty", "Price", "Total"],
         note="column names read from the real header row"),
    Case("add_table", "orders", {"sheet": "Orders", "range": "A1:D4",
                                 "name": "Styled", "style": "TableStyleMedium9",
                                 "totals_row": True},
         expect=lambda f: load(f)["Orders"].tables["Styled"].tableStyleInfo is not None),
    Case("describe_table", "table", {"name": "InvoiceData"},
         expect=lambda f: True,
         check=lambda d: d["name"] == "InvoiceData"
         and [c["name"] for c in d["columns"]] == ["Invoice", "Merchant", "Total"],
         note="describes the table a real file already had"),

    Case("set_gradient_fill", "orders", {"sheet": "Orders", "range": "A1:D1",
                                         "start_color": "FF1F3864",
                                         "end_color": "FF4F81BD"},
         expect=lambda f: load(f)["Orders"]["A1"].fill.fill_type == "linear"),

    Case("search_values", "orders", {"sheet": "Orders", "query": "Bolt"},
         check=lambda d: [m["cell"] for m in d["matches"]] == ["A2"],
         note="a real hit at A2"),
    Case("search_values", "orders", {"sheet": "Orders", "query": "bolt"},
         check=lambda d: len(d["matches"]) == 1,
         note="case-insensitive by default"),
    Case("search_values", "orders", {"sheet": "Orders", "query": "BOLT",
                                     "case_sensitive": True},
         check=lambda d: d["matches"] == [],
         note="case_sensitive narrows it to nothing"),
    Case("search_values", "orders", {"sheet": "Orders", "query": "=B",
                                     "match_formulas": True},
         check=lambda d: [m["cell"] for m in d["matches"]] == ["D2", "D3"],
         note="formula text searched as well as values"),
    Case("search_values", "orders", {"sheet": "Orders", "query": "o",
                                     "max_results": 2},
         check=lambda d: len(d["matches"]) == 2 and d["truncated"] is True,
         note="max_results honoured, truncation reported"),

    Case("summarize_range", "orders", {"sheet": "Orders", "range": "B2:C3"},
         check=lambda d: d["numeric"] == 4 and d["sum"] == 37.25
         and d["mean"] == 9.3125 and d["min"] == 0.75 and d["max"] == 25.0,
         note="10+25+1.5+0.75 = 37.25 over the real numbers"),
    Case("summarize_range", "orders", {"sheet": "Orders"},
         check=lambda d: d["cells"] == 24 and d["numeric"] == 4
         and d["text"] == 7 and d["blanks"] == 11,
         note="whole used range, 13 populated and 11 blank cells"),

    Case("list_comments", "comments", {"sheet": "Sheet1"},
         check=lambda d: len(d["comments"]) == 6
         and d["comments"][0]["author"] == "Cuke",
         note="6 real comments with their authors"),
    Case("list_comments", "comments", {"sheet": "Sheet3"},
         check=lambda d: len(d["comments"]) == 1),
    Case("list_comments", "comments", {"sheet": "Sheet2"},
         check=lambda d: d["comments"] == [],
         note="a sheet with no comments is empty, not an error"),

    Case("list_named_ranges", "orders", {},
         check=lambda d: d["names"] == [],
         note="a workbook with no defined names returns an empty list"),
    Case("list_named_ranges", "orders", {},
         setup=[("add_named_range", {"name": "Qtys", "sheet": "Orders",
                                     "range": "B2:B4"})],
         expect=lambda f: "Qtys" in load(f).defined_names,
         check=lambda d: any(n["name"] == "Qtys" for n in d["names"]),
         note="a name added to a real file is listed back"),

    Case("export_csv", "orders", {"sheet": "Orders", "range": "A1:D4"},
         check=lambda d: d["csv"] == "Item,Qty,Price,Total\r\n"
                                    "Bolt,10,1.5,=B2*C2\r\nNut,25,0.75,=B3*C3\r\n,,,",
         note="RFC 4180 CRLF line endings, formulas not evaluated"),
    Case("export_csv", "orders", {"sheet": "Orders"},
         check=lambda d: d["row_count"] == 6 and "Two line items" in d["csv"],
         note="the whole used range, gaps kept"),

    # -- Structure ----------------------------------------------------------------------
    Case("create_workbook", "orders", {"path": "fresh.xlsx", "sheets": ["Alpha", "Beta"]},
         verify_as="fresh.xlsx",
         expect=lambda f: sheets(f) == ["Alpha", "Beta"],
         note="a new file, checked by reading back what it wrote"),
    Case("create_workbook", "orders", {"path": "fresh.xlsx"},
         refuses="already exists",
         note="refuses to clobber without overwrite"),
    Case("add_sheet", "orders", {"title": "Extra"},
         expect=lambda f: sheets(f) == ["Sheet1", "Orders", "Extra"]),
    Case("add_sheet", "orders", {"title": "Front", "index": 0},
         expect=lambda f: sheets(f) == ["Front", "Sheet1", "Orders"],
         note="index inserts rather than appends"),
    Case("remove_sheet", "orders", {"sheet": "Sheet1"},
         expect=lambda f: sheets(f) == ["Orders"]),
    Case("remove_sheet", "orders", {"sheet": "Orders"},
         setup=[("remove_sheet", {"sheet": "Sheet1"})],
         refuses="at least one sheet",
         note="emptying a workbook is refused, as openpyxl refuses it"),
    # A removal followed by an ordinary edit is what turns a stale relationship into a sheet in
    # the file: the phantom is in the model by then, so the next save writes it out with a copy of
    # its neighbour's cells. The two cases here are the two halves -- what the server reports, and
    # what is actually on disk.
    Case("list_sheets", "orders", {},
         setup=[("remove_sheet", {"sheet": "Sheet1"})],
         expect=lambda f: sheets(f) == ["Orders"],
         check=lambda d: [s["name"] for s in d["sheets"]] == ["Orders"],
         note="after a removal the report matches the file, with no invented sheet"),
    Case("set_cell", "orders", {"sheet": "Orders", "cell": "Z9", "value": "edited"},
         setup=[("remove_sheet", {"sheet": "Sheet1"})],
         expect=lambda f: sheets(f) == ["Orders"] and cells(f, "Orders", "Z9") == "edited",
         note="an edit after a removal does not add a sheet back"),
    Case("rename_sheet", "orders", {"sheet": "Orders", "title": "Sales"},
         expect=lambda f: sheets(f) == ["Sheet1", "Sales"]
         and cells(f, "Sales", "A1") == "Item",
         note="the real sheet keeps its position and its data"),

    Case("merge_cells", "orders", {"sheet": "Orders", "range": "A8:C8"},
         expect=lambda f: merged(f, "Orders") == ["A6:D6", "A8:C8"],
         note="added alongside the merge the real file already had"),
    Case("unmerge_cells", "orders", {"sheet": "Orders", "range": "A6:D6"},
         expect=lambda f: merged(f, "Orders") == [],
         note="the real merged range is unmerged"),
    Case("unmerge_cells", "orders", {"sheet": "Orders", "range": "H1:J1"},
         refuses="not known as merged",
         note="unmerging a range that is not merged is an error"),

    Case("freeze_panes", "orders", {"sheet": "Orders", "cell": "C3"},
         expect=lambda f: load(f)["Orders"].freeze_panes == "C3",
         note="moves the freeze the real file had at A2"),
    Case("freeze_panes", "orders", {"sheet": "Orders"},
         expect=lambda f: load(f)["Orders"].freeze_panes in (None, "A1"),
         note="no cell unfreezes"),

    Case("set_auto_filter", "orders", {"sheet": "Orders", "range": "A1:D4"},
         expect=lambda f: load(f)["Orders"].auto_filter.ref == "A1:D4"),
    Case("set_auto_filter", "orders", {"sheet": "Orders"},
         expect=lambda f: load(f)["Orders"].auto_filter.ref is None,
         note="omitting the range clears it"),

    Case("add_named_range", "orders", {"name": "Qtys", "sheet": "Orders",
                                       "range": "B2:B4"},
         expect=lambda f: load(f).defined_names["Qtys"].value == "'Orders'!$B$2:$B$4",
         note="workbook-scoped, absolute"),
    Case("add_named_range", "orders", {"name": "Local", "sheet": "Orders",
                                       "range": "C2:C4", "scope": "Orders"},
         expect=lambda f: "Local" in load(f)["Orders"].defined_names
         and "Local" not in load(f).defined_names,
         note="a sheet-scoped name lands on the sheet, not the workbook"),

    # -- Values -------------------------------------------------------------------------
    Case("set_cell", "orders", {"sheet": "Orders", "cell": "B2", "value": 999},
         expect=lambda f: cells(f, "Orders", "B2") == 999),
    Case("set_cell", "orders", {"sheet": "Orders", "cell": "A2", "value": "=1+1"},
         expect=lambda f: cells(f, "Orders", "A2") == "=1+1",
         note="a leading = is stored as a formula"),
    Case("set_cell", "orders", {"sheet": "Orders", "cell": "A2",
                                "value": "2010-01-18"},
         expect=lambda f: str(cells(f, "Orders", "A2")) == "2010-01-18 00:00:00"),
    Case("set_cell", "orders", {"sheet": "Orders", "cell": "A2", "value": "50%"},
         expect=lambda f: cells(f, "Orders", "A2") == 0.5),
    Case("set_cell", "orders", {"sheet": "Orders", "cell": "A2", "value": None},
         expect=lambda f: cells(f, "Orders", "A2") is None),
    Case("set_cell", "orders", {"sheet": "Orders", "cell": "A2", "value": True},
         expect=lambda f: cells(f, "Orders", "A2") is True),
    Case("set_cell", "orders", {"sheet": "Orders", "cell": "A2", "value": "42"},
         note="a numeric string stays text, as openpyxl writes it",
         expect=lambda f: cells(f, "Orders", "A2") == "42"),

    Case("write_cells", "orders", {"sheet": "Orders", "start": "A8",
                                   "rows": [["Washer", 5], ["Screw", 100]]},
         expect=lambda f: [[c.value for c in row] for row in
                           load(f)["Orders"]["A8:B9"]] == [["Washer", 5], ["Screw", 100]],
         note="a 2x2 block written left to right, rows outward"),
    Case("append_row", "orders", {"sheet": "Orders", "values": ["Cog", 7, 0.5]},
         expect=lambda f: cells(f, "Orders", "A7") == "Cog"
         and cells(f, "Orders", "B7") == 7
         and cells(f, "Orders", "A6") == "Two line items",
         note="lands below the real last row, leaving the A6 note alone"),
    Case("clear_cells", "orders", {"sheet": "Orders", "range": "B2:B4"},
         expect=lambda f: [cells(f, "Orders", f"B{r}") for r in (2, 3, 4)] == [None] * 3
         and cells(f, "Orders", "A2") == "Bolt"
         and load(f)["Orders"]["B2"].number_format == "0.00",
         note="values go, neighbours and number formats stay"),

    # -- Layout -------------------------------------------------------------------------
    Case("set_column_width", "orders", {"sheet": "Orders", "columns": "B:D", "width": 22.0},
         expect=lambda f: all(load(f)["Orders"].column_dimensions[c].width == 22.0
                              for c in "BCD")),
    Case("set_column_width", "orders", {"sheet": "Orders", "columns": "Z:AB",
                                        "width": 25.0},
         expect=lambda f: all(load(f)["Orders"].column_dimensions[c].width == 25.0
                              for c in ("Z", "AA", "AB")),
         note="empty columns take a width"),
    Case("set_column_width", "orders", {"sheet": "Orders", "columns": "Z:AB",
                                        "width": 25.0},
         expect=lambda f: col_order(f, "Orders") == sorted(col_order(f, "Orders")),
         note="<col> elements come out in ascending order"),
    Case("set_row_height", "orders", {"sheet": "Orders", "rows": "2:4", "height": 30.0},
         expect=lambda f: all(load(f)["Orders"].row_dimensions[r].height == 30.0
                              for r in (2, 3, 4)),
         note="a row with no cells still takes a height"),
    Case("set_row_height", "orders", {"sheet": "Orders", "rows": "9", "height": 40.0},
         expect=lambda f: row_heights(f, "Orders").get(9) == 40.0,
         note="a wholly empty row still takes a height"),
    Case("set_row_height", "orders", {"sheet": "Orders", "rows": "1:2", "height": 30.0},
         expect=lambda f: all(load(f)["Orders"].row_dimensions[r].height == 30.0
                              for r in (1, 2)),
         note="a span of rows that all hold data"),
    Case("set_header_footer", "orders", {"sheet": "Orders", "left_header": "ACME",
                                         "right_footer": "Page &P of &N"},
         expect=lambda f: load(f)["Orders"].oddHeader.left.text == "ACME"
         and "&P" in load(f)["Orders"].oddFooter.right.text),
    Case("add_hyperlink", "orders", {"sheet": "Orders", "cell": "A7",
                                     "target": "https://example.com/orders"},
         expect=lambda f: load(f)["Orders"]["A7"].hyperlink.target
         == "https://example.com/orders"
         and cells(f, "Orders", "A7") == "https://example.com/orders",
         note="an empty cell is given the target as its display text"),
    Case("add_hyperlink", "orders", {"sheet": "Orders", "cell": "A1",
                                     "target": "https://example.com"},
         expect=lambda f: load(f)["Orders"]["A1"].hyperlink.target == "https://example.com"
         and cells(f, "Orders", "A1") == "Item",
         note="a cell that already has text keeps it"),

    # -- Formatting ---------------------------------------------------------------------
    Case("style_cells", "orders", {"sheet": "Orders", "range": "A1:D1",
                                   "style": {"bold": True, "font_color": "FFFFFFFF",
                                             "fill_color": "FF1F3864",
                                             "horizontal": "center",
                                             "border": "thin"}},
         expect=lambda f: load(f)["Orders"]["A1"].font.b
         and load(f)["Orders"]["A1"].font.color.rgb == "FFFFFFFF"
         and load(f)["Orders"]["A1"].fill.fgColor.rgb == "FF1F3864"
         and load(f)["Orders"]["A1"].alignment.horizontal == "center"
         and load(f)["Orders"]["A1"].border.left.style == "thin",
         note="five style keys applied to the real header row"),
    Case("style_cells", "orders", {"sheet": "Orders", "range": "A2:D4",
                                   "style": {"italic": True, "font_size": 14,
                                             "font_name": "Georgia", "wrap_text": True,
                                             "text_rotation": 45, "indent": 2,
                                             "vertical": "top"}},
         expect=lambda f: load(f)["Orders"]["A2"].font.i
         and load(f)["Orders"]["A2"].font.sz == 14
         and load(f)["Orders"]["A2"].font.name == "Georgia"
         and load(f)["Orders"]["A2"].alignment.wrap_text
         and load(f)["Orders"]["A2"].alignment.text_rotation == 45
         and load(f)["Orders"]["A2"].alignment.indent == 2),
    Case("style_cells", "orders", {"sheet": "Orders", "range": "A1:D1",
                                   "style": {"nonsense": True}},
         refuses="not a style property",
         note="a key the schema does not list is refused, not silently ignored"),

    Case("set_number_format", "orders", {"sheet": "Orders", "range": "B2:B4",
                                         "format": "#,##0.00"},
         expect=lambda f: load(f)["Orders"]["B2"].number_format == "#,##0.00"),
    Case("set_number_format", "orders", {"sheet": "Orders", "range": "C2:C4",
                                         "format": "yyyy-mm-dd"},
         expect=lambda f: load(f)["Orders"]["C2"].number_format == "yyyy-mm-dd"
         and stored_value(f, "Orders", "C2") == "1.5"
         and stored_value(f, "Orders", "C3") == "0.75",
         note="a date format over a plain number: the stored value is untouched"),

    Case("add_data_validation", "orders", {"sheet": "Orders", "range": "E2:E4",
                                           "type": "list", "formula1": '"Open,Closed"',
                                           "error_message": "Pick one"},
         expect=lambda f: any(dv.type == "list" and "Open" in str(dv.formula1)
                              and dv.error == "Pick one"
                              for dv in load(f)["Orders"].data_validations.dataValidation)),
    Case("add_data_validation", "orders", {"sheet": "Orders", "range": "F2:F4",
                                           "type": "whole", "formula1": "1",
                                           "formula2": "100", "operator": "between",
                                           "allow_blank": False},
         expect=lambda f: any(dv.type == "whole" and dv.operator == "between"
                              and not dv.allow_blank
                              for dv in load(f)["Orders"].data_validations.dataValidation)),
    Case("add_data_validation", "orders", {"sheet": "Orders", "range": "G2:G4",
                                           "type": "custom",
                                           "formula1": "=ISNUMBER(B2)"},
         expect=lambda f: any(dv.type == "custom"
                              for dv in load(f)["Orders"].data_validations.dataValidation)),

    Case("add_conditional_format", "orders", {"sheet": "Orders", "range": "B2:B4",
                                              "kind": "cellIs", "operator": "greaterThan",
                                              "formula": "20", "fill_color": "FFFFC7CE",
                                              "font_color": "FF9C0006", "bold": True},
         expect=lambda f: rule_types(f, "Orders") == ["cellIs"]),
    Case("add_conditional_format", "orders", {"sheet": "Orders", "range": "B2:B4",
                                              "kind": "formula",
                                              "formula": "=$B2>20"},
         expect=lambda f: rule_types(f, "Orders") == ["expression"]),
    Case("add_conditional_format", "orders", {"sheet": "Orders", "range": "B2:B4",
                                              "kind": "colorScale",
                                              "start_color": "FFF8696B",
                                              "mid_color": "FFFFEB84",
                                              "end_color": "FF63BE7B"},
         expect=lambda f: rule_types(f, "Orders") == ["colorScale"]),
    Case("add_conditional_format", "condfmt", {"sheet": "Sheet1", "range": "A1:A5",
                                               "kind": "cellIs", "operator": "between",
                                               "formula": "1", "second_formula": "3"},
         expect=lambda f: len(list(load(f)["Sheet1"].conditional_formatting)) == 31,
         note="a rule added to a sheet that already has 30"),
    Case("add_conditional_format", "orders", {"sheet": "Orders", "range": "B2:B4",
                                              "kind": "cellIs"},
         refuses="needs an operator",
         note="a cellIs rule without an operator names the missing argument"),

    # -- A second write after a chart or an image ---------------------------------------
    # The defect these cover was invisible to the cases above: each added its drawing as the
    # only call on a fresh copy, so nothing ever wrote to the workbook afterwards. The tool
    # reported success, and the next call deleted the drawing part -- leaving the sheet's
    # relationship naming a part that was no longer there, which is a file no reader can open.
    Case("add_chart", "orders",
         {"sheet": "Orders", "type": "bar", "anchor": "G2",
          "categories": "A2:A4",
          "series": [{"name": "Qty", "values": "B2:B4"}]},
         setup=[("set_cell", {"sheet": "Orders", "cell": "Z9", "value": 1})],
         expect=lambda f: "xl/charts/chart1.xml" in zip_names(f)
         and "xl/drawings/drawing1.xml" in zip_names(f)
         and no_dangling(f),
         note="a chart survives an unrelated write"),
    Case("add_image", "orders",
         {"sheet": "Orders", "image_path": "pixel.png", "anchor": "H2"},
         setup=[("set_cell", {"sheet": "Orders", "cell": "Z9", "value": 1})],
         expect=lambda f: any(n.startswith("xl/media/") for n in zip_names(f))
         and "xl/drawings/drawing1.xml" in zip_names(f)
         and no_dangling(f),
         note="an image survives an unrelated write"),
    Case("add_chart", "orders",
         {"sheet": "Orders", "type": "bar", "anchor": "G20",
          "categories": "A2:A4",
          "series": [{"name": "Qty", "values": "B2:B4"}]},
         setup=[("set_cell", {"sheet": "Orders", "cell": "Z9", "value": 1}),
                ("set_cell", {"sheet": "Orders", "cell": "Z8", "value": 2}),
                ("set_cell", {"sheet": "Orders", "cell": "Z7", "value": 3})],
         expect=lambda f: "xl/drawings/drawing1.xml" in zip_names(f)
         and no_dangling(f),
         note="survives three further writes"),
    Case("add_image", "orders",
         {"sheet": "Sheet1", "image_path": "pixel.png", "anchor": "A1"},
         setup=[("add_chart", {"sheet": "Orders", "type": "bar", "anchor": "G2",
                               "categories": "A2:A4",
                               "series": [{"name": "Qty", "values": "B2:B4"}]})],
         expect=lambda f: len([n for n in zip_names(f)
                               if re.fullmatch(r"xl/drawings/drawing\d+\.xml", n)]) >= 1
         and no_dangling(f),
         note="a chart on one sheet and an image on another do not collide"),
    Case("add_image", "orders",
         {"sheet": "Orders", "image_path": "pixel.png", "anchor": "H2"},
         setup=[("add_comment", {"sheet": "Orders", "cell": "Z1", "text": "a note"}),
                ("set_cell", {"sheet": "Orders", "cell": "Z9", "value": 1})],
         expect=lambda f: "xl/drawings/drawing1.xml" in zip_names(f)
         and "xl/drawings/commentsDrawing1.vml" in zip_names(f)
         and no_dangling(f),
         note="an image and a comment coexist"),
    Case("set_cell", "images",
         {"sheet": "Sheet1", "cell": "B5", "value": "hello"},
         expect=lambda f: len(load(f)["Sheet1"]._images) == 3
         and "xl/drawings/drawing1.xml" in zip_names(f)
         and no_dangling(f),
         note="EDITING A REAL FILE THAT ALREADY HAS IMAGES: all three survive"),
    # -- Text that XML has to escape ------------------------------------------------
    # `&`, `<`, `>`, `"` and `'` are written as entity references. The reader was handed each
    # one as a separate event and discarded them, so a cell holding any of these came back empty
    # after a save. The checks read the file back rather than trusting what the tool reported.
    Case("set_cell", "orders", {"sheet": "Orders", "cell": "A10", "value": "Tom & Jerry"},
         expect=lambda f: cells(f, "Orders", "A10") == "Tom & Jerry",
         note="an ampersand survives a round trip"),
    Case("write_cells", "orders",
         {"sheet": "Orders", "start": "A11",
          "rows": [["a < b"], ["c > d"], ['say "hi"'], ["it's"], ["R&D / P&L"]]},
         expect=lambda f: [cells(f, "Orders", f"A{row}") for row in range(11, 16)]
         == ["a < b", "c > d", 'say "hi"', "it's", "R&D / P&L"],
         note="angle brackets, quotes and apostrophes all survive"),
    Case("set_cell", "orders", {"sheet": "Orders", "cell": "A16",
                                "value": "<b>bold</b> & <i>it</i>"},
         expect=lambda f: cells(f, "Orders", "A16") == "<b>bold</b> & <i>it</i>",
         note="markup-looking text is text, not markup"),
    Case("set_header_footer", "orders",
         {"sheet": "Orders", "center_header": "R&D report",
          "right_footer": "Page &P of &N"},
         setup=[("set_cell", {"sheet": "Orders", "cell": "Z9", "value": 1})],
         expect=lambda f: load(f)["Orders"].oddHeader.center.text == "R&D report"
         and "&P" in (load(f)["Orders"].oddFooter.right.text or ""),
         note="a header set, then kept through a further write"),

    # -- Media --------------------------------------------------------------------------
    Case("add_chart", "orders", {"sheet": "Orders", "type": "bar", "anchor": "F2",
                                 "title": "Qty by item", "categories": "A2:A4",
                                 "series": [{"name": "Qty", "values": "B2:B4"}]},
         expect=lambda f: "xl/charts/chart1.xml" in zip_names(f),
         note="the chart part is in the zip; openpyxl cannot read charts back"),
    Case("add_chart", "orders", {"sheet": "Orders", "type": "line", "anchor": "F20",
                                 "categories": "A2:A4",
                                 "series": [{"name": "Qty", "values": "B2:B4"},
                                            {"name": "Price", "values": "C2:C4"}]},
         expect=lambda f: "xl/charts/chart1.xml" in zip_names(f),
         note="two series"),
    Case("add_chart", "orders", {"sheet": "Orders", "type": "pie", "anchor": "F38",
                                 "categories": "A2:A4",
                                 "series": [{"values": "B2:B4"}]},
         expect=lambda f: "xl/charts/chart1.xml" in zip_names(f),
         note="a series with no literal name"),
    Case("add_chart", "orders", {"sheet": "Orders", "type": "sparkline", "anchor": "F2",
                                 "series": [{"values": "B2:B4"}]},
         refuses="not a chart type",
         note="a type outside the enum is refused, naming the valid ones"),

    Case("add_image", "orders", {"sheet": "Orders", "image_path": "pixel.png",
                                 "anchor": "H2"},
         expect=lambda f: any(n.startswith("xl/media/") for n in zip_names(f))
         and len(load(f)["Orders"]._images) == 1,
         note="a real PNG embedded and read back by openpyxl"),
    Case("add_image", "orders", {"sheet": "Orders", "image_path": "nope.png",
                                 "anchor": "H2"},
         refuses="could not read",
         note="a missing image is an error the model can act on"),

    Case("add_comment", "comments", {"sheet": "Sheet1", "cell": "A1",
                                     "text": "Overwritten", "author": "tester"},
         expect=lambda f: load(f)["Sheet1"]["A1"].comment is not None
         and load(f)["Sheet1"]["A1"].comment.text.endswith("Overwritten")
         and load(f)["Sheet1"]["A1"].comment.author == "tester",
         note="replaces the real comment that was there"),
    Case("add_comment", "comments", {"sheet": "Sheet2", "cell": "A1",
                                     "text": "New note on a sheet that had none"},
         expect=lambda f: load(f)["Sheet2"]["A1"].comment is not None,
         note="a sheet that had none gains one"),
    Case("add_comment", "comments", {"sheet": "Sheet1", "cell": "A1", "text": ""},
         expect=lambda f: load(f)["Sheet1"]["A1"].comment is None
         and load(f)["Sheet1"]["A2"].comment is not None,
         note="an empty text removes that one comment and leaves the others"),
    Case("list_comments", "comments", {"sheet": "Sheet1"},
         setup=[("add_comment", {"sheet": "Sheet1", "cell": "E5", "text": "added",
                                 "author": "tester"})],
         expect=lambda f: load(f)["Sheet1"]["E5"].comment is not None,
         check=lambda d: len(d["comments"]) == 7
         and any(c["cell"] == "E5" for c in d["comments"]),
         note="a comment added to a real file is listed back"),
]


def text_of(result):
    return " ".join(p.get("text", "") for p in result.get("content", []))


def complaint_of(result):
    """The message a call reports, or None when it succeeded."""
    if "rpcError" in result:
        return f"rpcError {result['rpcError']}"
    if result.get("isError"):
        return text_of(result)
    return None


def main() -> int:
    shutil.rmtree(WORK, ignore_errors=True)
    WORK.mkdir(parents=True)
    for name, src in SOURCES.items():
        if not src.is_file():
            # Say where it looked and how to move it, rather than naming one path: on CI the
            # fixtures come from the openpyxl checkout the workflow fetches.
            print(f"missing fixture {name}: {src}")
            print(f"  the openpyxl checkout is taken from FERROXL_OPENPYXL, "
                  f"currently {OPENPYXL}")
            return 2
    (WORK / "pixel.png").write_bytes(PNG)

    seen, failures, notes, crashes = [], [], [], []
    made = 0

    # One long-lived server is how an agent talks to it, and most of the value of this
    # harness is that it does. But the environment this runs under will not keep a child
    # process alive indefinitely -- it ends one after a few hundred calls, silently, and
    # then refuses to start another. So the server is recycled every RECYCLE calls: often
    # enough that the run finishes, rarely enough that consecutive calls still share a
    # process, which is the case where a tool could corrupt state for the next one.
    RECYCLE = 12

    server = Server(WORK)
    made += 1
    try:
        server.initialize()
        catalogue = [t["name"] for t in server.list_tools()]

        def restart(quiet=False):
            """Bring the server back after a crash, so one death ends one case.

            A restart that itself fails is reported rather than raised: losing the rest of
            the run would be worse than reporting the cases that never got to answer.
            """
            nonlocal server, made
            try:
                server.close()
            except Exception:
                pass
            try:
                server = Server(WORK)
                made += 1
                server.initialize()
                return True
            except Exception as exc:  # noqa: BLE001
                if not quiet:
                    print(f"note  the server could not be restarted: {exc}")
                return False

        def call_or_die(name, arguments):
            """One tool call. Returns (result, None) or (None, why-it-died)."""
            nonlocal made
            try:
                result = server.call(name, **arguments)
            except Exception as exc:  # noqa: BLE001
                return None, f"{name}: {type(exc).__name__}: {exc}"
            made += 1
            if made % RECYCLE == 0:
                restart(quiet=True)
            return result, None

        for case in CASES:
            if case.tool not in seen:
                seen.append(case.tool)
            # Every case works on a file of its own, named after the tool.
            target = f"{case.tool}{case.extension}"
            shutil.copy(fixture_path(case.fixture), WORK / target)

            # `path` is the tool's own argument, so a case that names its own file (only
            # create_workbook does) keeps it; everything else writes to its own copy.
            calls = [(name, {**args, "path": args.get("path", target)})
                     for name, args in case.setup]
            calls.append((case.tool,
                          {**case.args, "path": case.args.get("path", target)}))
            result, died = None, None
            for name, arguments in calls:
                result, died = call_or_die(name, arguments)
                if died:
                    break

            label = f"{case.tool:24s} {case.fixture:8s}"
            if died:
                # A server that dies takes the rest of the run with it, so the process is
                # restarted. A crash is a finding about that tool, not a reason to stop.
                # Whether the binary is still on disk separates a fault in the tool from
                # the environment taking the executable away underneath it.
                still_there = "binary still on disk"
                try:
                    from mcp_client import find_binary
                    still_there = ("binary still on disk" if find_binary().is_file()
                                   else "BINARY GONE FROM DISK")
                except SystemExit:
                    still_there = "BINARY GONE FROM DISK"
                crashes.append((case.tool, died))
                failures.append((case.tool, f"the server died -- {died} [{still_there}]"))
                print(f"CRASH {label} {died[:100]} [{still_there}]")
                restart()
                continue

            complaint = complaint_of(result)

            # A call meant to fail is a pass only when it says the useful thing.
            if case.refuses:
                if complaint and case.refuses in complaint:
                    print(f"pass {label} refused: {complaint[:110]}")
                    if case.note:
                        notes.append(f"{case.tool}: {case.note}")
                    continue
                if complaint:
                    detail = f"refused, but not with {case.refuses!r}: {complaint}"
                else:
                    detail = f"should have refused with {case.refuses!r}, but it succeeded"
                failures.append((case.tool, detail))
                print(f"FAIL {label} {detail[:140]}")
                continue

            if complaint:
                failures.append((case.tool, complaint))
                print(f"FAIL {label} {complaint[:140]}")
                continue

            # `structuredContent` first: it is what a model reads before the file.
            checked, detail = True, ""
            if case.check:
                try:
                    checked = case.check(result.get("structuredContent"))
                except Exception as exc:  # noqa: BLE001
                    checked, detail = False, f"[data check raised {exc!r}]"
            written = case.verify_as or target
            if checked and case.expect:
                try:
                    checked = case.expect(written)
                except Exception as exc:  # noqa: BLE001
                    checked = False
                    detail = f"[readback raised {type(exc).__name__}: {exc}]"

            print(f"{'pass' if checked else 'FAIL'} {label} {(case.note or '')[:90]}"
                  f" {detail}")
            if checked and case.note:
                notes.append(f"{case.tool}: {case.note}")
            if not checked:
                failures.append((case.tool, f"{case.note or 'readback disagreed'} {detail}"))
    finally:
        server.close()

    missing = [t for t in catalogue if t not in seen]
    print()
    print(f"{len(CASES)} calls over {len(seen)}/{len(catalogue)} tools, "
          f"{len(failures)} failed, {len(crashes)} crashed the server")
    if missing:
        print(f"NOT COVERED: {', '.join(missing)}")
    for line in notes:
        print(f"note  {line}")
    for tool, why in failures:
        print(f"\nFAIL {tool}: {why}")
    return 1 if failures or missing else 0


if __name__ == "__main__":
    sys.exit(main())