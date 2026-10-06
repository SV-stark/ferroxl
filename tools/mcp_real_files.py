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

# Real workbooks, each the best honest example of the feature it is used for. Taken from
# the openpyxl test corpus and from this repo, so they are files real software produced.
SOURCES = {
    "orders": ROOT / "orders.xlsx",
    "sample": ROOT.parent / "openpyxl" / "openpyxl" / "reader" / "tests" / "data" / "sample.xlsx",
    "styles": ROOT.parent / "openpyxl" / "openpyxl" / "reader" / "tests" / "data" / "complex-styles.xlsx",
    "condfmt": ROOT.parent / "openpyxl" / "openpyxl" / "formatting" / "tests" / "data" / "conditional-formatting.xlsx",
    "comments": ROOT.parent / "openpyxl" / "openpyxl" / "comments" / "tests" / "data" / "comments.xlsx",
    "images": ROOT.parent / "openpyxl" / "openpyxl" / "reader" / "tests" / "data" / "sample_with_images.xlsx",
    "table": ROOT.parent / "openpyxl" / "openpyxl" / "reader" / "tests" / "data" / "print_area_table_defined_name.xlsx",
    "vba": ROOT.parent / "openpyxl" / "openpyxl" / "tests" / "data" / "reader" / "vba-test.xlsm",
    "bigfoot": ROOT.parent / "openpyxl" / "openpyxl" / "tests" / "data" / "reader" / "bigfoot.xlsx",
}

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
            print(f"missing fixture: {src}")
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
    RECYCLE = 25

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
            shutil.copy(SOURCES[case.fixture], WORK / target)

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