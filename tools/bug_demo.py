"""Minimal demonstration of the two writer defects `mcp_real_files.py` found.

Run in the same shell invocation as the build, because the built binary does not
survive long in this environment:

    cargo build -p ferroxl-mcp
    python tools/bug_demo.py
"""

import re
import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from mcp_client import Server  # noqa: E402

import openpyxl  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT / "target" / "bug_demo"
SRC = ROOT / "orders.xlsx"

# orders.xlsx holds data in rows 1, 2, 3 and 6. Rows 4 and 5 are entirely empty, which
# is what makes it the right fixture: a row is present enough to be named, and absent
# enough that nothing else will emit it.


def sheet_xml(path, title="Orders"):
    """The raw XML of one sheet's part, located through the package relationships."""
    with zipfile.ZipFile(path) as zf:
        book = zf.read("xl/workbook.xml").decode()
        rels = zf.read("xl/_rels/workbook.xml.rels").decode()
        targets = {}
        for tag in re.findall(r"<(?:\w+:)?Relationship [^>]*/?>", rels):
            rid_attr = re.search(r'Id="([^"]+)"', tag)
            target_attr = re.search(r'Target="([^"]+)"', tag)
            if rid_attr and target_attr:
                targets[rid_attr.group(1)] = target_attr.group(1)
        rid = None
        for tag in re.findall(r"<(?:\w+:)?sheet [^>]*/?>", book):
            if f'name="{title}"' in tag:
                rid = re.search(r'r:id="([^"]+)"', tag).group(1)
                break
        if rid is None:
            raise LookupError(f"no sheet named {title!r} in {path.name}")
        # A Target may be package-absolute ("/xl/worksheets/sheet1.xml", what openpyxl
        # writes) or relative to xl/ ("worksheets/sheet1.xml", what ferroxl writes).
        target = targets[rid]
        part = target if target.startswith("/") else "xl/" + target
        return zf.read(part.lstrip("/")).decode()


def row_heights(path):
    out = {}
    for tag in re.findall(r"<row [^>]*/?>", sheet_xml(path)):
        row = re.search(r'r="(\d+)"', tag)
        height = re.search(r'ht="([\d.]+)"', tag)
        if row:
            out[int(row.group(1))] = height.group(1) if height else None
    return out


def col_mins(path):
    """The `min` of every <col>, in the order they appear.

    Attribute order differs between the two writers, so each tag is parsed on its own
    rather than matched with one pattern.
    """
    out = []
    for tag in re.findall(r"<col [^>]*/?>", sheet_xml(path)):
        found = re.search(r'min="(\d+)"', tag)
        if found:
            out.append(int(found.group(1)))
    return out


shutil.rmtree(WORK, ignore_errors=True)
WORK.mkdir(parents=True)

print("=" * 74)
print("DEFECT 1   set_row_height silently loses any row that holds no cells")
print("=" * 74)
with Server(WORK) as s:
    s.initialize()
    for rows, asked in (("2:4", [2, 3, 4]), ("9", [9])):
        name = f"rows_{rows.replace(':', '_')}.xlsx"
        shutil.copy(SRC, WORK / name)
        s.call("set_row_height", path=name, sheet="Orders", rows=rows, height=30.0)
        written = row_heights(WORK / name)
        print(f"\n  set_row_height rows={rows!r} height=30")
        print(f"    reported by the tool : rows {asked}, height 30")
        print(f"    rows in the file     : {written}")
        for row in asked:
            got = written.get(row)
            print(f"      row {row}: {got if got else 'ABSENT -- the height is gone'}")

print()
print("=" * 74)
print("DEFECT 2   set_column_width writes <col> elements out of column order")
print("=" * 74)
name = "cols.xlsx"
shutil.copy(SRC, WORK / name)
with Server(WORK) as s:
    s.initialize()
    s.call("set_column_width", path=name, sheet="Orders", columns="Z:AB", width=25.0)

ferroxl_order = col_mins(WORK / name)

wb = openpyxl.load_workbook(SRC)
for letter in ("Z", "AA", "AB"):
    wb["Orders"].column_dimensions[letter].width = 25.0
wb.save(WORK / "opx_reference.xlsx")
openpyxl_order = col_mins(WORK / "opx_reference.xlsx")

print("\n  set_column_width columns='Z:AB' width=25, then reading the sheet XML back")
print(f"    ferroxl  <col> min order : {ferroxl_order}")
print(f"    openpyxl <col> min order : {openpyxl_order}")
print(f"    ferroxl  ascending       : {ferroxl_order == sorted(ferroxl_order)}")
print(f"    openpyxl ascending       : {openpyxl_order == sorted(openpyxl_order)}")
print("\n  The widths themselves are correct and openpyxl reads all three back; what")
print("  differs is the order the parts appear in.")
print(f"    openpyxl reads back      : "
      f"{ {c: openpyxl.load_workbook(WORK / name)['Orders'].column_dimensions[c].width for c in 'Z AA AB'.split()} }")