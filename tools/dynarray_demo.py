"""Show the three-way formula comparison behind Pending item 4 in PARITY.md.

A diagnostic, not an assertion: it prints what each writer stored and asserts nothing, because
the difference it shows is one only Excel can adjudicate. No harness here can fail on it, which
is the point of the record -- see `PARITY.md`, "A dynamic array formula is written as a bare
function name".

Excel's format wants three separate things for a post-2007 function, and only XlsxWriter does
all three: an `_xlfn.` prefix on the stored name, `cm="1"` on the cell, and an
`xl/metadata.xml` part holding the cell-metadata record that `cm` points at.

Needs `xlsxwriter` as well as `openpyxl`; neither is in `tools/requirements.txt`, because this
does not run in CI the way the four harnesses do. A missing one is reported and skipped.
"""

import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from mcp_client import Server  # noqa: E402

# Three shapes of formula, and they are not the same problem:
#   SUM       -- in the file format since 2007, right in all three
#   IFS       -- added in 2016, so the stored name should be _xlfn.IFS; XlsxWriter writes this
#                one bare as well, so no library prefixes it and the caller has to
#   SEQUENCE  -- a dynamic array, so it needs the prefix *and* cm="1" *and* the metadata part,
#                which is the row the three writers actually disagree on
FORMULAS = ["=SUM(B1:B3)", "=IFS(A1>0,\"y\",\"n\")", "=SEQUENCE(3)"]

OUT = Path(__file__).parent.parent / "target" / "dynarray"


def sheet_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/worksheets/sheet1.xml").decode("utf-8", "replace")


def parts(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        return sorted(archive.namelist())


def cell(path: Path, reference: str) -> str:
    xml = sheet_xml(path)
    found = re.search(rf"<c [^>]*r=\"{reference}\".*?</c>|<c [^>]*r=\"{reference}\"[^>]*/>", xml)
    return found.group(0) if found else "<no cell>"


def report(label: str, target: Path) -> None:
    metadata = [p for p in parts(target) if "metadata" in p]
    print(f"  {label}")
    for row, formula in enumerate(FORMULAS):
        print(f"    {formula:<22} {cell(target, f'A{row + 1}')}")
    print(f"    metadata part        {metadata or 'none'}")


def ferroxl() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    target = OUT / "ferroxl.xlsx"
    if target.exists():
        target.unlink()
    with Server(Path(__file__).parent.parent) as server:
        server.ok("create_workbook", path=str(target))
        sheets = server.data("list_sheets", path=str(target))
        first = sheets["sheets"][0] if isinstance(sheets, dict) else sheets[0]
        for row, formula in enumerate(FORMULAS):
            server.ok("set_cell", path=str(target), sheet=first, cell=f"A{row + 1}", value=formula)
    report("ferroxl", target)


def xlsxwriter() -> None:
    import xlsxwriter

    target = OUT / "xlsxwriter.xlsx"
    book = xlsxwriter.Workbook(str(target))
    sheet = book.add_worksheet()
    for row, formula in enumerate(FORMULAS):
        sheet.write_formula(row, 0, formula)
    book.close()
    report("xlsxwriter", target)


def openpyxl_() -> None:
    import openpyxl

    target = OUT / "openpyxl.xlsx"
    book = openpyxl.Workbook()
    for row, formula in enumerate(FORMULAS):
        book.active[f"A{row + 1}"] = formula
    book.save(target)
    report("openpyxl", target)


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    for name, run in (("ferroxl", ferroxl), ("xlsxwriter", xlsxwriter), ("openpyxl", openpyxl_)):
        try:
            run()
        except Exception as error:  # noqa: BLE001
            print(f"{name:10} failed: {error}")