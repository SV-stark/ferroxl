"""What is in openpyxl's fixture corpus, and can ferroxl read it?

A survey, not a test: it loads every fixture openpyxl ships and reports what each side can
make of it. The point is to find the files worth asserting on before writing assertions,
and to be honest about which ones ferroxl cannot read at all.

    python tools/upstream/survey.py path/to/openpyxl
"""

import sys
import traceback
import warnings
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from mcp_client import Server  # noqa: E402

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def fixtures(root: Path) -> list[Path]:
    """Every workbook fixture in a checkout, under `root/openpyxl/tests/data`."""
    data = root / "openpyxl" / "tests" / "data"
    return sorted(
        path
        for path in data.rglob("*")
        if path.is_file()
        and "__pycache__" not in path.parts
        and path.suffix.lower() in (".xlsx", ".xlsm")
    )


# A fixture this large exists to be read, not to be counted cell by cell through an MCP
# round trip, and one of openpyxl's is 400 kB. The survey reports the cap rather than
# quietly truncating, so a truncated row is never mistaken for a clean comparison.
MAX_ROWS = 2000


def python_view(path: Path):
    """What openpyxl makes of a fixture, as plain data."""
    book = openpyxl.load_workbook(path)
    sheets = []
    for sheet in book.worksheets:
        cells = 0
        formulas = 0
        truncated = False
        for row in sheet.iter_rows(max_row=MAX_ROWS):
            for cell in row:
                if cell.value is not None:
                    cells += 1
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formulas += 1
        if sheet.max_row and sheet.max_row > MAX_ROWS:
            truncated = True
        sheets.append(
            {
                "title": sheet.title,
                "cells": cells,
                "formulas": formulas,
                "truncated": truncated,
                "merged": len(sheet.merged_cells.ranges),
                "freeze": sheet.freeze_panes,
                "dimension": sheet.calculate_dimension(),
            }
        )
    return book.sheetnames, sheets


def ferroxl_view(server: Server, path: Path, names):
    """What the MCP server makes of the same fixture."""
    out = []
    for name in names:
        described = server.data("describe_sheet", path=path.name, sheet=name)
        out.append(
            {
                "title": name,
                "dimension": described.get("dimension"),
                "merged": described.get("merged", 0),
                "freeze": described.get("freeze_panes"),
            }
        )
    return out


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else "../openpyxl").resolve()
    if not (root / "openpyxl" / "tests" / "data").is_dir():
        print(f"no fixture corpus at {root / 'openpyxl' / 'tests' / 'data'}")
        print("pass the path to an openpyxl checkout")
        return 2

    work = Path(__file__).resolve().parents[2] / "target" / "upstream_survey"
    work.mkdir(parents=True, exist_ok=True)

    files = fixtures(root)
    print(f"{len(files)} fixtures under {root / 'openpyxl' / 'tests' / 'data'}\n")

    readable = disagreeing = unreadable = 0

    with Server(work) as server:
        server.initialize()

        for path in files:
            try:
                names, sheets = python_view(path)
            except Exception as error:
                print(f"  openpyxl cannot read {path.name}: {type(error).__name__}")
                unreadable += 1
                continue

            target = work / path.name
            target.write_bytes(path.read_bytes())
            try:
                ferroxl = ferroxl_view(server, target, names)
            except Exception as error:
                print(
                    f"  ferroxl cannot read {path.name}: "
                    f"{' '.join(str(p.get('text','')) for p in error.get('content', []))[:90]}"
                )
                unreadable += 1
                continue

            readable += 1
            problems = []
            for name, mine, theirs in zip(names, sheets, ferroxl):
                if mine["dimension"] != theirs["dimension"]:
                    problems.append(
                        f"{name}: dimension {theirs['dimension']} vs {mine['dimension']}"
                    )
                if mine["merged"] != theirs["merged"]:
                    problems.append(f"{name}: merged {theirs['merged']} vs {mine['merged']}")
                if (mine["freeze"] or None) != (theirs["freeze"] or None):
                    problems.append(f"{name}: freeze {theirs['freeze']} vs {mine['freeze']}")

            if problems:
                disagreeing += 1
                print(f"DIFF {path.name}")
                for problem in problems:
                    print(f"       {problem}")
            else:
                capped = sum(1 for s in sheets if s["truncated"])
                note = f"  ({len(names)} sheets" + (f", {capped} capped" if capped else "") + ")"
                print(f"ok   {path.name}{note}")

    print(
        f"\n{readable} readable by both, {disagreeing} disagreeing, {unreadable} unreadable"
    )
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(2)