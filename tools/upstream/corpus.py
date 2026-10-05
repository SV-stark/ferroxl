"""Every real workbook openpyxl ships, read by both implementations and compared cell by cell.

openpyxl's `tests/data/genuine/` holds files produced by Excel, LibreOffice and Mac Excel --
not files written by openpyxl. They exist precisely because they contain combinations nobody
imagined, which is the class of file that found all five of ferroxl's silent-loss bugs. This
runs openpyxl and ferroxl over every one of them and asserts they agree.

openpyxl is the oracle and ferroxl the system under test: every value here was produced by
the reference implementation, not by this project, so a failure means a divergence rather than
a disagreement about what the answer should be.

    python tools/upstream/corpus.py path/to/openpyxl
"""

import shutil
import sys
import traceback
import warnings
from datetime import date, datetime, time, timedelta
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from mcp_client import Server  # noqa: E402

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

WORK = Path(__file__).resolve().parents[2] / "target" / "upstream_corpus"

# Reading a 400 kB fixture cell by cell through a JSON-RPC round trip is not a comparison,
# it is a timeout. The cap is reported per sheet so a truncated row is never mistaken for a
# clean one.
MAX_ROWS = 500
MAX_COLUMNS = 60

# openpyxl's `reader/bigfoot.xlsx` has 1024 sheets on purpose. Every MCP call reloads the
# whole workbook, so sweeping every sheet of it costs 1024 full loads -- which is a real
# characteristic of the server, recorded in tools/mcp_limits.py, and not something this
# comparison should pay for. The genuine corpus is far smaller.
MAX_SHEETS = 40

RESULTS = []


def check(name, ok, detail=""):
    RESULTS.append((name, bool(ok)))
    mark = "ok   " if ok else "FAIL "
    print(f"{mark}{name}" + (f"\n        {detail}" if detail and not ok else ""))


def genuine(root: Path) -> list[Path]:
    """The real-world workbooks from a checkout's `tests/data/genuine`.

    `root` is the checkout, not the package, so the fixtures live under `root/openpyxl/`.
    """
    base = root / "openpyxl" / "tests" / "data" / "genuine"
    return sorted(p for p in base.glob("*") if p.suffix.lower() in (".xlsx", ".xlsm"))


def normalise(value):
    """A cell value reduced to something comparable across the two implementations.

    Both sides are put through the *same* normalisation, which is the whole point: comparing
    openpyxl's `datetime` against a server string is a comparison of two representations, not
    of two values. A leading `=` is a formula on both sides -- openpyxl without `data_only`
    hands back the formula as a plain string, so it has to be classified here or every
    formula in every fixture reads as a text divergence.
    """
    if isinstance(value, str) and value.startswith("="):
        return ("formula", value)
    if isinstance(value, datetime):
        return ("datetime", value.replace(tzinfo=None).isoformat(sep=" "))
    if isinstance(value, date):
        return ("date", value.isoformat())
    if isinstance(value, time):
        return ("time", value.isoformat())
    if isinstance(value, timedelta):
        return ("timedelta", value.total_seconds())
    if isinstance(value, bool):
        return ("bool", value)
    if isinstance(value, (int, float)):
        # An int in one and a float in the other is not a divergence, it is the same number.
        return ("number", float(value))
    if isinstance(value, str):
        return ("text", value)
    return ("other", None if value is None else str(value))


def server_view(server: Server, name: str, sheet: str, rows: int, columns: int):
    """The server's view of a rectangle, put through the same normalisation as openpyxl's.

    The server renders a temporal as `YYYY-MM-DD HH:MM:SS` with a space, where Python's
    `isoformat` uses a `T`. Left alone that is a difference in representation rather than in
    value, so the string is parsed back into a `datetime` and normalised like the other side.
    """
    end = openpyxl.utils.get_column_letter(min(columns, MAX_COLUMNS))
    got = server.data(
        "read_cells",
        path=name,
        sheet=sheet,
        range=f"A1:{end}{min(rows, MAX_ROWS)}",
        include_empty=False,
    )
    out = {}
    for cell in got["cells"]:
        coordinate = cell["cell"]
        row = int("".join(ch for ch in coordinate if ch.isdigit()))
        letters = "".join(ch for ch in coordinate if ch.isalpha())
        column = openpyxl.utils.column_index_from_string(letters)
        value = cell["value"]
        if not isinstance(value, dict):
            out[(row, column)] = normalise(value)
            continue
        kind, raw = value.get("type"), value.get("value")
        if kind in ("date", "datetime", "time"):
            text = str(raw).replace(" ", "T", 1)
            # A time-only cell is a bare `HH:MM:SS`, which `fromisoformat` only accepts from
            # 3.11 on, and whose `.isoformat` takes no separator. Branching on the result
            # rather than assuming is what keeps a time cell from being reported as a
            # divergence when it is only a different repr.
            try:
                parsed = datetime.fromisoformat(text)
            except ValueError:
                parsed = None
            if isinstance(parsed, datetime):
                out[(row, column)] = ("datetime", parsed.replace(tzinfo=None).isoformat(sep=" "))
            elif isinstance(parsed, time):
                out[(row, column)] = ("time", parsed.isoformat())
            else:
                try:
                    out[(row, column)] = (
                        "time",
                        time.fromisoformat(str(raw)).isoformat(),
                    )
                except ValueError:
                    out[(row, column)] = ("other", str(raw))
        elif kind == "number":
            out[(row, column)] = ("number", float(raw))
        elif kind == "boolean":
            out[(row, column)] = ("bool", bool(raw))
        elif kind in ("string", "text"):
            out[(row, column)] = ("text", raw)
        elif kind == "formula":
            out[(row, column)] = ("formula", raw)
        else:
            out[(row, column)] = ("other", None if raw is None else str(raw))
    return out


def compare(path: Path, server: Server) -> None:
    label = path.name
    book = openpyxl.load_workbook(path)
    target = WORK / label
    target.write_bytes(path.read_bytes())

    total_cells = 0
    total_diff = 0

    for sheet in book.worksheets[:MAX_SHEETS]:
        rows = min(sheet.max_row or 0, MAX_ROWS)
        columns = min(sheet.max_column or 0, MAX_COLUMNS)
        if rows == 0 or columns == 0:
            continue

        expected = {}
        for row in sheet.iter_rows(max_row=rows, max_col=columns):
            for cell in row:
                if cell.value is not None:
                    expected[(cell.row, cell.column)] = normalise(cell.value)

        try:
            actual = server_view(server, label, sheet.title, rows, columns)
        except AssertionError as error:
            check(f"{label} / {sheet.title}: readable by the server", False, str(error)[:200])
            return

        differences = []
        for coordinate, mine in sorted(expected.items()):
            theirs = actual.get(coordinate)
            if theirs is None:
                differences.append(f"{coordinate}: openpyxl has {mine}, ferroxl has nothing")
                continue
            if mine != theirs:
                differences.append(f"{coordinate}: openpyxl {mine} vs ferroxl {theirs}")
        # And the other direction: something the server reports that openpyxl does not.
        for coordinate, theirs in sorted(actual.items()):
            if coordinate not in expected:
                differences.append(f"{coordinate}: ferroxl has {theirs}, openpyxl has nothing")

        total_cells += len(expected)
        total_diff += len(differences)

        if differences:
            shown = differences[:4]
            more = f" (+{len(differences) - len(shown)} more)" if len(differences) > 4 else ""
            check(
                f"{label} / {sheet.title}: {len(expected)} cells agree",
                False,
                "; ".join(shown) + more,
            )

    if total_diff == 0:
        sheets = min(len(book.worksheets), MAX_SHEETS)
        note = f" (capped from {len(book.worksheets)})" if len(book.worksheets) > MAX_SHEETS else ""
        check(f"{label}: {total_cells} cells across {sheets} sheets agree{note}", True)


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else "../openpyxl").resolve()
    corpus = root / "openpyxl" / "tests" / "data" / "genuine"
    if not corpus.is_dir():
        print(f"no genuine fixture corpus at {corpus}")
        return 2

    shutil.rmtree(WORK, ignore_errors=True)
    WORK.mkdir(parents=True)

    files = genuine(root)
    print(f"{len(files)} genuine workbooks in {corpus}\n")

    with Server(WORK) as server:
        server.initialize()
        for path in files:
            try:
                compare(path, server)
            except Exception:
                check(f"{path.name}: no exception escaped", False, traceback.format_exc(limit=3))

    failed = [name for name, ok in RESULTS if not ok]
    print(f"\n{len(RESULTS) - len(failed)}/{len(RESULTS)} fixture checks passed")

    print(
        "\nnot asserted here:\n"
        f"  - rows past {MAX_ROWS} or columns past {MAX_COLUMNS}, for speed\n"
        "  - formulas' cached values: openpyxl without data_only returns the formula string,\n"
        "    which is compared, but the arithmetic is not\n"
        "  - charts and images: openpyxl 3.1.5's own reader does not return them either, so\n"
        "    there is no oracle for them -- the writer side is covered by tools/mcp_parity.py"
    )
    if failed:
        print("\nfailed:")
        for name in failed:
            print(f"  {name}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())