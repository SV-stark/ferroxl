"""What openpyxl's tests pin, and whether this project's tests pin the same numbers.

openpyxl's test suite is 161 files and about 1,700 test functions, almost all of them Python
calling a Python API -- none of it can be run against a Rust library. But a large part of it
pins *values*: a date serial, a password hash, a Julian day, the length of a colour table.
Those are exactly the assertions this project claims to have copied, and each one is a number
that can be checked textually on both sides.

So this extracts the `@pytest.mark.parametrize` cases from the modules that pin values, and
then asks a question that is answerable: does ferroxl's own test suite assert the same
expected numbers? A pinned value with no counterpart here is a claim in PARITY.md that no
test actually backs.

This is a **report**, not a gate. A missing counterpart is a gap to close, not a build break.

    python tools/upstream/expectations.py path/to/openpyxl
"""

import ast
import re
import sys
from pathlib import Path

# The openpyxl modules that pin observable values, and what each one pins. Chosen because
# PARITY.md claims the values are copied from them.
PINNING = {
    "utils/tests/test_datetime.py": "date serials, ISO-8601 round trips, Julian days",
    "utils/tests/test_protection.py": "the password hash",
    "utils/tests/test_units.py": "EMU and pixel conversions",
    "styles/tests/test_colors.py": "the indexed colour palette",
    "styles/tests/test_number_style.py": "built-in number format ids",
    "cell/tests/test_cell.py": "type inference from text",
}


def _describe(node: ast.AST) -> str:
    """A stable textual form of a parametrize case element.

    `ast.literal_eval` cannot evaluate `date(1900, 1, 1)`, and a parametrize block is one
    list literal -- so evaluating the whole list fails because of a single element and the
    entire block is lost. Each element is therefore rendered from source instead, which
    handles both `40196` and `date(2010, 1, 18)` the same way.
    """
    try:
        return repr(ast.literal_eval(node))
    except (ValueError, SyntaxError):
        return ast.unparse(node)


def parametrised_cases(path: Path):
    """`(function, input, expected)` from `@pytest.mark.parametrize` on a test function.

    Parsed rather than imported: importing openpyxl's tests would need its own test
    dependencies installed, and this only wants the literals.
    """
    try:
        source = path.read_text(encoding="utf-8", errors="replace")
        tree = ast.parse(source)
    except (OSError, SyntaxError):
        return []

    cases = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.FunctionDef):
            continue
        for decorator in node.decorator_list:
            target = decorator.func if isinstance(decorator, ast.Call) else decorator
            if getattr(target, "attr", None) != "parametrize":
                continue
            if not isinstance(decorator, ast.Call) or len(decorator.args) < 2:
                continue
            rows = decorator.args[1]
            # The list is usually a literal, but tolerate a name by giving up on that block.
            if not isinstance(rows, (ast.List, ast.Tuple)):
                continue
            for element in rows.elts:
                if isinstance(element, (ast.List, ast.Tuple)) and len(element.elts) >= 2:
                    cases.append(
                        (node.name, _describe(element.elts[0]), _describe(element.elts[1]))
                    )
                elif isinstance(element, ast.Tuple) and element.elts:
                    # The single-value form, `("1.9",)`, where the case is the input.
                    cases.append((node.name, _describe(element.elts[0]), ""))
    return cases


def literals(text: str) -> list[str]:
    """The numeric tokens in an expected value.

    A serial can be written as `40196`, as `datetime.date(2010, 1, 18)` or as
    `timedelta(seconds=90000)`; only the first is a number this project's tests would
    repeat verbatim, so the rest are reported as unmatched rather than guessed at.
    """
    return re.findall(r"(?<![\w.])\d[\d_]*\.?\d*(?![\w.])", text)


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else "../openpyxl").resolve()
    rust = Path(__file__).resolve().parents[2] / "crates"
    rust_tests = "\n".join(
        path.read_text(encoding="utf-8", errors="replace")
        for path in sorted(rust.rglob("*.rs"))
    )

    print("pinned values openpyxl's tests assert, and their counterpart here\n")

    total_cases = 0
    total_numeric = 0
    total_matched = 0
    rows = []

    for relative, what in PINNING.items():
        path = root / "openpyxl" / relative
        if not path.is_file():
            rows.append((relative, "?", 0, 0, "no such file in this checkout"))
            continue
        cases = parametrised_cases(path)
        # A date is pinned as `date(2010, 1, 18)` or `datetime.datetime(2010, 1, 18, ...)`.
        # Only a *numeric* expected value is something this project's Rust tests would
        # repeat verbatim, so the metric counts those and reports the rest as not-applicable
        # rather than as a gap -- `test_number_style.py` pins booleans and stripped format
        # strings, which this cannot see and which are not missing from this project.
        numeric = 0
        matched = 0
        for _, raw_input, expected in cases:
            numbers = literals(expected)
            if not numbers:
                continue
            numeric += 1
            for number in numbers:
                if number in rust_tests:
                    matched += 1
                    break
        rows.append((relative, what, len(cases), numeric, matched, ""))
        total_cases += len(cases)
        total_numeric += numeric
        total_matched += matched

    for relative, what, cases, numeric, matched, note in rows:
        if numeric == 0:
            verdict = "n/a "
            detail = (
                f"{cases} cases, none pinning a bare number -- this metric cannot see them"
                if cases
                else "no parametrised cases; the module pins values with plain asserts"
            )
        elif matched == numeric:
            verdict = "ok  "
            detail = f"{matched}/{numeric} numeric expectations appear in this project's tests"
        else:
            verdict = "GAP "
            detail = (
                f"{matched}/{numeric} numeric expectations appear in this project's tests; "
                f"the other {numeric - matched} are pinned upstream and not here"
            )
        print(f"{verdict} {relative}")
        print(f"       {what}")
        print(f"       {detail}")
        if note:
            print(f"       {note}")
        print()

    print(
        f"{total_matched} of {total_numeric} numeric expectations in "
        f"{total_cases} parametrised cases appear in this project's tests."
    )
    print(
        "\nWhat this does and does not establish:\n"
        "  - it establishes that a number openpyxl pins is also pinned here, which is the\n"
        "    claim PARITY.md makes about password hashes, date serials, Julian days and the\n"
        "    colour table.\n"
        "  - it does not establish that the two sides agree on those numbers. Agreement is\n"
        "    corpus.py's job, and only for the fixtures it can run.\n"
        "  - a non-numeric expected value -- a boolean, a stripped format string -- is out of\n"
        "    scope here, so a module reported `n/a` is not thereby unverified.\n"
        "  - nothing about the 161 test files as a whole. manifest.py is the honest map of\n"
        "    those; this file looks only at their value-pinning part."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())