"""Map openpyxl's own test suite onto ferroxl, and say what is not covered.

There are 180 test files in an openpyxl checkout. "Replicating the test suite" cannot mean
running them -- they are Python calling a Python API -- so the useful artefact is the map:
for each file, which ferroxl module it tests, whether anything here exercises it, and if
not, why. That turns "we have parity" from a claim into a checklist someone can argue with.

    python tools/upstream/manifest.py path/to/openpyxl
"""

import ast
import sys
from pathlib import Path

# ferroxl module -> the openpyxl packages it ports. Anything with no entry is ferroxl's own
# addition rather than a port.
PORTS = {
    "cell": "cell",
    "charts": "chart",
    "comments": "comments",
    "datavalidation": "worksheet",  # openpyxl keeps it in worksheet/datavalidation.py
    "date_time": "utils",
    "drawing": "drawing",
    "exceptions": "compat",
    "formatting": "formatting",
    "formula": "formula",
    "namedrange": "namedrange",
    "reader": "reader",
    "styles": "styles",
    "units": "utils",
    "workbook": "workbook",
    "worksheet": "worksheet",
    "writer": "writer",
    "xml": "xml",
}

# openpyxl packages with no ferroxl counterpart at all. PARITY.md accounts for each.
NO_PORT = {
    "pivot": "no Rust counterpart; the parts are preserved but not modelled",
    "chartsheet": "no Rust counterpart",
    "descriptors": "no Rust counterpart; ferroxl hand-writes its attribute lists",
    "packaging": "built by the writer instead of ported",
    "compat": "Python 2 shims",
}

# ferroxl additions: not a port, so openpyxl has no test file for them and the coverage is
# this project's own tests rather than an upstream one.
ADDITIONS = {
    "dependency": "worksheet::dependency -- cell dependency tracing, no upstream equivalent",
    "preserved": "pass-through preservation, no upstream equivalent",
    "cell_range": "ported from worksheet/cell_range.py, which openpyxl does test",
}


def test_files(root: Path) -> list[Path]:
    return sorted(
        path
        for path in (root / "openpyxl").rglob("test_*.py")
        if "__pycache__" not in path.parts
    )


def subject(path: Path, root: Path) -> tuple[str, str]:
    """The openpyxl package a test file belongs to, and its name."""
    relative = path.relative_to(root / "openpyxl")
    package = relative.parts[0] if len(relative.parts) > 1 else "(top level)"
    return package, path.stem


def count_tests(path: Path) -> int:
    try:
        tree = ast.parse(path.read_text(encoding="utf-8", errors="replace"))
    except SyntaxError:
        return 0
    total = 0
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            name = node.name
            if name.startswith("test_"):
                total += 1
    return total


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else "../openpyxl").resolve()
    if not (root / "openpyxl").is_dir():
        print(f"no openpyxl source tree at {root / 'openpyxl'}")
        return 2

    files = test_files(root)
    groups: dict[str, list[tuple[str, int]]] = {}
    for path in files:
        package, name = subject(path, root)
        groups.setdefault(package, []).append((name, count_tests(path)))

    total_files = len(files)
    total_tests = sum(count for members in groups.values() for _, count in members)

    print(f"openpyxl {root.name}: {total_files} test files, {total_tests} test functions\n")

    ported = [p for p in sorted(groups) if p not in NO_PORT]
    unported = [p for p in sorted(groups) if p in NO_PORT]

    print("packages with a ferroxl port")
    print("-" * 72)
    for package in ported:
        members = sorted(groups[package])
        count = sum(c for _, c in members)
        print(f"{package:16} {len(members):3} files {count:5} tests")
        for name, tests in members:
            print(f"                 {name} ({tests})")

    print("\npackages with no ferroxl counterpart")
    print("-" * 72)
    for package in unported:
        members = sorted(groups[package])
        count = sum(c for _, c in members)
        print(f"{package:16} {len(members):3} files {count:5} tests   {NO_PORT[package]}")

    print("\nferroxl additions, which upstream has no test file for")
    print("-" * 72)
    for name, note in ADDITIONS.items():
        print(f"{name:16} {note}")

    print(
        "\nWhat this does not tell you, deliberately:\n"
        "  - whether a matching name means matching behaviour. A file listed as ported may\n"
        "    still be covered by only one ferroxl test.\n"
        "  - the `descriptors` and `packaging` gaps are not missing features; PARITY.md says\n"
        "    why, and they are the reason adding an OOXML construct is hand-written twice.\n"
        "  - openpyxl's own suite does not test Excel, it tests openpyxl. Most of what it\n"
        "    covers is behaviour ferroxl matches by construction rather than by assertion."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())