#!/usr/bin/env python
"""Audit ferroxl against the openpyxl 1.9.0 source tree, module by module.

This is what PARITY.md is written from. Point it at a checkout of the Python original and
it reports, for every module, which of openpyxl's public names have a ferroxl counterpart.

    python tools/parity.py ../openpyxl/openpyxl

Names are compared after lowercasing and dropping underscores, because openpyxl uses
`snake_case` functions where ferroxl has methods on a type and where a faithful port
renamed a few things for Rust's conventions. A name that matches only after that
normalisation is reported separately: it is a naming difference, not a missing feature.

Exit status is 0 whether or not there are gaps -- this is a report, not a gate.
"""

import ast
import re
import sys
from pathlib import Path

# openpyxl's own test, sample and benchmark trees, and the lxml/Python compatibility shims.
SKIP_PARTS = {"tests", "benchmarks", "sample", "compat", "long"}

# Classes upstream that exist only to hold a dict of slot names, or are empty placeholders.
# Their content is covered by a ferroxl type with a different shape.
NOTABLE_SHAPE = {
    "formatting/rules.py:FormatRule",
    "styles/hashable.py:HashableObject",
    "worksheet/worksheet.py:SheetView",
}

RUST_ITEM = re.compile(
    r"^\s*(?:pub(?:\([^)]*\))?\s+)?(?:struct|enum|trait|type|fn|const|static)\s+"
    r"([A-Za-z_][A-Za-z0-9_]*)",
    re.M,
)
RUST_FN = re.compile(r"^\s*(?:pub(?:\([^)]*\))?\s+)?fn\s+([A-Za-z_][A-Za-z0-9_]*)", re.M)


def normalise(name: str) -> str:
    return name.lower().replace("_", "")


def python_items(path: Path) -> list[tuple[str, str]]:
    """Every top-level class, function and constant in a Python module."""
    tree = ast.parse(path.read_text(encoding="utf-8", errors="replace"))
    found = []
    for node in tree.body:
        if isinstance(node, ast.ClassDef):
            found.append(("class", node.name))
        elif isinstance(node, ast.FunctionDef) and not node.name.startswith("_"):
            found.append(("fn", node.name))
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and not target.id.startswith("_"):
                    found.append(("const", target.id))
    return found


def load_rust(src: Path) -> tuple[dict[str, set[str]], dict[str, set[str]]]:
    """Every item name and every function name in the Rust tree, normalised."""
    items: dict[str, set[str]] = {}
    functions: dict[str, set[str]] = {}
    for path in src.rglob("*.rs"):
        text = path.read_text(encoding="utf-8", errors="replace")
        where = path.relative_to(src).as_posix()
        for name in RUST_ITEM.findall(text):
            items.setdefault(normalise(name), set()).add(where)
        for name in RUST_FN.findall(text):
            functions.setdefault(normalise(name), set()).add(where)
    return items, functions


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print(__doc__)
        return 2
    python = Path(argv[1]).resolve()
    src = Path(__file__).resolve().parents[1] / "crates" / "ferroxl" / "src"
    if not python.is_dir():
        print(f"{python} is not a directory")
        return 2

    items, functions = load_rust(src)

    exact_gaps: list[str] = []
    renamed: list[str] = []
    rows = []

    for path in sorted(python.rglob("*.py")):
        rel = path.relative_to(python).as_posix()
        if set(rel.split("/")) & SKIP_PARTS or rel == "__init__.py":
            continue
        names = python_items(path)
        missing = []
        for kind, name in names:
            if name in NOTABLE_SHAPE:
                continue
            key = normalise(name)
            if key in items or key in functions:
                continue
            missing.append(f"{kind} {name}")
            (exact_gaps if key not in functions else renamed).append(f"{rel}: {name}")
        rows.append((rel, len(names), missing))

    for rel, count, missing in rows:
        status = "OK  " if not missing else "GAP "
        print(f"{status} {rel:38} {count:4}  {', '.join(missing)}")

    matched = sum(count for _, count, missing in rows if not missing)
    total = sum(count for _, count, _ in rows)
    print()
    print(f"openpyxl top-level names : {total}")
    print(f"fully matched modules    : {sum(1 for _, _, m in rows if not m)}/{len(rows)}")
    print(f"names in a matched module: {matched}")
    print(f"unmatched names          : {len(exact_gaps) + len(renamed)}")
    for entry in sorted(set(exact_gaps)):
        print(f"  unmatched  {entry}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))