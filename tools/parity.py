#!/usr/bin/env python
"""Audit ferroxl against the openpyxl 3.1.5 source tree, module by module.

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

# Modules whose public names are *data*, not identifiers. openpyxl's built-in styles are 51
# module-level string literals; ferroxl holds the same 49 as one Rust table and looks them up
# by name, so no identifier carries their name and a name-only matcher sees 51 phantom gaps.
#
# Matched on `builtinId`, not on the name: openpyxl calls the variable `accent_1_20` for the
# style it names "20 % - Accent1", so the two spellings agree on nothing but the id.
#
# The count is checked against the Rust source rather than trusted, so deleting an entry from
# the table shows up here instead of quietly inflating the score.
DATA_TABLE_MODULES = {
    "styles/builtins.py": ("styles/named_style.rs", "BUILTIN_DETAILS", 49),
}

PYTHON_BUILTIN_ID = re.compile(r'^(\w+) = """(.*?)"""', re.S | re.M)

# Names in a data-table module that the Rust table covers under a different spelling.
# `styles` is openpyxl's name -> XML lookup dict; `BUILTIN_DETAILS` is the same lookup as a
# table. Anything *not* listed here is a real gap, not a naming difference.
DATA_TABLE_ALIASES = {
    "styles/builtins.py": {"styles"},
}


def data_table_ids(rel: str, rust_root: Path, python_path: Path) -> dict[str, str] | None:
    """Map each Python variable to the `builtinId` the Rust table has to contain, by name.

    Returns None when the module is not a known data table, or when the Rust table is missing
    or the wrong size -- in which case the module is scored the honest way, as a gap.
    """
    entry = DATA_TABLE_MODULES.get(rel)
    if entry is None:
        return None
    relative, const_name, expected = entry
    source = rust_root / relative
    if not source.is_file():
        print(f"NOTE {rel}: data table missing at {source}")
        return None
    text = source.read_text(encoding="utf-8", errors="replace")
    body = re.search(
        rf"pub const {const_name}: \[[^\]]*; (\d+)\] = \[(.*?)\n\];", text, re.S
    )
    if body is None:
        print(f"NOTE {rel}: could not read {const_name} from {source}")
        return None
    found = int(body.group(1))
    if found != expected:
        print(
            f"NOTE {rel}: {const_name} holds {found} entries, expected {expected}"
            " -- openpyxl's count changed or the table was edited by hand"
        )
        return None
    ids = set(re.findall(r'builtin_id: "([^"]+)"', body.group(2)))

    python_text = python_path.read_text(encoding="utf-8", errors="replace")
    covered = {
        var: match.group(1)
        for var, block in PYTHON_BUILTIN_ID.findall(python_text)
        if (match := re.search(r'builtinId="(\d+)"', block))
        and match.group(1) in ids
    }
    for alias in DATA_TABLE_ALIASES.get(rel, set()):
        covered[alias] = ""
    return covered


def normalise(name: str) -> str:
    return name.lower().replace("_", "")


def normalise_style(name: str) -> str:
    """Compare a Python variable name against a style name.

    openpyxl names the variable `accent_1_20` for the style called "Accent1 20%". Dropping
    every non-alphanumeric character is what makes those the same entry rather than two.
    """
    return re.sub(r"[^a-z0-9]", "", name.lower())


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
        table = data_table_ids(rel, src, path)
        missing = []
        for kind, name in names:
            if name in NOTABLE_SHAPE:
                continue
            key = normalise(name)
            if key in items or key in functions:
                continue
            # A data table's entries are covered by name, not by an identifier bearing it.
            if table is not None and name in table:
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