#!/usr/bin/env python
"""Check that the README's Rust examples are the ones the library's own doctests compile.

The README is the first thing a reader sees, and an example that no longer compiles is
worse than no example. Rather than compiling the README separately — which means
rebuilding the dependency graph by hand — this asserts that every fenced ```` ```rust ````
block in README.md appears verbatim inside a doc comment in the library. The doctests are
run by `cargo test`, so an example that drifts out of date fails the build.

    python tools/check_readme.py
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
README = ROOT / "README.md"
SOURCES = sorted((ROOT / "crates" / "ferroxl" / "src").rglob("*.rs"))


def undoc(text: str) -> str:
    """Strip what only differs between a README block and a doc comment.

    A README fence may say ```rust where a doctest says ```no_run, a doc comment carries a
    `//!` or `///` prefix and whatever indent it sits at, and rustdoc hides a line from the
    rendered docs behind `# `. None of that is part of the example.
    """
    lines = []
    for line in text.splitlines():
        stripped = line.lstrip()
        if stripped.startswith(("//! ", "//!")):
            line = stripped[3:]
            if line.startswith(" "):
                line = line[1:]
        elif stripped.startswith("///"):
            # An indented doc comment, which is what a doctest inside an `impl` block has.
            # `///` with nothing after it is a blank line, and the indent is not part of
            # the example either.
            line = stripped[3:]
            if line.startswith(" "):
                line = line[1:]
        if line.startswith("# "):
            line = line[2:]
        if line.strip().startswith("```"):
            continue
        lines.append(line)
    return "\n".join(lines)


def main() -> int:
    blocks = re.findall(r"```rust\n(.*?)```", README.read_text(encoding="utf-8"), re.S)
    if not blocks:
        print("no Rust examples found in README.md")
        return 1
    docs = "\n".join(undoc(path.read_text(encoding="utf-8")) for path in SOURCES)

    failures = 0
    for index, block in enumerate(blocks):
        body = undoc(block).strip()
        if body in docs:
            print(f"example {index}: ok")
        else:
            failures += 1
            print(f"example {index}: NOT IN THE DOCTESTS")
    if failures:
        print(
            "\nEach example above must also appear verbatim in a doc comment under "
            "crates/ferroxl/src, so that `cargo test` compiles it."
        )
        return 1
    print(f"all {len(blocks)} example(s) are covered by the doctests")
    return 0


if __name__ == "__main__":
    sys.exit(main())