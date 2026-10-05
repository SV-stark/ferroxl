"""Where can the library panic on input it did not write?

A library that panics is a library whose caller has to wrap every call in `catch_unwind`,
because a malformed workbook is not a programmer error. This lists the `unwrap`, `expect`
and `panic!` calls that sit outside `#[cfg(test)]`, so the surface can be judged rather
than assumed.

    python tools/panic_surface.py
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1] / "crates" / "ferroxl" / "src"

PANIC = re.compile(r"\.unwrap\(\)|\.expect\(|panic!|unreachable!|unimplemented!")
ALLOW = re.compile(r"//\s*(?:allow|justified|invariant|internal)\b", re.I)


def main() -> int:
    findings = []
    for path in sorted(ROOT.rglob("*.rs")):
        in_test = False
        for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            if re.match(r"\s*#\[cfg\(test\)\]", line):
                in_test = True
            if in_test or not PANIC.search(line):
                continue
            # A doc comment is an example, not a panic in the shipped code.
            if line.lstrip().startswith(("///", "//!", "//!  ")):
                continue
            if ALLOW.search(line):
                continue
            findings.append((path.relative_to(ROOT).as_posix(), number, line.strip()))

    by_file = {}
    for file, _, _ in findings:
        by_file[file] = by_file.get(file, 0) + 1

    print(f"{len(findings)} panic sites outside tests, in {len(by_file)} files\n")
    for file, count in sorted(by_file.items(), key=lambda pair: -pair[1]):
        print(f"{count:3}  {file}")
    print("\nthe sites themselves:")
    for file, number, line in findings:
        print(f"  {file}:{number}  {line}")

    print(
        "\nNone of these are wrong by themselves -- most guard an invariant the writer\n"
        "establishes. The point of the list is that 'none' would be a stronger claim than\n"
        "'most', and a library reading hostile files wants the stronger one."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())