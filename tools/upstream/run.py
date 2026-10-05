"""Run everything that needs an openpyxl checkout, and say what it did and did not cover.

Four steps, in order of how much they prove:

1. `corpus.py`   -- every real workbook openpyxl ships, read by both and compared cell by
                    cell. openpyxl is the oracle; this is where ferroxl is most likely to
                    disagree, and where the 1904 date-system bug was found.
2. `manifest.py` -- maps openpyxl's own 161 test files onto ferroxl, so "we have parity" is
                    a checklist rather than a claim.
3. `survey.py`   -- what is in the wider fixture corpus and what ferroxl can read of it.
4. `expectations.py` -- the values openpyxl's tests pin, checked against this project.

Exit status is non-zero if a *comparison* fails. The manifest and survey are reports: a gap
they describe is information, not a build break, and failing CI on "this is incomplete"
would train everyone to ignore the output.

    python tools/upstream/run.py path/to/openpyxl
"""

import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent

STEPS = [
    (
        "corpus",
        ["corpus.py"],
        "real workbooks read by both implementations, compared cell by cell",
        True,
    ),
    (
        "expectations",
        ["expectations.py"],
        "values openpyxl's own tests pin, checked against ferroxl",
        True,
    ),
    (
        "manifest",
        ["manifest.py"],
        "openpyxl's test suite mapped onto ferroxl (a report, not a gate)",
        False,
    ),
]


def main() -> int:
    checkout = sys.argv[1] if len(sys.argv) > 1 else "../openpyxl"
    root = Path(checkout).resolve()
    if not (root / "openpyxl").is_dir():
        print(f"no openpyxl source tree at {root / 'openpyxl'}")
        print()
        print("These harnesses read openpyxl's test suite and its fixture corpus, neither of")
        print("which ships in the wheel -- `pip install openpyxl` gives the library but not")
        print("either. So this needs a checkout, pinned:")
        print()
        print("    git clone --depth 1 --branch 3.1.5 \\")
        print("        https://github.com/theorchard/openpyxl ../openpyxl")
        print()
        print("The same reason `tools/parity.py` needs one. The pinned tag is the release")
        print("PARITY.md's figures were measured against.")
        return 2

    failures = []
    for name, argv, blurb, is_gate in STEPS:
        print("=" * 74)
        print(f"{name}: {blurb}")
        print("=" * 74)
        result = subprocess.run(
            [sys.executable, str(HERE / argv[0]), str(root), *argv[1:]],
            cwd=HERE.parent.parent,
        )
        if result.returncode != 0:
            failures.append(name if is_gate else f"{name} (reported, not gated)")
        print()

    print("=" * 74)
    if failures:
        print("gating failures:", ", ".join(failures))
        return 1
    print("every gating check passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())