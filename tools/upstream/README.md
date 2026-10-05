# Running openpyxl's own test suite against ferroxl

openpyxl 3.1.5 ships **161 test files and about 1,700 test functions**. They are Python
calling a Python API, so they cannot be *run* against a Rust library. "Replicating the test
suite" therefore cannot mean executing it, and pretending otherwise would be the easiest way
to make this look covered when it is not.

What is possible is three other things, and this directory does all three.

## Why a checkout is needed

Neither the tests nor the fixtures ship in the wheel — `pip install openpyxl` gives you the
library and neither of those. So these harnesses need a source checkout, pinned:

```console
$ git clone --depth 1 --branch 3.1.5 https://github.com/theorchard/openpyxl ../openpyxl
$ python tools/upstream/run.py ../openpyxl
```

The same requirement `tools/parity.py` already has, and for the same reason: the pinned tag
is the release `PARITY.md`'s figures were measured against.

## What each tool does

| | what it is | gate? |
| --- | --- | --- |
| `corpus.py` | Every real workbook in `tests/data/genuine/`, read by openpyxl **and** by the server, compared cell by cell. openpyxl is the oracle. | yes |
| `expectations.py` | The numbers openpyxl's tests pin, cross-checked against this project's tests. | report |
| `manifest.py` | All 161 test files mapped onto ferroxl modules, with what has no counterpart. | report |
| `survey.py` | What is in the wider fixture corpus and what ferroxl can read of it. | report |
| `run.py` | Runs the above in order; non-zero if a *gate* fails. | — |

A report does not fail the build on a gap. A gap it describes is information, and a CI step
that fails on "this is incomplete" teaches everyone to ignore the output.

## What corpus.py found

It found a bug nothing else had: **the 1904 date system was ignored on read.**
`Cell::display_value` converted every serial against the 1900 epoch, so a real Mac Excel
workbook — `tests/data/genuine/mac_date.xlsx`, which carries `date1904="true"` — read every
date four years and one day early. That is 1462 days, the exact gap between the two epochs,
and `PARITY.md` claimed 1904 workbooks "round-trip exactly".

The write direction was already correct, so the damage was confined to reading: a date read
from a 1904 workbook and shown to a user was wrong, and nothing anywhere said so.

## What is *not* covered, deliberately

- **Charts and images.** openpyxl 3.1.5's own reader does not return them, so there is no
  oracle for what a reader should see. The writer side is covered by `tools/mcp_parity.py`.
- **The `pivot/`, `chartsheet/`, `descriptors/` and `packaging` suites.** 316 of openpyxl's
  tests target packages with no ferroxl counterpart. `manifest.py` accounts for each by name;
  `PARITY.md` explains why, and they are the reason adding an OOXML construct is hand-written
  twice rather than declared once.
- **Cell values past row 500 or column 60**, and the first 40 sheets of a workbook. Both caps
  are reported rather than applied silently, so a truncated row is never mistaken for a clean
  comparison. `reader/bigfoot.xlsx` has 1024 sheets on purpose, and every MCP call reloads the
  workbook, so sweeping it would cost 1024 full loads for no extra coverage.
- **Formula arithmetic.** A formula is compared as a string. openpyxl without `data_only`
  returns the formula text, so there is no cached value to compare against.

## The honest limit

This directory raises the floor; it does not reach the ceiling. Five bugs got past 732 unit
tests, and corpus.py found a sixth — so the classes these tools catch are real and the
classes they miss are certainly still there. The manifest is the checklist for closing that
gap; a row marked "ported" is not yet a row with comparable coverage.