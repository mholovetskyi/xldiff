"""Smoke test. Runs the Python engine against the example models and checks
that the seeded review bugs are all caught, and that a file compared with
itself produces nothing.

    python tests/test_smoke.py
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "cli"))

from engine import compare, to_r1c1, precedents  # noqa: E402

A = os.path.join(ROOT, "examples", "model_v14.xlsx")
B = os.path.join(ROOT, "examples", "model_v15.xlsx")
SA = os.path.join(ROOT, "examples", "sample_before.xlsx")
SB = os.path.join(ROOT, "examples", "sample_after.xlsx")

EXPECTED = [
    ("high", "run_break", "Revenue build", "F13"),
    ("high", "hardcode", "Revenue build", "C14"),
    ("medium", "constant_changed", "Revenue build", "C5"),
    ("medium", "sheet_added", "Sensitivity", ""),
    ("low", "row_added", "Opex", "row 5"),
    ("low", "run_repaired", "Revenue build", "E12"),
]

failures = []


def check(cond, label):
    print(("  ok   " if cond else "  FAIL ") + label)
    if not cond:
        failures.append(label)


def oc_all(result):
    return [f for f in result["findings"] if f.sheet == "Operating costs" and f.ref]


def main():
    result = compare(A, B)
    got = set((f.severity, f.kind, f.sheet, f.ref) for f in result["findings"])

    print("comparing the example models")
    for exp in EXPECTED:
        check(exp in got, "%s %s at %s!%s" % exp)

    check(not result["stale"], "cached values are present in both files")

    impacts = [f for f in result["findings"] if f.kind == "impact"]
    check(len(impacts) >= 4, "downstream value moves detected (%d)" % len(impacts))

    hardcode = [f for f in result["findings"] if f.kind == "hardcode"][0]
    check(hardcode.val_before != hardcode.val_after,
          "the hardcode finding carries a value delta")

    order = [f.severity for f in result["findings"]]
    rank = {"high": 0, "medium": 1, "low": 2}
    check(order == sorted(order, key=lambda s: rank[s]),
          "findings are sorted by severity")

    print("comparing a file with itself")
    same = compare(B, B)
    check(len(same["findings"]) == 0,
          "identical files produce no findings (got %d)" % len(same["findings"]))

    print("errors introduced by the revision")
    sample = compare(SA, SB)
    got = set((f.severity, f.kind, f.sheet, f.ref) for f in sample["findings"])
    for exp in [
        ("high", "error_introduced", "Revenue build", "B14"),
        ("high", "broken_reference", "Revenue build", "B16"),
        ("high", "sheet_deleted", "Prior year", ""),
    ]:
        check(exp in got, "%s %s at %s!%s" % exp)

    print("dependency chains")
    check(precedents("='Revenue build'!B9*1.5", "Sensitivity") == [("Revenue build", 9, 2)],
          "a cross-sheet reference resolves to the sheet it names")
    check(len(precedents("=SUM(B4:B6)", "S")) == 3, "a range expands to the cells it covers")
    check(precedents("=LOG10(B4)", "S") == [("S", 4, 2)],
          "a function name that looks like a reference is not one")
    named = precedents("=B7*TAX1", "Costs", {"TAX1": "'Revenue build'!$B$24"})
    check(("Revenue build", 24, 2) in named, "a name resolves to what it points at")
    check(all(c[2] < 100 for c in named),
          "and is not also read as a cell 13,570 columns out")

    growth = [f for f in sample["findings"]
              if f.sheet == "Revenue build" and f.ref == "B4"][0]
    check(growth.downstream >= 8,
          "the growth rate is known to feed %d cells" % growth.downstream)
    check(growth.chain and growth.chain[0] == "Revenue build!B4",
          "its chain starts at the cell that changed")
    check(len(growth.chain) >= 3,
          "and runs forward to a number that moved: %s" % " -> ".join(growth.chain))
    impacts = [f for f in sample["findings"] if f.kind == "impact"]
    check(all(not f.chain for f in impacts),
          "impact findings carry no chain: they are the far end of someone else's")

    print("defined names")
    kinds = dict((f.kind, f) for f in sample["findings"])
    check("name_changed" in kinds and kinds["name_changed"].ref == "GrowthRate",
          "a name repointed at another cell is high severity")
    check("name_deleted" in kinds and kinds["name_deleted"].severity == "high",
          "a deleted name is high: formulas using it can no longer resolve")
    check("name_added" in kinds and kinds["name_added"].severity == "low",
          "a new name is low on its own")

    # TAX1 is shaped exactly like a cell reference. Normalising it relative to
    # its own position makes identical formulas look different the moment a row
    # is inserted above them.
    check(to_r1c1("=B7*TAX1", 8, 2, {"TAX1"}) == to_r1c1("=B8*TAX1", 9, 2, {"TAX1"}),
          "a formula using a name compares equal after it moves down a row")
    check("TAX1" in to_r1c1("=B7*TAX1", 8, 2, {"TAX1"}),
          "the name survives R1C1 normalisation intact")
    check("TAX1" not in to_r1c1("=B7*TAX1", 8, 2, set()),
          "and is only spared because it is known to be a name")
    aftertax = [f for f in oc_all(sample) if f.ref[1:] == "9" and f.kind == "formula_changed"]
    check(not aftertax,
          "the After tax row, identical text one row down, reports no formula change")

    print("an inserted column does not shift the diff")
    oc = [f for f in sample["findings"] if f.sheet == "Operating costs"]
    check(any(f.kind == "column_added" and f.ref == "column B" for f in oc),
          "the inserted column is reported once, by letter and header")
    # Salaries, Rent and Software only moved one column right. Every one of
    # those cells would report as edited without column alignment, which is
    # the failure mode that made naive comparison tools unusable.
    shifted = [f for f in oc if f.ref[1:] in ("4", "5", "7") and f.ref[:1].isalpha()]
    check(not shifted,
          "cells that only moved right report nothing (got %d)" % len(shifted))
    al = sample["align"]["Operating costs"]
    check(al["col_added"] == [2] and not al["col_deleted"],
          "exactly one column is reported as inserted, at B")
    # The sheet shifts on both axes at once. Row alignment must survive the
    # column insertion and vice versa: each axis used to need the other solved
    # first, and got zero when it did not.
    check(al["added"] == [6] and not al["deleted"],
          "row alignment survives it: one insertion at row 6, nothing deleted")
    check(al["ratio"] > 0.35 and al["col_ratio"] > 0.35,
          "both confidences stay well clear of the low-confidence threshold "
          "(rows %.2f, columns %.2f)" % (al["ratio"], al["col_ratio"]))

    print("run breaks on both axes")
    check(("high", "run_break", "Revenue build", "B20") in got,
          "an overridden step mid-column is a vertical run break")
    vertical = [f for f in sample["findings"] if f.ref == "B20"][0]
    check("column" in vertical.summary, "the finding names the axis that broke")
    horizontal = [f for f in sample["findings"] if f.ref == "C7"][0]
    check("row" in horizontal.summary, "a horizontal break still reads as a row")

    # A column of identical formulas nearly always ends in a total that is
    # meant to differ. Reporting that as a break would make the whole axis
    # useless, so it is worth a test of its own.
    totals = compare(A, B)
    d11 = [f for f in totals["findings"] if f.sheet == "Opex" and f.ref == "D11"]
    check(d11 and d11[0].kind != "run_break",
          "a totals row below a uniform column is not a run break")

    err = [f for f in sample["findings"] if f.kind == "error_introduced"][0]
    check("#DIV/0!" in err.summary, "the error finding names the error it found")
    check(not sample["stale"],
          "an error value is a result, not a missing one, so nothing reads as stale")

    print("errors cleared by the revision")
    back = compare(SB, SA)
    cleared = [f for f in back["findings"] if f.kind == "error_cleared"]
    check(len(cleared) == 2, "reversing the pair reports both fixes (got %d)" % len(cleared))
    check(all(f.severity == "low" for f in cleared),
          "a fix is low severity, so it cannot crowd out a real problem")

    print("")
    if failures:
        print("%d check(s) failed" % len(failures))
        return 1
    print("all checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
