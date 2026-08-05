"""Smoke test. Runs the Python engine against the example models and checks
that the seeded review bugs are all caught, and that a file compared with
itself produces nothing.

    python tests/test_smoke.py
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "cli"))

from engine import compare  # noqa: E402

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
