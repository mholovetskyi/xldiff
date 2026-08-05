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

    print("")
    if failures:
        print("%d check(s) failed" % len(failures))
        return 1
    print("all checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
