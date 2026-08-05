"""Lineage: models live as dated copies in a folder, not as commits.

The question a reviewer asks is not "what changed since the last save" but
"when did this first appear". These check the walk across versions, and in
particular that a cell is followed through alignment rather than by address.

    python tests/test_lineage.py
"""

import io
import json
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "cli"))

import lineage  # noqa: E402

V = [os.path.join(ROOT, "examples", n) for n in
     ("sample_before.xlsx", "sample_interim.xlsx", "sample_after.xlsx")]

failures = []


def check(cond, label):
    print(("  ok   " if cond else "  FAIL ") + label)
    if not cond:
        failures.append(label)


def main():
    print("walking three versions")
    tmp = tempfile.mkdtemp()
    out = os.path.join(tmp, "lineage.json")
    code = lineage.main(V + ["--kind", "hardcode", "--json", out])
    check(code == 0, "the summary walk succeeds")
    data = json.load(io.open(out, encoding="utf-8"))
    check(len(data["steps"]) == 2, "three versions produce two steps")

    early, late = data["steps"]
    check(early["high"] == 0 and late["high"] > 0,
          "the high findings all land in the second step, where the damage was done")
    check(not early["first_of_kind"] and late["first_of_kind"] == ["Revenue build!D9"],
          "the plug is reported as first appearing in that step, once")

    print("following a cell by address")
    out2 = os.path.join(tmp, "cell.json")
    check(lineage.main(V + ["--cell", "Revenue build!D9", "--json", out2]) == 0,
          "tracing a cell succeeds")
    hist = json.load(io.open(out2, encoding="utf-8"))["history"]
    check([h["ref"] for h in hist] == ["D9", "D9", "D9"],
          "the plug sits at D9 in every version")
    check([h["calculated"] for h in hist] == [True, True, False],
          "it was calculated in the first two versions and is not in the third")
    check(hist[0]["formula"] == "=D7*D8" and hist[2]["formula"] is None,
          "so the last version carries no formula at all")
    check(hist[0]["value"] == hist[2]["value"] != hist[1]["value"],
          "and the value it froze at is the one it held two versions ago (%s)"
          % hist[0]["value"])

    print("following a cell that moved")
    # Operating costs gains a row and a column between the interim and final
    # files, so the After tax cell is B8 early and C9 late. Reading the same
    # address in every file would report a change that never happened.
    out3 = os.path.join(tmp, "moved.json")
    check(lineage.main(V + ["--cell", "Operating costs!C9", "--json", out3]) == 0,
          "tracing a cell that moved succeeds")
    moved = json.load(io.open(out3, encoding="utf-8"))["history"]
    check([h["ref"] for h in moved] == ["B8", "B8", "C9"],
          "it is followed back through the insertions to B8, not read at C9")
    check(moved[0]["formula"] == "=B7*TAX1" and moved[2]["formula"] == "=C8*TAX1",
          "and the formula is the same one, shifted")

    print("bad input")
    check(lineage.main([V[0]]) == 2, "one file is not a lineage")
    check(lineage.main(V + ["--cell", "not a cell"]) == 2,
          "a malformed cell reference is rejected rather than guessed at")
    check(lineage.main([V[0], os.path.join(tmp, "missing.xlsx")]) == 2,
          "a missing file is reported")

    print("")
    if failures:
        print("%d check(s) failed" % len(failures))
        return 1
    print("all checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
