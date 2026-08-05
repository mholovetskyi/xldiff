#!/usr/bin/env python3
"""xldiff lineage: walk a folder of dated versions instead of a single pair.

Models do not live in git. They live as model_v12.xlsx ... model_v15.xlsx in
one directory, and the question a reviewer actually asks is not "what changed
since the last save" but "when did this hardcode first appear, and who was I
talking to that week".

    python cli/lineage.py models/*.xlsx
    python cli/lineage.py models/*.xlsx --cell "Revenue build!D9"
    python cli/lineage.py models/*.xlsx --kind hardcode

Files are compared in the order given, so a shell glob sorts them. Everything
runs locally. No network calls.
"""

import argparse
import json
import os
import re
import sys

from engine import (build_inverse, col_to_num, compare, fmt,
                    get_column_letter)

_CELL = re.compile(r'^(.*)!\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})$')


def parse_cell(spec):
    m = _CELL.match(spec.strip())
    if not m:
        raise ValueError('expected a cell like "Sheet!B12", got %r' % spec)
    return m.group(1).strip("'"), int(m.group(3)), col_to_num(m.group(2))


def steps(paths, outputs=None):
    """Compare each consecutive pair, oldest first."""
    for older, newer in zip(paths[:-1], paths[1:]):
        yield older, newer, compare(older, newer, outputs)


def trace_cell(paths, target):
    """Follow one cell backwards through the versions, honouring alignment.

    A cell does not keep its address. Insert a row above it and D9 becomes D10,
    so walking the same reference through every file reads a different cell
    each time and reports changes that never happened. The alignment already
    computed for each pair is what maps it back.
    """
    sheet, row, col = target
    history = [None] * len(paths)
    where = [None] * len(paths)
    cur = (row, col)

    for i in range(len(paths) - 1, 0, -1):
        result = compare(paths[i - 1], paths[i])
        right, left = result["right"], result["left"]
        if sheet not in right:
            break
        where[i] = cur
        history[i] = right[sheet].get(*cur)
        inv = build_inverse(result["align"])
        rmap, cmap = inv.get(sheet, ({}, {}))
        prev = (rmap.get(cur[0]), cmap.get(cur[1]))
        if prev[0] is None or prev[1] is None:
            break
        cur = prev
        if i == 1 and sheet in left:
            where[0] = cur
            history[0] = left[sheet].get(*cur)
    return history, where


def render_cell(paths, target, history, where):
    sheet = target[0]
    print("%s through %d version%s" %
          (sheet, len(paths), "" if len(paths) == 1 else "s"))
    print("")
    prev = None
    for i, path in enumerate(paths):
        cell = history[i]
        pos = where[i]
        ref = ("%s%d" % (get_column_letter(pos[1]), pos[0])) if pos else "--"
        if cell is None:
            print("  %-24s %-8s not present" % (os.path.basename(path), ref))
            prev = None
            continue
        shown = cell.formula if cell.is_formula else fmt(cell.value)
        mark = " "
        if prev is not None and prev.r1c1 != cell.r1c1:
            mark = "*"
        print("  %s %-22s %-8s %-28s %s"
              % (mark, os.path.basename(path), ref, shown, fmt(cell.value)))
        prev = cell
    print("")
    print("  * marks the version where the formula changed")


def render_summary(paths, rows, kind):
    print("%d version%s, %d step%s" %
          (len(paths), "" if len(paths) == 1 else "s",
           len(rows), "" if len(rows) == 1 else "s"))
    print("")
    for row in rows:
        print("  %s -> %s" % (row["from"], row["to"]))
        print("     %d findings: %d high, %d medium, %d low"
              % (row["total"], row["high"], row["medium"], row["low"]))
        for f in row["notable"]:
            print("     %-6s %-22s %s"
                  % (f["severity"].upper(), f["ref"], f["summary"]))
        if kind:
            first = row["first_of_kind"]
            if first:
                print("     %s appears here: %s" % (kind, ", ".join(first)))
    print("")


def main(argv=None):
    ap = argparse.ArgumentParser(prog="xldiff-lineage", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+")
    ap.add_argument("--cell", default="",
                    help='follow one cell through every version, e.g. "Sheet!B12"')
    ap.add_argument("--kind", default="",
                    help="report the step where findings of this kind first appear")
    ap.add_argument("--json", dest="json_out")
    a = ap.parse_args(argv)

    for p in a.files:
        if not os.path.exists(p):
            sys.stderr.write("no such file: %s\n" % p)
            return 2
    if len(a.files) < 2:
        sys.stderr.write("lineage needs at least two versions to compare\n")
        return 2

    if a.cell:
        try:
            target = parse_cell(a.cell)
        except ValueError as exc:
            sys.stderr.write("%s\n" % exc)
            return 2
        history, where = trace_cell(a.files, target)
        render_cell(a.files, target, history, where)
        if a.json_out:
            # "formula" carries a formula or nothing. A cell holding a bare
            # number is the whole point of a plug, so it reads as calculated
            # false rather than as a formula that happens to be a number.
            payload = [{
                "file": os.path.basename(p),
                "ref": ("%s%d" % (get_column_letter(where[i][1]), where[i][0]))
                       if where[i] else None,
                "calculated": bool(history[i] and history[i].is_formula),
                "formula": (history[i].formula
                            if history[i] and history[i].is_formula else None),
                "value": fmt(history[i].value) if history[i] else None,
            } for i, p in enumerate(a.files)]
            with open(a.json_out, "w", encoding="utf-8") as fh:
                json.dump({"cell": a.cell, "history": payload}, fh, indent=2)
        return 0

    rows = []
    seen_kind = False
    for older, newer, result in steps(a.files):
        findings = result["findings"]
        counts = dict((s, sum(1 for f in findings if f.severity == s))
                      for s in ("high", "medium", "low"))
        of_kind = [f for f in findings if a.kind and f.kind == a.kind]
        first = []
        if of_kind and not seen_kind:
            first = [(f.sheet + "!" + f.ref) if f.ref else f.sheet for f in of_kind]
            seen_kind = True
        rows.append({
            "from": os.path.basename(older), "to": os.path.basename(newer),
            "total": len(findings), "high": counts["high"],
            "medium": counts["medium"], "low": counts["low"],
            "notable": [{"severity": f.severity,
                         "ref": (f.sheet + "!" + f.ref) if f.ref else f.sheet,
                         "summary": f.summary}
                        for f in findings if f.severity == "high"][:4],
            "first_of_kind": first,
        })

    render_summary(a.files, rows, a.kind)
    if a.json_out:
        with open(a.json_out, "w", encoding="utf-8") as fh:
            json.dump({"versions": [os.path.basename(p) for p in a.files],
                       "steps": rows}, fh, indent=2)
    return 0


if __name__ == "__main__":
    sys.exit(main())
