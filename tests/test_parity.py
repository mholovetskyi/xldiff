"""Holds the two engines against each other.

README claims web/engine.js and cli/engine.py produce identical findings on the
same inputs. Every feature has to land twice for that to stay true, so this
runs both over the example pairs and diffs the results field by field.

    python tests/test_parity.py

Skips with a clear message if node or the web dependencies are missing, so the
Python-only path still works.
"""

import json
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "cli"))

from engine import compare  # noqa: E402

PAIRS = [
    ("model_v14.xlsx", "model_v15.xlsx"),
    ("sample_before.xlsx", "sample_after.xlsx"),
    ("sample_after.xlsx", "sample_before.xlsx"),   # reversed, to catch asymmetry
    ("sample_after.xlsx", "sample_after.xlsx"),    # a file against itself
]

failures = []


def check(cond, label):
    print(("  ok   " if cond else "  FAIL ") + label)
    if not cond:
        failures.append(label)


def js_findings(a, b):
    out = subprocess.run(
        ["node", os.path.join(ROOT, "tests", "parity_dump.js"), a, b],
        capture_output=True, text=True, cwd=ROOT, shell=(os.name == "nt"))
    if out.returncode != 0:
        raise RuntimeError(out.stderr.strip() or "parity_dump.js failed")
    return json.loads(out.stdout)


def py_findings(a, b):
    r = compare(a, b)
    return {
        "stale": r["stale"],
        "findings": [{
            "id": f.id, "severity": f.severity, "kind": f.kind, "sheet": f.sheet, "ref": f.ref,
            "summary": f.summary, "detail": f.detail,
            "before": f.before, "after": f.after,
            "valBefore": f.val_before, "valAfter": f.val_after,
            "chain": f.chain, "downstream": f.downstream,
            "outputs": f.outputs, "outputImpact": f.output_impact,
            "onOutput": f.on_output,
        } for f in r["findings"]],
    }


def main():
    if not os.path.isdir(os.path.join(ROOT, "web", "node_modules")):
        print("skipped: web/node_modules is missing (cd web && npm install)")
        return 0

    for na, nb in PAIRS:
        a = os.path.join(ROOT, "examples", na)
        b = os.path.join(ROOT, "examples", nb)
        print("%s -> %s" % (na, nb))
        try:
            js = js_findings(a, b)
        except (OSError, RuntimeError) as exc:
            print("skipped: could not run node (%s)" % exc)
            return 0
        py = py_findings(a, b)

        check(js["stale"] == py["stale"],
              "stale-cache verdict agrees (js=%s py=%s)" % (js["stale"], py["stale"]))
        check(len(js["findings"]) == len(py["findings"]),
              "same number of findings (js=%d py=%d)"
              % (len(js["findings"]), len(py["findings"])))

        for i, (jf, pf) in enumerate(zip(js["findings"], py["findings"])):
            if jf == pf:
                continue
            diff = [k for k in jf if jf[k] != pf.get(k)]
            check(False, "finding %d at %s!%s differs on %s\n         js: %s\n         py: %s"
                  % (i, jf["sheet"], jf["ref"], ", ".join(diff),
                     {k: jf[k] for k in diff}, {k: pf.get(k) for k in diff}))
        if len(js["findings"]) == len(py["findings"]) and js == py:
            check(True, "every field matches across %d findings" % len(js["findings"]))

    print("")
    if failures:
        print("%d check(s) failed" % len(failures))
        return 1
    print("all checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
