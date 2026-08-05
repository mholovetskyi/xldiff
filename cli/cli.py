#!/usr/bin/env python3
"""xldiff: compare two Excel models and report what changed and what it did.

    python cli.py old.xlsx new.xlsx [-o report.html] [--json out.json]
                                    [--fail-on high|medium|low|never]

Runs entirely on this machine. No network calls.
"""

import argparse
import json
import os
import sys

from engine import compare
from report import render

SEV_RANK = {"high": 0, "medium": 1, "low": 2}
COLOR = {"high": "\033[31m", "medium": "\033[33m", "low": "\033[34m"}
RESET = "\033[0m"


def main(argv=None):
    ap = argparse.ArgumentParser(prog="xldiff", description=__doc__)
    ap.add_argument("old")
    ap.add_argument("new")
    ap.add_argument("-o", "--out", default="xldiff-report.html")
    ap.add_argument("--json", dest="json_out")
    ap.add_argument("--fail-on", default="never",
                    choices=["high", "medium", "low", "never"])
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args(argv)

    for p in (a.old, a.new):
        if not os.path.exists(p):
            sys.stderr.write("no such file: %s\n" % p)
            return 2

    result = compare(a.old, a.new)
    findings = result["findings"]
    tty = sys.stdout.isatty()

    if not a.quiet:
        counts = {s: sum(1 for f in findings if f.severity == s)
                  for s in ("high", "medium", "low")}
        print("%s -> %s" % (os.path.basename(a.old), os.path.basename(a.new)))
        print("%d findings: %d high, %d medium, %d low"
              % (len(findings), counts["high"], counts["medium"], counts["low"]))
        if result["stale"]:
            print("warning: a file has no cached values, so value impact is incomplete")
        print("")
        for f in findings[:40]:
            tag = f.severity.upper().ljust(6)
            if tty:
                tag = COLOR[f.severity] + tag + RESET
            loc = ("%s!%s" % (f.sheet, f.ref)) if f.ref else (f.sheet or "workbook")
            print("%s %-22s %s" % (tag, loc, f.summary))
            if f.before or f.after:
                print("%s %-22s   %s  ->  %s" % (" " * 6, "", f.before, f.after))
            if f.val_before != f.val_after and (f.val_before or f.val_after):
                print("%s %-22s   value %s -> %s"
                      % (" " * 6, "", f.val_before, f.val_after))
        if len(findings) > 40:
            print("... %d more, see the report" % (len(findings) - 40))

    render(result, a.out)
    if not a.quiet:
        print("\nreport: %s" % os.path.abspath(a.out))

    if a.json_out:
        payload = [{
            "severity": f.severity, "kind": f.kind, "sheet": f.sheet, "ref": f.ref,
            "summary": f.summary, "detail": f.detail,
            "before": f.before, "after": f.after,
            "value_before": f.val_before, "value_after": f.val_after,
        } for f in findings]
        with open(a.json_out, "w") as fh:
            json.dump({"findings": payload, "stale_values": result["stale"]},
                      fh, indent=2)

    if a.fail_on != "never":
        limit = SEV_RANK[a.fail_on]
        if any(SEV_RANK[f.severity] <= limit for f in findings):
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
