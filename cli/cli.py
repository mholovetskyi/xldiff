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

import waivers as waiver_file
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
    ap.add_argument("--outputs", default="",
                    help="cells to rank findings against, comma separated, "
                         "e.g. \"Summary!B10,Model!C4\". Detected automatically "
                         "if not given.")
    ap.add_argument("--waivers", default="",
                    help="JSON file of findings already reviewed and accepted. "
                         "They are reported separately and never fail the gate.")
    ap.add_argument("--write-waivers", dest="write_waivers", default="",
                    help="write a waiver file covering every finding in this "
                         "run, to edit down and commit.")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args(argv)

    for p in (a.old, a.new):
        if not os.path.exists(p):
            sys.stderr.write("no such file: %s\n" % p)
            return 2

    declared = [x.strip() for x in a.outputs.split(",") if x.strip()]
    result = compare(a.old, a.new, declared)
    all_findings = result["findings"]

    try:
        accepted = waiver_file.load(a.waivers)
    except waiver_file.WaiverError as exc:
        sys.stderr.write("%s\n" % exc)
        return 2
    findings = waiver_file.apply(all_findings, accepted)
    waived = [f for f in all_findings if f.waived]
    orphaned = waiver_file.stale(all_findings, accepted)
    result["findings"] = findings
    tty = sys.stdout.isatty()

    if a.write_waivers:
        with open(a.write_waivers, "w", encoding="utf-8") as fh:
            json.dump(waiver_file.dump(all_findings), fh, indent=2)
        if not a.quiet:
            print("waiver template: %s (%d findings, add a reason to each)"
                  % (os.path.abspath(a.write_waivers), len(all_findings)))

    if not a.quiet:
        counts = {s: sum(1 for f in findings if f.severity == s)
                  for s in ("high", "medium", "low")}
        print("%s -> %s" % (os.path.basename(a.old), os.path.basename(a.new)))
        print("%d findings: %d high, %d medium, %d low"
              % (len(findings), counts["high"], counts["medium"], counts["low"]))
        if waived:
            print("%d finding%s waived" % (len(waived), "" if len(waived) == 1 else "s"))
        if orphaned:
            # A waiver that matches nothing usually means the change it covered
            # was edited again, so the acceptance no longer describes what is
            # there now. Silence would let it rot unnoticed.
            print("%d waiver%s no longer match%s anything: %s"
                  % (len(orphaned), "" if len(orphaned) == 1 else "s",
                     "es" if len(orphaned) == 1 else "",
                     ", ".join(w.get("ref") or w["id"] for w in orphaned[:5])))
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
            if len(f.chain) > 1:
                print("%s %-22s   moves %s"
                      % (" " * 6, "", "  ->  ".join(f.chain[1:])))
            for o in f.outputs[:2]:
                print("%s %-22s   %s %s: %s -> %s (%.1f%%)"
                      % (" " * 6, "", o["ref"], o["label"] or "output",
                         o["before"], o["after"], o["move"] * 100))
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
            "chain": f.chain, "downstream": f.downstream,
            "outputs": f.outputs, "output_impact": f.output_impact,
        } for f in findings]
        with open(a.json_out, "w", encoding="utf-8") as fh:
            json.dump({"findings": payload, "stale_values": result["stale"],
                       "outputs": result["outputs"],
                       "outputs_declared": result["outputs_declared"]},
                      fh, indent=2)

    if a.fail_on != "never":
        limit = SEV_RANK[a.fail_on]
        if any(SEV_RANK[f.severity] <= limit for f in findings):
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
