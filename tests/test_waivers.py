"""Waivers: the decision to accept a finding has to outlive the session.

The load-bearing property is that a waiver written in the browser matches the
same finding in CI. Both engines compute the id, so this pins the id itself as
well as the behaviour built on it.

    python tests/test_waivers.py
"""

import io
import json
import os
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "cli"))

import waivers as waiver_file        # noqa: E402
from engine import compare, fingerprint   # noqa: E402

A = os.path.join(ROOT, "examples", "sample_before.xlsx")
B = os.path.join(ROOT, "examples", "sample_after.xlsx")

failures = []


def check(cond, label):
    print(("  ok   " if cond else "  FAIL ") + label)
    if not cond:
        failures.append(label)


def main():
    result = compare(A, B)
    findings = result["findings"]
    plug = [f for f in findings if f.kind == "hardcode"][0]

    print("finding ids")
    check(len(plug.id) == 8 and all(c in "0123456789abcdef" for c in plug.id),
          "an id is eight hex characters (%s)" % plug.id)
    check(len(set(f.id for f in findings)) == len(findings),
          "every finding in this comparison has a distinct id")
    check(compare(A, B)["findings"][0].id == findings[0].id,
          "ids are stable across runs")

    # A waiver records that a human accepted a specific change. Edit the change
    # and the acceptance no longer applies, so the id has to move with it.
    check(fingerprint("hardcode", "S", "C14", "=A1*2", "500") !=
          fingerprint("hardcode", "S", "C14", "=A1*2", "600"),
          "editing the change itself produces a different id")
    check(fingerprint("hardcode", "S", "C14", "=A1*2", "500") ==
          fingerprint("hardcode", "S", "C14", "=A1*2", "500"),
          "and an identical change produces the same one")

    print("applying a waiver file")
    tmp = tempfile.mkdtemp()
    path = os.path.join(tmp, "waivers.json")
    with io.open(path, "w", encoding="utf-8") as fh:
        json.dump({"version": 1, "waivers": [
            {"id": plug.id, "ref": "Revenue build!D9", "kind": "hardcode",
             "reason": "Q3 plug agreed with the deal team"},
            {"id": "deadbeef", "ref": "Revenue build!Z99", "kind": "hardcode",
             "reason": "covers something that no longer exists"},
        ]}, fh)

    loaded = waiver_file.load(path)
    fresh = compare(A, B)["findings"]
    live = waiver_file.apply(fresh, loaded)
    check(len(live) == len(fresh) - 1, "the waived finding drops out of the live list")
    check([f for f in fresh if f.waived][0].kind == "hardcode",
          "and is marked with the waiver that covered it")
    check([w["id"] for w in waiver_file.stale(fresh, loaded)] == ["deadbeef"],
          "a waiver matching nothing is reported rather than ignored")

    check(waiver_file.load(os.path.join(tmp, "absent.json")) == {},
          "a missing waiver file is not an error")
    with io.open(os.path.join(tmp, "bad.json"), "w", encoding="utf-8") as fh:
        fh.write("{\"nope\": 1}")
    try:
        waiver_file.load(os.path.join(tmp, "bad.json"))
        check(False, "a file that is not a waiver file is rejected")
    except waiver_file.WaiverError:
        check(True, "a file that is not a waiver file is rejected")

    print("the gate")
    cli = os.path.join(ROOT, "cli", "cli.py")
    out = os.path.join(tmp, "r.html")

    def run(*extra):
        return subprocess.run(
            [sys.executable, cli, A, B, "--quiet", "--out", out,
             "--fail-on", "high"] + list(extra),
            capture_output=True, text=True, cwd=ROOT).returncode

    check(run() == 1, "high findings fail the gate")
    every = os.path.join(tmp, "all.json")
    subprocess.run([sys.executable, cli, A, B, "--quiet", "--out", out,
                    "--write-waivers", every], cwd=ROOT, capture_output=True)
    check(run("--waivers", every) == 0,
          "waiving everything passes it, which is what makes the gate usable")
    check(run("--waivers", path) == 1,
          "waiving one plug does not waive the rest")

    print("")
    if failures:
        print("%d check(s) failed" % len(failures))
        return 1
    print("all checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
