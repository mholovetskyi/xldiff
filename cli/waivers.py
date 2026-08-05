"""Waivers: findings a human looked at and accepted, recorded in a file.

Suppressing a finding in the browser lasts as long as the tab does. A gate that
re-reports the same accepted exception every run gets switched off within a
fortnight, so the decision has to outlive the session and live next to the
model in version control.

The file is JSON so it diffs and reviews like anything else:

    {
      "version": 1,
      "waivers": [
        {"id": "3f2a91cc",
         "ref": "Revenue build!C14",
         "kind": "hardcode",
         "reason": "Q3 plug agreed with the deal team, unwinds at close",
         "who": "mh",
         "when": "2026-08-04"}
      ]
    }

Only "id" is load-bearing. The rest is there so a reader can tell what was
waived and why without running anything.
"""

import json
import os


class WaiverError(Exception):
    pass


def load(path):
    """Read a waiver file. Missing file is not an error; an unreadable one is."""
    if not path:
        return {}
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
    except (ValueError, OSError) as exc:
        raise WaiverError("could not read waiver file %s: %s" % (path, exc))
    if not isinstance(data, dict) or not isinstance(data.get("waivers"), list):
        raise WaiverError("%s is not a waiver file: expected a top-level "
                          '"waivers" list' % path)
    out = {}
    for entry in data["waivers"]:
        if isinstance(entry, dict) and entry.get("id"):
            out[str(entry["id"])] = entry
    return out


def apply(findings, waivers):
    """Mark findings a waiver covers. Returns the ones that survive."""
    live = []
    for f in findings:
        entry = waivers.get(f.id)
        if entry is None:
            live.append(f)
            continue
        f.waived = True
        f.waiver = entry
    return live


def stale(findings, waivers):
    """Waivers matching nothing in this comparison.

    Worth surfacing: a waiver that stopped matching usually means the change it
    covered was edited again, and the acceptance no longer applies to what is
    there now.
    """
    seen = set(f.id for f in findings)
    return [w for wid, w in sorted(waivers.items()) if wid not in seen]


def dump(findings, who="", when=""):
    """A waiver file covering the findings given, ready to edit and commit."""
    return {
        "version": 1,
        "waivers": [{
            "id": f.id,
            "ref": (f.sheet + "!" + f.ref) if f.ref else f.sheet,
            "kind": f.kind,
            "reason": "",
            "who": who,
            "when": when,
        } for f in findings],
    }
