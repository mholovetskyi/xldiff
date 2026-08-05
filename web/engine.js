/* xldiff engine, browser build. Pure function over two SheetJS workbooks.
   No network calls, no state outside the returned object. */
(function (root) {
"use strict";

var MAX_COL = 60, MAX_ROW = 5000, IMPACT_CAP = 25, ALIGN_LIMIT = 2500;
var BORING = [0, 1, -1, 2, 12, 100, 365, 360, 1000, 0.5];
var SEV_ORDER = { high: 0, medium: 1, low: 2 };

var QUOTED = /"[^"]*"/g;
var REF = /(^|[^A-Za-z0-9_$.])(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]{0,6})(?![0-9(])/g;

function colToNum(s) {
  var n = 0;
  for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}
function colLetter(n) {
  var s = "";
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}

function toR1C1(formula, row, col) {
  if (typeof formula !== "string" || formula.charAt(0) !== "=") return formula;
  var stash = [];
  var body = formula.replace(QUOTED, function (m) { stash.push(m); return "\u0000" + (stash.length - 1) + "\u0000"; });
  body = body.replace(REF, function (all, pre, cabs, letters, rabs, digits) {
    var c = colToNum(letters), r = parseInt(digits, 10);
    var rp = rabs ? "R" + r : (r !== row ? "R[" + (r - row) + "]" : "R");
    var cp = cabs ? "C" + c : (c !== col ? "C[" + (c - col) + "]" : "C");
    return pre + rp + cp;
  });
  for (var i = 0; i < stash.length; i++) body = body.split("\u0000" + i + "\u0000").join(stash[i]);
  return body;
}

function literalsIn(formula) {
  if (typeof formula !== "string" || formula.charAt(0) !== "=") return [];
  var body = formula.replace(QUOTED, "").replace(REF, "$1");
  var out = [], m, re = /(^|[^A-Za-z0-9_.\]])(\d+(?:\.\d+)?)/g;
  while ((m = re.exec(body))) {
    var v = parseFloat(m[2]);
    if (BORING.indexOf(v) === -1) out.push(v);
  }
  return out;
}

function fmt(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") {
    if (Math.abs(v) < 1e15 && v === Math.round(v)) return v.toLocaleString("en-US");
    return parseFloat(v.toFixed(4)).toLocaleString("en-US", { maximumFractionDigits: 4 });
  }
  return String(v);
}

function isFormula(cell) { return cell && typeof cell.f === "string" && cell.f.charAt(0) === "="; }
function isEmpty(cell) { return !cell || (cell.f === null && cell.v === null); }
var EMPTY = { f: null, v: null, r1c1: null };

function readSheet(ws, name) {
  var grid = { name: name, cells: {}, maxRow: 0, maxCol: 0 };
  if (!ws || !ws["!ref"]) return grid;
  for (var key in ws) {
    if (key.charAt(0) === "!") continue;
    var a = key.match(/^([A-Z]{1,3})([1-9][0-9]{0,6})$/);
    if (!a) continue;
    var c = colToNum(a[1]), r = parseInt(a[2], 10);
    if (r > MAX_ROW || c > MAX_COL) continue;
    var src = ws[key];
    if (src.f === undefined && src.v === undefined) continue;
    var formula = src.f !== undefined ? "=" + src.f : (src.v !== undefined ? src.v : null);
    var cell = { f: formula, v: src.v === undefined ? null : src.v, r1c1: null };
    cell.r1c1 = isFormula(cell) ? toR1C1(formula, r, c) : formula;
    grid.cells[r + ":" + c] = cell;
    if (r > grid.maxRow) grid.maxRow = r;
    if (c > grid.maxCol) grid.maxCol = c;
  }
  return grid;
}

function get(grid, r, c) { return grid.cells[r + ":" + c] || EMPTY; }

function rowSignature(grid, row) {
  var label = "", shape = "";
  for (var c = 1; c <= grid.maxCol; c++) {
    var cell = grid.cells[row + ":" + c];
    if (!cell || isEmpty(cell)) { shape += "."; continue; }
    if (isFormula(cell)) shape += "f";
    else if (typeof cell.f === "string") {
      shape += "t";
      if (!label) label = cell.f.trim().slice(0, 40);
    } else shape += "n";
  }
  return label.toLowerCase() + "|" + shape;
}

/* LCS over row fingerprints, after stripping the common prefix and suffix. */
function lcsOpcodes(a, b) {
  var n = a.length, m = b.length, lo = 0;
  while (lo < n && lo < m && a[lo] === b[lo]) lo++;
  var hiA = n, hiB = m;
  while (hiA > lo && hiB > lo && a[hiA - 1] === b[hiB - 1]) { hiA--; hiB--; }
  var A = a.slice(lo, hiA), B = b.slice(lo, hiB);
  var pairs = [], added = [], deleted = [], matches = lo + (n - hiA);

  for (var k = 0; k < lo; k++) pairs.push([k + 1, k + 1]);

  if (A.length && B.length) {
    if (A.length * B.length > ALIGN_LIMIT * ALIGN_LIMIT) {
      var q = Math.min(A.length, B.length);
      for (var i = 0; i < q; i++) { pairs.push([lo + i + 1, lo + i + 1]); if (A[i] === B[i]) matches++; }
      for (var i2 = q; i2 < A.length; i2++) deleted.push(lo + i2 + 1);
      for (var j2 = q; j2 < B.length; j2++) added.push(lo + j2 + 1);
    } else {
      var w = B.length + 1;
      var dp = new Int32Array((A.length + 1) * w);
      for (var i3 = A.length - 1; i3 >= 0; i3--)
        for (var j3 = B.length - 1; j3 >= 0; j3--)
          dp[i3 * w + j3] = A[i3] === B[j3]
            ? dp[(i3 + 1) * w + j3 + 1] + 1
            : Math.max(dp[(i3 + 1) * w + j3], dp[i3 * w + j3 + 1]);
      var i4 = 0, j4 = 0, pendD = [], pendA = [];
      function flush() {
        var q2 = Math.min(pendD.length, pendA.length);
        for (var t = 0; t < q2; t++) pairs.push([pendD[t], pendA[t]]);
        for (var t2 = q2; t2 < pendD.length; t2++) deleted.push(pendD[t2]);
        for (var t3 = q2; t3 < pendA.length; t3++) added.push(pendA[t3]);
        pendD = []; pendA = [];
      }
      while (i4 < A.length && j4 < B.length) {
        if (A[i4] === B[j4]) {
          flush(); pairs.push([lo + i4 + 1, lo + j4 + 1]); matches++; i4++; j4++;
        } else if (dp[(i4 + 1) * w + j4] >= dp[i4 * w + j4 + 1]) {
          pendD.push(lo + i4 + 1); i4++;
        } else { pendA.push(lo + j4 + 1); j4++; }
      }
      while (i4 < A.length) { pendD.push(lo + i4 + 1); i4++; }
      while (j4 < B.length) { pendA.push(lo + j4 + 1); j4++; }
      flush();
    }
  } else {
    for (var i5 = 0; i5 < A.length; i5++) deleted.push(lo + i5 + 1);
    for (var j5 = 0; j5 < B.length; j5++) added.push(lo + j5 + 1);
  }

  for (var k2 = 0; k2 < n - hiA; k2++) pairs.push([hiA + k2 + 1, hiB + k2 + 1]);
  pairs.sort(function (x, y) { return (x[1] - y[1]) || (x[0] - y[0]); });
  var ratio = (n + m) ? (2 * matches) / (n + m) : 1;
  return { pairs: pairs, added: added, deleted: deleted, ratio: ratio };
}

function alignRows(left, right) {
  var ls = [], rs = [];
  for (var r = 1; r <= left.maxRow; r++) ls.push(rowSignature(left, r));
  for (var r2 = 1; r2 <= right.maxRow; r2++) rs.push(rowSignature(right, r2));
  return lcsOpcodes(ls, rs);
}

function findRunBreak(grid, row, col) {
  var cell = get(grid, row, col), nb = [];
  for (var c = Math.max(1, col - 4); c <= Math.min(grid.maxCol, col + 4); c++) {
    if (c === col) continue;
    var n = get(grid, row, c);
    if (!isEmpty(n) && isFormula(n)) nb.push(n.r1c1);
  }
  if (nb.length < 2) return null;
  for (var i = 1; i < nb.length; i++) if (nb[i] !== nb[0]) return null;
  if (isFormula(cell) && cell.r1c1 === nb[0]) return null;
  return nb[0];
}

function valueDelta(a, b) {
  if (typeof a === "number" && typeof b === "number") {
    if (a === b) return 0;
    return Math.abs(b - a) / Math.max(Math.abs(a), 1e-9);
  }
  return a === b ? 0 : 1;
}

function F(sev, kind, sheet, ref, summary, o) {
  o = o || {};
  return {
    severity: sev, kind: kind, sheet: sheet, ref: ref, summary: summary,
    detail: o.detail || "", before: o.before || "", after: o.after || "",
    valBefore: o.vb || "", valAfter: o.va || "", magnitude: o.mag || 0,
    rowLeft: o.lr || 0, rowRight: o.rr || 0
  };
}

function diffSheet(name, left, right, findings, align) {
  var al = alignRows(left, right);
  align[name] = al;
  if (al.ratio < 0.35)
    findings.push(F("low", "alignment", name, "",
      "Row alignment confidence is low (" + Math.round(al.ratio * 100) + "%)",
      { detail: "The two sheets may not be versions of each other. Findings below may be noisy." }));

  al.added.forEach(function (r) {
    var label = rowSignature(right, r).split("|")[0];
    findings.push(F("low", "row_added", name, "row " + r,
      "Row inserted" + (label ? ": " + label : ""), { rr: r }));
  });
  al.deleted.forEach(function (r) {
    var label = rowSignature(left, r).split("|")[0];
    findings.push(F("low", "row_deleted", name, "row " + r,
      "Row deleted" + (label ? ": " + label : ""), { lr: r }));
  });

  al.pairs.forEach(function (p) {
    var lr = p[0], rr = p[1], cols = {};
    for (var c = 1; c <= Math.max(left.maxCol, right.maxCol); c++)
      if (left.cells[lr + ":" + c] || right.cells[rr + ":" + c]) cols[c] = 1;
    Object.keys(cols).map(Number).sort(function (a, b) { return a - b; }).forEach(function (c) {
      var lc = get(left, lr, c), rc = get(right, rr, c);
      var ref = colLetter(c) + rr;
      var vb = fmt(lc.v), va = fmt(rc.v), mag = valueDelta(lc.v, rc.v);
      var base = { vb: vb, va: va, mag: mag, lr: lr, rr: rr };

      if (lc.r1c1 === rc.r1c1) {
        if (lc.v !== rc.v && !isEmpty(lc))
          findings.push(F("medium", "impact", name, ref,
            "Value moved with no edit to this cell",
            Object.assign({ detail: "Downstream of a change elsewhere." }, base)));
        return;
      }
      if (isEmpty(lc)) {
        findings.push(F("low", "cell_added", name, ref, "Cell added",
          Object.assign({ after: fmt(rc.f) }, base)));
        return;
      }
      if (isEmpty(rc)) {
        findings.push(F("medium", "cell_deleted", name, ref, "Cell cleared",
          Object.assign({ before: fmt(lc.f) }, base)));
        return;
      }
      if (isFormula(lc) && !isFormula(rc)) {
        var expected = findRunBreak(right, rr, c);
        findings.push(F("high", "hardcode", name, ref, "Hardcode replaced a formula",
          Object.assign({
            detail: expected ? "The row it sits in is still calculated." : "This cell no longer recalculates.",
            before: fmt(lc.f), after: fmt(rc.f)
          }, base)));
        return;
      }
      if (!isFormula(lc) && isFormula(rc)) {
        findings.push(F("medium", "constant_to_formula", name, ref,
          "Constant replaced by a formula",
          Object.assign({ before: fmt(lc.f), after: fmt(rc.f) }, base)));
        return;
      }
      if (isFormula(lc) && isFormula(rc)) {
        var exp = findRunBreak(right, rr, c), was = findRunBreak(left, lr, c);
        if (exp !== null) {
          findings.push(F("high", "run_break", name, ref, "Formula breaks the run in its row",
            Object.assign({ detail: "Neighbours use " + exp, before: fmt(lc.f), after: fmt(rc.f) }, base)));
        } else if (was !== null) {
          findings.push(F("low", "run_repaired", name, ref, "Formula now matches the rest of its row",
            Object.assign({ before: fmt(lc.f), after: fmt(rc.f) }, base)));
        } else {
          var old = literalsIn(lc.f), lits = literalsIn(rc.f).filter(function (x) { return old.indexOf(x) === -1; });
          lits = lits.filter(function (x, i) { return lits.indexOf(x) === i; }).sort(function (a, b) { return a - b; });
          findings.push(F(lits.length ? "high" : "medium", "formula_changed", name, ref,
            "Formula changed" + (lits.length ? " and gained a hardcoded number" : ""),
            Object.assign({
              detail: lits.length ? "New numeric literal: " + lits.map(fmt).join(", ") : "",
              before: fmt(lc.f), after: fmt(rc.f)
            }, base)));
        }
        return;
      }
      findings.push(F("medium", "constant_changed", name, ref, "Constant changed",
        Object.assign({ before: fmt(lc.f), after: fmt(rc.f) }, base)));
    });
  });
}

function compare(wbA, wbB) {
  var left = {}, right = {}, findings = [], align = {}, stale = false;
  wbA.SheetNames.forEach(function (n) { left[n] = readSheet(wbA.Sheets[n], n); });
  wbB.SheetNames.forEach(function (n) { right[n] = readSheet(wbB.Sheets[n], n); });

  [left, right].forEach(function (side) {
    for (var n in side) for (var k in side[n].cells) {
      var c = side[n].cells[k];
      if (isFormula(c) && (c.v === null || c.v === undefined)) { stale = true; return; }
    }
  });

  wbB.SheetNames.forEach(function (n) {
    if (!left[n]) findings.push(F("medium", "sheet_added", n, "", "Sheet added: " + n));
  });
  wbA.SheetNames.forEach(function (n) {
    if (!right[n]) findings.push(F("high", "sheet_deleted", n, "", "Sheet deleted: " + n));
  });
  wbA.SheetNames.forEach(function (n) {
    if (right[n]) diffSheet(n, left[n], right[n], findings, align);
  });

  var impacts = findings.filter(function (f) { return f.kind === "impact"; });
  if (impacts.length > IMPACT_CAP) {
    impacts.sort(function (a, b) { return b.magnitude - a.magnitude; });
    var keep = impacts.slice(0, IMPACT_CAP);
    var supp = impacts.length - IMPACT_CAP;
    findings = findings.filter(function (f) { return f.kind !== "impact" || keep.indexOf(f) !== -1; });
    findings.push(F("low", "impact_capped", "", "",
      supp + " further cells moved in value with no edit",
      { detail: "Showing the " + IMPACT_CAP + " largest by relative change." }));
  }

  findings.sort(function (a, b) {
    return (SEV_ORDER[a.severity] - SEV_ORDER[b.severity]) ||
      (a.sheet < b.sheet ? -1 : a.sheet > b.sheet ? 1 : 0) ||
      ((a.rowRight || a.rowLeft) - (b.rowRight || b.rowLeft));
  });

  return { left: left, right: right, findings: findings, align: align, stale: stale };
}

root.xldiff = { compare: compare, toR1C1: toR1C1, fmt: fmt, colLetter: colLetter, get: get };

})(typeof module !== "undefined" && module.exports ? module.exports : (typeof window !== "undefined" ? window : this));
