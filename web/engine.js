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

var EXTLINK = /\[[^\[\]]+\]/g;

/* Workbook references embedded in a formula, as [book.xlsx] tokens. */
function externalLinks(f) {
  if (typeof f !== "string") return [];
  var m = f.match(EXTLINK);
  if (!m) return [];
  return m.filter(function (x, i) { return m.indexOf(x) === i; }).sort();
}

/* Normalise an A1 formula to R1C1 so a copied formula compares equal.
   Defined names are left alone: a name like TAX1 is shaped exactly like a cell
   reference, and rewriting it relative to its own position makes the same
   formula look different once it moves down a row — a false change on every
   formula that uses a name, in any file where a row was inserted above it. */
function toR1C1(formula, row, col, names) {
  if (typeof formula !== "string" || formula.charAt(0) !== "=") return formula;
  var stash = [];
  var body = formula.replace(QUOTED, function (m) { stash.push(m); return "\u0000" + (stash.length - 1) + "\u0000"; });
  body = body.replace(REF, function (all, pre, cabs, letters, rabs, digits) {
    if (names && names[(letters + digits).toUpperCase()]) return all;
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

var ERR_CODE = { 0: "#NULL!", 7: "#DIV/0!", 15: "#VALUE!", 23: "#REF!",
  29: "#NAME?", 36: "#NUM!", 42: "#N/A", 43: "#GETTING_DATA" };

function isFormula(cell) { return cell && typeof cell.f === "string" && cell.f.charAt(0) === "="; }
function isEmpty(cell) { return !cell || (cell.f === null && cell.v === null); }
var EMPTY = { f: null, v: null, r1c1: null };

/* Workbook-level defined names, as NAME -> the reference it points at. Sheet
   local names are out of scope: they cannot be referenced from elsewhere, so a
   change to one is already visible as a formula change. */
function readNames(wb) {
  var out = {};
  var list = wb && wb.Workbook && wb.Workbook.Names;
  if (!list) return out;
  for (var i = 0; i < list.length; i++) {
    var n = list[i];
    if (!n || !n.Name || n.Sheet !== undefined && n.Sheet !== null) continue;
    out[n.Name] = String(n.Ref === undefined ? "" : n.Ref);
  }
  return out;
}

function readSheet(ws, name, names) {
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
    /* SheetJS hands back an error as a numeric code; openpyxl hands back the
       text. Normalise to the text here so both engines see the same cell. */
    var val = src.v === undefined ? null : src.v;
    if (src.t === "e") val = typeof src.w === "string" ? src.w : (ERR_CODE[src.v] || "#N/A");
    var formula = src.f !== undefined ? "=" + src.f : val;
    var cell = { f: formula, v: val, r1c1: null };
    cell.r1c1 = isFormula(cell) ? toR1C1(formula, r, c, names) : formula;
    grid.cells[r + ":" + c] = cell;
    if (r > grid.maxRow) grid.maxRow = r;
    if (c > grid.maxCol) grid.maxCol = c;
  }
  return grid;
}

function get(grid, r, c) { return grid.cells[r + ":" + c] || EMPTY; }

/* Run-length normalise a shape string: "tnnnn" and "tnnnnn" both read "tn".
   The fingerprint records the pattern of cell kinds along the row, not how
   many columns each run happens to span. Without this an inserted column
   changes the shape of every row at once and row alignment collapses, so the
   two axes would each need the other to have been solved first. */
function collapse(shape) {
  var out = "";
  for (var i = 0; i < shape.length; i++)
    if (shape.charAt(i) !== out.charAt(out.length - 1)) out += shape.charAt(i);
  return out;
}

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
  return label.toLowerCase() + "|" + collapse(shape);
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

/* Fingerprint for a column: its header text plus a shape string, read over the
   rows that already aligned rather than over every row in the sheet. A single
   inserted row would otherwise perturb every column signature at once, and two
   sequences that share nothing align by position — which is exactly the
   over-reporting that column alignment exists to prevent. */
function columnSignature(grid, col, rows) {
  var label = "", shape = "";
  for (var i = 0; i < rows.length; i++) {
    var cell = grid.cells[rows[i] + ":" + col];
    if (!cell || isEmpty(cell)) { shape += "."; continue; }
    if (isFormula(cell)) shape += "f";
    else if (typeof cell.f === "string") {
      shape += "t";
      if (!label) label = cell.f.trim().slice(0, 40);
    } else shape += "n";
  }
  return label.toLowerCase() + "|" + shape;
}

function alignColumns(left, right, rowPairs) {
  var lrows = rowPairs.map(function (p) { return p[0]; });
  var rrows = rowPairs.map(function (p) { return p[1]; });
  var ls = [], rs = [];
  for (var c = 1; c <= left.maxCol; c++) ls.push(columnSignature(left, c, lrows));
  for (var c2 = 1; c2 <= right.maxCol; c2++) rs.push(columnSignature(right, c2, rrows));
  return lcsOpcodes(ls, rs);
}

/* Formulas adjacent to a cell in one direction, stopping at the first gap. A
   run is contiguous. Reading past a blank row or a label into the next block
   of the sheet is how unrelated formulas end up treated as neighbours, which
   reads either as noise or as a break that is not there. */
function walk(grid, row, col, dr, dc, limit) {
  var out = [];
  for (var k = 1; k <= (limit || 4); k++) {
    var n = get(grid, row + dr * k, col + dc * k);
    if (isEmpty(n) || !isFormula(n)) break;
    out.push(n.r1c1);
  }
  return out;
}

/* Is this cell the odd one out in an otherwise uniform run? axis "row" reads
   the horizontal neighbours, "column" the vertical ones. Both matter: filling
   across a row and dragging down a column break a deliberate exception in
   exactly the same way. */
function findRunBreak(grid, row, col, axis) {
  var cell = get(grid, row, col);
  var dr = axis === "row" ? 0 : 1, dc = axis === "row" ? 1 : 0;
  var fb = walk(grid, row, col, -dr, -dc), fa = walk(grid, row, col, dr, dc);
  var nb = fb.concat(fa), i;
  if (nb.length < 2) return null;
  for (i = 1; i < nb.length; i++) if (nb[i] !== nb[0]) return null;

  /* A column of identical formulas nearly always ends in a total that is
     meant to differ, so a vertical break only counts when the run continues
     on both sides of the odd cell. Rows do not need this: a row of quarters
     is a homogeneous series to its last column, and requiring both sides
     there would blind the tool to a break in the final period. */
  if (axis === "column" && (!fb.length || !fa.length)) return null;

  if (isFormula(cell) && cell.r1c1 === nb[0]) return null;
  return nb[0];
}

/* What the neighbours agree on, and which axis they lie along. Horizontal is
   tried first: it is the more common accident, and settling on one axis keeps
   a single cell from producing two findings for the same mistake. */
function runContext(grid, row, col) {
  var axes = ["row", "column"];
  for (var i = 0; i < axes.length; i++) {
    var expected = findRunBreak(grid, row, col, axes[i]);
    if (expected !== null) return { expected: expected, axis: axes[i] };
  }
  return { expected: null, axis: null };
}

var ERROR = /^#(REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NULL!|NUM!|SPILL!|CALC!)$/;

/* The Excel error a cached value carries, or null if it holds a result. */
function errorText(v) {
  if (typeof v !== "string") return null;
  var s = v.trim();
  return ERROR.test(s) ? s : null;
}
/* A formula whose text has lost a reference reads #REF! literally. */
function brokenRef(f) { return typeof f === "string" && f.indexOf("#REF!") !== -1; }

/* Errors outrank every other reading of a cell. A formula that changed and a
   formula that now returns #DIV/0! are the same edit, but only one of them is
   worth waking someone for, so the error is what gets reported and the formula
   change rides along in before/after. */
function errorFinding(name, ref, lc, rc, base) {
  var lb = brokenRef(lc.f), rb = brokenRef(rc.f);
  if (rb && !lb)
    return F("high", "broken_reference", name, ref,
      "Formula lost a reference and now reads #REF!",
      Object.assign({ detail: "Whatever this pointed at was deleted. The cell cannot recalculate.",
        before: fmt(lc.f), after: fmt(rc.f) }, base));

  var le = errorText(lc.v), re = errorText(rc.v);
  if (re && !le)
    return F("high", "error_introduced", name, ref,
      "This cell now evaluates to " + re,
      Object.assign({ detail: "It held a value in the baseline.",
        before: fmt(lc.f), after: fmt(rc.f) }, base));
  if (le && !re)
    return F("low", "error_cleared", name, ref, "An error here was fixed",
      Object.assign({ detail: "The baseline returned " + le + ".",
        before: fmt(lc.f), after: fmt(rc.f) }, base));
  if (lb && !rb)
    return F("low", "error_cleared", name, ref, "A broken reference here was repaired",
      Object.assign({ before: fmt(lc.f), after: fmt(rc.f) }, base));
  return null;
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
  var cal = alignColumns(left, right, al.pairs);
  al.colPairs = cal.pairs; al.colAdded = cal.added;
  al.colDeleted = cal.deleted; al.colRatio = cal.ratio;
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

  var lrows = al.pairs.map(function (p) { return p[0]; });
  var rrows = al.pairs.map(function (p) { return p[1]; });
  cal.added.forEach(function (c) {
    var label = columnSignature(right, c, rrows).split("|")[0];
    findings.push(F("low", "column_added", name, "column " + colLetter(c),
      "Column inserted" + (label ? ": " + label : "")));
  });
  cal.deleted.forEach(function (c) {
    var label = columnSignature(left, c, lrows).split("|")[0];
    findings.push(F("medium", "column_deleted", name, "column " + colLetter(c),
      "Column deleted" + (label ? ": " + label : "")));
  });

  al.pairs.forEach(function (p) {
    var lr = p[0], rr = p[1];
    cal.pairs.forEach(function (cp) {
      var lcol = cp[0], rcol = cp[1];
      var lc = get(left, lr, lcol), rc = get(right, rr, rcol);
      if (isEmpty(lc) && isEmpty(rc)) return;
      var c = rcol;
      var ref = colLetter(rcol) + rr;
      var vb = fmt(lc.v), va = fmt(rc.v), mag = valueDelta(lc.v, rc.v);
      var base = { vb: vb, va: va, mag: mag, lr: lr, rr: rr };

      var err = errorFinding(name, ref, lc, rc, base);
      if (err) { findings.push(err); return; }

      /* A formula that now reads from a different workbook is a change of
         source, not of logic, and the diff of the text buries it. */
      var lx = externalLinks(lc.f).join(", "), rx = externalLinks(rc.f).join(", ");
      if (lx !== rx && (lx || rx)) {
        findings.push(F("high", "external_link_changed", name, ref,
          "Formula reads from a different workbook",
          Object.assign({
            detail: "Was " + (lx || "no external link") + ", now " + (rx || "no external link") + ".",
            before: fmt(lc.f), after: fmt(rc.f)
          }, base)));
        return;
      }

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
        var hc = runContext(right, rr, rcol);
        findings.push(F("high", "hardcode", name, ref, "Hardcode replaced a formula",
          Object.assign({
            detail: hc.expected ? "The " + hc.axis + " it sits in is still calculated."
              : "This cell no longer recalculates.",
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
        var now = runContext(right, rr, rcol), was = runContext(left, lr, lcol);
        if (now.expected !== null) {
          findings.push(F("high", "run_break", name, ref,
            "Formula breaks the run in its " + now.axis,
            Object.assign({ detail: "Neighbours use " + now.expected,
              before: fmt(lc.f), after: fmt(rc.f) }, base)));
        } else if (was.expected !== null) {
          findings.push(F("low", "run_repaired", name, ref,
            "Formula now matches the rest of its " + was.axis,
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

/* Defined names break formulas silently when they move. A deleted name leaves
   every formula that used it unresolvable, and a repointed name changes what
   those formulas mean without changing a character of their text — the one
   edit a formula diff cannot see. */
function diffNames(a, b, findings) {
  var an = Object.keys(a).sort(), bn = Object.keys(b).sort();
  bn.forEach(function (n) {
    if (!(n in a)) findings.push(F("low", "name_added", "", n,
      "Named range added: " + n, { after: b[n] }));
  });
  an.forEach(function (n) {
    if (!(n in b)) findings.push(F("high", "name_deleted", "", n,
      "Named range deleted: " + n,
      { detail: "Any formula that referred to it can no longer resolve.", before: a[n] }));
  });
  an.forEach(function (n) {
    if (n in b && a[n] !== b[n]) findings.push(F("high", "name_changed", "", n,
      "Named range " + n + " now points somewhere else",
      { detail: "Formulas using it changed meaning without changing text.",
        before: a[n], after: b[n] }));
  });
}

function compare(wbA, wbB) {
  var left = {}, right = {}, findings = [], align = {}, stale = false;
  var namesA = readNames(wbA), namesB = readNames(wbB);
  var upA = {}, upB = {};
  for (var ka in namesA) upA[ka.toUpperCase()] = 1;
  for (var kb in namesB) upB[kb.toUpperCase()] = 1;
  wbA.SheetNames.forEach(function (n) { left[n] = readSheet(wbA.Sheets[n], n, upA); });
  wbB.SheetNames.forEach(function (n) { right[n] = readSheet(wbB.Sheets[n], n, upB); });

  diffNames(namesA, namesB, findings);

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

  return { left: left, right: right, findings: findings, align: align, stale: stale,
    names: [namesA, namesB] };
}

root.xldiff = { compare: compare, toR1C1: toR1C1, fmt: fmt, colLetter: colLetter, get: get };

})(typeof module !== "undefined" && module.exports ? module.exports : (typeof window !== "undefined" ? window : this));
