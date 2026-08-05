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

var PREC = new RegExp(
  "(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?" +
  "\\$?([A-Z]{1,3})\\$?([1-9][0-9]{0,6})" +
  "(?::\\$?([A-Z]{1,3})\\$?([1-9][0-9]{0,6}))?" +
  "(?![0-9(])", "g");
var RANGE_CAP = 400;
/* Cell keys join sheet, row and column with a character that cannot appear
   in a sheet name. Without a separator, row 4 column 2 and row 42 column 1
   produce the same key. */
var SEP = "\u0001";

/* Cells a formula reads, as "sheet\u0001row\u0001col", ranges expanded. Enough to
   answer "what does this feed", which is the question the impact findings
   could not answer before: they said a value moved, never what moved it. Not a
   parser — array formulas, structured table references and INDIRECT are all
   invisible here, and a missed edge only ever shortens a chain rather than
   inventing one. */
function precedents(formula, sheet, names) {
  if (typeof formula !== "string" || formula.charAt(0) !== "=") return [];
  var body = formula.replace(QUOTED, ""), out = [], total = 0, n;
  if (names) {
    for (n in names) {
      var pat = new RegExp("(^|[^A-Za-z0-9_.])" + n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") +
        "(?![A-Za-z0-9_.])", "gi");
      if (pat.test(body)) {
        out = out.concat(precedents("=" + names[n], sheet));
        /* Blank the token out. A name shaped like a reference — TAX1 — would
           otherwise also be read as cell TAX1, inventing an edge to a cell
           13,570 columns out. */
        body = body.replace(pat, "$1 ");
      }
    }
  }
  var m;
  PREC.lastIndex = 0;
  while ((m = PREC.exec(body)) !== null) {
    if (m.index === PREC.lastIndex) PREC.lastIndex++;
    var pre = m.index ? body.charAt(m.index - 1) : "";
    if (!m[1] && !m[2] && /[A-Za-z0-9_.]/.test(pre)) continue;
    var where = m[1] || m[2] || sheet;
    var r1 = parseInt(m[4], 10), c1 = colToNum(m[3]);
    var r2 = m[6] ? parseInt(m[6], 10) : r1, c2 = m[5] ? colToNum(m[5]) : c1;
    var rlo = Math.min(r1, r2), rhi = Math.max(r1, r2);
    var clo = Math.min(c1, c2), chi = Math.max(c1, c2);
    if ((rhi - rlo + 1) * (chi - clo + 1) > RANGE_CAP) continue;
    for (var r = rlo; r <= rhi; r++)
      for (var c = clo; c <= chi; c++) {
        out.push(where + SEP + r + SEP + c);
        if (++total > RANGE_CAP * 4) return out;
      }
  }
  return out;
}

/* Reverse dependency map: a cell -> the cells that read it. */
function buildGraph(grids, names) {
  var dependents = {};
  for (var sname in grids) {
    var g = grids[sname];
    for (var k in g.cells) {
      var cell = g.cells[k];
      if (!isFormula(cell)) continue;
      var parts = k.split(":");
      var self = sname + SEP + parts[0] + SEP + parts[1];
      var ps = precedents(cell.f, sname, names), seen = {};
      for (var i = 0; i < ps.length; i++) {
        if (seen[ps[i]]) continue;
        seen[ps[i]] = 1;
        (dependents[ps[i]] = dependents[ps[i]] || []).push(self);
      }
    }
  }
  for (var key in dependents) dependents[key].sort(cellOrder);
  return dependents;
}

/* Sheet, then row, then column — the same order the Python engine sorts
   tuples in, so both walk the graph identically. */
function cellOrder(a, b) {
  var x = a.split(SEP), y = b.split(SEP);
  if (x[0] !== y[0]) return x[0] < y[0] ? -1 : 1;
  if (+x[1] !== +y[1]) return +x[1] - +y[1];
  return +x[2] - +y[2];
}

/* Breadth-first walk downstream of a cell. Returns how many cells it feeds and
   the path to the furthest one whose value actually moved — the difference
   between "this cell changed" and "this cell changed, and here is the number
   it landed on". */
function trace(start, dependents, moved, limit) {
  var parent = {}, order = [start], i = 0;
  parent[start] = null;
  limit = limit || 2000;
  while (i < order.length && order.length < limit) {
    var nb = dependents[order[i]] || [];
    for (var j = 0; j < nb.length; j++)
      if (!(nb[j] in parent)) { parent[nb[j]] = order[i]; order.push(nb[j]); }
    i++;
  }
  var target = null;
  for (var k = order.length - 1; k >= 1; k--)
    if (moved[order[k]]) { target = order[k]; break; }
  var path = [];
  while (target !== null && target !== undefined) { path.push(target); target = parent[target]; }
  path.reverse();
  return { count: order.length - 1, path: path };
}

function cellLabel(key) {
  var p = key.split(SEP);
  return p[0] + "!" + colLetter(+p[2]) + p[1];
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
    rowLeft: o.lr || 0, rowRight: o.rr || 0,
    chain: [], downstream: 0,
    outputs: [], outputImpact: 0, onOutput: false
  };
}

var CELLREF = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/;
var OUTPUT_CAP = 60;

function rowLabel(grid, row) { return rowSignature(grid, row).split("|")[0]; }
function isNum(v) { return typeof v === "number" && isFinite(v); }

/* Right-hand row and column back to their aligned left-hand counterparts. */
function buildInverse(align) {
  var inv = {};
  for (var sheet in align) {
    var rmap = {}, cmap = {};
    align[sheet].pairs.forEach(function (p) { rmap[p[1]] = p[0]; });
    (align[sheet].colPairs || []).forEach(function (p) { cmap[p[1]] = p[0]; });
    inv[sheet] = [rmap, cmap];
  }
  return inv;
}

/* Calculated cells that nothing else reads. A leaf in the dependency graph is
   what a model exists to produce — the number someone reads out loud. Ranking
   against these is what makes severity economic rather than syntactic.

   A cell counts if it was calculated in either version. Requiring a formula in
   the revised file only would drop every cell someone replaced with a plug,
   which is the one kind of output most worth ranking. */
function detectOutputs(right, graph, left, inv) {
  var leaves = [], sheets = Object.keys(right).sort();
  for (var i = 0; i < sheets.length; i++) {
    var sname = sheets[i], g = right[sname];
    var maps = inv[sname] || [{}, {}], rmap = maps[0], cmap = maps[1];
    var keys = Object.keys(g.cells).map(function (k) {
      var p = k.split(":"); return [+p[0], +p[1]];
    }).sort(function (a, b) { return (a[0] - b[0]) || (a[1] - b[1]); });
    for (var j = 0; j < keys.length; j++) {
      var r = keys[j][0], c = keys[j][1], cell = g.cells[r + ":" + c];
      if (graph[sname + SEP + r + SEP + c]) continue;
      if (!isNum(cell.v)) continue;
      var was = (left[sname] && rmap[r] && cmap[c]) ? get(left[sname], rmap[r], cmap[c]) : EMPTY;
      if (!isFormula(cell) && !isFormula(was)) continue;
      leaves.push(sname + SEP + r + SEP + c);
      if (leaves.length >= OUTPUT_CAP) return leaves;
    }
  }
  return leaves;
}

/* Explicit outputs, as "Sheet!B10" strings. Overrides detection. */
function parseOutputSpec(spec, grids) {
  var out = [];
  (spec || []).forEach(function (item) {
    var at = String(item).lastIndexOf("!");
    if (at < 0) return;
    var sheet = item.slice(0, at).replace(/^'|'$/g, "");
    var m = CELLREF.exec(item.slice(at + 1).replace(/\$/g, "").trim().toUpperCase());
    if (!m || !grids[sheet]) return;
    out.push(sheet + SEP + parseInt(m[2], 10) + SEP + colToNum(m[1]));
  });
  return out;
}

function reachable(start, dependents, limit) {
  var seen = {}, order = [start], i = 0;
  seen[start] = 1;
  limit = limit || 2000;
  while (i < order.length && order.length < limit) {
    var nb = dependents[order[i]] || [];
    for (var j = 0; j < nb.length; j++)
      if (!seen[nb[j]]) { seen[nb[j]] = 1; order.push(nb[j]); }
    i++;
  }
  return seen;
}

/* Measure each finding by how far it moves the model's outputs. */
function rankByOutputs(findings, left, right, inv, graph, outputs) {
  var oset = {};
  outputs.forEach(function (k) { oset[k] = 1; });

  function outputMove(key) {
    var p = key.split(SEP), sheet = p[0], r = +p[1], c = +p[2];
    var maps = inv[sheet] || [{}, {}];
    var rv = right[sheet] ? get(right[sheet], r, c).v : null;
    var lr = maps[0][r], lc = maps[1][c];
    var lv = (left[sheet] && lr && lc) ? get(left[sheet], lr, lc).v : null;
    if (!isNum(rv) || !isNum(lv) || rv === lv) return null;
    return [lv, rv, Math.round(Math.abs(rv - lv) / Math.max(Math.abs(lv), 1e-9) * 1e6) / 1e6];
  }

  findings.forEach(function (f) {
    var m = CELLREF.exec(f.ref || "");
    if (!m || !outputs.length) return;
    var start = f.sheet + SEP + parseInt(m[2], 10) + SEP + colToNum(m[1]);
    var seen = reachable(start, graph), hits = [];
    Object.keys(oset).sort(cellOrder).forEach(function (key) {
      if (!seen[key]) return;
      var move = outputMove(key);
      if (!move) return;
      var p = key.split(SEP);
      hits.push({ ref: cellLabel(key), label: rowLabel(right[p[0]], +p[1]),
        before: fmt(move[0]), after: fmt(move[1]), move: move[2] });
    });
    hits.sort(function (a, b) { return (b.move - a.move) || (a.ref < b.ref ? -1 : 1); });
    f.outputs = hits.slice(0, 5);
    f.outputImpact = hits.length ? hits[0].move : 0;
    f.onOutput = !!oset[start];
  });
}

/* Give every cell finding the path from it to the number it moved. Impact
   findings already said a value moved with no edit to the cell. The graph is
   what turns that into a cause: the same edit, traced forward, ends on the
   output someone is about to read out loud. */
function attachChains(findings, right, namesB) {
  var graph = buildGraph(right, namesB), moved = {}, i, m;
  for (i = 0; i < findings.length; i++) {
    m = CELLREF.exec(findings[i].ref || "");
    if (m && findings[i].valBefore !== findings[i].valAfter)
      moved[findings[i].sheet + SEP + parseInt(m[2], 10) + SEP + colToNum(m[1])] = 1;
  }
  for (i = 0; i < findings.length; i++) {
    var f = findings[i];
    m = CELLREF.exec(f.ref || "");
    if (!m || f.kind === "impact") continue;
    var r = trace(f.sheet + SEP + parseInt(m[2], 10) + SEP + colToNum(m[1]), graph, moved);
    f.downstream = r.count;
    f.chain = r.path.map(cellLabel);
  }
  return graph;
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

function compare(wbA, wbB, outputSpec) {
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

  var graph = attachChains(findings, right, namesB);
  var inv = buildInverse(align);
  var declared = parseOutputSpec(outputSpec, right);
  var picked = declared.length ? declared : detectOutputs(right, graph, left, inv);
  rankByOutputs(findings, left, right, inv, graph, picked);

  /* Severity first, then how far the change moved a model output, then whether
     it sits on an output at all. That last one matters because a plug freezes a
     number rather than moving it: the most dangerous finding in the tool
     measures zero movement by construction, and without this it would sort
     below every unrelated finding. */
  findings.sort(function (a, b) {
    return (SEV_ORDER[a.severity] - SEV_ORDER[b.severity]) ||
      (b.outputImpact - a.outputImpact) ||
      ((a.onOutput ? 0 : 1) - (b.onOutput ? 0 : 1)) ||
      (a.sheet < b.sheet ? -1 : a.sheet > b.sheet ? 1 : 0) ||
      ((a.rowRight || a.rowLeft) - (b.rowRight || b.rowLeft));
  });

  return { left: left, right: right, findings: findings, align: align, stale: stale,
    names: [namesA, namesB], outputs: picked.map(cellLabel),
    outputsDeclared: declared.length > 0 };
}

root.xldiff = { compare: compare, toR1C1: toR1C1, fmt: fmt, colLetter: colLetter,
  get: get, precedents: precedents, SEP: SEP };

})(typeof module !== "undefined" && module.exports ? module.exports : (typeof window !== "undefined" ? window : this));
