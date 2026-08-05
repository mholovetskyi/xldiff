#!/usr/bin/env node
/* Writes examples/sample_before.xlsx and examples/sample_after.xlsx: a small
   two-sheet model with one seeded bug per finding class, so the web app can be
   exercised without a real workbook to hand.

   Uses the SheetJS already installed for the web build, so there is nothing
   extra to install:

     cd web && npm install
     node ../examples/make_samples.js

   Cached values are written alongside every formula. That matters: xldiff
   reads the last value Excel stored to work out downstream impact, and a file
   written without them would trip the stale-cache warning instead. */

const fs = require("fs");
const path = require("path");
const ROOT = path.dirname(__dirname);
const XLSX = require(path.join(ROOT, "web/node_modules/xlsx"));

const COL = (n) => XLSX.utils.encode_col(n - 1);

/* SheetJS stores an error as a code, not its text. */
const ERR = { "#NULL!": 0, "#DIV/0!": 7, "#VALUE!": 15, "#REF!": 23,
  "#NAME?": 29, "#NUM!": 36, "#N/A": 42 };
const err = (text, formula) => ({ error: text, f: formula });

/* A sheet is described as a map of "A1" -> value, -> [formula, cachedValue],
   or -> err("#DIV/0!", formula). Formulas are written without the leading "=",
   as SheetJS expects. */
function sheet(cells) {
  const ws = {};
  let maxR = 1, maxC = 1;
  for (const ref of Object.keys(cells)) {
    const v = cells[ref];
    ws[ref] = Array.isArray(v)
      ? { t: "n", f: v[0], v: v[1] }
      : (v && v.error ? { t: "e", v: ERR[v.error], w: v.error, f: v.f }
      : (typeof v === "number" ? { t: "n", v: v } : { t: "s", v: String(v) }));
    const a = XLSX.utils.decode_cell(ref);
    if (a.r + 1 > maxR) maxR = a.r + 1;
    if (a.c + 1 > maxC) maxC = a.c + 1;
  }
  ws["!ref"] = "A1:" + COL(maxC) + maxR;
  return ws;
}

function book(pairs, names) {
  const wb = { SheetNames: [], Sheets: {} };
  for (const [name, cells] of pairs) {
    wb.SheetNames.push(name);
    wb.Sheets[name] = sheet(cells);
  }
  if (names) {
    wb.Workbook = { Names: Object.keys(names).map((n) => ({ Name: n, Ref: names[n] })) };
  }
  return wb;
}

/* ── Revenue build ──────────────────────────────────────────────────────
   Four quarters across columns B–E. Units grow off a base at the rate in
   B4; revenue is units × price. */
const BASE = [1000, 1100, 1200, 1300];
const PRICE = [25, 25, 26, 26];

function revenue(growth, opts) {
  const c = {
    A1: "Revenue build",
    A3: "Quarter", B3: "Q1", C3: "Q2", D3: "Q3", E3: "Q4",
    A4: "Growth rate", B4: growth,
    A6: "Base units",
    A7: "Units",
    A8: "Price",
    A9: "Revenue",
    A11: "Discounted revenue",
    A13: "Cohorts", B13: opts.cohorts,
    A14: "Revenue per cohort",
    A16: "Prior year revenue"
  };

  /* Someone zeroes an input and the division below it blows up. */
  c.B14 = opts.cohorts === 0
    ? err("#DIV/0!", "B9/B13")
    : ["B9/B13", null];  /* filled in once revenue is known */

  /* The revised file drops the Prior year sheet, so this reference dies with
     it — Excel rewrites the formula text itself, which is the tell. */
  c.B16 = opts.priorYearGone
    ? err("#REF!", "#REF!B9")
    : ["'Prior year'!B9", 24000];

  const units = [], rev = [];
  for (let i = 0; i < 4; i++) {
    const L = COL(i + 2);
    c[L + "6"] = BASE[i];
    c[L + "8"] = PRICE[i];

    /* Q2 in the revised file: someone widened the growth assumption for one
       quarter only, breaking a run that is otherwise uniform. */
    const bumped = opts.bumpQ2 && i === 1;
    /* Q4 in the baseline: a deliberate exception that the revision removes,
       which xldiff reports as a repair at low severity. */
    const exception = opts.q4Exception && i === 3;

    const u = BASE[i] * (1 + growth + (bumped ? 0.02 : 0)) + (exception ? 5 : 0);
    units.push(u);
    c[L + "7"] = [
      L + "6*(1+$B$4" + (bumped ? "+0.02" : "") + ")" + (exception ? "+5" : ""),
      u
    ];

    /* Q3 in the revised file: the formula is gone, replaced by the number it
       used to produce in the baseline. The classic plug. */
    if (opts.plugQ3 && i === 2) {
      c[L + "9"] = BASE[2] * (1 + 0.04) * PRICE[2];
      rev.push(BASE[2] * (1 + 0.04) * PRICE[2]);
    } else {
      const r = u * PRICE[i];
      rev.push(r);
      c[L + "9"] = [L + "7*" + L + "8", r];
    }
  }

  /* No horizontal neighbours, so an edit here reads as an ordinary formula
     change — and the new literal is what makes it high severity. */
  const disc = opts.discount;
  c.B11 = ["B9*" + disc, rev[0] * disc];
  if (opts.cohorts) c.B14 = ["B9/B13", rev[0] / opts.cohorts];

  /* A monthly ramp down column B. Every step is the same formula, so
     overriding one in the middle is a vertical run break — the cell has the
     run continuing above and below it, which is what separates a real
     drag-fill accident from a totals row that is meant to differ. */
  c.A24 = "Tax rate";
  c.B24 = 0.2;

  c.A18 = "Monthly ramp";
  c.B18 = 100;
  let u = 100;
  for (let r = 19; r <= 22; r++) {
    const rate = (opts.overrideMonth === r) ? 1.05 : 1.01;
    u = u * rate;
    c["B" + r] = ["B" + (r - 1) + "*" + rate, Number(u.toFixed(6))];
  }
  return c;
}

/* ── Operating costs ────────────────────────────────────────────────────
   The revision inserts a Travel line and a prior-year column, so this sheet
   shifts on both axes at once. Nothing about the quarters actually changed:
   without alignment on both axes every cell here reports as edited, which is
   the failure mode the whole approach exists to avoid. */
function costs(opts) {
  const c = { A1: "Operating costs", A3: "Item", A4: "Salaries", A5: "Rent" };
  const SAL = [420, 430, 440, 450], RENT = [90, 90, 95, 95];
  const TRAVEL = [12, 14, 11, 16], SOFT = [60, 62, 64, 66];
  const PRIOR = { sal: 400, rent: 85, travel: 9, soft: 55 };

  let row = 6;
  if (opts.withTravel) { c["A" + row] = "Travel"; row++; }
  const softRow = row;
  c["A" + softRow] = "Software";
  const totalRow = softRow + 1;
  c["A" + totalRow] = "Total";

  /* The inserted column lands at B, pushing every quarter one to the right. */
  let col = 2;
  if (opts.withPriorYear) {
    const P = COL(col);
    c[P + "3"] = "FY prior";
    c[P + "4"] = PRIOR.sal;
    c[P + "5"] = PRIOR.rent;
    if (opts.withTravel) c[P + "6"] = PRIOR.travel;
    c[P + softRow] = PRIOR.soft;
    const pt = PRIOR.sal + PRIOR.rent + PRIOR.soft + (opts.withTravel ? PRIOR.travel : 0);
    c[P + totalRow] = ["SUM(" + P + "4:" + P + softRow + ")", pt];
    col++;
  }

  for (let i = 0; i < 4; i++) {
    const L = COL(col + i);
    c[L + "3"] = "Q" + (i + 1);
    c[L + "4"] = SAL[i];
    c[L + "5"] = RENT[i];
    if (opts.withTravel) c[L + "6"] = TRAVEL[i];
    c[L + softRow] = SOFT[i];
    const total = SAL[i] + RENT[i] + SOFT[i] + (opts.withTravel ? TRAVEL[i] : 0);
    c[L + totalRow] = ["SUM(" + L + "4:" + L + softRow + ")", total];

    /* Identical text in both files, but the inserted Travel row moves it down
       one. TAX1 is shaped exactly like a cell reference, so normalising it to
       R1C1 relative to its own position would make these two look different
       and report a change that never happened. */
    c[L + (totalRow + 1)] = ["" + L + totalRow + "*TAX1", total * 0.2];
  }
  c["A" + (totalRow + 1)] = "After tax";
  return c;
}

/* A sheet that exists only in the revised file. */
const sensitivity = {
  A1: "Sensitivity",
  A3: "Growth", B3: "Revenue",
  A4: 0.04, B4: ["'Revenue build'!B9", 26500],
  A5: 0.06, B5: ["'Revenue build'!B9*1.5", 39750]
};

/* A sheet the revised file deletes, taking a reference down with it. */
const priorYear = {
  A1: "Prior year", A3: "Revenue", B9: 24000
};

const before = book([
  /* The baseline carries the Q4 exception that the revision repairs. */
  ["Revenue build", revenue(0.04, { q4Exception: true, discount: 0.9, cohorts: 4 })],
  ["Operating costs", costs({})],
  ["Prior year", priorYear]
], {
  GrowthRate: "'Revenue build'!$B$4",
  TAX1: "'Revenue build'!$B$24",
  LegacyOpex: "'Operating costs'!$B$7"
});

const after = book([
  ["Revenue build", revenue(0.06, {
    bumpQ2: true, plugQ3: true, discount: 0.85, cohorts: 0, priorYearGone: true,
    overrideMonth: 20
  })],
  ["Operating costs", costs({ withTravel: true, withPriorYear: true })],
  ["Sensitivity", sensitivity]
], {
  /* Repointed at the cohort count instead of the growth rate: every formula
     using it changed meaning without changing text. */
  GrowthRate: "'Revenue build'!$B$13",
  TAX1: "'Revenue build'!$B$24",
  /* LegacyOpex is gone, so anything referring to it can no longer resolve. */
  Headroom: "'Sensitivity'!$B$5"
});

/* A version between the two, so lineage has three points to walk. The growth
   rate has already moved here; the plug and the deleted sheet have not landed
   yet, which is what lets "when did this first appear" have an answer. */
const interim = book([
  ["Revenue build", revenue(0.06, { discount: 0.9, cohorts: 4 })],
  ["Operating costs", costs({})],
  ["Prior year", priorYear]
], {
  GrowthRate: "'Revenue build'!$B$4",
  TAX1: "'Revenue build'!$B$24",
  LegacyOpex: "'Operating costs'!$B$7"
});

for (const [name, wb] of [["sample_before.xlsx", before],
                          ["sample_interim.xlsx", interim],
                          ["sample_after.xlsx", after]]) {
  const dest = path.join(__dirname, name);
  XLSX.writeFile(wb, dest, { bookType: "xlsx" });
  console.log("wrote %s (%d KB)", dest, Math.round(fs.statSync(dest).size / 1024));
}
