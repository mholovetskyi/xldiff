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

/* A sheet is described as a map of "A1" -> value, or -> [formula, cachedValue].
   Formulas are written without the leading "=", as SheetJS expects. */
function sheet(cells) {
  const ws = {};
  let maxR = 1, maxC = 1;
  for (const ref of Object.keys(cells)) {
    const v = cells[ref];
    ws[ref] = Array.isArray(v)
      ? { t: "n", f: v[0], v: v[1] }
      : (typeof v === "number" ? { t: "n", v: v } : { t: "s", v: String(v) });
    const a = XLSX.utils.decode_cell(ref);
    if (a.r + 1 > maxR) maxR = a.r + 1;
    if (a.c + 1 > maxC) maxC = a.c + 1;
  }
  ws["!ref"] = "A1:" + COL(maxC) + maxR;
  return ws;
}

function book(pairs) {
  const wb = { SheetNames: [], Sheets: {} };
  for (const [name, cells] of pairs) {
    wb.SheetNames.push(name);
    wb.Sheets[name] = sheet(cells);
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
    A11: "Discounted revenue"
  };

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
  return c;
}

/* ── Operating costs ────────────────────────────────────────────────────
   The revision inserts a Travel line, which shifts every row beneath it. */
function costs(withTravel) {
  const c = {
    A1: "Operating costs",
    A3: "Item", B3: "Q1", C3: "Q2", D3: "Q3", E3: "Q4",
    A4: "Salaries", A5: "Rent"
  };
  const SAL = [420, 430, 440, 450], RENT = [90, 90, 95, 95];
  const TRAVEL = [12, 14, 11, 16], SOFT = [60, 62, 64, 66];

  let row = 6;
  if (withTravel) { c["A" + row] = "Travel"; row++; }
  const softRow = row;
  c["A" + softRow] = "Software";
  const totalRow = softRow + 1;
  c["A" + totalRow] = "Total";

  for (let i = 0; i < 4; i++) {
    const L = COL(i + 2);
    c[L + "4"] = SAL[i];
    c[L + "5"] = RENT[i];
    if (withTravel) c[L + "6"] = TRAVEL[i];
    c[L + softRow] = SOFT[i];
    const total = SAL[i] + RENT[i] + SOFT[i] + (withTravel ? TRAVEL[i] : 0);
    c[L + totalRow] = ["SUM(" + L + "4:" + L + softRow + ")", total];
  }
  return c;
}

/* A sheet that exists only in the revised file. */
const sensitivity = {
  A1: "Sensitivity",
  A3: "Growth", B3: "Revenue",
  A4: 0.04, B4: ["'Revenue build'!B9", 26500],
  A5: 0.06, B5: ["'Revenue build'!B9*1.5", 39750]
};

const before = book([
  ["Revenue build", revenue(0.04, { discount: 0.9 })],
  ["Operating costs", costs(false)]
]);

const after = book([
  ["Revenue build", revenue(0.06, { bumpQ2: true, q4Exception: false, plugQ3: true, discount: 0.85 })],
  ["Operating costs", costs(true)],
  ["Sensitivity", sensitivity]
]);

/* The baseline carries the Q4 exception that the revision repairs. */
before.Sheets["Revenue build"] = sheet(revenue(0.04, { q4Exception: true, discount: 0.9 }));

for (const [name, wb] of [["sample_before.xlsx", before], ["sample_after.xlsx", after]]) {
  const dest = path.join(__dirname, name);
  XLSX.writeFile(wb, dest, { bookType: "xlsx" });
  console.log("wrote %s (%d KB)", dest, Math.round(fs.statSync(dest).size / 1024));
}
