/* Checks the browser engine against the same example models, and that the
   committed index.html is an up-to-date build of the sources.

     cd web && npm install xlsx && node ../tests/test_engine.js            */
const fs = require("fs");
const path = require("path");
const ROOT = path.dirname(__dirname);

const XLSX = require(path.join(ROOT, "web/node_modules/xlsx"));
const { xldiff } = require(path.join(ROOT, "web/engine.js"));

let failed = 0;
function check(cond, label) {
  console.log((cond ? "  ok   " : "  FAIL ") + label);
  if (!cond) failed++;
}

const opts = { type: "buffer", cellFormula: true };
const a = XLSX.read(fs.readFileSync(path.join(ROOT, "examples/model_v14.xlsx")), opts);
const b = XLSX.read(fs.readFileSync(path.join(ROOT, "examples/model_v15.xlsx")), opts);
const r = xldiff.compare(a, b);

const key = (f) => [f.severity, f.kind, f.sheet, f.ref].join("|");
const got = new Set(r.findings.map(key));

console.log("comparing the example models");
[
  "high|run_break|Revenue build|F13",
  "high|hardcode|Revenue build|C14",
  "medium|constant_changed|Revenue build|C5",
  "medium|sheet_added|Sensitivity|",
  "low|row_added|Opex|row 5",
  "low|run_repaired|Revenue build|E12"
].forEach((k) => check(got.has(k), k.replace(/\|/g, " ")));

check(!r.stale, "cached values are present in both files");
check(r.findings.filter((f) => f.kind === "impact").length >= 4, "downstream value moves detected");

const opex = r.align["Opex"];
check(opex.added.length === 1 && opex.added[0] === 5,
  "the inserted Opex row aligns as a single insertion");
check(opex.ratio > 0.9, "Opex alignment confidence is high (" + opex.ratio.toFixed(2) + ")");

console.log("comparing a file with itself");
check(xldiff.compare(b, b).findings.length === 0, "identical files produce no findings");

console.log("errors introduced by the revision");
const sa = XLSX.read(fs.readFileSync(path.join(ROOT, "examples/sample_before.xlsx")), opts);
const sb = XLSX.read(fs.readFileSync(path.join(ROOT, "examples/sample_after.xlsx")), opts);
const s = xldiff.compare(sa, sb);
const sgot = new Set(s.findings.map(key));
[
  "high|error_introduced|Revenue build|B14",
  "high|broken_reference|Revenue build|B16",
  "high|sheet_deleted|Prior year|"
].forEach((k) => check(sgot.has(k), k.replace(/\|/g, " ")));

console.log("dependency chains");
const P = (f, sheet, names) => xldiff.precedents(f, sheet, names)
  .map((k) => k.split(xldiff.SEP).join("!"));
check(String(P("='Revenue build'!B9*1.5", "Sensitivity")) === "Revenue build!9!2",
  "a cross-sheet reference resolves to the sheet it names");
check(P("=SUM(B4:B6)", "S").length === 3, "a range expands to the cells it covers");
check(String(P("=LOG10(B4)", "S")) === "S!4!2",
  "a function name that looks like a reference is not one");
const namedP = P("=B7*TAX1", "Costs", { TAX1: "'Revenue build'!$B$24" });
check(namedP.indexOf("Revenue build!24!2") !== -1, "a name resolves to what it points at");
check(namedP.every((k) => +k.split("!").pop() < 100),
  "and is not also read as a cell 13,570 columns out");

const growth = s.findings.find((f) => f.sheet === "Revenue build" && f.ref === "B4");
check(growth.downstream >= 8,
  "the growth rate is known to feed " + growth.downstream + " cells");
check(growth.chain.length && growth.chain[0] === "Revenue build!B4",
  "its chain starts at the cell that changed");
check(growth.chain.length >= 3,
  "and runs forward to a number that moved: " + growth.chain.join(" -> "));
check(s.findings.filter((f) => f.kind === "impact").every((f) => !f.chain.length),
  "impact findings carry no chain: they are the far end of someone else's");

console.log("finding ids");
// A waiver written in the browser has to match the same finding in CI, so the
// id is a cross-engine contract. tests/test_parity.py compares them run for
// run; this pins the shape and the rule the id encodes.
check(/^[0-9a-f]{8}$/.test(s.findings[0].id), "an id is eight hex characters");
check(new Set(s.findings.map((f) => f.id)).size === s.findings.length,
  "every finding in this comparison has a distinct id");
check(xldiff.fingerprint("hardcode", "S", "C14", "=A1*2", "500") !==
      xldiff.fingerprint("hardcode", "S", "C14", "=A1*2", "600"),
  "editing the change itself produces a different id");
check(xldiff.fingerprint("hardcode", "S", "C14", "=A1*2", "500") ===
      xldiff.fingerprint("hardcode", "S", "C14", "=A1*2", "500"),
  "and an identical change produces the same one");

console.log("ranking against model outputs");
check(s.outputs.length && !s.outputsDeclared,
  "outputs are detected when none are declared (" + s.outputs.length + " found)");
const ranked = s.findings.filter((f) => f.severity === "high");
const impacts = ranked.map((f) => f.outputImpact);
check(String(impacts) === String(impacts.slice().sort((a, b) => b - a)),
  "high findings are ordered by how far they move an output");
check(ranked[0].outputImpact > 0 && ranked[0].outputs.length,
  "the finding at the top moves one: " + ranked[0].outputs[0].ref +
  " by " + (ranked[0].outputImpact * 100).toFixed(1) + "%");

// A plug freezes a number instead of moving it, so it measures zero by
// construction, and without the on-output tiebreak it sorts below every
// unrelated finding.
const plug = ranked.find((f) => f.kind === "hardcode");
check(plug.outputImpact === 0 && plug.onOutput,
  "a plug moves nothing measurable but sits on an output");
const inert = ranked.findIndex((f) => !f.onOutput && f.outputImpact === 0);
check(inert >= 0 && ranked.indexOf(plug) < inert,
  "and still outranks every high finding that moves nothing and sits on nothing");

const declaredRun = xldiff.compare(sa, sb, ["'Revenue build'!B11"]);
check(String(declaredRun.outputs) === "Revenue build!B11" && declaredRun.outputsDeclared,
  "an explicit output list replaces detection entirely");

console.log("defined names");
const byKind = {};
s.findings.forEach((f) => { if (!byKind[f.kind]) byKind[f.kind] = f; });
check(byKind.name_changed && byKind.name_changed.ref === "GrowthRate",
  "a name repointed at another cell is high severity");
check(byKind.name_deleted && byKind.name_deleted.severity === "high",
  "a deleted name is high: formulas using it can no longer resolve");
check(byKind.name_added && byKind.name_added.severity === "low",
  "a new name is low on its own");

// TAX1 is shaped exactly like a cell reference. Normalising it relative to its
// own position makes identical formulas look different the moment a row is
// inserted above them.
const named = { TAX1: 1 };
check(xldiff.toR1C1("=B7*TAX1", 8, 2, named) === xldiff.toR1C1("=B8*TAX1", 9, 2, named),
  "a formula using a name compares equal after it moves down a row");
check(xldiff.toR1C1("=B7*TAX1", 8, 2, named).indexOf("TAX1") !== -1,
  "the name survives R1C1 normalisation intact");
check(xldiff.toR1C1("=B7*TAX1", 8, 2, {}).indexOf("TAX1") === -1,
  "and is only spared because it is known to be a name");
check(!s.findings.some((f) => f.sheet === "Operating costs" &&
        /^[A-Z]9$/.test(f.ref) && f.kind === "formula_changed"),
  "the After tax row, identical text one row down, reports no formula change");

console.log("an inserted column does not shift the diff");
const oc = s.findings.filter((f) => f.sheet === "Operating costs");
check(oc.some((f) => f.kind === "column_added" && f.ref === "column B"),
  "the inserted column is reported once, by letter and header");
// Salaries, Rent and Software only moved one column right. Every one of those
// cells would report as edited without column alignment.
const shifted = oc.filter((f) => /^[A-Z][457]$/.test(f.ref));
check(!shifted.length,
  "cells that only moved right report nothing (got " + shifted.length + ")");
const ocAl = s.align["Operating costs"];
check(String(ocAl.colAdded) === "2" && !ocAl.colDeleted.length,
  "exactly one column is reported as inserted, at B");
// The sheet shifts on both axes at once. Row alignment must survive the column
// insertion and vice versa: each axis used to need the other solved first.
check(String(ocAl.added) === "6" && !ocAl.deleted.length,
  "row alignment survives it: one insertion at row 6, nothing deleted");
check(ocAl.ratio > 0.35 && ocAl.colRatio > 0.35,
  "both confidences stay well clear of the low-confidence threshold (rows " +
  ocAl.ratio.toFixed(2) + ", columns " + ocAl.colRatio.toFixed(2) + ")");

console.log("run breaks on both axes");
check(sgot.has("high|run_break|Revenue build|B20"),
  "an overridden step mid-column is a vertical run break");
check(s.findings.find((f) => f.ref === "B20").summary.indexOf("column") !== -1,
  "the finding names the axis that broke");
check(s.findings.find((f) => f.ref === "C7").summary.indexOf("row") !== -1,
  "a horizontal break still reads as a row");

// A column of identical formulas nearly always ends in a total that is meant
// to differ. Reporting that as a break would make the whole axis useless.
const d11 = r.findings.find((f) => f.sheet === "Opex" && f.ref === "D11");
check(d11 && d11.kind !== "run_break",
  "a totals row below a uniform column is not a run break");

const errf = s.findings.find((f) => f.kind === "error_introduced");
check(errf.summary.indexOf("#DIV/0!") !== -1, "the error finding names the error it found");
check(!s.stale, "an error value is a result, not a missing one, so nothing reads as stale");

console.log("errors cleared by the revision");
const cleared = xldiff.compare(sb, sa).findings.filter((f) => f.kind === "error_cleared");
check(cleared.length === 2, "reversing the pair reports both fixes (" + cleared.length + ")");
check(cleared.every((f) => f.severity === "low"),
  "a fix is low severity, so it cannot crowd out a real problem");

console.log("checking the committed build");
const tpl = fs.readFileSync(path.join(ROOT, "web/app.template.html"), "utf8");
const expected = tpl
  .replace("__XLSX__", () => fs.readFileSync(path.join(ROOT, "web/node_modules/xlsx/dist/xlsx.core.min.js"), "utf8"))
  .replace("__ENGINE__", () => fs.readFileSync(path.join(ROOT, "web/engine.js"), "utf8"));
const committed = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
check(expected === committed, "index.html matches a fresh build of the sources");
check(!/<script[^>]+src=/i.test(committed), "index.html loads nothing from the network");

// The drift check proves the build matches its sources. It cannot tell whether
// those sources parse, and a syntax error in the app script ships a page that
// loads and does nothing at all.
const scripts = committed.match(/<script>([\s\S]*?)<\/script>/g) || [];
const appScript = scripts[scripts.length - 1].replace(/^<script>|<\/script>$/g, "");
let parses = true;
try {
  new (require("vm").Script)(appScript, { filename: "index.html app script" });
} catch (e) {
  parses = false;
  console.log("       " + e.message);
}
check(parses, "the app script in index.html parses");

console.log("");
console.log(failed ? failed + " check(s) failed" : "all checks passed");
process.exit(failed ? 1 : 0);
