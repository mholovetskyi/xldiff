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

console.log("checking the committed build");
const tpl = fs.readFileSync(path.join(ROOT, "web/app.template.html"), "utf8");
const expected = tpl
  .replace("__XLSX__", () => fs.readFileSync(path.join(ROOT, "web/node_modules/xlsx/dist/xlsx.core.min.js"), "utf8"))
  .replace("__ENGINE__", () => fs.readFileSync(path.join(ROOT, "web/engine.js"), "utf8"));
const committed = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
check(expected === committed, "index.html matches a fresh build of the sources");
check(!/<script[^>]+src=/i.test(committed), "index.html loads nothing from the network");

console.log("");
console.log(failed ? failed + " check(s) failed" : "all checks passed");
process.exit(failed ? 1 : 0);
