#!/usr/bin/env node
/* Dumps the browser engine's findings as JSON so test_parity.py can hold the
   two implementations against each other.

     node tests/parity_dump.js old.xlsx new.xlsx                            */
const fs = require("fs");
const path = require("path");
const ROOT = path.dirname(__dirname);
const XLSX = require(path.join(ROOT, "web/node_modules/xlsx"));
const { xldiff } = require(path.join(ROOT, "web/engine.js"));

const opts = { type: "buffer", cellFormula: true };
const a = XLSX.read(fs.readFileSync(process.argv[2]), opts);
const b = XLSX.read(fs.readFileSync(process.argv[3]), opts);
const r = xldiff.compare(a, b);

process.stdout.write(JSON.stringify({
  stale: r.stale,
  findings: r.findings.map((f) => ({
    severity: f.severity, kind: f.kind, sheet: f.sheet, ref: f.ref,
    summary: f.summary, detail: f.detail,
    before: f.before, after: f.after,
    valBefore: f.valBefore, valAfter: f.valAfter,
    chain: f.chain, downstream: f.downstream,
    outputs: f.outputs, outputImpact: f.outputImpact, onOutput: f.onOutput
  }))
}, null, 1));
