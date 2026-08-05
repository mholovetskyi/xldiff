#!/usr/bin/env node
/* Inlines SheetJS and the engine into one self-contained index.html.
   Run: npm install xlsx && node build.js */
const fs = require("fs");
const path = require("path");

const here = __dirname;
const sheetjs = path.join(here, "node_modules/xlsx/dist/xlsx.core.min.js");

if (!fs.existsSync(sheetjs)) {
  console.error("Missing SheetJS. Run: npm install xlsx");
  process.exit(1);
}

let out = fs.readFileSync(path.join(here, "app.template.html"), "utf8");
out = out
  .replace("__XLSX__", () => fs.readFileSync(sheetjs, "utf8"))
  .replace("__ENGINE__", () => fs.readFileSync(path.join(here, "engine.js"), "utf8"));

const dest = path.join(here, "index.html");
fs.writeFileSync(dest, out);
console.log("wrote %s (%d KB)", dest, Math.round(out.length / 1024));
