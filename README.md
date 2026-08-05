# xldiff

Compare two Excel models. See which cells changed, which changes are hardcodes or
broken formula runs, and what moved downstream as a result.

Reviewing a spreadsheet takes longer than editing one, which is why review is where
models break. The errors that survive review are rarely dramatic: someone drag-fills a
formula across a row that held one deliberate exception, or drops a plug into a
calculated row. Nobody catches it until the number is in front of a client. xldiff
looks for those two things specifically, ranks what it finds, and tells you what the
change did to the outputs.

**Nothing you compare is uploaded anywhere.** The web app reads both files in the
browser and makes no network requests at all, which you can verify by loading the page
once and then disconnecting. The command line tool is equally offline.

## Two ways to run it

**Browser.** Open `index.html`, drop the old file on the left and the new one on the
right. That file is self-contained: SheetJS and the engine are inlined, so it works
from disk, from a USB stick, or on a machine with no internet. Host it anywhere static.

The review screen opens on the worst finding and leads with a verdict derived from the
counts, so the question "does this need me?" is answered before you scroll. Findings sit
beside the two sheets; selecting one moves both panes to it and shows the formula diff
with the changed tokens marked. The panes read as **Values**, **Formulas**, or **Impact**
— the last showing the value delta per cell, scaled against the largest move on screen.
A finding you have judged and dismissed can be suppressed, which drops it out of the
counts and the verdict without deleting it. **Export** writes the review to disk as a
single HTML file that needs no scripting to read, or as JSON for a CI gate.

**Command line.** For pre-commit hooks and CI:

```bash
pip install -r cli/requirements.txt
python cli/cli.py old.xlsx new.xlsx -o report.html
python cli/cli.py base.xlsx head.xlsx --quiet --fail-on high   # exits 1 on high findings
```

Try it on the included examples, which have four bugs seeded into them:

```bash
python cli/cli.py examples/model_v14.xlsx examples/model_v15.xlsx
```

`examples/sample_before.xlsx` and `examples/sample_after.xlsx` are a second pair,
built to fire one finding of every class at once — drag them onto the web app to see
what a full report looks like. Regenerate or edit them with
`node examples/make_samples.js`, which needs nothing beyond the SheetJS already
installed for the web build.

## What it catches

| Finding | Severity | Why it matters |
|---|---|---|
| Hardcode replaced a formula | high | The cell stopped recalculating. The classic plug. |
| Formula breaks the run in its row | high | Someone drag-filled over a deliberate exception, or created one by accident. |
| Formula changed and gained a numeric literal | high | An assumption got buried inside a calculation. |
| Formula changed, constant changed, cell cleared | medium | Ordinary edits, shown with the value delta. |
| Value moved with no edit to this cell | medium | Downstream impact. Nobody touched this cell; something else moved it. |
| Sheet added or deleted, row inserted or deleted | low to high | Structural context, reported before cell findings. |

Findings are ordered by severity, then by position. Most spreadsheet comparison tools
order by position, which buries the one thing that matters under two hundred formatting
changes.

## How it works

**R1C1 normalisation.** Formulas are compared in R1C1 form, so a formula that moved to
a new position without changing its logic does not register as a change.

**Row alignment before diffing.** Insert a row at the top of a sheet and a naive
comparison reports every cell below it as changed. That is why Microsoft's own
Spreadsheet Compare is unusable on real models. xldiff fingerprints each row by its
label text plus a shape string (which cells hold formulas, text, numbers, nothing) and
aligns the two sequences with LCS. The fingerprint deliberately ignores formula
*content*, so a row whose formula was edited still matches its counterpart.

**Impact without a formula engine.** An `.xlsx` stores the last value Excel calculated
for every cell. Reading formulas and cached values together gives you value deltas for
free, and the useful signal falls out of it: a cell whose formula is unchanged but
whose value moved is downstream of an edit somewhere else. If a file was saved with
calculation set to manual the cache is stale, so the tool detects that case and warns
rather than reporting a number it cannot stand behind.

**Run-break detection.** For any changed cell, read its horizontal neighbours in R1C1.
Uniform neighbours plus a different cell is a break. The reverse, where a cell used to
differ and now matches, is reported as a repair at low severity so a real fix does not
crowd out real problems.

## Layout

```
index.html          the built web app, committed so Pages can serve it directly
web/                app template, engine.js, build script
cli/                Python engine, HTML report writer, CLI
examples/           two pairs of small models with seeded bugs, and their generators
tests/              smoke tests for both engines
```

`web/engine.js` and `cli/engine.py` implement the same algorithm and produce identical
findings on the same inputs. The tests check both.

## Building the web app

```bash
cd web
npm install
node build.js        # regenerates ../index.html
node ../tests/test_engine.js
```

CI fails if `index.html` drifts from a fresh build of the sources.

## Performance

25 sheets, 700 rows, 14 columns per side, roughly 245,000 cells: about 2.3 seconds.
The alignment strips the common prefix and suffix before doing any quadratic work, so
cost scales with how much actually diverged rather than with file size.

## Known limits

- Horizontal formula runs only. Vertical run detection is not implemented.
- Rows are aligned, columns are not. An inserted column will over-report.
- No cross-sheet dependency tracing, so impact says what moved, not what caused it.
- `.xlsx` and `.xlsm` only. No `.xlsb`, no `.xls`, no Google Sheets.
- Formatting changes are ignored rather than collapsed into a category.
- Capped at 60 columns, 5000 rows per sheet, and 25 impact findings.
- Merging two versions is out of scope and likely to stay that way.

## Contributing

The most useful contribution right now is not code. Run it on a real model you have two
versions of and open an issue about what it got wrong: a finding that should have been
high and was not, noise that should have been suppressed, an alignment that fell apart.
The algorithm was tuned against synthetic fixtures, which hide exactly the problems that
matter.
