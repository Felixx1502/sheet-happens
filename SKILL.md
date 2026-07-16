---
name: sheet-happens
description: >
  Map any Excel financial model (.xlsx/.xlsm) into a single self-contained
  interactive HTML dependency graph with three zoom levels (sheets → functional
  blocks → line items), data-flow arrows, click-through formula navigation, and
  a built-in audit (hardcoded constants, error cells, circular references,
  unmapped refs). ALWAYS trigger when the user wants to: understand/visualize
  how an Excel model works, map model structure or dependencies, trace
  precedents/dependents across sheets, audit a financial model for hardcodes,
  "model map", "model x-ray", "dependency graph of my spreadsheet", or asks how
  sheets/tabs in a workbook link together. Input is an Excel file; output is
  one HTML file that opens offline in any browser.
---

# sheet-happens

Turn an Excel financial model into an interactive dependency map.

## Workflow

1. **Locate the model file** (.xlsx or .xlsm). Confirm with the user if
   multiple candidates exist.

2. **Run the pipeline** (pure Python; needs `lxml`, plus `openpyxl` only for
   tests/demo):

   ```bash
   python map_model.py "MODEL.xlsx" -o "MODEL - Model Map.html" --title "Client model"
   ```

   Useful flags:
   - `--anonymize` — strips all labels/formulas, pseudonymises sheet names.
     Use whenever the output may be shared outside the model's owner.
   - `--blocks-override blocks.json` — pin block boundaries per sheet
     (`{"Sheet": [[1, "Revenue"], [40, "COGS"]]}`) when auto-detection
     misses the model's real structure.
   - `--max-blocks N` / `--min-block-rows N` — tune auto-detection.

   Sandbox tips: if a stale bytecode cache interferes, set
   `PYTHONPYCACHEPREFIX=/tmp/pyc`. Very large workbooks (100MB+) parse in
   seconds — the extractor switches to text-scan mode for oversized sheets
   automatically — but keep shell timeouts in mind.

3. **Read the run summary** printed by the CLI (block/link counts, hardcode
   rows, error cells, circular groups, unresolved names). Report the audit
   headlines to the user alongside the file.

4. **Review block quality**: open the map's Blocks view mentally via the data —
   if auto-detected blocks look wrong for key sheets (generic names, giant
   blocks), write a `blocks_override.json` for those sheets from their row
   labels and re-run. One iteration is usually enough.

5. **Deliver** the HTML. Always tell the user:
   - the map embeds labels and formulas from the model → same confidentiality
     as the model itself; use `--anonymize` for anything shared publicly;
   - known accuracy limits (in the map's About panel): dynamic refs
     (INDIRECT/OFFSET) aren't traceable, item-level range links point to the
     range's head row, row-level "circular groups" are often legitimate
     roll-forwards.

## What the output contains

Single HTML file, offline, no upload anywhere: sheet-level flow graph laid out
in calculation order → auto-detected functional blocks → row/item drill-down
with per-formula navigation; search, weight filters, upstream/downstream
tracing; Audit panel (hardcodes, error cells, cycles, unmapped external/table
refs, defined-name health).

## Repo layout

`map_model.py` (CLI) · `xlmap/` (refs, extract, blocks, graph, render) ·
`template/app.html` (UI) · `assets/cytoscape.min.js` (MIT, vendored) ·
`tests/test_refs.py` (parser unit tests) · `tests/make_demo.py` (synthetic
DemoCo model) · `demo/` (generated demo map).
