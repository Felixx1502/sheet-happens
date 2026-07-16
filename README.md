# sheet-happens — an Excel model X-ray

> *For all the analyst monkeys, working on their models.*

**Point it at any `.xlsx` financial model → get an interactive map of how it
actually works.** Sheet flows, functional blocks, line-item dependencies,
formula-by-formula navigation — in one self-contained HTML file. No add-ins,
no upload: your model never leaves your machine.

![demo](docs/screenshot.png)

## Use it as a Claude skill

Install once, then just ask.

**Claude Code:**

```bash
git clone https://github.com/Felixx1502/sheet-happens ~/.claude/skills/sheet-happens
```

**Claude Desktop / Cowork:** download `sheet-happens.skill` from the
[Releases](https://github.com/Felixx1502/sheet-happens/releases) page,
open it, hit "Save skill".

Then type **`/sheet-happens`** in any chat, or simply say *"map this Excel
model"* and point Claude at your .xlsx. Claude runs the pipeline locally and
hands back one interactive HTML file. No API keys, no upload, no setup.

## Why

Every inherited model is a black box: 30 tabs, 30,000 formulas, no
documentation. Excel's *Trace Precedents* shows one cell at a time. Commercial
model-audit add-ins are paid, Windows-only, and IT-approval-heavy.
`sheet-happens` is free, local, and produces something you can open in any
browser and hand to a colleague.

## Quick start

```bash
pip install lxml
python map_model.py "My Model.xlsx"
# -> "My Model - Model Map.html"
```

Try it on the synthetic demo model:

```bash
pip install openpyxl   # only needed to build the demo/tests
python tests/make_demo.py
python map_model.py demo/DemoCo_model.xlsx -o demo/DemoCo_map.html
```

## What you get

- **Three zoom levels** — sheet graph (laid out in calculation order) →
  auto-detected functional blocks (Revenue, COGS, Debt schedule…) → line-item
  view: every row, its formulas, and row-to-row links with external feeders
  and consumers.
- **Data-flow arrows** — source → consumer, so you read the model the way it
  calculates.
- **Click-through tracing** — click any dependency to jump to it; trace the
  full upstream/downstream closure of any block.
- **Built-in audit** — hardcoded constants inside formulas (the #1 thing
  model reviewers hunt), error cells (#REF!, #DIV/0!…), row-level circular
  groups, defined-name health, and an explicit list of what could **not** be
  mapped (external links, table refs) so you know what you're not seeing.
- **`--anonymize`** — strips labels & formulas, pseudonymises sheets, keeps
  structure. For sharing a model's shape without its content.

## Accuracy notes (read this)

The parser handles quoted/escaped sheet names, 3D refs, full-column refs,
ranges (mapped to **every** block they cross), shared formulas, defined
names, and never confuses `FY26_tax` with a cell ref or `#VALUE!` with a
sheet. But static analysis has limits:

| Limitation | Impact |
|---|---|
| INDIRECT / OFFSET / CHOOSE | dynamic targets not traceable; OFFSET maps to its anchor |
| Item-level range links | point to the range's head row (block level is exact); marked dashed |
| Row-level circular groups | often *legitimate* corkscrews (opening = prior closing) — verify at cell level |
| External workbooks, table refs | counted in the Audit panel, not drawn |
| Block boundaries | heuristic; pin them with `--blocks-override blocks.json` |
| Column dimension | rows are the atom; a hardcoded override in one quarter of an otherwise clean row is not separately flagged |

## Confidentiality

The generated HTML **embeds row labels and formula text from your model**.
Treat it with the same care as the model itself. For anything public, use
`--anonymize` and check the output before sharing. Never commit maps of real
client models to a repo.

## Options

```
python map_model.py MODEL.xlsx [-o out.html] [--title "..."]
                    [--blocks-override blocks.json] [--anonymize]
                    [--max-blocks 14] [--min-block-rows 3]
```

`blocks.json` example:

```json
{ "WORKING": [[1, "Volume build"], [19, "Revenue"], [33, "COGS"]] }
```

## Tests

```bash
python tests/test_refs.py
```

18 parser cases: escaped-quote sheet names, defined-name resolution,
prefix-match traps (FY26_tax), function-name traps (LOG10/ATAN2), error
literals, external refs, 3D refs, table refs, range endpoints, absolute-row
flags, string stripping, constant extraction.

## License

MIT. Bundles [cytoscape.js](https://js.cytoscape.org/) (MIT) — see NOTICE.
