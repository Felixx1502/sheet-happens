"""Automatic functional-block detection per sheet.

Heuristics (in priority order):
1. Section-header rows: label whose letters are >=70% uppercase, len>=3,
   sitting at/after a blank-row gap or holding no formulas on that row.
2. Blank-gap boundaries: >=2 consecutive empty rows followed by a labeled row.
3. Fallback: one block spanning the whole sheet.

Post-processing: blocks shorter than `min_rows` merge into the previous
block; sheets are capped at `max_blocks` (smallest-formula blocks merged).
Users can override any sheet via blocks_override.json:
    {"Sheet name": [[1, "Block A"], [40, "Block B"]], ...}
"""
import json, re


def _is_headerish(label):
    letters = [c for c in label if c.isalpha()]
    if len(letters) < 3:
        return False
    return sum(1 for c in letters if c.isupper()) / len(letters) >= 0.7


def _clean(label, fallback):
    lab = re.sub(r'\s+', ' ', label).strip(' :·-')
    if not lab:
        return fallback
    if lab.isupper():
        lab = lab.title()
    return lab[:48]


def detect_blocks(sheet, min_rows=3, max_blocks=14):
    """sheet: extract.Sheet -> [(start_row, name), ...] sorted."""
    if sheet.max_row <= 1:
        return [(1, sheet.name)]
    frows = {}
    for rnum, col, ftext, refs, pr in sheet.formulas:
        frows[rnum] = frows.get(rnum, 0) + 1
    occupied = set(frows) | set(sheet.row_labels)

    cands = []
    for r, lab in sorted(sheet.row_labels.items()):
        gap_before = (r - 1 not in occupied) and (r - 2 not in occupied)
        if _is_headerish(lab) or (gap_before and r > min(occupied, default=1)):
            cands.append((r, _clean(lab, f'Section @{r}')))

    if not cands:
        return [(1, sheet.name)]
    if cands[0][0] > 1:
        first_lab = sheet.row_labels.get(1) or sheet.row_labels.get(2) or 'Header'
        cands.insert(0, (1, _clean(first_lab, 'Header')))

    # merge blocks that are too short
    merged = [cands[0]]
    for r, nm in cands[1:]:
        if r - merged[-1][0] < min_rows:
            continue
        merged.append((r, nm))

    # cap count: merge blocks with fewest formulas into predecessor
    def block_formulas(blks):
        starts = [b[0] for b in blks] + [sheet.max_row + 1]
        return [sum(v for rr, v in frows.items() if starts[i] <= rr < starts[i + 1])
                for i in range(len(blks))]

    while len(merged) > max_blocks:
        counts = block_formulas(merged)
        idx = max(1, counts.index(min(counts[1:])))   # never drop the first
        merged.pop(idx)

    # dedupe names
    seen = {}
    out = []
    for r, nm in merged:
        if nm in seen:
            seen[nm] += 1
            nm = f'{nm} ({seen[nm]})'
        else:
            seen[nm] = 1
        out.append((r, nm))
    return out


def load_blocks(wb, override_path=None, min_rows=3, max_blocks=14):
    override = {}
    if override_path:
        with open(override_path, encoding='utf-8') as f:
            override = json.load(f)
    blocks = {}
    for s in wb.sheets:
        if s.name in override:
            blocks[s.name] = sorted((int(r), str(n)) for r, n in override[s.name])
        else:
            blocks[s.name] = detect_blocks(s, min_rows, max_blocks)
    return blocks
