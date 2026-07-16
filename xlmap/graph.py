"""Graph assembly: block & row edges, roles, layered layout, cycle report,
hardcode audit, error scan.

Edge semantics: an edge (provider_block -> consumer_block) with weight =
number of DISTINCT (formula-cell, provider-block) pairs — i.e. one formula
referencing the same block five times counts once, but five filled columns
of that formula count five (they are five live cells).
Ranges add an edge to EVERY block they intersect.
"""
import re
from collections import Counter, defaultdict
from bisect import bisect_right


def _block_lookup(block_list, max_row):
    starts = [r for r, _ in block_list]
    names = [n for _, n in block_list]

    def of(row):
        i = bisect_right(starts, max(row, 1)) - 1
        return names[max(i, 0)]

    def span(r0, r1):
        if r1 == -1 or r1 >= max_row:
            r1 = max_row if max_row else r0
        i0 = max(bisect_right(starts, max(r0, 1)) - 1, 0)
        i1 = max(bisect_right(starts, max(r1, r0, 1)) - 1, 0)
        return names[i0:i1 + 1] or [names[0]]
    return of, span


TRIVIAL = {'2', '3', '4', '12', '100', '1000', '10', '365', '360', '30'}


def build(wb, blocks, max_formula_len=200, samples_per_block=6):
    sheets = {s.name: s for s in wb.sheets}
    lookups = {s.name: _block_lookup(blocks[s.name], s.max_row) for s in wb.sheets}

    bedges = Counter(); bedge_ex = {}
    redges = Counter()
    binfo = defaultdict(lambda: {'nf': 0, 'tmpl': Counter(), 'sample': {}})
    rinfo = defaultdict(lambda: {'nf': 0, 'f': []})
    hardcodes = defaultdict(lambda: {'n': 0, 'consts': Counter(), 'sample': ''})
    dropped = Counter()   # external / table / error refs per sheet

    tmpl_re = re.compile(r'(\$?[A-Z]{1,3}\$?)\d+')

    for s in wb.sheets:
        of_own, _ = lookups[s.name]
        for rnum, col, ftext, refs, pr in s.formulas:
            sb = of_own(rnum)
            key = (s.name, sb)
            bi = binfo[key]
            bi['nf'] += 1
            tm = tmpl_re.sub(r'\1#', ftext)
            bi['tmpl'][tm] += 1
            if tm not in bi['sample'] and len(bi['sample']) < 40:
                bi['sample'][tm] = ftext[:max_formula_len]
            rk = (s.name, rnum)
            ri = rinfo[rk]
            ri['nf'] += 1
            if len(ri['f']) < 2 and ftext[:max_formula_len] not in ri['f']:
                ri['f'].append(ftext[:max_formula_len])
            if pr and pr.consts:
                hc = hardcodes[rk]
                hc['n'] += 1
                for cst in pr.consts:
                    if cst not in TRIVIAL:
                        hc['consts'][cst] += 1
                if not hc['sample']:
                    hc['sample'] = ftext[:max_formula_len]
            if pr:
                if pr.n_external:
                    dropped[(s.name, 'external workbook refs')] += pr.n_external
                if pr.n_table:
                    dropped[(s.name, 'table (structured) refs')] += pr.n_table
                if pr.n_error_ref:
                    dropped[(s.name, '#REF!/error refs')] += pr.n_error_ref
                    ek = (s.name, sb, '#REF!', '#REF! broken')
                    bedges[ek] += 1
                    bedge_ex.setdefault(ek, ftext[:max_formula_len])

            seen = set()
            for ref in refs:
                if ref.sheet in ('(external)', '#REF!'):
                    continue
                dsh = ref.sheet
                if dsh not in sheets:
                    dropped[(s.name, f'refs to unknown sheet {dsh!r}')] += 1
                    continue
                _, span_d = lookups[dsh]
                for db in span_d(ref.row0, ref.row1):
                    if dsh == s.name and db == sb:
                        continue
                    ek = (s.name, sb, dsh, db)
                    if ek not in seen:
                        seen.add(ek)
                        bedges[ek] += 1
                        bedge_ex.setdefault(ek, ftext[:max_formula_len])
                drow = ref.row0   # row-level: head row (ranges noted in UI)
                if not (dsh == s.name and drow == rnum):
                    redges[(s.name, rnum, dsh, drow, ref.row1 != ref.row0)] += 1

    # ---- roles & layers from sheet DAG (arrows provider -> consumer) ----
    sagg = Counter()
    for (ss, sb, ds, db), w in bedges.items():
        if ds != ss and ds != '#REF!':
            sagg[(ds, ss)] += w        # provider -> consumer
    consumers = defaultdict(set); providers = defaultdict(set)
    for (p, c), w in sagg.items():
        consumers[p].add(c); providers[c].add(p)

    layers = {}
    def layer_of(n, stack=()):
        if n in layers:
            return layers[n]
        if n in stack:
            return 0
        l = 0 if not providers[n] else 1 + max(
            layer_of(p, stack + (n,)) for p in providers[n])
        layers[n] = l
        return l
    for s in wb.sheets:
        layer_of(s.name)

    roles = {}
    for s in wb.sheets:
        n = s.name
        has_out, has_in = bool(consumers[n]), bool(providers[n])
        dens = s.n_formula / max(s.n_formula + s.n_num, 1)
        if not has_in and not has_out:
            roles[n] = 'isolated'
        elif not has_in:
            roles[n] = 'source' if dens < 0.5 else 'engine'
        elif not has_out:
            roles[n] = 'output'
        else:
            roles[n] = 'drivers' if dens < 0.25 else 'engine'

    # ---- cycles on the row graph (Tarjan, iterative) ----
    radj = defaultdict(list)
    for (ss, sr, ds, dr, isrng), w in redges.items():
        radj[(ds, dr)].append((ss, sr))     # provider -> consumer
    cycles = _sccs(radj)

    # ---- assemble ----
    blocks_out, sheet_out = [], []
    for s in wb.sheets:
        bl = blocks[s.name]
        sheet_out.append({'name': s.name, 'role': roles[s.name], 'layer': layers[s.name],
                          'state': s.state, 'nf': s.n_formula, 'ncells': s.n_formula + s.n_num + s.n_text,
                          'maxr': s.max_row, 'maxc': s.max_col,
                          'uniq': len({tmpl_re.sub(r'\1#', f[2]) for f in s.formulas})})
        for i, (r0, bn) in enumerate(bl):
            r1 = bl[i + 1][0] - 1 if i + 1 < len(bl) else max(s.max_row, r0)
            st = binfo.get((s.name, bn), {'nf': 0, 'tmpl': Counter(), 'sample': {}})
            labs, seenl = [], set()
            for rr, lab in sorted(s.row_labels.items()):
                if r0 <= rr <= r1 and lab not in seenl and len(lab) > 2:
                    seenl.add(lab); labs.append(lab)
                if len(labs) >= 10:
                    break
            tops = [{'f': st['sample'].get(t, ''), 'n': n}
                    for t, n in st['tmpl'].most_common(samples_per_block) if st['sample'].get(t)]
            blocks_out.append({'id': f'{s.name}ǀ{bn}', 'sheet': s.name, 'name': bn,
                               'r0': r0, 'r1': r1, 'nf': st['nf'], 'labels': labs, 'formulas': tops})

    if any(k[2] == '#REF!' for k in bedges):
        sheet_out.append({'name': '#REF!', 'role': 'isolated', 'layer': max(layers.values(), default=0) + 1,
                          'state': 'visible', 'nf': 0, 'ncells': 0, 'maxr': 0, 'maxc': 0, 'uniq': 0})
        blocks_out.append({'id': '#REF!ǀ#REF! broken', 'sheet': '#REF!', 'name': '#REF! broken',
                           'r0': 0, 'r1': 0, 'nf': 0, 'labels': [], 'formulas': []})

    edges_out = [{'s': f'{a}ǀ{b}', 'd': f'{c}ǀ{d}', 'w': w,
                  'ex': bedge_ex.get((a, b, c, d), '')}
                 for (a, b, c, d), w in bedges.most_common()]

    rows_out = {}
    for s in wb.sheets:
        rows_out[s.name] = {}
    for (sh, r), v in rinfo.items():
        rows_out[sh][str(r)] = {'l': sheets[sh].row_labels.get(r, ''), 'nf': v['nf'], 'f': v['f']}
    for (ss, sr, ds, dr, isrng), w in redges.items():
        rows_out[ds].setdefault(str(dr), {'l': sheets[ds].row_labels.get(dr, ''), 'nf': 0, 'f': []})
    redges_out = [{'s': f'{a}ǀ{b}', 'd': f'{c}ǀ{d}', 'w': w, 'rng': 1 if rng else 0}
                  for (a, b, c, d, rng), w in redges.most_common()]

    hc_out = []
    for (sh, r), v in sorted(hardcodes.items(), key=lambda kv: -kv[1]['n']):
        if not v['consts']:
            continue
        hc_out.append({'sheet': sh, 'row': r, 'label': sheets[sh].row_labels.get(r, ''),
                       'n': v['n'], 'consts': [c for c, _ in v['consts'].most_common(6)],
                       'sample': v['sample']})

    err_out = [{'sheet': s.name, 'cell': c, 'val': v}
               for s in wb.sheets for c, v in s.error_cells]

    audit = {
        'hardcodes': hc_out[:400],
        'errors': err_out[:400],
        'cycles': [[f'{a}ǀ{b}' for a, b in comp] for comp in cycles[:20]],
        'dropped': [{'sheet': k[0], 'what': k[1], 'n': v} for k, v in dropped.most_common()],
        'names_total': wb.n_names_total, 'names_broken': wb.n_names_broken,
        'names_resolved': len(wb.names),
        'skipped_sheets': [{'name': n, 'why': r} for n, r in wb.skipped],
    }
    return {'sheets': sheet_out, 'blocks': blocks_out, 'edges': edges_out,
            'rows': rows_out, 'redges': redges_out, 'audit': audit}


def _sccs(adj):
    """Iterative Tarjan; returns components with size>1 (true cycles)."""
    idx, low, onstk = {}, {}, set()
    stack, out, counter = [], [], [0]
    for start in list(adj):
        if start in idx:
            continue
        work = [(start, 0)]
        while work:
            node, pi = work[-1]
            if node not in idx:
                idx[node] = low[node] = counter[0]; counter[0] += 1
                stack.append(node); onstk.add(node)
            recursed = False
            children = adj.get(node, [])
            for i in range(pi, len(children)):
                ch = children[i]
                if ch not in idx:
                    work[-1] = (node, i + 1)
                    work.append((ch, 0))
                    recursed = True
                    break
                elif ch in onstk:
                    low[node] = min(low[node], idx[ch])
            if recursed:
                continue
            if low[node] == idx[node]:
                comp = []
                while True:
                    n2 = stack.pop(); onstk.discard(n2); comp.append(n2)
                    if n2 == node:
                        break
                if len(comp) > 1:
                    out.append(comp)
            work.pop()
            if work:
                parent = work[-1][0]
                low[parent] = min(low[parent], low[node])
    return out
