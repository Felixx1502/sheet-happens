"""Render the interactive HTML from graph data + template. Optional
anonymization for sharing the map's structure without model content."""
import json, os, datetime, hashlib

SEP = 'ǀ'


def _find(name, pkg_dir):
    for d in (os.path.join(pkg_dir, '..', 'template'), os.path.join(pkg_dir, '..', 'assets'),
              os.path.join(pkg_dir, 'template'), os.path.join(pkg_dir, 'assets')):
        p = os.path.join(d, name)
        if os.path.isfile(p):
            return p
    raise FileNotFoundError(name)


def anonymize(data):
    """Strip labels/formulas; pseudonymise sheet & block names.
    Structure (nodes, edges, weights, row numbers) is preserved."""
    def pseud(s):
        return 'Sheet-' + hashlib.sha1(s.encode()).hexdigest()[:6]
    smap = {s['name']: pseud(s['name']) for s in data['sheets']}
    idmap = {}
    for b in data['blocks']:
        nsh = smap[b['sheet']]
        nnm = f'Block r{b["r0"]}-{b["r1"]}'
        idmap[b['id']] = nsh + SEP + nnm
        b['sheet'], b['name'], b['id'] = nsh, nnm, idmap[b['id']]
        b['labels'], b['formulas'] = [], []
    for s in data['sheets']:
        s['name'] = smap[s['name']]
    for e in data['edges']:
        e['s'] = idmap.get(e['s'], e['s'])
        e['d'] = idmap.get(e['d'], e['d'])
        e['ex'] = ''
    data['rows'] = {smap.get(sh, sh): {r: {'l': '', 'nf': v['nf'], 'f': []}
                                       for r, v in rows.items()}
                    for sh, rows in data['rows'].items()}
    for e in data['redges']:
        for k in ('s', 'd'):
            i = e[k].index(SEP)
            e[k] = smap.get(e[k][:i], e[k][:i]) + SEP + e[k][i + 1:]
    a = data['audit']
    a['hardcodes'] = [{'sheet': smap.get(x['sheet'], x['sheet']), 'row': x['row'],
                       'label': '', 'n': x['n'], 'consts': ['…'], 'sample': ''}
                      for x in a['hardcodes']]
    a['errors'] = [{'sheet': smap.get(x['sheet'], x['sheet']), 'cell': x['cell'], 'val': x['val']}
                   for x in a['errors']]
    a['dropped'] = [{'sheet': smap.get(x['sheet'], x['sheet']), 'what': x['what'], 'n': x['n']}
                    for x in a['dropped']]
    a['cycles'] = [[smap.get(k[:k.index(SEP)], k[:k.index(SEP)]) + SEP + k[k.index(SEP) + 1:]
                    for k in comp] for comp in a['cycles']]
    a['skipped_sheets'] = []
    return smap


def render(data, out_path, title, source_name, pkg_dir):
    tpl = open(_find('app.html', pkg_dir), encoding='utf-8').read()
    cyto = open(_find('cytoscape.min.js', pkg_dir), encoding='utf-8').read()
    payload = json.dumps(data, ensure_ascii=False).replace('</', '<\\/')
    html = (tpl.replace('__CYTO__', cyto)
               .replace('__DATA__', payload)
               .replace('__TITLE__', title.replace('<', '‹').replace('>', '›').replace('"', "'"))
               .replace('__SOURCE__', source_name.replace('<', '‹'))
               .replace('__GENERATED__', datetime.date.today().isoformat()))
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    return out_path
