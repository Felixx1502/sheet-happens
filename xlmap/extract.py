"""Streaming workbook extraction — pure Python, no temp dirs, no shell tools.

Reads the xlsx as a zip; parses workbook/rels/definedNames with lxml by
attribute NAME (never attribute order); streams every worksheet with a
row-clearing iterparse so 400MB ghost-formatted sheets don't blow memory.
Handles: missing sharedStrings, inline strings, shared formulas (row-delta
resolution), error cells, hidden sheets, chartsheets (skipped, reported).
"""
import re, zipfile
from collections import Counter
from lxml import etree

from .refs import parse_formula, shift_rows, ERROR_TOKENS

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'
AREA = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]{0,30}))!\$?[A-Z]{0,3}\$?(\d+)(?::\$?[A-Z]{0,3}\$?(\d+))?")


def col_to_num(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


CREF = re.compile(r'^([A-Z]{1,3})(\d+)$')


class Sheet:
    def __init__(self, name, state):
        self.name, self.state = name, state
        self.formulas = []      # (row, col, formula_text, Parsed)
        self.row_labels = {}    # row -> first text label (cols A..H)
        self.error_cells = []   # (cellref, errval)
        self.n_num = self.n_text = self.n_formula = 0
        self.max_row = self.max_col = 0


class Workbook:
    def __init__(self):
        self.sheets = []        # ordered Sheet objects (worksheets only)
        self.names = {}         # NAME_UPPER -> [(sheet,row0,row1)]
        self.skipped = []       # (name, reason) e.g. chartsheets
        self.n_names_total = 0
        self.n_names_broken = 0


def _load_shared_strings(zf):
    sst = []
    if 'xl/sharedStrings.xml' not in zf.namelist():
        return sst
    with zf.open('xl/sharedStrings.xml') as f:
        for _, si in etree.iterparse(f, tag=NS + 'si'):
            # plain text runs only; skip phonetic runs (rPh)
            txt = ''.join(t.text or '' for t in si.iter(NS + 't')
                          if t.getparent().tag != NS + 'rPh')
            sst.append(txt)
            si.clear()
    return sst


def _defined_names(root, valid_sheets):
    names, total, broken = {}, 0, 0
    dn = root.find(NS + 'definedNames')
    if dn is None:
        return names, total, broken
    for el in dn.findall(NS + 'definedName'):
        total += 1
        val = el.text or ''
        if ERROR_TOKENS.search(val):
            broken += 1
            continue
        areas = []
        for m in AREA.finditer(val):
            sh = (m.group(1) or m.group(2) or '').replace("''", "'")
            if sh not in valid_sheets:
                continue
            r0 = int(m.group(3)); r1 = int(m.group(4)) if m.group(4) else r0
            areas.append((sh, min(r0, r1), max(r0, r1)))
        if areas:
            nm = (el.get('name') or '').upper()
            if nm and not nm.startswith('_XLNM'):
                names.setdefault(nm, []).extend(areas)
    return names, total, broken


BIG_SHEET_BYTES = 40_000_000   # uncompressed; above this use text scanning

F_CELL = re.compile(
    r'<c r="([A-Z]{1,3})(\d{1,7})"[^>]{0,200}?>'
    r'<f([^>]{0,200}?)(?:/>|>([^<]{0,8000})</f>)')
L_CELL = re.compile(
    r'<c r="([A-H])(\d{1,7})"[^>]{0,200}?t="s"[^>]{0,200}?><v>(\d{1,10})</v>')
E_CELL = re.compile(
    r'<c r="([A-Z]{1,3}\d{1,7})"[^>]{0,200}?t="e"[^>]{0,200}?><v>([^<]{0,100})</v>')
ROW_TAG = re.compile(r'<row r="(\d{1,7})"')
SI_ATTR = re.compile(r'si="(\d+)"')
import html as _html


def _scan_big_sheet(zf, tgt, name, state, sst, wbnames, valid, log):
    """Regex text-scan for pathologically large sheets (e.g. millions of
    empty styled cells). C-level scanning instead of per-cell Python loops."""
    s = Sheet(name, state)
    shared, pending = {}, []
    seen_cells = set()
    CH, OV = 32 * 1024 * 1024, 64 * 1024

    def _cell_at_tail(seg):
        """cell (colnum,row) from the '<c r=\"..\"' nearest the END of seg."""
        j = seg.rfind('<c r="')
        if j < 0:
            return None, None, ''
        m = CREF.match(seg[j + 6:j + 16].split('"', 1)[0])
        if not m:
            return None, None, ''
        return col_to_num(m.group(1)), int(m.group(2)), seg[j:]

    def process(text, count_upto):
        for m in ROW_TAG.finditer(text):
            r = int(m.group(1))
            if r > s.max_row:
                s.max_row = r
        # formulas: C-speed split on '<f' (attr chars or '>' always follow)
        parts = text.split('<f')
        for i in range(1, len(parts)):
            head = parts[i]
            if head[:1] not in (' ', '>', '/'):
                continue                      # '<frame...' etc. — not a formula
            col, rnum, _ = _cell_at_tail(parts[i - 1])
            if col is None or (rnum, col) in seen_cells:
                continue
            seen_cells.add((rnum, col))
            if col > s.max_col:
                s.max_col = col
            gt = head.find('>')
            if gt < 0:
                continue
            attrs = head[:gt]
            if attrs.endswith('/'):
                attrs, body = attrs[:-1], None
            else:
                end = head.find('</f>', gt)
                body = head[gt + 1:end] if end >= 0 else None
            sim = SI_ATTR.search(attrs)
            si = sim.group(1) if sim else None
            if body:
                ftext = _html.unescape(body)
                pr = parse_formula(ftext, name, wbnames, valid)
                s.formulas.append((rnum, col, ftext, pr.refs, pr))
                if 'shared' in attrs and si is not None:
                    shared[si] = (rnum, ftext, pr.refs, pr)
            elif 'shared' in attrs and si is not None:
                pending.append((rnum, col, si))
        # labels (cols A-H, shared strings)
        for m in L_CELL.finditer(text):
            rnum = int(m.group(2))
            if rnum not in s.row_labels:
                try:
                    txt = sst[int(m.group(3))].strip()
                    if txt:
                        s.row_labels[rnum] = txt[:120]
                except (ValueError, IndexError):
                    pass
        # error cells: split on the rare t="e" marker
        eparts = text.split('t="e"')
        for i in range(1, len(eparts)):
            col, rnum, seg = _cell_at_tail(eparts[i - 1])
            if col is None:
                continue
            vpos = eparts[i].find('<v>')
            if 0 <= vpos < 200:
                vend = eparts[i].find('</v>', vpos)
                val = eparts[i][vpos + 3:vend][:100] if vend > 0 else ''
                s.error_cells.append((seg[6:].split('"', 1)[0], _html.unescape(val)))
        s.n_num += text.count('<v>', 0, count_upto)

    with zf.open(tgt) as f:
        tail = ''
        while True:
            chunk = f.read(CH)
            if not chunk:
                if tail:
                    process(tail, len(tail))
                break
            text = tail + chunk.decode('utf-8', errors='replace')
            if len(chunk) == CH:              # more coming: keep overlap
                cut = len(text) - OV
                process(text, cut)
                tail = text[cut:]
            else:                             # final partial chunk
                process(text, len(text))
                tail = ''
                break
    s.error_cells = list(dict.fromkeys(s.error_cells))
    for rnum, col, si in pending:
        mm = shared.get(si)
        if mm:
            s.formulas.append((rnum, col, mm[1], shift_rows(mm[2], rnum - mm[0]), mm[3]))
    s.n_formula = len(s.formulas)
    s.n_num = max(s.n_num - len(s.row_labels) - len(s.error_cells) - len(s.formulas), 0)
    log(f"  scanned {name!r} (large sheet, text mode): {s.n_formula:,} formulas, "
        f"{len(s.row_labels):,} labels, max row {s.max_row}")
    return s


def extract(path, log=print):
    zf = zipfile.ZipFile(path)
    wbroot = etree.fromstring(zf.read('xl/workbook.xml'))
    rels = etree.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
    rid2target = {}
    for rel in rels.iter(PKG + 'Relationship'):
        t = rel.get('Type', '')
        tgt = rel.get('Target', '')
        if tgt.startswith('/'):
            tgt = tgt[1:]
        elif not tgt.startswith('xl/'):
            tgt = 'xl/' + tgt
        rid2target[rel.get('Id')] = (t, tgt)

    wb = Workbook()
    order = []
    for sh in wbroot.find(NS + 'sheets').findall(NS + 'sheet'):
        name = sh.get('name')
        state = sh.get('state') or 'visible'
        rid = sh.get(RNS + 'id')
        typ, tgt = rid2target.get(rid, ('', ''))
        if 'worksheet' not in typ:
            wb.skipped.append((name, 'chartsheet/other'))
            continue
        if tgt not in zf.namelist():
            wb.skipped.append((name, 'missing part'))
            continue
        order.append((name, state, tgt))
    valid = {n for n, _, _ in order}
    wb.names, wb.n_names_total, wb.n_names_broken = _defined_names(wbroot, valid)
    sst = _load_shared_strings(zf)

    for name, state, tgt in order:
        if zf.getinfo(tgt).file_size > BIG_SHEET_BYTES:
            wb.sheets.append(_scan_big_sheet(zf, tgt, name, state, sst,
                                             wb.names, valid, log))
            continue
        s = Sheet(name, state)
        shared = {}     # si -> (master_row, formula, refs)
        pending = []    # (row, col, si) children seen before master (rare)
        with zf.open(tgt) as f:
            for _, rowel in etree.iterparse(f, tag=NS + 'row'):
                rnum = int(rowel.get('r') or 0)
                if rnum > s.max_row:
                    s.max_row = rnum
                for c in rowel.iter(NS + 'c'):
                    cref = c.get('r') or ''
                    m = CREF.match(cref)
                    col = col_to_num(m.group(1)) if m else 0
                    if col > s.max_col:
                        s.max_col = col
                    t = c.get('t')
                    fEl = c.find(NS + 'f')
                    vEl = c.find(NS + 'v')
                    if fEl is not None:
                        ftext, ftype, si = fEl.text, fEl.get('t'), fEl.get('si')
                        if ftype == 'shared' and not ftext and si is not None:
                            mm = shared.get(si)
                            if mm:
                                s.n_formula += 1
                                s.formulas.append((rnum, col, mm[1],
                                                   shift_rows(mm[2], rnum - mm[0]), mm[3]))
                            else:
                                pending.append((rnum, col, si))
                        elif ftext:
                            pr = parse_formula(ftext, name, wb.names, valid)
                            s.n_formula += 1
                            s.formulas.append((rnum, col, ftext, pr.refs, pr))
                            if ftype == 'shared' and si is not None:
                                shared[si] = (rnum, ftext, pr.refs, pr)
                        if t == 'e' and vEl is not None:
                            s.error_cells.append((cref, vEl.text or ''))
                    elif t == 'e' and vEl is not None:
                        s.error_cells.append((cref, vEl.text or ''))
                    elif t == 'inlineStr':
                        isEl = c.find(NS + 'is')
                        txt = ''.join(x.text or '' for x in isEl.iter(NS + 't')) if isEl is not None else ''
                        s.n_text += 1
                        if txt.strip() and col <= 8 and rnum not in s.row_labels:
                            s.row_labels[rnum] = txt.strip()[:120]
                    elif vEl is not None:
                        if t == 's':
                            s.n_text += 1
                            if col <= 8 and rnum not in s.row_labels:
                                try:
                                    txt = sst[int(vEl.text)].strip()
                                    if txt:
                                        s.row_labels[rnum] = txt[:120]
                                except (ValueError, IndexError):
                                    pass
                        elif t == 'str':
                            s.n_text += 1
                        else:
                            s.n_num += 1
                rowel.clear()
                while rowel.getprevious() is not None:
                    del rowel.getparent()[0]
        for rnum, col, si in pending:
            mm = shared.get(si)
            if mm:
                s.n_formula += 1
                s.formulas.append((rnum, col, mm[1], shift_rows(mm[2], rnum - mm[0]), mm[3]))
        wb.sheets.append(s)
        log(f"  parsed {name!r}: {s.n_formula:,} formulas, "
            f"{len(s.row_labels):,} labels, max {s.max_row}x{s.max_col}"
            + (f", {len(s.error_cells)} error cells" if s.error_cells else ""))
    zf.close()
    return wb

# end of module
