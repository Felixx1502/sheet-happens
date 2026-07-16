"""Canonical Excel formula reference parser.

Fixes applied vs. naive parsing (see README "Accuracy notes"):
- escaped quotes in sheet names ('It''s'!A1)
- external workbook refs [1]Sheet!A1 -> classified external, never local
- 3D refs Sheet1:Sheet3!A1 -> both endpoint sheets
- error literals (#REF!, #VALUE!, ...) never become sheet names
- ranges keep BOTH endpoints (A5:A200 -> rows 5..200, not just 5)
- full-column (A:A) and full-row (5:10) refs
- defined names resolved via workbook name table
- structured table refs detected & counted (not mapped)
- bare-ref regex cannot prefix-match names like FY26_tax
- absolute-row flags captured positionally (not substring search)
- per-formula constant extraction for hardcode audit
"""
import re
from dataclasses import dataclass, field

ERROR_TOKENS = re.compile(r'#(?:REF|VALUE|NAME\??|NUM|NULL|N/A|DIV/0)!?', re.I)
STRINGS = re.compile(r'"(?:[^"]|"")*"')
EXTERNAL = re.compile(r"(?:'\[\d+\][^']*'|\[\d+\][A-Za-z_0-9. ]*)!")
TABLE_REF = re.compile(r'[A-Za-z_\\][\w.]*\[(?:[^\[\]]|\[[^\]]*\])*\]|\[@[^\]]*\]')
QSHEET = re.compile(r"'((?:[^']|'')+)'!")
USHEET = re.compile(r"(?<![A-Za-z0-9_.\]'])([A-Za-z_][\w.]{0,30})(?::([A-Za-z_][\w.]{0,30}))?!")
# range after a sheet marker or bare; captures abs flags positionally
CELL = r'(\$?)([A-Z]{1,3})(\$?)(\d{1,7})'
RANGE_AFTER = re.compile(
    r'(?:' + CELL + r'(?::' + CELL + r')?'      # A1 or A1:B2
    r'|(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})'      # A:B
    r'|(\$?)(\d{1,7}):(\$?)(\d{1,7})'            # 5:10
    r')')
BARE = re.compile(
    r'(?<![A-Za-z0-9_$!:.\]])'
    r'(\$?)([A-Z]{1,3})(\$?)(\d{1,7})'
    r'(?::(\$?)([A-Z]{1,3})(\$?)(\d{1,7}))?'
    r'(?![A-Za-z0-9_(])')
NAME_TOKEN = re.compile(r"(?<![A-Za-z0-9_.'\[])([A-Za-z_\\][\w.]{0,254})(?![\w.(])")
NUMBER = re.compile(r'(?<![\w.$])(\d+(?:\.\d+)?)(?:[eE][+-]\d+)?(?![\w$])')
FUNC_WORDS = {'TRUE', 'FALSE'}


@dataclass
class Ref:
    sheet: str          # resolved sheet name; '#REF!' for broken; '' impossible
    row0: int
    row1: int           # inclusive; row1 == -1 means "to sheet end" (full column)
    abs0: bool = False
    kind: str = 'local' # local | xsheet | external | name | error | d3


@dataclass
class Parsed:
    refs: list = field(default_factory=list)
    n_external: int = 0
    n_table: int = 0
    n_error_ref: int = 0
    names_used: list = field(default_factory=list)
    consts: list = field(default_factory=list)   # non-trivial numeric literals


def unescape_sheet(s):
    return s.replace("''", "'")


def _rows_from_range(m, base):
    """m: RANGE_AFTER match -> (row0,row1,abs0) or None"""
    g = m.groups()
    if g[3]:                       # cell form
        r0 = int(g[3]); a0 = g[2] == '$'
        r1 = int(g[7]) if g[7] else r0
        return (min(r0, r1), max(r0, r1), a0)
    if g[9]:                       # col:col  A:B  -> whole column
        return (1, -1, True)
    if g[13]:                      # row:row
        r0, r1 = int(g[13]), int(g[16])
        return (min(r0, r1), max(r0, r1), True)
    return None


def parse_formula(formula, own_sheet, names=None, known_sheets=None):
    """Parse one formula. names: {NAME_UPPER: [(sheet,row0,row1),...]}.
    known_sheets: set of workbook sheet names (case preserved) for validating
    unquoted matches; unknown unquoted 'sheets' fall through to name lookup."""
    p = Parsed()
    s = STRINGS.sub('""', formula)

    n_err = len(ERROR_TOKENS.findall(s))
    if n_err:
        # broken refs like #REF!!A1 or SUM(#REF!) -> record, remove
        p.n_error_ref = n_err
        s = ERROR_TOKENS.sub(' ', s)

    ext = EXTERNAL.findall(s)
    if ext:
        p.n_external = len(ext)
        p.refs += [Ref('(external)', 1, 1, True, 'external')] * len(ext)
        s = EXTERNAL.sub(' ', s)
        s = RANGE_AFTER.sub(' ', s, count=p.n_external)  # consume their ranges

    tbl = TABLE_REF.findall(s)
    if tbl:
        p.n_table = len(tbl)
        s = TABLE_REF.sub(' ', s)

    out_sheeted = []

    def eat_qsheet(m):
        sheets = [unescape_sheet(m.group(1))]
        if ':' in sheets[0]:                      # quoted 3D 'S1:S3'!
            a, _, b = sheets[0].partition(':')
            sheets = [a, b]
        out_sheeted.append((sheets, m.end()))
        return '\x01' * (m.end() - m.start())     # keep offsets stable

    s = QSHEET.sub(eat_qsheet, s)

    def eat_usheet(m):
        names_ = [m.group(1)] + ([m.group(2)] if m.group(2) else [])
        if known_sheets is not None and m.group(1) not in known_sheets and \
           (len(names_) == 1 or m.group(2) not in known_sheets):
            return m.group(0)                     # not a sheet -> maybe a name; leave
        out_sheeted.append((names_, m.end()))
        return '\x01' * (m.end() - m.start())

    s = USHEET.sub(eat_usheet, s)

    # attach ranges that directly follow each sheet marker
    consumed = []
    for sheets, endpos in out_sheeted:
        m = RANGE_AFTER.match(s, endpos)
        rows = _rows_from_range(m, endpos) if m else None
        if m:
            consumed.append((m.start(), m.end()))
        for sh in sheets:
            kind = 'd3' if len(sheets) > 1 else ('xsheet' if sh != own_sheet else 'local')
            if rows:
                p.refs.append(Ref(sh, rows[0], rows[1], rows[2], kind))
            else:
                p.refs.append(Ref(sh, 1, -1, True, kind))   # sheet-level (rare)
    # blank consumed ranges so BARE doesn't re-match them
    if consumed:
        sl = list(s)
        for a, b in consumed:
            sl[a:b] = '\x01' * (b - a)
        s = ''.join(sl)

    for m in BARE.finditer(s):
        r0 = int(m.group(4)); a0 = m.group(3) == '$'
        r1 = int(m.group(8)) if m.group(8) else r0
        p.refs.append(Ref(own_sheet, min(r0, r1), max(r0, r1), a0, 'local'))
    s = BARE.sub(lambda m: '\x01' * (m.end() - m.start()), s)

    if names:
        for m in NAME_TOKEN.finditer(s):
            tok = m.group(1)
            if tok.upper() in FUNC_WORDS:
                continue
            hit = names.get(tok.upper())
            if hit:
                p.names_used.append(tok)
                for sh, r0, r1 in hit:
                    p.refs.append(Ref(sh, r0, r1, True, 'name'))

    # constants for hardcode audit (strip ref placeholders first)
    body = s.replace('\x01', ' ')
    for m in NUMBER.finditer(body):
        v = float(m.group(1))
        if v not in (0.0, 1.0):
            p.consts.append(m.group(0))
    return p


def shift_rows(refs, delta):
    """Shared-formula child: shift relative rows by delta (rows only; a
    horizontal shared fill keeps rows identical, which this preserves)."""
    out = []
    for r in refs:
        if r.abs0 or r.row1 == -1:
            out.append(r)
        else:
            out.append(Ref(r.sheet, r.row0 + delta, r.row1 + delta, r.abs0, r.kind))
    return out
