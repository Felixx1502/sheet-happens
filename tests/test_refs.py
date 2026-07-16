#!/usr/bin/env python3
"""Unit tests for the canonical reference parser (run: python tests/test_refs.py)."""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from xlmap.refs import parse_formula

SHEETS = {'Data', 'P&L', "It's", 'Q1', 'FY2026', 'Rev Build', 'S1', 'S3'}
FAILS = []


def check(desc, cond):
    print(('PASS ' if cond else 'FAIL ') + desc)
    if not cond:
        FAILS.append(desc)


def refs_of(f, own='Own', names=None):
    return parse_formula(f, own, names, SHEETS).refs


# 1. simple cross-sheet
r = refs_of('=Data!B5*2')
check('cross-sheet ref', len(r) == 1 and r[0].sheet == 'Data' and r[0].row0 == 5)
# 2. quoted with space
r = refs_of("='Rev Build'!C10")
check('quoted sheet with space', r and r[0].sheet == 'Rev Build' and r[0].row0 == 10)
# 3. escaped quote in sheet name
r = refs_of("='It''s'!A1+1")
check("escaped quote sheet ('It''s')", r and r[0].sheet == "It's" and r[0].row0 == 1)
# 4. range keeps both endpoints
r = refs_of('=SUM(Data!A5:A200)')
check('range endpoints 5..200', r and r[0].row0 == 5 and r[0].row1 == 200)
# 5. bare range on own sheet
r = refs_of('=SUM(B10:B20)')
check('bare range 10..20 on own sheet', r and r[0].sheet == 'Own' and (r[0].row0, r[0].row1) == (10, 20))
# 6. full column
r = refs_of('=SUM(Data!A:A)')
check('full column -> rows 1..end', r and r[0].row0 == 1 and r[0].row1 == -1)
# 7. defined name resolution
names = {'TAX_RATE': [('Data', 99, 99)]}
p = parse_formula('=B5*TAX_RATE', 'Own', names, SHEETS)
check('defined name resolved', any(x.sheet == 'Data' and x.row0 == 99 for x in p.refs)
      and 'TAX_RATE' in p.names_used)
# 8. name NOT prefix-matched as bare ref (FY26_tax must not become FY26 / row 26)
p = parse_formula('=FY26_tax*2', 'Own', {}, SHEETS)
check('no prefix-match of FY26_tax', not p.refs)
# 9. function names not matched (LOG10, ATAN2)
p = parse_formula('=LOG10(B5)+ATAN2(1,2)', 'Own', {}, SHEETS)
check('LOG10/ATAN2 not refs', len(p.refs) == 1 and p.refs[0].row0 == 5)
# 10. error literal never a sheet
p = parse_formula('=#REF!+Data!A1', 'Own', {}, SHEETS)
check('#REF! flagged, not a sheet', p.n_error_ref == 1
      and all(x.sheet != 'REF' for x in p.refs)
      and any(x.sheet == 'Data' for x in p.refs))
# 11. external workbook ref classified external
p = parse_formula("='[1]Ext Sheet'!B2+[2]Other!C3", 'Own', {}, SHEETS)
check('external refs not local', p.n_external == 2
      and all(x.kind == 'external' for x in p.refs if x.sheet == '(external)'))
# 12. 3D ref hits both endpoint sheets
p = parse_formula('=SUM(S1:S3!B2)', 'Own', {}, SHEETS)
shs = {x.sheet for x in p.refs}
check('3D ref endpoints', shs == {'S1', 'S3'})
# 13. structured table ref counted, not mapped
p = parse_formula('=SUM(Table1[Amount])+B2', 'Own', {}, SHEETS)
check('table ref counted', p.n_table == 1 and len(p.refs) == 1)
# 14. sheet named like a cell ref must stay quoted-resolvable
r = refs_of("='Q1'!B2")
check('sheet named Q1', r and r[0].sheet == 'Q1' and r[0].row0 == 2)
# 15. strings stripped (no refs from text)
p = parse_formula('=IF(A1>0,"see B99","")', 'Own', {}, SHEETS)
check('no refs from strings', all(x.row0 != 99 for x in p.refs))
# 16. constants extracted for hardcode audit; 0/1 excluded
p = parse_formula('=B5*8.5+0+1', 'Own', {}, SHEETS)
check('consts: 8.5 kept, 0/1 dropped', p.consts == ['8.5'])
# 17. absolute-row flag positional (B2:C$25 -> relative start row)
r = refs_of('=SUM(B2:C$25)')
check('abs flag not substring-matched', r and r[0].row0 == 2 and r[0].abs0 is False)
# 18. TRUE/FALSE not names
p = parse_formula('=IF(TRUE,1,FALSE)', 'Own', {'TRUE': [('Data', 1, 1)]}, SHEETS)
check('TRUE/FALSE excluded from names', not p.refs)

print()
if FAILS:
    print(f'{len(FAILS)} FAILURES'); sys.exit(1)
print('all tests passed')
