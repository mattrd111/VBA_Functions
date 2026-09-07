#!/usr/bin/env python3
"""UAT DATA CHECK - source-table integrity as the VBA (Sheet730.cls / Main.bas) assumes it.

Runs entirely offline against HP_Flat_File_Test_v3.3.xlsb (pyxlsb for cell values, raw BIFF12
record parsing for defined names, data validations and formula-vs-constant classification).
Output is plain text; RESULT.md is written by hand from this output.
"""
import sys, os, re, ast, struct, mmap, collections, datetime

SCR = '/tmp/claude-0/-home-user-VBA-Functions/4411ca91-ef32-5ffa-9c78-fab36847f3b2/scratchpad'
sys.path.insert(0, SCR)
os.chdir(SCR)
from parse_biff import records, wstr
from pyxlsb import open_workbook

XLSB = os.path.join(SCR, 'HP_Flat_File_Test_v3.3.xlsb')
WS = os.path.join(SCR, 'xlsb/xl/worksheets')

# ----------------------------------------------------------------------------- helpers
def colL(c0):
    s = ''; c = c0 + 1
    while c:
        c, r = divmod(c - 1, 26); s = chr(65 + r) + s
    return s

def colN(letters):
    n = 0
    for ch in letters: n = n * 26 + (ord(ch) - 64)
    return n - 1

def serial_to_date(v):
    try:
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).strftime('%Y-%m-%d')
    except Exception:
        return None

def is_blank(v):
    return v is None or (isinstance(v, str) and v == '')

def section(title):
    print('\n' + '=' * 100); print(title); print('=' * 100)

# ----------------------------------------------------------------------------- workbook structure (BIFF12)
wbbin = open('xlsb/xl/workbook.bin', 'rb').read()
sheet_order = []          # 0-based sheet index -> name (BrtBundleSh order)
externsheet = []          # ixti -> (iSup, itabFirst, itabLast)
names = []                # BrtName in file order: (name, itab, flags, rgce)
for rid, body in records(wbbin):
    if rid == 0x9C:       # BrtBundleSh
        hs, itab = struct.unpack_from('<II', body, 0)
        rel, off = wstr(body, 8); nm, off = wstr(body, off)
        sheet_order.append(nm)
    elif rid == 0x16A:    # BrtExternSheet
        cnt = struct.unpack_from('<I', body, 0)[0]
        for i in range(cnt):
            externsheet.append(struct.unpack_from('<III', body, 4 + 12 * i))
    elif rid == 0x27:     # BrtName
        flags = struct.unpack_from('<I', body, 0)[0]
        itab = struct.unpack_from('<I', body, 5)[0]
        nm, off = wstr(body, 9)
        cce = struct.unpack_from('<I', body, off)[0]
        rgce = body[off + 4: off + 4 + cce]
        names.append((nm, itab, flags, rgce))

def sheet_for_ixti(ixti):
    if ixti >= len(externsheet): return f'<ixti {ixti} out of range>'
    iSup, a, b = externsheet[ixti]
    if iSup != 0: return f'<external iSup={iSup} itab={a}>'
    if a == 0xFFFFFFFF: return '<no sheet>'
    if a == 0xFFFFFFFE: return '<deleted sheet>'
    nm = sheet_order[a] if a < len(sheet_order) else f'<sheet {a}>'
    if a != b: nm += ':' + (sheet_order[b] if b < len(sheet_order) else f'<sheet {b}>')
    return nm

def qs(nm):
    return nm if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_.]*', nm) else "'" + nm.replace("'", "''") + "'"

FUNC = {0:'COUNT',1:'IF',2:'ISNA',3:'ISERROR',4:'SUM',5:'AVERAGE',6:'MIN',7:'MAX',8:'ROW',9:'COLUMN',10:'NA',
        29:'INDEX',64:'MATCH',78:'OFFSET',82:'SEARCH',100:'CHOOSE',101:'HLOOKUP',102:'VLOOKUP',105:'ISREF',
        115:'LEFT',116:'RIGHT',117:'MID',118:'TRIM',124:'FIND',148:'INDIRECT',162:'CLEAN',169:'COUNTA',
        183:'PRODUCT',184:'FACT',219:'ADDRESS',220:'DAYS360',221:'TODAY',255:'USER_DEFINED',
        ' ':''}
OPS = {0x03:'+',0x04:'-',0x05:'*',0x06:'/',0x07:'^',0x08:'&',0x09:'<',0x0A:'<=',0x0B:'=',0x0C:'>=',0x0D:'>',0x0E:'<>',
       0x0F:' ',0x10:',',0x11:':'}

def ref_txt(row, colfield):
    col = colfield & 0x3FFF; cRel = bool(colfield & 0x4000); rRel = bool(colfield & 0x8000)
    return ('' if cRel else '$') + colL(col) + ('' if rRel else '$') + str(row + 1)

def decode_rgce(rgce):
    """Decode a BIFF12 rgce token stream into formula text (subset of tokens; unknown -> <ptg xx>)."""
    st = []; i = 0; n = len(rgce); notes = []
    while i < n:
        ptg = rgce[i]; i += 1
        base = ptg & 0x1F if ptg >= 0x20 else ptg
        if ptg in OPS:
            b = st.pop() if st else '?'; a = st.pop() if st else '?'
            st.append(f'{a}{OPS[ptg]}{b}')
        elif ptg == 0x13: st.append('+' + st.pop())
        elif ptg == 0x14: st.append('-' + st.pop())
        elif ptg == 0x15: st.append(st.pop() + '%')
        elif ptg == 0x12: st.append('(' + st.pop() + ')')
        elif ptg == 0x16: st.append('')
        elif ptg == 0x17:
            cch = struct.unpack_from('<H', rgce, i)[0]; s = rgce[i+2:i+2+2*cch].decode('utf-16le', 'replace'); i += 2 + 2*cch
            st.append('"' + s.replace('"', '""') + '"')
        elif ptg == 0x19:
            t = rgce[i]; i += 1
            if t == 0x04:
                cnt = struct.unpack_from('<H', rgce, i)[0]; i += 2 + 2*(cnt+1)
            elif t == 0x10:
                i += 2; st.append('SUM(' + st.pop() + ')')
            else:
                i += 2
                if t == 0x01: notes.append('PtgAttrSemi(volatile)')
        elif ptg == 0x1C: st.append(f'#ERR{rgce[i]:02x}'); i += 1
        elif ptg == 0x1D: st.append('TRUE' if rgce[i] else 'FALSE'); i += 1
        elif ptg == 0x1E: st.append(str(struct.unpack_from('<H', rgce, i)[0])); i += 2
        elif ptg == 0x1F: st.append(repr(struct.unpack_from('<d', rgce, i)[0])); i += 8
        elif base == 0x01 and ptg in (0x21, 0x41, 0x61):
            iftab = struct.unpack_from('<H', rgce, i)[0]; i += 2
            nm = FUNC.get(iftab, f'FUNC{iftab}')
            # fixed-arg functions: pop by known arity (approximate)
            arity = {8:0,9:0,10:0,221:0}.get(iftab, 1)
            args = [st.pop() for _ in range(arity)][::-1] if arity else []
            st.append(f'{nm}({",".join(args)})')
        elif base == 0x02 and ptg in (0x22, 0x42, 0x62):
            cp = rgce[i]; iftab = struct.unpack_from('<H', rgce, i+1)[0]; i += 3
            nm = FUNC.get(iftab & 0x7FFF, f'FUNC{iftab & 0x7FFF}')
            args = [st.pop() for _ in range(cp)][::-1]
            st.append(f'{nm}({",".join(args)})')
        elif base == 0x03 and ptg in (0x23, 0x43, 0x63):
            idx = struct.unpack_from('<I', rgce, i)[0]; i += 4
            nm = names[idx-1][0] if 1 <= idx <= len(names) else f'<name#{idx}>'
            st.append(f'{nm}<PtgName#{idx}>')
        elif base == 0x04 and ptg in (0x24, 0x44, 0x64):
            row, cf = struct.unpack_from('<IH', rgce, i); i += 6
            st.append(ref_txt(row, cf))
        elif base == 0x05 and ptg in (0x25, 0x45, 0x65):
            r1, r2, c1, c2 = struct.unpack_from('<IIHH', rgce, i); i += 12
            st.append(ref_txt(r1, c1) + ':' + ref_txt(r2, c2))
        elif base == 0x0A and ptg in (0x2A, 0x4A, 0x6A): i += 6; st.append('#REF!')
        elif base == 0x0B and ptg in (0x2B, 0x4B, 0x6B): i += 12; st.append('#REF!')
        elif base == 0x19 and ptg in (0x39, 0x59, 0x79):
            ixti, idx = struct.unpack_from('<HI', rgce, i); i += 6
            st.append(f'{sheet_for_ixti(ixti)}!<nameX#{idx}>')
        elif base == 0x1A and ptg in (0x3A, 0x5A, 0x7A):
            ixti, row, cf = struct.unpack_from('<HIH', rgce, i); i += 8
            st.append(qs(sheet_for_ixti(ixti)) + '!' + ref_txt(row, cf))
        elif base == 0x1B and ptg in (0x3B, 0x5B, 0x7B):
            ixti, r1, r2, c1, c2 = struct.unpack_from('<HIIHH', rgce, i); i += 14
            st.append(qs(sheet_for_ixti(ixti)) + '!' + ref_txt(r1, c1) + ':' + ref_txt(r2, c2))
        elif base == 0x1C and ptg in (0x3C, 0x5C, 0x7C): i += 8; st.append('#REF!')
        elif base == 0x1D and ptg in (0x3D, 0x5D, 0x7D): i += 14; st.append('#REF!')
        else:
            st.append(f'<ptg {ptg:02x}>'); notes.append(f'unknown ptg {ptg:02x} at {i-1}')
            break
    txt = st[-1] if len(st) == 1 else ' ; '.join(st)
    return txt, notes

def area3d_of(rgce):
    """If rgce is a single PtgArea3d/PtgRef3d, return (sheet, r1, r2, c1, c2) 1-based rows, 0-based cols."""
    if not rgce: return None
    p = rgce[0]
    if p in (0x3B, 0x5B, 0x7B) and len(rgce) == 15:
        ixti, r1, r2, c1, c2 = struct.unpack_from('<HIIHH', rgce, 1)
        return (sheet_for_ixti(ixti), r1 + 1, r2 + 1, c1 & 0x3FFF, c2 & 0x3FFF)
    if p in (0x3A, 0x5A, 0x7A) and len(rgce) == 9:
        ixti, r, c = struct.unpack_from('<HIH', rgce, 1)
        return (sheet_for_ixti(ixti), r + 1, r + 1, c & 0x3FFF, c & 0x3FFF)
    return None

# ----------------------------------------------------------------------------- load tables
TABLES = ['fund', 'position', 'company', 'position_company', 'company_nav', 'cashflow', 'valuation_history', 'company_breakdown']
SHEETBIN = {'Fund View': 'sheet3', 'fund': 'sheet5', 'position': 'sheet6', 'company': 'sheet7', 'position_company': 'sheet8',
            'company_nav': 'sheet9', 'cashflow': 'sheet10', 'valuation_history': 'sheet11', 'company_breakdown': 'sheet12'}
T = {}   # name -> dict(header=[...], rows={rownum1: [values]}, ncols, last_any, last_key ...)
fundview = {}
with open_workbook(XLSB) as wb:
    with wb.get_sheet('Fund View') as sh:
        for row in sh.rows():
            for c in row:
                if c.v is not None: fundview[f'{colL(c.c)}{c.r+1}'] = c.v
    for name in TABLES:
        with wb.get_sheet(name) as sh:
            header = None; rows = {}; ncols = 0
            for row in sh.rows():
                vals = [c.v for c in row]
                if row and row[0].r == 0:
                    header = vals; ncols = len(vals); continue
                if row and any(not is_blank(v) for v in vals):
                    rows[row[0].r + 1] = vals
                    ncols = max(ncols, len(vals))
            T[name] = dict(header=header, rows=rows, ncols=ncols)
            print(f'loaded {name}: header cols={len(header)} data rows(non-empty)={len(rows)} last non-empty row={max(rows) if rows else None}')

def hdr_index(tbl):
    return {h: i for i, h in enumerate(T[tbl]['header']) if h is not None}

def hdr_index_ci(tbl):
    return {str(h).casefold(): i for i, h in enumerate(T[tbl]['header']) if h is not None}

def col_values(tbl, c0):
    return {r: (v[c0] if c0 < len(v) else None) for r, v in T[tbl]['rows'].items()}

# ============================================================================= CHECK 1: header names
section('CHECK 1 - header names the VBA looks up by text (Application.Match on row 1: case-insensitive exact)')
NEEDED = {
    'company_nav':       ['lookup_key', 'nav_value', 'company_id', 'company_name', 'currency', 'source_tab', 'position_id', 'source_row'],
    'company':           ['source_tab', 'position_id', 'source_row', 'company_id'],
    'position_company':  ['source_tab', 'position_id', 'source_row', 'company_id'],
    'cashflow':          ['source_tab', 'position_id', 'source_row', 'company_id'],
    'valuation_history': ['source_tab', 'position_id', 'source_row', 'company_id'],
    'company_breakdown': ['source_tab', 'position_id', 'source_row', 'company_id'],
    'position':          ['source_tab', 'position_id', 'source_row', 'company_id'],
    'fund':              ['source_tab', 'position_id', 'source_row', 'company_id'],
}
c1_fail = []
for tbl, need in NEEDED.items():
    hx = hdr_index(tbl); hci = hdr_index_ci(tbl)
    out = []
    for h in need:
        if h in hx: out.append(f'{h}=col {colL(hx[h])} (exact)')
        elif h.casefold() in hci: out.append(f'{h}=col {colL(hci[h.casefold()])} (CASE-INSENSITIVE ONLY: header is {T[tbl]["header"][hci[h.casefold()]]!r})')
        else: out.append(f'{h}=ABSENT')
    print(f'{tbl}: ' + '; '.join(out))
    # header hygiene
    for i, h in enumerate(T[tbl]['header']):
        if isinstance(h, str) and (h != h.strip() or '  ' in h): print(f'  WARN header {colL(i)} has stray whitespace: {h!r}')
        if h is None: print(f'  WARN header {colL(i)} is blank')
    dups = [h for h, n in collections.Counter(T[tbl]['header']).items() if n > 1 and h is not None]
    if dups: print(f'  WARN duplicate header names: {dups}')

# ============================================================================= CHECK 2: hard-coded columns
section('CHECK 2 - hard-coded column letters vs. header row')
EXPECT = [('position', 'AB', 'updated_by', 'StampFundUpdated writes Environ("USERNAME")'),
          ('position', 'S', 'date_of_update', 'StampFundUpdated writes Date'),
          ('company_nav', 'A', 'source_tab', 'PushCompanyQuarterOverride insert'),
          ('company_nav', 'B', 'position_id', 'PushCompanyQuarterOverride insert'),
          ('company_nav', 'C', 'source_row', 'PushCompanyQuarterOverride insert'),
          ('company_nav', 'D', 'lookup_key', 'PushCompanyQuarterOverride insert (comment says Column D = lookup_key)'),
          ('company_nav', 'E', 'company_id', 'PushCompanyQuarterOverride insert'),
          ('company_nav', 'F', 'company_name', 'PushCompanyQuarterOverride insert'),
          ('company_nav', 'G', 'currency', 'PushCompanyQuarterOverride insert'),
          ('company_nav', 'H', 'Nav_date', 'PushCompanyQuarterOverride insert (CDate)'),
          ('fund', 'A', 'fund_key', 'AddNewFund writes newId'),
          ('fund', 'B', 'source_tab', 'AddNewFund writes newId'),
          ('position', 'A', 'position_id', 'AddNewFund writes newId'),
          ('position', 'B', 'source_tab', 'AddNewFund writes newId')]
for tbl, col, exp, why in EXPECT:
    h = T[tbl]['header'][colN(col)] if colN(col) < len(T[tbl]['header']) else None
    ok = (h == exp)
    print(f'{"PASS" if ok else "FAIL"} {tbl}!{col} header={h!r} expected={exp!r}  [{why}]')
# convention checks that make AddNewFund's A/B fill consistent
for tbl, ca, cb in [('fund', 'A', 'B'), ('position', 'A', 'B'), ('company_nav', 'A', 'B'), ('position_company', 'A', 'B'),
                    ('cashflow', 'A', 'B'), ('valuation_history', 'A', 'B'), ('company_breakdown', 'A', 'B')]:
    a = col_values(tbl, colN(ca)); b = col_values(tbl, colN(cb))
    diff = [r for r in a if a[r] != b[r]]
    print(f'{tbl}: rows where {ca} ({T[tbl]["header"][colN(ca)]}) != {cb} ({T[tbl]["header"][colN(cb)]}): {len(diff)} {diff[:5]}')
# the columns AddNewFund leaves blank that Fund View lookups need
print('AddNewFund fills only A,B; other columns the dashboard looks up on the new row remain blank: fund C..N =',
      T['fund']['header'][2:], '; position C..AG =', T['position']['header'][2:])

# ---- 2b FormulaBank TargetCol vs header named in the formula; KeyRange name sheet vs SourceSheet
section('CHECK 2b - FormulaBank TargetCol / KeyRange consistency with the source tables (what PushOverrideToSource writes to)')
bank = []
for line in open('ctx/formulabank.txt'):
    m = re.match(r'^(\d+) (\[.*\])$', line.rstrip('\n'))
    if m and int(m.group(1)) > 1:
        r = ast.literal_eval(m.group(2)); r = (r + [None] * 8)[:8]; bank.append((int(m.group(1)), r))
name_map = {}
for nm, itab, flags, rgce in names:
    a = area3d_of(rgce)
    if a and itab == 0xFFFFFFFF: name_map[nm] = a
mism = []; ok = 0; skipped = 0
bj = {k: v for k, v in fundview.items() if k.startswith('BJ')}
for bankrow, (ref, formula, isarr, push, src, keyexpr, keyrange, tcol) in bank:
    if not src: skipped += 1; continue
    hx = hdr_index(src)
    m = re.search(r'MATCH\("([^"]+)",' + re.escape(src) + r'!\$A\$1:\$[A-Z]+\$1,0\)', str(formula))
    field = None
    if m: field = m.group(1)
    else:
        m2 = re.search(r',\$BJ\$(\d+)\)', str(formula))
        if m2:
            idx = fundview.get(f'BJ{m2.group(1)}')
            if isinstance(idx, (int, float)): field = T[src]['header'][int(idx) - 1]
    if field is None: mism.append((bankrow, ref, src, tcol, 'could not derive field from formula')); continue
    actual = T[src]['header'][colN(tcol)] if colN(tcol) < len(T[src]['header']) else None
    if actual == field: ok += 1
    else: mism.append((bankrow, ref, src, tcol, f'formula field={field!r} but header at {tcol} is {actual!r}'))
    # key range checks
    nmr = name_map.get(keyrange)
    if not nmr: mism.append((bankrow, ref, src, keyrange, 'KeyRange name not found as workbook-level area name'))
    elif nmr[0] != src: mism.append((bankrow, ref, src, keyrange, f'KeyRange name points at sheet {nmr[0]!r}'))
    else:
        m3 = re.search(r'XMATCH\([^,]+(?:,"yyyy-mm-dd"\))?[^,]*,' + re.escape(src) + r'!\$([A-Z]+)\$2:\$([A-Z]+)\$(\d+),0,1\)', str(formula))
        if m3 and colN(m3.group(1)) != nmr[3]:
            mism.append((bankrow, ref, src, keyrange, f'formula XMATCH column {m3.group(1)} != named range column {colL(nmr[3])}'))
    # C5 vs C6 oddity
    if re.search(r'XMATCH\(C5&', str(formula)): mism.append((bankrow, ref, src, tcol, 'formula keys on C5 (not C6) while KeyExpr=' + str(keyexpr)))
print(f'FormulaBank rows checked={len(bank)} TargetCol==formula field OK={ok} skipped(no SourceSheet)={skipped} issues={len(mism)}')
for x in mism[:60]: print('  ', x)
if len(mism) > 60: print(f'  ... {len(mism)-60} more')

# ============================================================================= CHECK 3: named key ranges
section('CHECK 3 - named key ranges vs. actual data extent')
print('Defined names decoded from workbook.bin (BrtName order index is 0-based; PtgName uses 1-based):')
KEYNAMES = ['PositionKey', 'FundKey', 'CompanyKey', 'PositionCompanyKey', 'CompanyNavKey', 'CashflowKey', 'QuarterKey', 'CompanyBreakdownKey', 'ValuationHistoryKey', 'MacroFlag']
for i, (nm, itab, flags, rgce) in enumerate(names):
    if nm in KEYNAMES:
        a = area3d_of(rgce)
        print(f'  #{i} {nm}: itab={"workbook" if itab == 0xFFFFFFFF else itab} flags={flags:#x} -> {a[0]}!{colL(a[3])}{a[1]}:{colL(a[4])}{a[2]}' if a else f'  #{i} {nm}: rgce={rgce.hex()}')
EXPECT_NAMES = {'PositionKey': ('position', 'A', 2, 819), 'FundKey': ('fund', 'A', 2, 819), 'CompanyKey': ('company', 'D', 2, 4979),
                'PositionCompanyKey': ('position_company', 'E', 2, 87982), 'CompanyNavKey': ('company_nav', 'D', 2, 15134),
                'CashflowKey': ('cashflow', 'D', 2, 33498), 'QuarterKey': ('valuation_history', 'K', 2, 32608),
                'CompanyBreakdownKey': ('company_breakdown', 'D', 2, 1459)}
coverage = {}
for nm, (tbl, col, r1, r2) in EXPECT_NAMES.items():
    a = name_map.get(nm)
    print(f'\n{nm}: decoded={a[0]}!{colL(a[3])}{a[1]}:{colL(a[4])}{a[2]}  expected {tbl}!{col}{r1}:{col}{r2} -> {"MATCH" if a and a[0]==tbl and colL(a[3])==col and a[1]==r1 and a[2]==r2 else "DIFFERENT"}')
    rows = T[tbl]['rows']; c0 = colN(col)
    keyvals = {r: (v[c0] if c0 < len(v) else None) for r, v in rows.items()}
    last_any = max(rows); first_any = min(rows)
    nonblank_key_rows = [r for r, v in keyvals.items() if not is_blank(v)]
    last_key = max(nonblank_key_rows)
    blank_in_range = sorted(r for r in range(a[1], a[2] + 1) if is_blank(keyvals.get(r)))
    data_outside = sorted(r for r in rows if r < a[1] or r > a[2])
    gaps = [r for r in range(a[1], a[2] + 1) if r not in rows]
    print(f'  data rows: first={first_any} last(any col)={last_any} last(non-blank key)={last_key} count={len(rows)}; range rows={a[2]-a[1]+1}')
    print(f'  blank key cells inside range: {len(blank_in_range)} {blank_in_range[:20]}')
    print(f'  completely empty rows inside range: {len(gaps)} {gaps[:20]}')
    print(f'  data rows outside range: {len(data_outside)} {data_outside[:20]}')
    fv_ext = sorted({m for m in re.findall(re.escape(tbl) + r'!\$[A-Z]+\$2:\$[A-Z]+\$(\d+)', open('ctx/formulabank.txt').read())})
    print(f'  Fund View formulas reference {tbl} rows up to: {fv_ext}')
    coverage[nm] = (a, last_any, last_key, blank_in_range, data_outside)
# VBA hard-coded bound for company_nav
rows = T['company_nav']['rows']
beyond = sorted(r for r in rows if r > 15131)
print(f'\ncompany_nav rows beyond the VBA hard-coded search bound 2..15131 (PushCompanyQuarterOverride): {len(beyond)} ->',
      [(r, rows[r][3]) for r in beyond])

# ============================================================================= CHECK 4: uniqueness
section('CHECK 4 - key uniqueness per table (exact and case-insensitive, as MATCH/XMATCH are case-insensitive)')
KEYCOLS = [('position', 'A'), ('fund', 'A'), ('company', 'D'), ('position_company', 'E'), ('company_nav', 'D'), ('cashflow', 'D'),
           ('valuation_history', 'K'), ('valuation_history', 'D'), ('company_breakdown', 'D')]
dupinfo = {}
for tbl, col in KEYCOLS:
    kv = col_values(tbl, colN(col))
    byk = collections.defaultdict(list); byci = collections.defaultdict(list)
    for r, v in kv.items():
        if is_blank(v): continue
        s = v if isinstance(v, str) else repr(v)
        byk[s].append(r); byci[s.casefold().strip()].append(r)
    d = {k: rs for k, rs in byk.items() if len(rs) > 1}
    dci = {k: rs for k, rs in byci.items() if len(rs) > 1}
    extra_ci = len(dci) - len(d)
    dist = collections.Counter(len(rs) for rs in d.values())
    print(f'{tbl}!{col} ({T[tbl]["header"][colN(col)]}): distinct keys={len(byk)} duplicated keys={len(d)} (rows involved={sum(len(x) for x in d.values())}, multiplicity={dict(dist)}); '
          f'additional case/space-insensitive collisions={extra_ci}')
    for k, rs in list(sorted(d.items(), key=lambda x: x[1][0]))[:12]:
        print(f'    {k!r}: rows {rs}')
    if len(d) > 12: print(f'    ... {len(d)-12} more')
    if extra_ci:
        for k, rs in list(dci.items())[:5]:
            if k not in {x.casefold().strip() for x in d}: print(f'    ci-collision {k!r}: rows {rs}')
    dupinfo[(tbl, col)] = d

# ---- 4b company_nav duplicate forensics
section('CHECK 4b - company_nav duplicated lookup_keys: forensics (VBA-inserted rows?)')
cn = T['company_nav']; H = hdr_index('company_nav'); rows = cn['rows']
def g(r, h):
    v = rows[r]; i = H[h]; return v[i] if i < len(v) else None
blank_asof = sum(1 for r in rows if is_blank(g(r, 'as_of_date')))
blank_name = sum(1 for r in rows if is_blank(g(r, 'company_name')))
blank_cur = sum(1 for r in rows if is_blank(g(r, 'currency')))
blank_nav = sum(1 for r in rows if is_blank(g(r, 'nav_value')))
blank_cid = sum(1 for r in rows if is_blank(g(r, 'company_id')))
print(f'baseline blank rates over {len(rows)} rows: as_of_date={blank_asof} company_name={blank_name} currency={blank_cur} nav_value={blank_nav} company_id={blank_cid}')
d = dupinfo[('company_nav', 'D')]
funds = collections.Counter(); summary = collections.Counter(); topflag = []
keys_sorted = sorted(d.items(), key=lambda x: x[1][0])
for k, rs in keys_sorted:
    fund = k.split('|')[0]; funds[fund] += 1
    recs = []
    for r in rs:
        rec = dict(row=r, src_tab=g(r, 'source_tab'), pos=g(r, 'position_id'), src_row=g(r, 'source_row'), cid=g(r, 'company_id'),
                   name=g(r, 'company_name'), cur=g(r, 'currency'), navdate=g(r, 'Nav_date'), nav=g(r, 'nav_value'), asof=g(r, 'as_of_date'))
        recs.append(rec)
    partial = [x['row'] for x in recs if is_blank(x['name']) or is_blank(x['cur']) or is_blank(x['nav'])]
    near_top = [x['row'] for x in recs if x['row'] <= 10]
    names_differ = len({x['name'] for x in recs}) > 1
    cur_differ = len({x['cur'] for x in recs}) > 1
    cid_differ = len({x['cid'] for x in recs}) > 1
    nav_differ = len({x['nav'] for x in recs}) > 1
    identical = all(recs[0][f] == x[f] for x in recs for f in ('src_tab', 'pos', 'src_row', 'cid', 'name', 'cur', 'navdate', 'nav', 'asof'))
    adjacent = all(rs[i+1] - rs[i] == 1 for i in range(len(rs) - 1))
    tag = 'IDENTICAL' if identical else 'DIFFER'
    summary[tag] += 1
    if partial: summary['partial'] += 1
    if near_top: summary['near_top'] += 1
    if adjacent: summary['adjacent'] += 1
    if names_differ: summary['name_differs'] += 1
    if cur_differ: summary['currency_differs'] += 1
    if nav_differ: summary['nav_differs'] += 1
    print(f'{k!r}: rows {rs} adjacent={adjacent} {tag} partial={partial} near_top={near_top} name_differs={names_differ} cur_differs={cur_differ} cid_differs={cid_differ} nav_differs={nav_differ}')
    for x in recs:
        print(f'      row {x["row"]}: cid={x["cid"]!r} name={x["name"]!r} cur={x["cur"]!r} Nav_date={serial_to_date(x["navdate"]) if isinstance(x["navdate"], float) else x["navdate"]!r} nav={x["nav"]!r} as_of={x["asof"]!r}')
print('\nfunds with duplicated keys:', dict(funds))
print('summary over duplicated keys:', dict(summary))

# ---- 4c sort order / contiguity / sparse rows (VBA inserts land at row 3 with only key + target columns filled)
section('CHECK 4c - sort order of KEY columns, per-fund contiguity, and sparse rows (signatures of AppendOverrideRow / PushCompanyQuarterOverride inserts)')
KEYCOL = {'position': 'A', 'fund': 'A', 'company': 'D', 'position_company': 'E', 'company_nav': 'D', 'cashflow': 'D', 'valuation_history': 'K', 'company_breakdown': 'D'}
for tbl, col in KEYCOL.items():
    kv = col_values(tbl, colN(col)); rs = sorted(kv)
    desc = []; prev = None; prevr = None
    for r in rs:
        v = kv[r]
        if is_blank(v): continue
        s = str(v)
        if prev is not None and s.casefold() < prev.casefold(): desc.append((r, kv[r], prevr, kv[prevr]))
        prev, prevr = s, r
    # contiguity of each source_tab block (rows with a non-blank key only)
    st = col_values(tbl, 0); blocks = collections.defaultdict(list)
    for r in rs:
        if not is_blank(kv[r]) and not is_blank(st[r]): blocks[st[r]].append(r)
    noncontig = {f: (b[0], b[-1], len(b)) for f, b in blocks.items() if b[-1] - b[0] + 1 != len(b)}
    print(f'{tbl}!{col}: rows 2-4 = {[kv.get(2), kv.get(3), kv.get(4)]}; key-order descents={len(desc)} {desc[:4]}; funds whose keyed rows are NOT contiguous={len(noncontig)} {dict(list(noncontig.items())[:5])}')
    if tbl in ('fund', 'position'):
        for r in (2, 3, 4): print(f'    row {r}: {[x for x in T[tbl]["rows"][r]][:8]} ...')
print('\nSparse rows: a column filled on >=99% of rows is blank on this row (3+ such columns, or 2+ for narrow tables). VBA inserts fill only key columns + the overridden target column.')
for tbl in TABLES:
    rows = T[tbl]['rows']; hdr = T[tbl]['header']; n = len(rows); ncols = len(hdr)
    fill = [sum(1 for v in rows.values() if i < len(v) and not is_blank(v[i])) for i in range(ncols)]
    dense = [i for i in range(ncols) if fill[i] >= 0.99 * n and hdr[i] is not None]
    thresh = 3 if len(dense) >= 6 else 2
    if tbl == 'position_company':   # component rows (blank key, P/Q filled) are a second record type - evaluate keyed rows only
        cand = {r: v for r, v in rows.items() if not is_blank(v[4])}
        fill = [sum(1 for v in cand.values() if i < len(v) and not is_blank(v[i])) for i in range(ncols)]
        dense = [i for i in range(ncols) if fill[i] >= 0.99 * len(cand) and hdr[i] is not None]
    else:
        cand = rows
    sparse = []
    for r, v in cand.items():
        missing = [hdr[i] for i in dense if i >= len(v) or is_blank(v[i])]
        if len(missing) >= thresh: sparse.append((r, missing))
    print(f'{tbl}: dense columns={[hdr[i] for i in dense]} -> sparse rows={len(sparse)}')
    for r, missing in sparse[:12]:
        print(f'    row {r}: missing {missing}; content={[(colL(i), x) for i, x in enumerate(rows[r]) if not is_blank(x)]}')
# the specific rows found above
print('\ncompany_breakdown rows 771-775 (row 773 has no T:U formulas) and row 1034 (source_row is text):')
for r in (771, 772, 773, 774, 775, 1034):
    v = T['company_breakdown']['rows'].get(r); print(f'    row {r}: {[(colL(i), x) for i, x in enumerate(v) if not is_blank(x)] if v else None}')
print('position_company rows 3929-3932 (row 3930: duplicate key, blank position_id/line_type):')
for r in (3929, 3930, 3931, 3932):
    v = T['position_company']['rows'].get(r); print(f'    row {r}: {[(colL(i), x) for i, x in enumerate(v) if not is_blank(x)] if v else None}')
print('company_nav rows with blank company_name or currency:')
for r, v in T['company_nav']['rows'].items():
    if is_blank(v[5]) or is_blank(v[6]): print(f'    row {r}: {[(colL(i), x) for i, x in enumerate(v) if not is_blank(x)]}')
print('Fund View rows the position_company/company_breakdown partial rows map to: position_company source_row 120 -> Fund View row 120 (company line 1: G=company_id, R=multiple, S=methodology, T=quality); company_breakdown source_row 33 -> Fund View row 33 (AS..AX carry/proceeds block)')

# ============================================================================= CHECK 5: formula columns
section('CHECK 5 - cashflow V:AH and company_breakdown T:U: formula vs blank/constant per data row (BIFF12 cell record kinds)')
FORMULA = {8: 'fmla_str', 9: 'fmla_num', 10: 'fmla_bool', 11: 'fmla_err'}
CONST = {1: 'blank', 2: 'rk', 3: 'err', 4: 'bool', 5: 'real', 6: 'str', 7: 'isst'}
def scan_cols(path, cols, last_row):
    """Stream the sheet part; return per-row dict col->kind for the requested 0-based cols, plus max row seen."""
    res = {}; maxrow = -1
    with open(path, 'rb') as fh, mmap.mmap(fh.fileno(), 0, access=mmap.ACCESS_READ) as mm:
        row = None
        for rid, body in records(mm):
            if rid == 0x00:
                row = struct.unpack_from('<I', body, 0)[0]; maxrow = max(maxrow, row)
            elif rid in FORMULA or rid in CONST:
                if row is None: continue
                col = struct.unpack_from('<I', body, 0)[0]
                if col in cols:
                    res.setdefault(row, {})[col] = FORMULA.get(rid) or CONST.get(rid)
    return res, maxrow
for tbl, c1, c2 in [('cashflow', 'V', 'AH'), ('company_breakdown', 'T', 'U')]:
    cols = set(range(colN(c1), colN(c2) + 1)); path = f'{WS}/{SHEETBIN[tbl]}.bin'
    last = max(T[tbl]['rows'])
    res, maxrow = scan_cols(path, cols, last)
    print(f'{tbl}: scanned {path} ({os.path.getsize(path)/1e6:.1f} MB); max row index seen={maxrow} (=> row {maxrow+1}); data rows 2..{last}')
    hdr = T[tbl]['header']
    bad_rows = []; percol = collections.Counter(); kinds = collections.Counter()
    for r in range(2, last + 1):
        k = res.get(r - 1, {})
        nonf = [c for c in sorted(cols) if not str(k.get(c, 'missing')).startswith('fmla')]
        for c in sorted(cols): kinds[k.get(c, 'missing')] += 1
        if nonf:
            bad_rows.append((r, [(colL(c), k.get(c, 'missing')) for c in nonf]))
            for c in nonf: percol[colL(c)] += 1
    allblank = [r for r, cs in bad_rows if len(cs) == len(cols) and all(x[1] in ('missing', 'blank') for x in cs)]
    consts = [(r, cs) for r, cs in bad_rows if any(x[1] not in ('missing', 'blank') for x in cs)]
    print(f'  cell kinds over {c1}:{c2} x rows 2..{last}: {dict(kinds)}')
    print(f'  rows with >=1 non-formula cell in {c1}:{c2}: {len(bad_rows)}; rows with ALL of {c1}:{c2} blank/missing: {len(allblank)} {allblank[:15]}')
    print(f'  rows holding CONSTANTS in {c1}:{c2}: {len(consts)} {consts[:10]}')
    print(f'  per-column non-formula counts: {dict(percol)}')
    for r, cs in bad_rows[:10]:
        print(f'    row {r}: key={T[tbl]["rows"][r][3]!r} -> {cs}')
    # formulas present beyond the last data row?
    beyond = sorted(r + 1 for r in res if r + 1 > last)
    print(f'  rows beyond {last} that hold cells in {c1}:{c2}: {len(beyond)} {beyond[:10]}')

# ============================================================================= CHECK 6: whitespace in ids
section('CHECK 6 - position/fund identifiers with leading/trailing/double spaces, other whitespace, or "|"')
def ws_issues(s):
    iss = []
    if s != s.strip(): iss.append('lead/trail space')
    if '  ' in s: iss.append('double space')
    if any(ch in s for ch in '\t\r\n'): iss.append('tab/newline')
    if ' ' in s: iss.append('NBSP')
    if '|' in s: iss.append('pipe')
    if s != s.strip(' ') and s.strip() == s: iss.append('other ws')
    return iss
for tbl, col in [('position', 'A'), ('position', 'B'), ('fund', 'A'), ('fund', 'B'), ('company', 'A'), ('position_company', 'A'), ('position_company', 'B'),
                 ('company_nav', 'A'), ('company_nav', 'B'), ('cashflow', 'A'), ('cashflow', 'B'), ('valuation_history', 'A'), ('valuation_history', 'B'),
                 ('company_breakdown', 'A'), ('company_breakdown', 'B')]:
    kv = col_values(tbl, colN(col)); found = collections.defaultdict(list); nonstr = []
    for r, v in kv.items():
        if is_blank(v): continue
        if not isinstance(v, str): nonstr.append((r, v)); continue
        iss = ws_issues(v)
        if iss: found[(v, tuple(iss))].append(r)
    print(f'{tbl}!{col} ({T[tbl]["header"][colN(col)]}): distinct problem ids={len(found)} non-string values={len(nonstr)} {nonstr[:3]}')
    for (v, iss), rs in found.items(): print(f'    {v!r} {iss}: {len(rs)} rows {rs[:8]}')
# how the trailing-space id is keyed in child tables
pid = [v for v in col_values('position', 0).values() if isinstance(v, str) and v != v.strip()]
for bad in pid:
    print(f'\nid {bad!r}: Trim() gives {bad.strip()!r}. Rows per table keyed with the space vs without:')
    for tbl in TABLES:
        a = col_values(tbl, 0); b = col_values(tbl, 1)
        w = sum(1 for v in a.values() if v == bad); wo = sum(1 for v in a.values() if v == bad.strip())
        kb = [v for v in col_values(tbl, colN({'company':'D','position_company':'E','company_nav':'D','cashflow':'D','valuation_history':'K','company_breakdown':'D'}.get(tbl, 'A'))).values()
              if isinstance(v, str) and v.startswith(bad.strip())]
        print(f'    {tbl}: col A == {bad!r}: {w}; col A == {bad.strip()!r}: {wo}; key col values starting with {bad.strip()!r}: {len(kb)} e.g. {kb[:2]}')

# ============================================================================= CHECK 6b: referential integrity + source_row reachability
section('CHECK 6b - referential integrity of position_id across tables, and source_row reachability vs. Fund View tracked rows')
pos_ids = {v for v in col_values('position', 0).values() if not is_blank(v)}
fund_ids = {v for v in col_values('fund', 0).values() if not is_blank(v)}
print(f'position ids={len(pos_ids)} fund keys={len(fund_ids)} pos-fund={sorted(pos_ids-fund_ids)[:5]} fund-pos={sorted(fund_ids-pos_ids)[:5]}')
tracked_rows = collections.defaultdict(set)
for bankrow, (ref, formula, isarr, push, src, keyexpr, keyrange, tcol) in bank:
    if src: tracked_rows[src].add(int(re.sub(r'[A-Z]+', '', ref)))
for tbl in ['company', 'position_company', 'company_nav', 'cashflow', 'valuation_history', 'company_breakdown']:
    hx = hdr_index(tbl)
    pidc = hx.get('position_id', hx.get('source_tab')); src_tab = col_values(tbl, hx['source_tab'])
    ids = {v for v in src_tab.values() if not is_blank(v)}
    orphan = ids - pos_ids
    srows = col_values(tbl, hx['source_row'])
    nums = [v for v in srows.values() if isinstance(v, (int, float)) and not isinstance(v, bool)]
    nonnum = [(r, v) for r, v in srows.items() if not is_blank(v) and not (isinstance(v, (int, float)) and not isinstance(v, bool))]
    nonint = [(r, v) for v_r, v in [(r, v) for r, v in srows.items()] for r in [v_r] if isinstance(v, float) and not v.is_integer()]
    tr = tracked_rows[tbl]
    unreach = collections.Counter(int(v) for v in nums if int(v) not in tr)
    print(f'{tbl}: distinct source_tab={len(ids)} orphans(not in position!A)={len(orphan)} {sorted(orphan)[:5]}; source_row min={min(nums)} max={max(nums)} non-numeric={len(nonnum)} {nonnum[:3]} non-integer={len(nonint)}; '
          f'Fund View tracked rows for this sheet={min(tr)}..{max(tr)} ({len(tr)} rows); rows whose source_row is NOT a tracked Fund View row={sum(unreach.values())} by source_row={dict(sorted(unreach.items()))}')
    # composite key convention (header-driven column indices)
    def srtxt(v): return str(int(v)) if isinstance(v, float) and v.is_integer() else str(v)
    isrc, irow = hx['source_tab'], hx['source_row']
    if tbl == 'company_nav':
        ik, idt = hx['lookup_key'], hx['Nav_date']
        bad = [r for r, v in T[tbl]['rows'].items() if v[ik] != f'{v[isrc]}|{srtxt(v[irow])}|{serial_to_date(v[idt])}']
        print(f'    lookup_key == source_tab|source_row|yyyy-mm-dd(Nav_date) violations: {len(bad)} {bad[:5]}')
    elif tbl == 'valuation_history':
        iq, ipid, iqe, ik = hx['quarter_key'], hx['position_id'], hx['quarter_end'], hx['lookup_key']
        bad = [r for r, v in T[tbl]['rows'].items() if v[iq] != f'{v[ipid]}|{serial_to_date(v[iqe])}']
        bad2 = [r for r, v in T[tbl]['rows'].items() if v[ik] != f'{v[isrc]}|{srtxt(v[irow])}']
        print(f'    quarter_key == position_id|yyyy-mm-dd(quarter_end) violations: {len(bad)} {bad[:5]}; lookup_key == source_tab|source_row violations: {len(bad2)} {bad2[:5]}')
    else:
        ik = hx['lookup_key']
        blankkey = [r for r, v in T[tbl]['rows'].items() if is_blank(v[ik])]
        bad = [r for r, v in T[tbl]['rows'].items() if not is_blank(v[ik]) and v[ik] != f'{v[isrc]}|{srtxt(v[irow])}']
        print(f'    lookup_key == source_tab|source_row violations (rows with a key): {len(bad)} {[(r, T[tbl]["rows"][r][ik], T[tbl]["rows"][r][irow]) for r in bad[:5]]}; rows with blank key: {len(blankkey)}')

# ============================================================================= CHECK 7: data validation on Fund View
section('CHECK 7 - Fund View data validations (BrtDVal 0x40 and x14 BrtDVal14 0x41D in sheet3.bin)')
def dv_flags(f):
    return dict(valType=f & 0xF, errStyle=(f >> 4) & 7, fStrLookup=(f >> 7) & 1, fAllowBlank=(f >> 8) & 1, fSuppressCombo=(f >> 9) & 1,
                mdImeMode=(f >> 10) & 0xFF, fShowInputMsg=(f >> 18) & 1, fShowErrorMsg=(f >> 19) & 1, typOperator=(f >> 20) & 0xF)
VALTYPE = {0: 'any', 1: 'whole', 2: 'decimal', 3: 'list', 4: 'date', 5: 'time', 6: 'textLength', 7: 'custom'}
def sqref_txt(rfx):
    return ','.join(f'{colL(c1)}{r1+1}' if (r1 == r2 and c1 == c2) else f'{colL(c1)}{r1+1}:{colL(c2)}{r2+1}' for r1, r2, c1, c2 in rfx)
def parse_strings(body, off, n=4):
    out = []
    for _ in range(n):
        s, off = wstr(body, off); out.append(s)
    return out, off
def parse_dvformula(body, off):
    cce = struct.unpack_from('<I', body, off)[0]; rgce = body[off+4: off+4+cce]; off += 4 + cce
    cb = struct.unpack_from('<I', body, off)[0]; off += 4 + cb
    return rgce, off
d = open(f'{WS}/sheet3.bin', 'rb').read()
c6_hits = []
for rid, body in records(d):
    if rid == 0x40:
        flags = struct.unpack_from('<I', body, 0)[0]; cnt = struct.unpack_from('<I', body, 4)[0]; off = 8; rfx = []
        for i in range(cnt): rfx.append(struct.unpack_from('<IIII', body, off)); off += 16
        strs, off = parse_strings(body, off)
        f1, off = parse_dvformula(body, off); f2, off = parse_dvformula(body, off)
        fl = dv_flags(flags); t1, n1 = decode_rgce(f1); t2, n2 = decode_rgce(f2)
        sq = sqref_txt(rfx)
        print(f'BrtDVal sqref={sq} type={VALTYPE.get(fl["valType"])} strLookup={fl["fStrLookup"]} allowBlank={fl["fAllowBlank"]} showErr={fl["fShowErrorMsg"]} errStyle={fl["errStyle"]} '
              f'formula1={t1} {n1 or ""} formula2={t2!r} strings={strs}')
        if any(r1 <= 5 <= r2 and c1 <= 2 <= c2 for r1, r2, c1, c2 in rfx): c6_hits.append(('BrtDVal', sq, t1))
    elif rid == 0x41D:
        hdr = struct.unpack_from('<I', body, 0)[0]; off = 4; rfx = []; formulas = []
        if hdr & 1:  # fRef
            n = struct.unpack_from('<I', body, off)[0]; off += 4
            for i in range(n): off += 4; rfx.append(struct.unpack_from('<IIII', body, off)); off += 16
        if hdr & 2:  # fSqref
            n = struct.unpack_from('<I', body, off)[0]; off += 4
            for i in range(n):
                off += 4; cnt = struct.unpack_from('<I', body, off)[0]; off += 4
                for j in range(cnt): rfx.append(struct.unpack_from('<IIII', body, off)); off += 16
        if hdr & 4:  # fFormula
            n = struct.unpack_from('<I', body, off)[0]; off += 4
            for i in range(n):
                off += 4; cce, cb = struct.unpack_from('<II', body, off); off += 8
                formulas.append(body[off: off+cce]); off += cce + cb
        flags = struct.unpack_from('<I', body, off)[0]; off += 4
        strs, off = parse_strings(body, off)
        fl = dv_flags(flags); sq = sqref_txt(rfx)
        ftxt = [decode_rgce(f) for f in formulas]
        print(f'BrtDVal14 (x14 ext) frtflags={hdr:#x} sqref={sq} type={VALTYPE.get(fl["valType"])} strLookup={fl["fStrLookup"]} allowBlank={fl["fAllowBlank"]} showErr={fl["fShowErrorMsg"]} '
              f'errStyle={fl["errStyle"]} flags={flags:#010x} formulas={[t for t, n in ftxt]} notes={[n for t, n in ftxt]} strings={strs} trailing_bytes={len(body)-off} rgce_hex={[f.hex() for f in formulas]}')
        for f in formulas:
            a = area3d_of(f)
            if a: print(f'   -> single 3D area: sheet={a[0]!r} rows {a[1]}..{a[2]} cols {colL(a[3])}..{colL(a[4])}')
        if any(r1 <= 5 <= r2 and c1 <= 2 <= c2 for r1, r2, c1, c2 in rfx): c6_hits.append(('BrtDVal14', sq, [t for t, n in ftxt]))
print('\nValidations covering C6:', c6_hits)
pk = name_map.get('PositionKey'); fk = name_map.get('FundKey')
print(f'PositionKey = {pk[0]}!{colL(pk[3])}{pk[1]}:{colL(pk[4])}{pk[2]};  FundKey = {fk[0]}!{colL(fk[3])}{fk[1]}:{colL(fk[4])}{fk[2]}')
print('PtgName tokens (0x23/0x43/0x63) present in any Fund View DV formula:', any('PtgName' in str(h) for h in c6_hits))
# does the dropdown list equal the position key list?
fv = [v for r, v in sorted(col_values('fund', 0).items()) if not is_blank(v)]
pv = [v for r, v in sorted(col_values('position', 0).items()) if not is_blank(v)]
print(f'fund!A2:A819 values == position!A2:A819 values (same order)? {fv == pv}; as sets? {set(fv) == set(pv)}; C6 current value {fundview.get("C6")!r} in fund list? {fundview.get("C6") in fv}')
print('\nDONE')
