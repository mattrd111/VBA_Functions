#!/usr/bin/env python3
"""UAT edge-case data checks for HP_Flat_File_Test_v3.3.xlsb (VBA write-back audit).

Reads the real source tables with pyxlsb (values) and the raw BIFF12 worksheet parts with
cellscan.scan (cell KINDS: formula / constant / error) and looks for inputs that break the
assumptions made by Sheet730.cls / Main.bas (Application.Match limits, wildcard chars,
text-vs-number keys, stray whitespace, text dates, non-quarter-end dates, source_tab !=
position_id, error values, blank company_id, blank J dates, ...).

Run:  python3 check.py > output.txt
"""
import sys, os, re, collections, unicodedata, datetime, time

SCR = '/tmp/claude-0/-home-user-VBA-Functions/4411ca91-ef32-5ffa-9c78-fab36847f3b2/scratchpad'
sys.path.insert(0, SCR)
from pyxlsb import open_workbook
from cellscan import scan, colL

XLSB = os.path.join(SCR, 'HP_Flat_File_Test_v3.3.xlsb')
PARTS = os.path.join(SCR, 'xlsb', 'xl', 'worksheets')

# sheet name -> BIFF12 part (from CONTEXT.md / parse_biff)
PART = {'Fund View': 'sheet3.bin', 'fund': 'sheet5.bin', 'position': 'sheet6.bin', 'company': 'sheet7.bin',
        'position_company': 'sheet8.bin', 'company_nav': 'sheet9.bin', 'cashflow': 'sheet10.bin',
        'valuation_history': 'sheet11.bin', 'company_breakdown': 'sheet12.bin'}

# lookup-key columns that Application.Match / XMATCH / CountIf run against
KEY_COLS = {
    'fund':              {'A': 'fund_key (FundKey)', 'B': 'source_tab'},
    'position':          {'A': 'position_id (PositionKey)', 'B': 'source_tab'},
    'company':           {'D': 'lookup_key (CompanyKey)', 'E': 'company_id'},
    'position_company':  {'E': 'lookup_key (PositionCompanyKey)', 'F': 'company_id'},
    'company_nav':       {'D': 'lookup_key (CompanyNavKey)', 'E': 'company_id (VBA Match target)'},
    'cashflow':          {'D': 'lookup_key (CashflowKey)'},
    'valuation_history': {'D': 'lookup_key (ValuationHistoryKey)', 'K': 'quarter_key (QuarterKey)'},
    'company_breakdown': {'D': 'lookup_key (CompanyBreakdownKey)'},
}
# TargetCol per SourceSheet, derived from FormulaBank (H column)
TARGET_COLS = {
    'position': 'C E F G H I J K L M N O P Q R S U V W X Y Z AA AB AC AD AE'.split(),
    'fund': 'D E F G H I J K L N'.split(),
    'valuation_history': 'G H I'.split(),
    'cashflow': 'I L M O P R'.split(),
    'company_breakdown': 'E F J K L M N O P Q R S'.split(),
    'position_company': 'F G H I K L M'.split(),
    'company': 'F G H I J M'.split(),
    'company_nav': ['I'],
}
DATE_COLS = {
    'company_nav': {'H': 'Nav_date', 'J': 'as_of_date'},
    'valuation_history': {'G': 'val_date', 'H': 'valuation_date', 'J': 'quarter_end', 'F': 'as_of_date'},
    'cashflow': {'I': 'cashflow_date', 'J': 'quarter_end', 'F': 'as_of_date'},
}
QE_REQUIRED = {('company_nav', 'H'), ('valuation_history', 'H'), ('valuation_history', 'J'), ('cashflow', 'J')}

WILD = set('*?~')
NBSP = ' '
UNI_WS = {' ': 'NBSP U+00A0', ' ': 'U+1680', ' ': 'U+2000', ' ': 'U+2001', ' ': 'U+2002', ' ': 'U+2003',
          ' ': 'U+2004', ' ': 'U+2005', ' ': 'U+2006', ' ': 'U+2007', ' ': 'U+2008', ' ': 'U+2009',
          ' ': 'U+200A', '​': 'ZWSP U+200B', ' ': 'U+2028', ' ': 'U+2029', ' ': 'NNBSP U+202F',
          ' ': 'U+205F', '　': 'U+3000', '﻿': 'BOM U+FEFF', '\t': 'TAB', '\n': 'LF', '\r': 'CR', '\x0b': 'VT', '\x0c': 'FF'}
NUMLIKE = re.compile(r'^\s*[-+]?(\d[\d,]*(\.\d*)?|\.\d+)([eE][-+]?\d+)?\s*%?\s*$')
ERRCODES = {'0x0': '#NULL!', '0x7': '#DIV/0!', '0xf': '#VALUE!', '0x17': '#REF!', '0x1d': '#NAME?', '0x24': '#NUM!', '0x2a': '#N/A', '0x2b': '#GETTING_DATA'}

def col_idx(letter):
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - 64)
    return n - 1

def xl_date(serial):
    try:
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(serial))
    except Exception:
        return None

def is_quarter_end(serial):
    d = xl_date(serial)
    if d is None:
        return False
    return (d.month, d.day) in ((3, 31), (6, 30), (9, 30), (12, 31))

def fmt_date(serial):
    d = xl_date(serial)
    return d.isoformat() if d else '?'

def load(wb, name):
    """Return (header, rows) where rows = list of (excel_row, [values...]) for data rows."""
    header, rows = None, []
    with wb.get_sheet(name) as sh:
        for r in sh.rows():
            vals = [c.v for c in r]
            xlrow = r[0].r + 1
            if xlrow == 1:
                header = vals
            else:
                rows.append((xlrow, vals))
    return header, rows

def get(vals, letter):
    i = col_idx(letter)
    return vals[i] if i < len(vals) else None

def show(v):
    return repr(v) if not isinstance(v, float) else (repr(int(v)) if v == int(v) else repr(v))

def ex(lst, n=8):
    return lst[:n]

section_no = 0
def H(title):
    global section_no
    section_no += 1
    print('\n' + '=' * 100)
    print(f'[{section_no}] {title}')
    print('=' * 100)

t0 = time.time()
print('check.py started', datetime.datetime.now().isoformat(timespec='seconds'))
print('workbook:', XLSB)

# ---------------------------------------------------------------------------------------------
# Pass 1: cell KINDS from the raw BIFF12 parts (authoritative for error cells / formula cells)
# ---------------------------------------------------------------------------------------------
H('CELL KINDS per sheet/column (BIFF12 record scan) - error and formula cells')
kinds = {}   # sheet -> Counter((colLetter, kind))
errcells = {}  # sheet -> list of (row, col, kind)
fv_kind = {}   # Fund View: ref -> kind
for name, part in PART.items():
    path = os.path.join(PARTS, part)
    cnt = collections.Counter()
    errs = []
    t = time.time()
    for row0, col0, kind, rid, body in scan(path):
        L = colL(col0)
        cnt[(L, kind)] += 1
        if kind in ('err', 'fmla_err'):
            errs.append((row0 + 1, L, kind))
        if name == 'Fund View':
            fv_kind[f'{L}{row0 + 1}'] = kind
    kinds[name] = cnt
    errcells[name] = errs
    bycol = collections.defaultdict(collections.Counter)
    for (L, k), c in cnt.items():
        bycol[L][k] += c
    fcols = sorted([L for L in bycol if any(k.startswith('fmla') for k in bycol[L])], key=lambda s: (len(s), s))
    print(f'{name}: scanned in {time.time() - t:.1f}s; error cells (err/fmla_err) = {len(errs)}; formula-bearing columns = {fcols}')
    if errs:
        percol = collections.Counter((L, k) for _, L, k in errs)
        print('   error cells by column/kind:', dict(percol))
        print('   examples (row, col, kind):', ex(errs, 15))

# ---------------------------------------------------------------------------------------------
# Pass 2: values via pyxlsb
# ---------------------------------------------------------------------------------------------
wb = open_workbook(XLSB)
tables = {}
for name in ['fund', 'position', 'company', 'position_company', 'company_nav', 'cashflow', 'valuation_history', 'company_breakdown']:
    t = time.time()
    hdr, rows = load(wb, name)
    tables[name] = (hdr, rows)
    print(f'loaded {name}: {len(rows)} data rows (last excel row {rows[-1][0] if rows else None}) in {time.time() - t:.1f}s')
fv_hdr, fv_rows = load(wb, 'Fund View')
fv = {}
for xlrow, vals in fv_rows:
    fv[xlrow] = vals
wb.close()

position_ids = [get(v, 'A') for _, v in tables['position'][1]]
position_id_set = set(x for x in position_ids if x is not None)

# ---------------------------------------------------------------------------------------------
H('KEY LENGTH: max lookup-key length per table/column and any key > 255 chars (Application.Match limit)')
maxlen_report = {}
for name, cols in KEY_COLS.items():
    hdr, rows = tables[name]
    for L, label in cols.items():
        lens = [(len(str(get(v, L))), xlrow, get(v, L)) for xlrow, v in rows if get(v, L) not in (None, '')]
        if not lens:
            print(f'{name}!{L} ({label}): no non-blank values'); continue
        mx = max(lens)
        over = [(xlrow, ln) for ln, xlrow, val in lens if ln > 255]
        maxlen_report[(name, L)] = mx[0]
        print(f'{name}!{L} ({label}): non-blank={len(lens)} max_len={mx[0]} at row {mx[1]} -> {show(mx[2])}; keys >255 chars = {len(over)} {ex(over)}')
# longest text anywhere in each table (would a write of chg.Value or a Match on it be affected?)
print('\nLongest text value anywhere per table (any column):')
for name, (hdr, rows) in tables.items():
    best = (0, None, None)
    over255 = collections.Counter()
    for xlrow, v in rows:
        for i, x in enumerate(v):
            if isinstance(x, str):
                if len(x) > best[0]:
                    best = (len(x), xlrow, colL(i))
                if len(x) > 255:
                    over255[colL(i)] += 1
    print(f'  {name}: max text len {best[0]} at {best[2]}{best[1]}; cells >255 chars by column: {dict(over255) or "none"}')

# ---------------------------------------------------------------------------------------------
H('WILDCARDS: keys containing * ? ~ (Application.Match / CountIf wildcard interpretation)')
for name, cols in KEY_COLS.items():
    hdr, rows = tables[name]
    for L, label in cols.items():
        hits = [(xlrow, get(v, L)) for xlrow, v in rows if isinstance(get(v, L), str) and (set(get(v, L)) & WILD)]
        print(f'{name}!{L} ({label}): {len(hits)} keys with * ? ~  {ex(hits)}')
# Fund View G120:G159 current values (companyId used as Match lookup_value in PushCompanyQuarterOverride)
g_hits = [(r, fv[r][col_idx('G')]) for r in range(120, 160) if isinstance(fv[r][col_idx('G')], str) and set(fv[r][col_idx('G')]) & WILD]
print(f'Fund View G120:G159 (current fund): {len(g_hits)} company ids with wildcards {g_hits}')

# ---------------------------------------------------------------------------------------------
H('NUMERIC / NUMERIC-LOOKING KEYS (text-vs-number mismatch in Application.Match)')
for name, cols in KEY_COLS.items():
    hdr, rows = tables[name]
    for L, label in cols.items():
        types = collections.Counter(type(get(v, L)).__name__ for xlrow, v in rows)
        numeric = [(xlrow, get(v, L)) for xlrow, v in rows if isinstance(get(v, L), (int, float)) and not isinstance(get(v, L), bool)]
        numlike = [(xlrow, get(v, L)) for xlrow, v in rows if isinstance(get(v, L), str) and NUMLIKE.match(get(v, L))]
        print(f'{name}!{L} ({label}): value types={dict(types)}; stored-as-number={len(numeric)} {ex(numeric)}; text-that-looks-numeric={len(numlike)} {ex(numlike)}')
# source_row columns are numeric by design; the VBA writes chg.Row (Long) - confirm they are whole numbers
for name in ['company', 'position_company', 'company_nav', 'cashflow', 'valuation_history', 'company_breakdown']:
    hdr, rows = tables[name]
    L = 'B' if name == 'company' else 'C'
    bad = [(xlrow, get(v, L)) for xlrow, v in rows if not (isinstance(get(v, L), float) and get(v, L) == int(get(v, L)))]
    print(f'{name}!{L} source_row: non-whole-number/non-numeric = {len(bad)} {ex(bad)}')

# ---------------------------------------------------------------------------------------------
H('WHITESPACE / NON-ASCII in keys: leading, trailing, double spaces, NBSP U+00A0, other Unicode whitespace')
ws_summary = {}
for name, cols in KEY_COLS.items():
    hdr, rows = tables[name]
    for L, label in cols.items():
        lead, trail, dbl, nb, uni, nonascii = [], [], [], [], [], []
        for xlrow, v in rows:
            s = get(v, L)
            if not isinstance(s, str) or s == '':
                continue
            if s[0] == ' ': lead.append((xlrow, s))
            if s[-1] == ' ': trail.append((xlrow, s))
            if '  ' in s: dbl.append((xlrow, s))
            if NBSP in s: nb.append((xlrow, s))
            odd = [UNI_WS[ch] for ch in s if ch in UNI_WS and ch != NBSP]
            if odd: uni.append((xlrow, s, odd))
            if any(ord(ch) > 127 for ch in s): nonascii.append((xlrow, s))
        ws_summary[(name, L)] = (len(lead), len(trail), len(dbl), len(nb), len(uni), len(nonascii))
        print(f'{name}!{L} ({label}): leading-space={len(lead)} {ex(lead,4)}; trailing-space={len(trail)} {ex(trail,6)}; '
              f'double-space={len(dbl)} {ex(dbl,4)}; NBSP={len(nb)} {ex(nb,4)}; other-unicode-ws/ctrl={len(uni)} {ex(uni,4)}; '
              f'non-ASCII chars={len(nonascii)} {ex(nonascii,6)}')
# the trailing-space id: where does it appear, and does a trimmed twin exist?
print("\nTrailing-space identifier follow-up ('IK VII No.4 '):")
for name, (hdr, rows) in tables.items():
    A = [(xlrow, get(v, 'A')) for xlrow, v in rows if isinstance(get(v, 'A'), str) and get(v, 'A').strip() == 'IK VII No.4']
    raw = [x for x in A if x[1] == 'IK VII No.4 ']
    trimmed = [x for x in A if x[1] == 'IK VII No.4']
    keycol = list(KEY_COLS[name].keys())[0]
    keys = [(xlrow, get(v, keycol)) for xlrow, v in rows if isinstance(get(v, keycol), str) and get(v, keycol).startswith('IK VII No.4')]
    print(f'  {name}: rows with A=="IK VII No.4 " (trailing space)={len(raw)}, A=="IK VII No.4" (trimmed)={len(trimmed)}; '
          f'{keycol}-keys starting "IK VII No.4"={len(keys)} e.g. {ex(keys,3)}')

# ---------------------------------------------------------------------------------------------
H('DATE COLUMNS: storage type (serial vs text) and quarter-end check (company_nav H, valuation_history H/J, cashflow J)')
for name, cols in DATE_COLS.items():
    hdr, rows = tables[name]
    for L, label in cols.items():
        types = collections.Counter()
        text_dates, non_qe, fractional, blanks = [], [], [], 0
        for xlrow, v in rows:
            x = get(v, L)
            if x is None or x == '':
                blanks += 1; types['blank'] += 1; continue
            if isinstance(x, bool):
                types['bool'] += 1; continue
            if isinstance(x, (int, float)):
                types['number'] += 1
                if x != int(x): fractional.append((xlrow, x))
                if not is_quarter_end(x): non_qe.append((xlrow, x, fmt_date(x)))
            elif isinstance(x, str):
                if x in ERRCODES: types['error'] += 1
                else:
                    types['text'] += 1; text_dates.append((xlrow, x))
            else:
                types[type(x).__name__] += 1
        need_qe = (name, L) in QE_REQUIRED
        print(f'{name}!{L} ({label}): types={dict(types)}; text-stored dates={len(text_dates)} {ex(text_dates)}; '
              f'serials with time fraction={len(fractional)} {ex(fractional,4)}; '
              f'NOT quarter-end={len(non_qe)}{" (quarter-end REQUIRED by key/TEXT logic)" if need_qe else " (informational)"} {ex(non_qe,8)}')
        if need_qe and non_qe:
            byfund = collections.Counter(get(v, 'A') for xlrow, v in rows if isinstance(get(v, L), (int, float)) and not is_quarter_end(get(v, L)))
            print(f'   non-quarter-end rows by fund (top 10): {byfund.most_common(10)}')
            mins = min(x for _, x, _ in non_qe); maxs = max(x for _, x, _ in non_qe)
            print(f'   non-quarter-end date range: {fmt_date(mins)} .. {fmt_date(maxs)}')

# key/date consistency: does the key's date segment equal the date column?
print('\nKey <-> date column consistency:')
hdr, rows = tables['company_nav']
mism = [(xlrow, get(v, 'D'), get(v, 'H')) for xlrow, v in rows
        if not (isinstance(get(v, 'D'), str) and isinstance(get(v, 'H'), (int, float)) and get(v, 'D') == f"{get(v,'A')}|{int(get(v,'C')) if isinstance(get(v,'C'),float) else get(v,'C')}|{fmt_date(get(v,'H'))}")]
print(f'  company_nav: lookup_key != source_tab|source_row|yyyy-mm-dd(Nav_date): {len(mism)} {ex(mism)}')
hdr, rows = tables['valuation_history']
mism = [(xlrow, get(v, 'K'), get(v, 'J')) for xlrow, v in rows
        if not (isinstance(get(v, 'K'), str) and isinstance(get(v, 'J'), (int, float)) and get(v, 'K') == f"{get(v,'B')}|{fmt_date(get(v,'J'))}")]
print(f'  valuation_history: quarter_key != position_id|yyyy-mm-dd(quarter_end): {len(mism)} {ex(mism)}')
mismD = [(xlrow, get(v, 'D')) for xlrow, v in rows if get(v, 'D') != f"{get(v,'A')}|{int(get(v,'C')) if isinstance(get(v,'C'),float) else get(v,'C')}"]
print(f'  valuation_history: lookup_key != source_tab|source_row: {len(mismD)} {ex(mismD)}')
# valuation_history: H valuation_date vs J quarter_end relationship
hdr, rows = tables['valuation_history']
h_ne_j = [(xlrow, fmt_date(get(v, 'H')), fmt_date(get(v, 'J'))) for xlrow, v in rows if isinstance(get(v, 'H'), (int, float)) and isinstance(get(v, 'J'), (int, float)) and get(v, 'H') != get(v, 'J')]
print(f'  valuation_history: rows where valuation_date(H) != quarter_end(J): {len(h_ne_j)} of {len(rows)} {ex(h_ne_j,5)}')
for name in ['company', 'position_company', 'cashflow', 'company_breakdown']:
    hdr, rows = tables[name]
    kc = 'D' if name != 'position_company' else 'E'
    rc = 'B' if name == 'company' else 'C'
    mism = [(xlrow, get(v, kc)) for xlrow, v in rows if get(v, kc) != f"{get(v,'A')}|{int(get(v,rc)) if isinstance(get(v,rc),float) else get(v,rc)}"]
    print(f'  {name}: lookup_key({kc}) != source_tab|source_row: {len(mism)} {ex(mism)}')

# ---------------------------------------------------------------------------------------------
H('source_tab != position_id (tables holding both) and orphan source_tab values')
for name in ['position', 'fund', 'position_company', 'company_nav', 'cashflow', 'valuation_history', 'company_breakdown']:
    hdr, rows = tables[name]
    a, b = hdr[0], hdr[1]
    diff = [(xlrow, get(v, 'A'), get(v, 'B')) for xlrow, v in rows if get(v, 'A') != get(v, 'B')]
    print(f'{name}: {a}(A) != {b}(B): {len(diff)} {ex(diff)}')
for name in ['company', 'position_company', 'company_nav', 'cashflow', 'valuation_history', 'company_breakdown']:
    hdr, rows = tables[name]
    orphan = collections.Counter(get(v, 'A') for xlrow, v in rows if get(v, 'A') not in position_id_set)
    print(f'{name}: source_tab values NOT present in position!A: {len(orphan)} distinct, {sum(orphan.values())} rows {orphan.most_common(6)}')

# ---------------------------------------------------------------------------------------------
H('ERROR VALUES in key / target columns (cellscan kinds err / fmla_err)')
for name in KEY_COLS:
    errs = errcells[name]
    keyc = set(KEY_COLS[name].keys())
    tgtc = set(TARGET_COLS[name])
    inkey = [e for e in errs if e[1] in keyc]
    intgt = [e for e in errs if e[1] in tgtc]
    other = [e for e in errs if e[1] not in keyc | tgtc]
    print(f'{name}: errors in key cols={len(inkey)} {ex(inkey)}; in target cols={len(intgt)} {ex(intgt)}; elsewhere={len(other)} '
          f'{dict(collections.Counter(e[1] for e in other))} {ex(other,6)}')
# pyxlsb view: hex error strings in key/target/date columns (cross-check)
for name in KEY_COLS:
    hdr, rows = tables[name]
    cols = set(KEY_COLS[name]) | set(TARGET_COLS[name]) | set(DATE_COLS.get(name, {}))
    hits = [(xlrow, L, ERRCODES.get(get(v, L))) for xlrow, v in rows for L in cols if isinstance(get(v, L), str) and get(v, L) in ERRCODES]
    print(f'{name}: pyxlsb error-coded values in key/target/date cols = {len(hits)} {ex(hits)}')

# ---------------------------------------------------------------------------------------------
H('FUND VIEW J23:J64 (valuation dates feeding C6&"|"&TEXT(J#,"yyyy-mm-dd") keys)')
jrows = []
for r in range(23, 65):
    v = fv[r][col_idx('J')]
    k = fv_kind.get(f'J{r}', 'missing')
    jrows.append((r, k, v))
print('row, kind, value:')
for r, k, v in jrows:
    extra = ''
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        extra = fmt_date(v) + ('' if is_quarter_end(v) else '  <-- NOT quarter-end')
    print(f'  J{r}: {k:10s} {show(v):>14s} {extra}')
blank = [r for r, k, v in jrows if v is None or v == '']
text = [(r, v) for r, k, v in jrows if isinstance(v, str) and v != '']
errk = [r for r, k, v in jrows if k in ('err', 'fmla_err')]
nonqe = [(r, fmt_date(v)) for r, k, v in jrows if isinstance(v, (int, float)) and not isinstance(v, bool) and not is_quarter_end(v)]
consts = [r for r, k, v in jrows if not k.startswith('fmla')]
dupes = collections.Counter(v for r, k, v in jrows if isinstance(v, (int, float)))
dupes = {fmt_date(k): c for k, c in dupes.items() if c > 1}
print(f'summary: blank={len(blank)} {blank}; text={len(text)} {text}; error={len(errk)} {errk}; not-quarter-end={len(nonqe)} {nonqe}; '
      f'non-formula(constant/blank) cells={len(consts)} {consts}; duplicate dates={dupes}')
print('kinds:', dict(collections.Counter(k for r, k, v in jrows)))
print('J65 (row after range):', show(fv[65][col_idx('J')]), fv_kind.get('J65'))
print('Blank J -> Evaluate(C6&"|"&TEXT(J#,"yyyy-mm-dd")) = "<fund>|1900-01-00" (TEXT of blank = 0 -> 1900-01-00); text J -> TEXT returns the text unchanged.')
# Does any valuation_history quarter_key end with 1900-01-00 already (evidence of earlier bad appends)?
hdr, rows = tables['valuation_history']
bad = [(xlrow, get(v, 'K')) for xlrow, v in rows if isinstance(get(v, 'K'), str) and (get(v, 'K').endswith('|1900-01-00') or get(v, 'K').endswith('|'))]
print(f'valuation_history K quarter_keys ending "|1900-01-00" or "|": {len(bad)} {ex(bad)}')
# tracked cells in rows 23..64 with row-keyed sources: are there cells beyond the data rows of the current fund (blank formulas)?
print('Fund View row 114 quarter headers AT:AZ kinds/values:', [(f'{L}114', fv_kind.get(f'{L}114'), show(fv[114][col_idx(L)]), fmt_date(fv[114][col_idx(L)]) if isinstance(fv[114][col_idx(L)], float) else '') for L in ['AT','AU','AV','AW','AX','AY','AZ']])

# ---------------------------------------------------------------------------------------------
H('FUND VIEW rows 120..159: G (company_id) blank while AT:AZ hold formulas -> override would create company_nav rows with blank company_id')
qcols = ['AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ']
blankG_with_formulas, nonblankG, gkinds = [], [], collections.Counter()
for r in range(120, 160):
    g = fv[r][col_idx('G')]
    gk = fv_kind.get(f'G{r}', 'missing')
    gkinds[gk] += 1
    fcount = sum(1 for L in qcols if fv_kind.get(f'{L}{r}', '').startswith('fmla'))
    if g is None or (isinstance(g, str) and g.strip() == ''):
        if fcount:
            blankG_with_formulas.append((r, gk, fcount))
    else:
        nonblankG.append((r, g, type(g).__name__))
print(f'G120:G159 kinds: {dict(gkinds)}')
print(f'rows with non-blank G (current fund {fv[6][col_idx("C")]!r}): {len(nonblankG)} {nonblankG}')
print(f'rows with BLANK G but AT:AZ formulas present: {len(blankG_with_formulas)} -> rows {[r for r,_,_ in blankG_with_formulas]}')
print('  (G is a formula returning "" so Len(Trim(CStr(G)))=0; PushCompanyQuarterOverride writes companyId="" into company_nav!E and the copy-from-existing-row Match runs with "")')
# how many company rows does each fund really have (position_company source_row 120..159 with company_id)?
hdr, rows = tables['position_company']
per_fund = collections.defaultdict(set)
for xlrow, v in rows:
    sr = get(v, 'C'); cid = get(v, 'F')
    if isinstance(sr, float) and 120 <= sr <= 159 and cid not in (None, ''):
        per_fund[get(v, 'A')].add(int(sr))
counts = sorted(len(s) for s in per_fund.values())
import statistics
print(f'position_company: funds with >=1 company row (source_row 120..159, company_id non-blank): {len(per_fund)} of {len(position_id_set)}; '
      f'company rows per fund min/median/max = {counts[0] if counts else 0}/{statistics.median(counts) if counts else 0}/{counts[-1] if counts else 0}; '
      f'funds with all 40 rows filled: {sum(1 for c in counts if c >= 40)}; funds with 0 company rows: {len(position_id_set) - len(per_fund)}')
print(f'  => for a typical fund, {40 - (statistics.median(counts) if counts else 0):.0f} of the 40 grid rows have blank G with live AT:AZ formulas.')

# ---------------------------------------------------------------------------------------------
H('company_nav rows with blank company_id: are company_name / currency also blank? (defeats VBA copy-from-existing-row)')
hdr, rows = tables['company_nav']
blankE = [(xlrow, v) for xlrow, v in rows if get(v, 'E') in (None, '')]
bothblank = [(xlrow, v) for xlrow, v in blankE if get(v, 'F') in (None, '') and get(v, 'G') in (None, '')]
nameblank = [x for x in blankE if get(x[1], 'F') in (None, '')]
currblank = [x for x in blankE if get(x[1], 'G') in (None, '')]
navpresent = [x for x in blankE if get(x[1], 'I') not in (None, '')]
asof = collections.Counter(type(get(v, 'J')).__name__ for _, v in blankE)
srdist = collections.Counter(int(get(v, 'C')) for _, v in blankE if isinstance(get(v, 'C'), float))
funds = collections.Counter(get(v, 'A') for _, v in blankE)
print(f'blank company_id rows: {len(blankE)}; of these company_name blank: {len(nameblank)}, currency blank: {len(currblank)}, BOTH blank: {len(bothblank)}, nav_value present: {len(navpresent)}')
print(f'  as_of_date types on those rows: {dict(asof)}; source_row distribution: {sorted(srdist.items())}')
print(f'  distinct funds: {len(funds)}; top: {funds.most_common(8)}')
print(f'  examples: {[(xlrow, get(v,"A"), int(get(v,"C")), get(v,"D"), get(v,"F"), get(v,"G"), get(v,"I")) for xlrow, v in blankE[:8]]}')
nonblank_name_or_curr = [(xlrow, get(v,'D'), get(v,'F'), get(v,'G')) for xlrow, v in blankE if get(v,'F') not in (None,'') or get(v,'G') not in (None,'')]
print(f'  blank-company_id rows that DO carry a name or currency: {len(nonblank_name_or_curr)} {ex(nonblank_name_or_curr,6)}')
# can the blank company_id be recovered from position_company / company for the same (fund,row)?
pc_map = {}
for xlrow, v in tables['position_company'][1]:
    if get(v, 'F') not in (None, ''):
        pc_map.setdefault((get(v, 'A'), int(get(v, 'C')) if isinstance(get(v, 'C'), float) else get(v, 'C')), get(v, 'F'))
co_map = {}
for xlrow, v in tables['company'][1]:
    if get(v, 'E') not in (None, ''):
        co_map.setdefault((get(v, 'A'), int(get(v, 'B')) if isinstance(get(v, 'B'), float) else get(v, 'B')), get(v, 'E'))
recover_pc = sum(1 for _, v in blankE if (get(v, 'A'), int(get(v, 'C'))) in pc_map)
recover_co = sum(1 for _, v in blankE if (get(v, 'A'), int(get(v, 'B')) if False else int(get(v, 'C'))) in co_map)
print(f'  blank-company_id rows whose (fund,row) has a company_id in position_company: {recover_pc}; in company: {recover_co}')
# company_ids in company_nav whose every row lacks name or currency (copy would yield Empty even when Match succeeds)
byid = collections.defaultdict(list)
for xlrow, v in rows:
    if get(v, 'E') not in (None, ''):
        byid[get(v, 'E')].append((get(v, 'F'), get(v, 'G'), xlrow))
noname = [cid for cid, lst in byid.items() if all(n in (None, '') for n, c, r in lst)]
nocurr = [cid for cid, lst in byid.items() if all(c in (None, '') for n, c, r in lst)]
firstrow_missing = [cid for cid, lst in byid.items() if (lst[0][0] in (None, '') or lst[0][1] in (None, ''))]
print(f'  distinct non-blank company_ids in company_nav: {len(byid)}; ids with NO name on any row: {len(noname)} {ex(noname,5)}; '
      f'ids with NO currency on any row: {len(nocurr)} {ex(nocurr,5)}; ids whose FIRST row (the one Match returns) lacks name or currency: {len(firstrow_missing)} {ex(firstrow_missing,5)}')
# company_ids known to position_company/company but absent from company_nav (first override for such a company can never copy name/currency)
pc_ids = set(pc_map.values()); co_ids = set(co_map.values())
print(f'  company_ids in position_company: {len(pc_ids)}, in company: {len(co_ids)}, in company_nav: {len(byid)}; '
      f'position_company ids with NO company_nav row at all: {len(pc_ids - set(byid))}')
# company_nav rows beyond the hard-coded search bound
beyond = [(xlrow, get(v, 'D')) for xlrow, v in rows if xlrow > 15131]
print(f'  company_nav data rows beyond VBA hard-coded bound row 15131: {len(beyond)} {beyond}')
# company_id types in company_nav vs Fund View G (CStr(G) vs cell type)
etypes = collections.Counter(type(get(v, 'E')).__name__ for _, v in rows)
print(f'  company_nav!E value types: {dict(etypes)}')

# ---------------------------------------------------------------------------------------------
H('DUPLICATE KEYS per key column (exact and case-insensitive; Application.Match/XMATCH are case-insensitive and return the FIRST hit)')
for name, cols in KEY_COLS.items():
    hdr, rows = tables[name]
    for L, label in cols.items():
        if 'company_id' in label:
            continue  # ids legitimately repeat
        vals = [get(v, L) for _, v in rows if get(v, L) not in (None, '')]
        exact = collections.Counter(vals)
        ci = collections.Counter(str(x).lower() for x in vals)
        dupe_exact = {k: c for k, c in exact.items() if c > 1}
        dupe_ci_only = {k: c for k, c in ci.items() if c > 1 and k not in {str(x).lower() for x in dupe_exact}}
        print(f'{name}!{L} ({label}): distinct={len(exact)} exact-duplicated keys={len(dupe_exact)} (rows involved {sum(dupe_exact.values())}) {ex(list(dupe_exact.items()),5)}; '
              f'extra case-insensitive-only duplicates={len(dupe_ci_only)} {ex(list(dupe_ci_only.items()),5)}')

# ---------------------------------------------------------------------------------------------
H('cashflow formula columns V:AH (Issue Log #13) - kinds and error counts')
cnt = kinds['cashflow']
for L in ['V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH']:
    ks = {k: c for (cl, k), c in cnt.items() if cl == L}
    print(f'  cashflow!{L}: {ks}')


# ---------------------------------------------------------------------------------------------
H('FOLLOW-UP DRILL-DOWNS on the anomalies found above')
def rowmap(name):
    return {xlrow: v for xlrow, v in tables[name][1]}
def dump(name, rs):
    rm = rowmap(name)
    for r in rs:
        v = rm.get(r)
        print(f'  {name}!{r}:', [(colL(i), x) for i, x in enumerate(v) if x not in (None, '')] if v else None)
def sr(v, L='C'):
    x = get(v, L)
    return int(x) if isinstance(x, float) and x == int(x) else x

print('--- (a) numeric / odd company_id rows and the duplicate position_company key')
dump('position_company', [3929, 3930, 3931, 6339]); dump('company_nav', [6622, 6623, 6624])
pc = rowmap('position_company')
print('  position_company rows with lookup_key "Index VI (Jersey) (A)|120":', [r for r, v in pc.items() if get(v, 'E') == 'Index VI (Jersey) (A)|120'],
      '-> XMATCH/Match return the FIRST (row 3930, company_id numeric 0, position_id blank), not the real row 3931 (NC_PQ/27718)')
print('  company rows for key Index VI (Jersey) (A)|120:', [(r, get(v, 'E'), get(v, 'F')) for r, v in rowmap('company').items() if get(v, 'D') == 'Index VI (Jersey) (A)|120'])

print('--- (b) position_company blank lookup_key rows (PositionCompanyKey = E2:E87982)')
blank = [(r, v) for r, v in pc.items() if get(v, 'E') in (None, '')]
lt = collections.Counter(get(v, 'D') for r, v in blank)
print(f'  blank-E rows: {len(blank)} of {len(pc)}; line_type: {lt.most_common(5)}; first blank-E row: {min(r for r, v in blank)}; last row WITH a key: {max(r for r, v in pc.items() if get(v, "E") not in (None, ""))}')
print('  example blank-E row:', [(colL(i), x) for i, x in enumerate(blank[0][1]) if x not in (None, '')])
print('  => rows 8323..87982 are component rows (P=component_label, Q=component_value) used only by the AT117:AZ117 SUMPRODUCT array formulas; they carry no lookup_key, so XMATCH/Match never see them.')
rollup = collections.Counter(sr(v) for r, v in pc.items() if get(v, 'D') == 'rollup')
hdrtext = [(r, get(v, 'A'), get(v, 'F')) for r, v in pc.items() if isinstance(get(v, 'F'), str) and get(v, 'F').strip().lower() == 'company id']
print(f'  rollup rows by source_row: {sorted(rollup.items())}; rows whose company_id(F) is the literal header text "Company ID": {len(hdrtext)} e.g. {hdrtext[:4]}')
print('  Fund View G116:G119 / H119:', [(f"G{r}", fv[r][col_idx("G")]) for r in (116, 117, 118, 119)], ('H119', fv[119][col_idx('H')]), '; kinds:', {f'{L}{r}': fv_kind.get(f'{L}{r}') for L in ('G', 'H') for r in (119,)},
      '; AT118:AZ118 kinds:', {f'{L}118': fv_kind.get(f'{L}118', 'blank') for L in ['AT','AU','AV','AW','AX','AY','AZ']},
      '; AT119:AZ119 kinds:', {f'{L}119': fv_kind.get(f'{L}119', 'blank') for L in ['AT','AU','AV','AW','AX','AY','AZ']})
print('  => rows 117-119 lie inside COMPANY_QUARTER_RANGE (AT117:AZ159) but are not company rows; a constant typed in AT118:AZ118 / AT119:AZ119 (untracked) goes down PushCompanyQuarterOverride with key <fund>|118|<date> or <fund>|119|<date> and companyId = CStr(G118)="" or CStr(G119)="Company ID".')

print('--- (c) error clusters -> which funds / which error codes')
for name in ['fund', 'position', 'company', 'position_company', 'company_nav', 'cashflow', 'valuation_history']:
    hdr, rows = tables[name]
    bycol = collections.defaultdict(collections.Counter); funds = collections.defaultdict(set); rws = collections.defaultdict(list)
    for xlrow, v in rows:
        for i, x in enumerate(v):
            if isinstance(x, str) and x in ERRCODES:
                L = colL(i); bycol[L][ERRCODES[x]] += 1; funds[L].add(get(v, 'A')); rws[L].append(xlrow)
    allf = set().union(*funds.values()) if funds else set()
    print(f'  {name}: rows with errors touch {len(allf)} funds: {sorted(allf)}')
    for L in sorted(bycol, key=lambda s: (len(s), s)):
        print(f'     {L} {hdr[col_idx(L)]}: {dict(bycol[L])} funds={len(funds[L])} rows {rws[L][:3]}..{rws[L][-1]}')
print('  => the #N/A/#REF! cells are CONSTANTS (kind "err"), i.e. pasted error values, not live formulas. Fund View lookups wrap INDEX in IFERROR(...,"") so such cells display as blank.')
print('  position!P percentage_of_fund_held errors:', dict(collections.Counter(ERRCODES[get(v, "P")] for _, v in tables["position"][1] if isinstance(get(v, "P"), str) and get(v, "P") in ERRCODES)), '-> Fund View F9/G9 show "" for those funds')

print('--- (d) #N/A company_id rows across company / position_company / company_nav')
cn = rowmap('company_nav')
na = [(r, v) for r, v in cn.items() if get(v, 'E') == '0x2a']
print(f'  company_nav #N/A company_id rows: {len(na)}; name blank: {sum(1 for r, v in na if get(v, "F") in (None, ""))}; currency blank: {sum(1 for r, v in na if get(v, "G") in (None, ""))}; funds: {sorted(set(get(v, "A") for r, v in na))}')
print('  examples:', [(r, get(v, 'D'), get(v, 'F'), get(v, 'G'), get(v, 'I')) for r, v in na[:4]])
pcna = {(get(v, 'A'), sr(v)) for r, v in pc.items() if get(v, 'F') == '0x2a'}
cona = {(get(v, 'A'), sr(v, 'B')) for r, v in rowmap('company').items() if get(v, 'E') == '0x2a'}
print(f'  (fund,row) pairs with #N/A company_id: position_company={len(pcna)}, company={len(cona)}, overlap company&position_company={len(cona & pcna)}; company_nav #N/A rows whose (fund,row) is #N/A in position_company: {sum(1 for r, v in na if (get(v, "A"), sr(v)) in pcna)}')
print('  => for these 5 funds Fund View G shows "" (IFERROR) although a company exists; an AT:AZ override there creates company_nav rows with company_id "" and cannot copy name/currency.')

print('--- (e) the 485 blank-company_id company_nav rows: does ANY table know an id for that (fund,row)?')
blank485 = [(r, v) for r, v in cn.items() if get(v, 'E') in (None, '')]
pcrows = {}
for r, v in pc.items():
    if get(v, 'E') not in (None, '') and get(v, 'D') != 'rollup':
        pcrows.setdefault((get(v, 'A'), sr(v)), get(v, 'F'))
corows = {(get(v, 'A'), sr(v, 'B')): get(v, 'E') for r, v in rowmap('company').items()}
in_pc = [(r, v) for r, v in blank485 if (get(v, 'A'), sr(v)) in pcrows]
in_co = [(r, v) for r, v in blank485 if (get(v, 'A'), sr(v)) in corows]
print(f'  (fund,row) present in position_company: {len(in_pc)} (of which company_id blank there too: {sum(1 for r, v in in_pc if pcrows[(get(v, "A"), sr(v))] in (None, ""))}); present in company: {len(in_co)} (blank there too: {sum(1 for r, v in in_co if corows[(get(v, "A"), sr(v))] in (None, ""))})')
print(f'  distinct company names among the 485: {len(set(get(v, "F") for r, v in blank485))} e.g. {sorted(set(get(v, "F") for r, v in blank485))[:6]}')
print('  => these are real companies (name+currency+NAV present) that have NO company_id anywhere in the flat file; the VBA copy logic keys on company_id so it can never find them.')

print('--- (f) valuation_history: quarter_end range vs the Fund View J ladder, text/zero valuation dates')
vh = rowmap('valuation_history')
qe = [get(v, 'J') for v in vh.values()]
print(f'  quarter_end(J) min {fmt_date(min(qe))} max {fmt_date(max(qe))}; rows before 2024-12-31: {sum(1 for x in qe if x < 45657)}; after 2034-12-31: {sum(1 for x in qe if x > 49309)} (Fund View J23..J63 ladder = 2024-12-31..2034-12-31, 41 quarters; J64 blank)')
cur = fv[6][col_idx('C')]
ab = [(r, fmt_date(get(v, 'J')), get(v, 'H'), get(v, 'I')) for r, v in vh.items() if get(v, 'A') == cur]
print(f'  current fund {cur!r}: {len(ab)} valuation_history rows (one per ladder quarter): first 4 {ab[:4]}; rows with valuation_date(H) blank: {sum(1 for x in ab if x[2] in (None, ""))}')
perfund = collections.Counter(get(v, 'A') for v in vh.values())
print(f'  rows per fund: min {min(perfund.values())} max {max(perfund.values())}; funds with != 40 rows: {[(f, c) for f, c in perfund.items() if c != 40][:8]} (count {sum(1 for c in perfund.values() if c != 40)})')
liq = [(r, get(v, 'A'), get(v, 'H'), fmt_date(get(v, 'J'))) for r, v in vh.items() if get(v, 'H') == 'Liquidated']
print(f'  H valuation_date == "Liquidated" (text in a date column): {len(liq)} rows / {len(set(x[1] for x in liq))} funds e.g. {liq[:3]}')
gl = collections.Counter(get(v, 'A') for v in vh.values() if get(v, 'G') == 'Liquidated')
print(f'  G val_date == "Liquidated": {sum(gl.values())} rows / {len(gl)} funds')
dump('valuation_history', [16413])
print('  => row 16413 (Index VI (Jersey) (A), 2026-09-30) has valuation_date H = 0 (=1899-12-30) with unadjusted_nav 29,000,000 - shape of a typed-over value.')

print('--- (g) company_breakdown rows whose source_row / lookup_key disagree')
dump('company_breakdown', [382, 1034])
cb = rowmap('company_breakdown')
print('  OATV II rows (C,D):', [(r, get(v, 'C'), get(v, 'D')) for r, v in cb.items() if get(v, 'A') == 'OATV II'])
print('  China Spec Opp III rows (C,D):', [(r, get(v, 'C'), get(v, 'D')) for r, v in cb.items() if get(v, 'A') == 'China Spec Opp III'])
print('  => keys D are what Match uses, so both rows still resolve; but source_row C is 0 / "nm" (text) - any logic keyed on source_row would miss them.')

print('--- (h) template / test funds present in production tables')
tmpl = sorted(x for x in position_id_set if re.search(r'template|test', str(x), re.I))
print(f'  position ids matching /template|test/i: {len(tmpl)} {tmpl}')
for name in ['fund', 'position', 'company', 'position_company', 'company_nav', 'cashflow', 'valuation_history', 'company_breakdown']:
    c = collections.Counter(get(v, 'A') for _, v in tables[name][1] if get(v, 'A') in set(tmpl))
    print(f'     {name}: rows for template/test funds = {sum(c.values())} {dict(c)}')

print('--- (i) cashflow error rows and quarter_end population')
cf = rowmap('cashflow')
fe = [(r, get(v, 'A')) for r, v in cf.items() if any(isinstance(x, str) and x in ERRCODES for x in v)]
print(f'  cashflow rows with any error cell: {len(fe)}; funds: {sorted(set(f for r, f in fe))}')
dump('cashflow', [904])
print(f'  cashflow rows with quarter_end(J) populated: {sum(1 for v in cf.values() if get(v, "J") not in (None, ""))}; with cashflow_date(I) populated: {sum(1 for v in cf.values() if get(v, "I") not in (None, ""))}; I set but J blank: {[(r, fmt_date(get(v, "I"))) for r, v in cf.items() if get(v, "J") in (None, "") and get(v, "I") not in (None, "")]}')
print('  => cashflow keys are <fund>|<row> (no date), so quarter_end is not part of any VBA key; the 287 #N/A fx_rate constants sit outside the tracked target columns.')

print('--- (j) company_id with trailing TAB (Trim() does not remove tabs)')
print('  company_nav rows:', [(r, get(v, 'D'), get(v, 'F')) for r, v in cn.items() if get(v, 'E') == 'HF/00301\t'])
print('  position_company rows:', [(r, get(v, 'E')) for r, v in pc.items() if get(v, 'F') == 'HF/00301\t'], '; company rows:', [(r, get(v, 'D')) for r, v in rowmap('company').items() if get(v, 'E') == 'HF/00301\t'])
print('  any "HF/00301" WITHOUT the tab in company/position_company/company_nav?', [r for r, v in cn.items() if get(v, 'E') == 'HF/00301'] + [r for r, v in pc.items() if get(v, 'F') == 'HF/00301'])
print('  => the tab is consistent across all three tables, so Match(CStr(G), company_nav!E) still succeeds; it only breaks if anyone re-keys the id by hand.')

print('--- (k) C6 dropdown source: BrtDVal records in Fund View part are tokenised formulas; the list source cannot be read offline -> BLOCKED (verify in Excel that it is position!A / PositionKey).')

print(f'\ncheck.py finished in {time.time() - t0:.1f}s')
