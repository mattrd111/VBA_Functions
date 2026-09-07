#!/usr/bin/env python3
"""
UAT data check - pushback key resolution for HP_Flat_File_Test_v3.3.xlsb

Simulates, for the current fund (Fund View!C6) and a sample of other funds, what
Sheet730.PushOverrideToSource / AppendOverrideRow / PushCompanyQuarterOverride
would do for EVERY tracked cell in FormulaBank, using the real table contents
read with pyxlsb (Excel is not available, so nothing is executed live).

Key construction mirrors the VBA exactly:
  * PushOverrideToSource:   keyExpr with "ROW()" replaced by the cell row, then
                            evaluated -> C6 / C6|row / C6|TEXT(J#,"yyyy-mm-dd")
                            (C6 is NOT trimmed on this path), looked up with
                            Application.Match(key, wsSource.Range(KeyRange), 0)
                            inside the named range bounds.
  * PushCompanyQuarterOverride (cells in AT117:AZ159): Trim(C6)|row|Format(row-114
                            header,"yyyy-mm-dd") looked up in company_nav lookup_key
                            rows 2..15131 (hard-coded bound).
Row-114 headers are fund dependent (decoded from the BIFF12 formula tokens):
  AT114 = LET(d, MAXIFS(company_nav!$H$2:$H$15134, company_nav!$A$2:$A$15134, $C$6,
                        company_nav!$I$2:$I$15134, "<>"), IF(d>0, MAX($AA$23:$AA$64), d))
  AU114:AZ114 = EOMONTH(<cell to the left>, -3)
Fund View J23 is a constant (2024-12-31) and J24:J63 = EOMONTH(J23,3); J64 is blank,
so the valuation_history keys are a fixed quarter series (TEXT(blank) -> "1900-01-00").
"""
import sys, os, re, struct, collections, datetime
SCR = '/tmp/claude-0/-home-user-VBA-Functions/4411ca91-ef32-5ffa-9c78-fab36847f3b2/scratchpad'
sys.path.insert(0, SCR)
from pyxlsb import open_workbook
from parse_biff import records, wstr

WB = os.path.join(SCR, 'HP_Flat_File_Test_v3.3.xlsb')
CURRENT_FUND = 'Abenex IV (A)'
SAMPLE = ['5AM Ventures IV', 'Bain Venture 2009', 'Clearview III', 'EQT VII', 'Harbour VI',
          'IK VII No.4 ', 'Matrix China II', 'Permira III (A)', 'TA SDF III', 'Yorktown IX']
FUNDS = [CURRENT_FUND] + SAMPLE
COMPANY_NAV_HARD_BOUND = 15131          # PushCompanyQuarterOverride: rows 2..15131
CQ_RANGE = ('AT', 'AZ', 117, 159)       # COMPANY_QUARTER_RANGE = "AT117:AZ159"

# ----------------------------------------------------------------- helpers
def col2num(s):
    n = 0
    for ch in s: n = n * 26 + (ord(ch) - 64)
    return n
def num2col(n):
    s = ''
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s
def split_ref(ref):
    m = re.fullmatch(r'([A-Z]+)(\d+)', ref); return m.group(1), int(m.group(2))
EPOCH = datetime.date(1899, 12, 30)
def serial_to_date(s):
    return EPOCH + datetime.timedelta(days=int(s))
def excel_text_ymd(v):
    """Excel TEXT(v,"yyyy-mm-dd"); blank/0 -> '1900-01-00' (Excel's day-zero quirk)."""
    if v is None or v == '': v = 0
    if isinstance(v, str): return v
    if int(v) == 0: return '1900-01-00'
    return serial_to_date(v).isoformat()
def vba_format_ymd(v):
    """VBA Format(CDate(v),'yyyy-mm-dd') of a date-formatted cell value. Serial 0 -> 1899-12-30."""
    return serial_to_date(v).isoformat()
def eomonth(serial, months):
    """Excel EOMONTH; returns '#NUM!' when the result falls before 1900-01-01."""
    if isinstance(serial, str): return '#NUM!'
    if int(serial) == 0:
        y, m = 1900, 1                         # Excel treats 0 as 1900-01-00 (January 1900)
    else:
        d = serial_to_date(serial); y, m = d.year, d.month
    m += months
    while m < 1: m += 12; y -= 1
    while m > 12: m -= 12; y += 1
    if m == 12: last = datetime.date(y, 12, 31)
    else: last = datetime.date(y, m + 1, 1) - datetime.timedelta(days=1)
    s = (last - EPOCH).days
    return s if s >= 1 else '#NUM!'
def match_key(k):
    """Application.Match exact match is case-insensitive."""
    return k.casefold() if isinstance(k, str) else k
def in_cq_range(ref):
    c, r = split_ref(ref)
    return col2num(CQ_RANGE[0]) <= col2num(c) <= col2num(CQ_RANGE[1]) and CQ_RANGE[2] <= r <= CQ_RANGE[3]

RE_KEY_ROW = re.compile(r'\$?C\$?6&"\|"&(\d+)')
RE_KEY_J = re.compile(r'\$?C\$?6&"\|"&TEXT\(J(\d+),"yyyy-mm-dd"\)')
RE_KEY_CN = re.compile(r'C6&"\|"&(\d+)&"\|"&TEXT\(([A-Z]+)\$114,"yyyy-mm-dd"\)&""')

out_lines = []
def P(*a):
    s = ' '.join(str(x) for x in a); print(s); out_lines.append(s)

# ----------------------------------------------------------------- 1. defined names from workbook.bin
def load_names():
    wbbin = open(os.path.join(SCR, 'xlsb/xl/workbook.bin'), 'rb').read()
    sheets = []; xti = []; names = {}
    for rid, body in records(wbbin):
        if rid == 0x9C:                                   # BrtBundleSh
            hs, itab = struct.unpack_from('<II', body, 0)
            rel, off = wstr(body, 8); name, off = wstr(body, off); sheets.append(name)
        elif rid == 0x16A:                                # BrtExternSheet
            cxti = struct.unpack_from('<I', body, 0)[0]
            for i in range(cxti):
                xti.append(struct.unpack_from('<III', body, 4 + 12 * i))
        elif rid == 0x27:                                 # BrtName
            name, off = wstr(body, 9)
            cce = struct.unpack_from('<I', body, off)[0]; off += 4
            rgce = body[off:off + cce]
            if cce and rgce[0] in (0x3B, 0x5B, 0x7B):
                ixti, r1, r2 = struct.unpack_from('<HII', rgce, 1); c1, c2 = struct.unpack_from('<HH', rgce, 11)
                names[name] = dict(ixti=ixti, r1=r1 + 1, r2=r2 + 1, c1=num2col((c1 & 0x3FFF) + 1), c2=num2col((c2 & 0x3FFF) + 1))
            elif cce and rgce[0] in (0x3A, 0x5A, 0x7A):
                ixti, r1 = struct.unpack_from('<HI', rgce, 1); c1 = struct.unpack_from('<H', rgce, 7)[0]
                names[name] = dict(ixti=ixti, r1=r1 + 1, r2=r1 + 1, c1=num2col((c1 & 0x3FFF) + 1), c2=num2col((c1 & 0x3FFF) + 1))
    for n, d in names.items():
        try:
            ext, first, last = xti[d['ixti']]
            d['sheet'] = sheets[first] if ext == 0 else f'<external {ext}>'
        except Exception:
            d['sheet'] = '?'
    return names

# ----------------------------------------------------------------- 2. load workbook data
def load_tables():
    T = {}
    with open_workbook(WB) as wb:
        # FormulaBank
        bank = []
        with wb.get_sheet('FormulaBank') as sh:
            for row in sh.rows():
                if row[0].r == 0: continue
                v = [(c.v if c.v is not None else '') for c in row]
                v = (v + [''] * 8)[:8]
                if v[0] == '': continue
                bank.append(dict(ref=v[0], formula=v[1], isarray=v[2], canpush=v[3], sheet=v[4], keyexpr=v[5], keyrange=v[6], targetcol=v[7]))
        T['bank'] = bank
        # Fund View cached values (current fund) : J22:J64, G116:G159, AA23:AA64, AT114:AZ114, C6
        fv = {}
        with wb.get_sheet('Fund View') as sh:
            for row in sh.rows():
                r = row[0].r + 1
                if r > 160: break
                for c in row:
                    if c.v is None: continue
                    col = num2col(c.c + 1)
                    if (col == 'J' and 22 <= r <= 64) or (col == 'G' and 116 <= r <= 159) or (col == 'AA' and 23 <= r <= 64) \
                       or (r == 114 and col in ('AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ')) or (r == 6 and col == 'C') \
                       or (col == 'BJ' and 22 <= r <= 24):
                        fv[f'{col}{r}'] = c.v
        T['fv'] = fv
        # source tables
        def load(name, keep):
            hdr = None; rows = []
            with wb.get_sheet(name) as sh:
                for row in sh.rows():
                    r = row[0].r + 1
                    vals = {}
                    for c in row:
                        if c.v is None: continue
                        vals[c.c + 1] = c.v
                    if r == 1:
                        hdr = {num2col(k): v for k, v in vals.items()}; continue
                    if not vals: continue
                    rows.append((r, {k: vals.get(col2num(k)) for k in keep}))
            return hdr, rows
        T['fund'] = load('fund', ['A'])
        T['position'] = load('position', ['A'])
        T['company'] = load('company', ['A', 'B', 'D', 'E'])
        T['position_company'] = load('position_company', ['E'])
        T['company_nav'] = load('company_nav', ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'])
        T['cashflow'] = load('cashflow', ['D'])
        T['valuation_history'] = load('valuation_history', ['A', 'C', 'D', 'H', 'J', 'K'])
        T['company_breakdown'] = load('company_breakdown', ['D'])
    return T

# ----------------------------------------------------------------- 3. index builders
def build_key_index(rows, col, r1, r2):
    """first-match index (Application.Match semantics) restricted to rows r1..r2; also duplicates and out-of-range keys."""
    idx = {}; dups = collections.defaultdict(list); outside = []
    for r, v in rows:
        k = v.get(col)
        if k is None or k == '': continue
        mk = match_key(k)
        if r1 <= r <= r2:
            if mk in idx: dups[mk].append(r)
            else: idx[mk] = r
        else:
            outside.append((r, k))
    return idx, dups, outside

def main():
    names = load_names()
    T = load_tables()
    bank = T['bank']; fv = T['fv']
    P('=' * 100)
    P('UAT DATA CHECK - PUSHBACK KEY RESOLUTION  (HP_Flat_File_Test_v3.3.xlsb)')
    P('=' * 100)
    P(f'FormulaBank rows loaded: {len(bank)}   Fund View C6 (current fund) = {fv.get("C6")!r}')
    assert fv.get('C6') == CURRENT_FUND

    # ---- named ranges
    P('\n--- [A] Named KeyRanges (parsed from workbook.bin BrtName records) vs table extents ---')
    key_defs = {}
    for nm in ['PositionKey', 'FundKey', 'CompanyKey', 'PositionCompanyKey', 'CompanyNavKey', 'CashflowKey', 'QuarterKey', 'CompanyBreakdownKey', 'ValuationHistoryKey', 'MacroFlag']:
        d = names.get(nm)
        if d: P(f'  {nm:22s} -> {d["sheet"]}!{d["c1"]}{d["r1"]}:{d["c2"]}{d["r2"]}')
        else: P(f'  {nm:22s} -> MISSING')
        key_defs[nm] = d
    table_of = {'PositionKey': 'position', 'FundKey': 'fund', 'CompanyKey': 'company', 'PositionCompanyKey': 'position_company',
                'CompanyNavKey': 'company_nav', 'CashflowKey': 'cashflow', 'QuarterKey': 'valuation_history', 'CompanyBreakdownKey': 'company_breakdown'}
    idx = {}; dups = {}; outside = {}; extents = {}
    for nm, tbl in table_of.items():
        d = key_defs[nm]; hdr, rows = T[tbl]
        last = max(r for r, v in rows if v.get(d['c1']) not in (None, ''))
        extents[tbl] = (last, len(rows))
        idx[nm], dups[nm], outside[nm] = build_key_index(rows, d['c1'], d['r1'], d['r2'])
        P(f'  {tbl:18s} key col {d["c1"]} ({hdr.get(d["c1"])}): last key row {last}, data rows {len(rows)}, named range rows {d["r1"]}..{d["r2"]}, '
          f'keys OUTSIDE named range: {len(outside[nm])}, duplicate keys inside: {len(dups[nm])}')
        if outside[nm]: P('     outside:', outside[nm][:10])
        if dups[nm]: P('     duplicates (first 8):', [(k, [idx[nm][k]] + v) for k, v in list(dups[nm].items())[:8]])
    # company_nav hard-coded bound
    hdr_cn, rows_cn = T['company_nav']
    cn_idx, cn_dups, cn_outside = build_key_index(rows_cn, 'D', 2, COMPANY_NAV_HARD_BOUND)
    P(f'  company_nav VBA hard-coded search rows 2..{COMPANY_NAV_HARD_BOUND}: keys BEYOND bound: {len(cn_outside)} -> {cn_outside}')
    P(f'  company_nav lookup_key column resolved by VBA Match("lookup_key", row1) = {[c for c, h in hdr_cn.items() if h == "lookup_key"]}, nav_value col = {[c for c, h in hdr_cn.items() if h == "nav_value"]}')
    # case-insensitive collisions: distinct raw keys mapping to same casefolded key
    P('\n--- [B] Case-insensitive collisions (Application.Match is case-insensitive) ---')
    for nm, tbl in table_of.items():
        d = key_defs[nm]; raw = collections.defaultdict(set)
        for r, v in T[tbl][1]:
            k = v.get(d['c1'])
            if k not in (None, ''): raw[match_key(k)].add(k)
        coll = {k: v for k, v in raw.items() if len(v) > 1}
        P(f'  {tbl:18s}: {len(coll)} casefold collisions' + (f' e.g. {list(coll.items())[:3]}' if coll else ''))

    # ---- FormulaBank structural checks
    P('\n--- [C] FormulaBank: KeyRange column vs formula XMATCH column, TargetCol vs formula header ---')
    hdrs = {tbl: T[tbl][0] for tbl in table_of.values()}
    bj = {k: fv.get(k) for k in ('BJ22', 'BJ23', 'BJ24')}
    P(f'  Fund View BJ22/BJ23/BJ24 cached = {bj}  (formulas decoded from BIFF: MATCH("val_date"|"valuation_date"|"unadjusted_nav", valuation_history!row1))')
    vh_hdr = hdrs['valuation_history']
    for k, hname in (('BJ22', 'val_date'), ('BJ23', 'valuation_date'), ('BJ24', 'unadjusted_nav')):
        want = [c for c, h in vh_hdr.items() if h == hname][0]
        P(f'    {k}={bj[k]} -> column {num2col(int(bj[k]))} header {vh_hdr.get(num2col(int(bj[k])))!r}; header "{hname}" is at column {want} -> {"OK" if num2col(int(bj[k])) == want else "MISMATCH"}')
    tc_ok = tc_bad = tc_skip = 0; tc_bad_list = []
    kr_ok = kr_bad = 0; kr_bad_list = []
    pat_counter = collections.Counter()
    for b in bank:
        if b['sheet'] == '':
            tc_skip += 1; continue
        f = b['formula']
        m = re.search(r'MATCH\("([^"]+)",\s*([A-Za-z_]+)!\$A\$1:\$([A-Z]+)\$1,0\)', f)
        if m:
            hname, shname = m.group(1), m.group(2)
            actual = hdrs[b['sheet']].get(b['targetcol'])
            if shname == b['sheet'] and actual == hname: tc_ok += 1
            else: tc_bad += 1; tc_bad_list.append((b['ref'], b['sheet'], b['targetcol'], hname, actual))
        else:
            m2 = re.search(r'\$BJ\$(2[234])\)', f)
            if m2 and b['sheet'] == 'valuation_history':
                colidx = int(bj['BJ' + m2.group(1)])
                if num2col(colidx) == b['targetcol']: tc_ok += 1
                else: tc_bad += 1; tc_bad_list.append((b['ref'], b['sheet'], b['targetcol'], 'BJ' + m2.group(1), num2col(colidx)))
            else:
                tc_skip += 1; tc_bad_list.append((b['ref'], b['sheet'], b['targetcol'], 'NO HEADER MATCH IN FORMULA', None))
        mx = re.search(r'XMATCH\((.+?),\s*([A-Za-z_]+)!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+),0,1\)', f)
        d = key_defs.get(b['keyrange'])
        if mx and d:
            if mx.group(2) == d['sheet'] and mx.group(3) == d['c1']: kr_ok += 1
            else: kr_bad += 1; kr_bad_list.append((b['ref'], b['keyrange'], f'{mx.group(2)}!{mx.group(3)}', f'{d["sheet"]}!{d["c1"]}'))
            pat_counter[(b['sheet'], f'{mx.group(2)}!${mx.group(3)}${mx.group(4)}:${mx.group(5)}${mx.group(6)}', f'{d["sheet"]}!{d["c1"]}{d["r1"]}:{d["c2"]}{d["r2"]}')] += 1
        elif d is None:
            kr_bad += 1; kr_bad_list.append((b['ref'], b['keyrange'], 'NAME NOT FOUND', None))
    P(f'  TargetCol vs formula header: OK={tc_ok} MISMATCH={tc_bad} skipped(no source)={tc_skip}')
    for x in tc_bad_list[:20]: P('     ', x)
    P(f'  KeyRange column vs formula XMATCH column: OK={kr_ok} MISMATCH={kr_bad}')
    for x in kr_bad_list[:20]: P('     ', x)
    P('  Formula lookup range vs named KeyRange (per source sheet):')
    for (sh, frng, nrng), n in sorted(pat_counter.items()): P(f'     {sh:18s} formula {frng:40s} named {nrng:40s} cells={n}')
    # cells in CQ range by source / canpush
    cq_cells = [b for b in bank if in_cq_range(b['ref'])]
    P(f'  Tracked cells inside COMPANY_QUARTER_RANGE AT117:AZ159: {len(cq_cells)} -> by (SourceSheet,CanPushback): {collections.Counter((b["sheet"] or "(blank)", b["canpush"]) for b in cq_cells)}')
    P(f'  Cells inside AT117:AZ159 NOT tracked in FormulaBank (also routed to PushCompanyQuarterOverride when typed over): '
      f'{[f"{c}{r}" for r in range(117,160) for c in ["AT","AU","AV","AW","AX","AY","AZ"] if f"{c}{r}" not in {b["ref"] for b in bank}]}')

    # ---- Fund View J column (fixed quarter series)
    P('\n--- [D] Fund View J23:J64 (keys for valuation_history) ---')
    J = {r: fv.get(f'J{r}') for r in range(23, 65)}
    chain_ok = all(J[r] == eomonth(J[r - 1], 3) for r in range(24, 64))
    P(f'  J23={J[23]} ({excel_text_ymd(J[23])}) constant; J24:J63 = EOMONTH(previous,3) chain verified against cached values: {chain_ok}; J63={J[63]} ({excel_text_ymd(J[63])}); J64={J[64]!r} -> TEXT(J64,"yyyy-mm-dd") = {excel_text_ymd(J[64])!r}')
    P('  => J dates are NOT fund-specific; the valuation_history KeyExpr yields the same 42 keys (C6|<quarter>) for every fund.')

    # ---- per-fund simulation
    hdr_vh, rows_vh = T['valuation_history']
    hdr_co, rows_co = T['company']
    vh_by_fund = collections.defaultdict(list)
    for r, v in rows_vh: vh_by_fund[v['A']].append((r, v))
    cn_by_fund = collections.defaultdict(list)
    for r, v in rows_cn: cn_by_fund[v['A']].append((r, v))
    co_idx = idx['CompanyKey']
    co_by_row = {r: v for r, v in rows_co}
    def company_id_for(fund, row):
        """Fund View G<row> = INDEX(company!E, XMATCH(C6|row, company!D)) -> '' when not found."""
        r = co_idx.get(match_key(f'{fund}|{row}'))
        return (co_by_row[r]['E'] if r and co_by_row[r]['E'] is not None else '') if r else ''
    def headers_for(fund):
        d = 0
        for r, v in cn_by_fund.get(fund, []):
            if v['I'] not in (None, '') and isinstance(v['H'], (int, float)) and 2 <= r <= 15134:
                d = max(d, v['H'])
        # AA23:AA64 = valuation_date (H) of the valuation_history row matched by quarter_key C6|TEXT(J#)
        aa = []
        vh_rows_map = {r: v for r, v in vh_by_fund.get(fund, [])}
        for r in range(23, 65):
            mr = idx['QuarterKey'].get(match_key(f'{fund}|{excel_text_ymd(J[r])}'))
            if mr is not None:
                h = vh_rows_map.get(mr, {}).get('H')
                aa.append(0 if h in (None, '') else h)    # INDEX of blank cell -> 0
        max_aa = max([x for x in aa if isinstance(x, (int, float))], default=0)
        at114 = max_aa if d > 0 else d
        hdrs_ = {'AT': at114}
        prev = at114
        for c in ('AU', 'AV', 'AW', 'AX', 'AY', 'AZ'):
            prev = eomonth(prev, -3); hdrs_[c] = prev
        return hdrs_, d, max_aa

    grand = collections.defaultdict(collections.Counter)
    fund_reports = {}
    for fund in FUNDS:
        P('\n' + '=' * 100)
        P(f'FUND {fund!r}  (position row {idx["PositionKey"].get(match_key(fund))}, fund row {idx["FundKey"].get(match_key(fund))})')
        hd, d, max_aa = headers_for(fund)
        hd_txt = {c: (excel_text_ymd(v) if not isinstance(v, str) else v) for c, v in hd.items()}
        P(f'  row-114 headers (recomputed per AT114 formula): d=MAXIFS nav date={d} ({excel_text_ymd(d) if d else 0}), MAX(AA23:AA64)={max_aa} -> {hd_txt}')
        if fund == CURRENT_FUND:
            cached = {c: fv.get(f'{c}114') for c in hd}
            P(f'  self-check vs cached Fund View AT114:AZ114 {cached}: {"MATCH" if all(cached[c] == hd[c] for c in hd) else "DIFFERENT"}')
            gchk = {r: fv.get(f'G{r}', '') for r in (117, 118, 119, 120, 121)}
            P(f'  self-check G117..G121 cached {gchk} vs computed G120={company_id_for(fund, 120)!r}')
        results = collections.Counter(); examples = collections.defaultdict(list); flags = collections.Counter(); flag_ex = collections.defaultdict(list)
        cq_inserts_in_order = []
        for b in bank:
            ref = b['ref']; col, row = split_ref(ref); sheet = b['sheet']
            if in_cq_range(ref):
                # ---------- PushCompanyQuarterOverride
                fund_id = fund.strip()                     # Trim(C6)
                if row in (117, 118, 119): gval = fv.get(f'G{row}', '')     # G117/G118 blank, G119 = 'Company ID' (constants)
                else: gval = company_id_for(fund, row)
                company_id = '' if gval is None else str(gval)
                raw = hd[col]
                if isinstance(raw, str):                   # #NUM! -> IsDate False -> PromptForQuarterHeader
                    outcome = 'COMPANY_QUARTER:PROMPT_FOR_DATE'
                    key = None
                else:
                    date_text = vba_format_ymd(raw)
                    key = f'{fund_id}|{row}|{date_text}'
                    mr = cn_idx.get(match_key(key))
                    if mr is not None:
                        outcome = 'COMPANY_QUARTER:UPDATE'
                        if match_key(key) in cn_dups: flags['cq_update_hits_duplicate_key(first row wins)'] += 1; flag_ex['cq_update_hits_duplicate_key(first row wins)'].append((ref, key, [mr] + cn_dups[match_key(key)]))
                    else:
                        outcome = 'COMPANY_QUARTER:INSERT'
                        beyond = [r for r, k in cn_outside if match_key(k) == match_key(key)]
                        if beyond: flags['cq_key_exists_BEYOND_row_15131 -> duplicate row inserted'] += 1; flag_ex['cq_key_exists_BEYOND_row_15131 -> duplicate row inserted'].append((ref, key, beyond))
                        full = idx['CompanyNavKey'].get(match_key(key))
                        if company_id == '': flags['cq_insert_with_blank_company_id'] += 1; flag_ex['cq_insert_with_blank_company_id'].append((ref, key))
                        elif company_id == 'Company ID': flags["cq_insert_company_id='Company ID' (G119 header text)"] += 1; flag_ex["cq_insert_company_id='Company ID' (G119 header text)"].append((ref, key))
                        if int(raw) == 0: flags['cq_insert_header_date_serial_0 (1899-12-30)'] += 1; flag_ex['cq_insert_header_date_serial_0 (1899-12-30)'].append((ref, key))
                        if fund_id != fund: flags['cq_key_uses_Trim(C6) != C6 -> orphan position_id'] += 1; flag_ex['cq_key_uses_Trim(C6) != C6 -> orphan position_id'].append((ref, key, f'table key would be {fund}|{row}|{date_text}', 'exists' if idx['CompanyNavKey'].get(match_key(f'{fund}|{row}|{date_text}')) else 'absent'))
                        # company_name/currency copy: Match(companyId, company_nav!E2:E15131)
                        if company_id == '' or not any(v['E'] == company_id for r, v in rows_cn if r <= COMPANY_NAV_HARD_BOUND):
                            flags['cq_insert_company_name/currency_not_copied (MsgBox asks to fill by hand)'] += 1
                    if b['canpush'] != 'Y': flags['cq_cell_has_CanPushback=N_but_is_pushed_anyway (AT117:AZ117 position_company rollup -> company_nav)'] += 1; flag_ex['cq_cell_has_CanPushback=N_but_is_pushed_anyway (AT117:AZ117 position_company rollup -> company_nav)'].append((ref, key, outcome))
                results[('company_nav[CQ]', outcome)] += 1
                if len(examples[('company_nav[CQ]', outcome)]) < 4: examples[('company_nav[CQ]', outcome)].append((ref, key))
                continue
            # ---------- PushOverrideToSource
            if b['canpush'] != 'Y':
                results[(sheet or '(none)', 'NO_PUSH(CanPushback<>Y)')] += 1; continue
            ke = b['keyexpr'].replace('ROW()', str(row))
            m_row = RE_KEY_ROW.fullmatch(ke); m_j = RE_KEY_J.fullmatch(ke); m_cn = RE_KEY_CN.fullmatch(ke)
            if ke in ('C6', '$C$6'): key = fund
            elif m_row: key = fund + '|' + m_row.group(1)
            elif m_j: key = fund + '|' + excel_text_ymd(J[int(m_j.group(1))])
            elif m_cn: key = fund + '|' + m_cn.group(1) + '|' + excel_text_ymd(hd[m_cn.group(2)])
            else:
                results[(sheet, 'UNPARSED_KEYEXPR ' + ke)] += 1; continue
            nm = b['keyrange']; mr = idx[nm].get(match_key(key))
            if mr is not None:
                outcome = 'UPDATE'
                if match_key(key) in dups[nm]: flags[f'{sheet}: update hits duplicate key (first row wins)'] += 1
            elif sheet in ('position', 'fund'):
                outcome = 'FAIL(MsgBox, not written)'
            else:
                outcome = 'APPEND(insert row 3)'
                if key.endswith('|1900-01-00'): flags['valuation_history: J64 blank -> key ends |1900-01-00 -> APPEND'] += 1; flag_ex['valuation_history: J64 blank -> key ends |1900-01-00 -> APPEND'].append((ref, key))
                if key.endswith('|64'): flags[f'{sheet}: row 64 has no source row in any fund -> APPEND'] += 1; flag_ex[f'{sheet}: row 64 has no source row in any fund -> APPEND'].append((ref, key))
                if sheet == 'valuation_history' and not key.endswith('|1900-01-00'): flags['valuation_history: quarter missing for this fund -> APPEND (D lookup_key & J quarter_end left blank)'] += 1; flag_ex['valuation_history: quarter missing for this fund -> APPEND (D lookup_key & J quarter_end left blank)'].append((ref, key))
                if sheet == 'cashflow': flags['cashflow: APPEND leaves formula cols V:AH blank (Issue Log #13)'] += 1
                if sheet in ('position_company', 'company') and row >= 120 and company_id_for(fund, row) == '': flags[f'{sheet}: APPEND with no company on this grid row (G blank)'] += 1
                if sheet == 'position_company' and row == 119: flags["position_company: APPEND row 119 copies G119='Company ID' into company_id"] += 1; flag_ex["position_company: APPEND row 119 copies G119='Company ID' into company_id"].append((ref, key))
            results[(sheet, outcome)] += 1
            if len(examples[(sheet, outcome)]) < 4: examples[(sheet, outcome)].append((ref, key))
        # print fund report
        P('  Outcome counts per source sheet:')
        for (sh, oc), n in sorted(results.items()):
            P(f'    {sh:20s} {oc:45s} {n:5d}   e.g. {examples[(sh, oc)][:3]}')
            grand[sh][oc] += n
        P('  Flags:')
        for fl, n in sorted(flags.items()):
            P(f'    [{n:3d}] {fl}   e.g. {flag_ex[fl][:3]}')
        fund_reports[fund] = (results, flags, hd_txt)

    # ---- all-funds aggregate (fast: only key lookups)
    P('\n' + '=' * 100)
    P('ALL 818 FUNDS - aggregate outcome counts (same simulation, examples omitted)')
    allf = [v['A'] for r, v in T['position'][1] if v.get('A') not in (None, '')]
    agg = collections.defaultdict(collections.Counter); at114_zero = []; prompt_funds = []; trim_funds = []; beyond_funds = collections.Counter()
    for fund in allf:
        hd, d, max_aa = headers_for(fund)
        if hd['AT'] == 0: at114_zero.append(fund)
        if any(isinstance(hd[c], str) for c in hd): prompt_funds.append(fund)
        if fund.strip() != fund: trim_funds.append(fund)
        for b in bank:
            ref = b['ref']; col, row = split_ref(ref); sheet = b['sheet']
            if in_cq_range(ref):
                raw = hd[col]
                if isinstance(raw, str): agg['company_nav[CQ]']['PROMPT_FOR_DATE'] += 1; continue
                key = f'{fund.strip()}|{row}|{vba_format_ymd(raw)}'
                if cn_idx.get(match_key(key)) is not None: agg['company_nav[CQ]']['UPDATE'] += 1
                else:
                    agg['company_nav[CQ]']['INSERT'] += 1
                    if any(match_key(k) == match_key(key) for r, k in cn_outside): beyond_funds[fund] += 1
                continue
            if b['canpush'] != 'Y': continue
            ke = b['keyexpr'].replace('ROW()', str(row))
            if ke in ('C6', '$C$6'): key = fund
            elif RE_KEY_J.fullmatch(ke): key = fund + '|' + excel_text_ymd(J[int(RE_KEY_J.fullmatch(ke).group(1))])
            else: key = f'{fund}|{row}'
            if idx[b['keyrange']].get(match_key(key)) is not None: agg[sheet]['UPDATE'] += 1
            elif sheet in ('position', 'fund'): agg[sheet]['FAIL'] += 1
            else: agg[sheet]['APPEND'] += 1
    for sh in sorted(agg):
        tot = sum(agg[sh].values()); P(f'  {sh:20s} ' + '  '.join(f'{k}={v} ({100*v/tot:.1f}%)' for k, v in sorted(agg[sh].items())))
    P(f'  funds whose AT114 evaluates to 0 (no company_nav nav_value): {len(at114_zero)} {at114_zero}')
    P(f'  funds whose AU114:AZ114 evaluate to #NUM! (PromptForQuarterHeader on every override in AU117:AZ159): {len(prompt_funds)} {prompt_funds}')
    P(f'  funds with leading/trailing space in id (Trim mismatch on company-quarter path & StampFundUpdated): {trim_funds}')
    P(f'  funds with company-quarter keys that exist only beyond row {COMPANY_NAV_HARD_BOUND}: {dict(beyond_funds)}')
    P('\nGRAND TOTALS over the 11 simulated funds:')
    for sh in sorted(grand): P(f'  {sh:20s} {dict(grand[sh])}')

    with open(os.path.join(SCR, 'agents/uat-pushback/output.txt'), 'w') as f:
        f.write('\n'.join(out_lines) + '\n')

if __name__ == '__main__':
    main()
