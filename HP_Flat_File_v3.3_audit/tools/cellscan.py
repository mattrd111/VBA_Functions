import struct, sys, re, ast, collections
from parse_biff import records
def colL(c):
    s=''; c+=1
    while c: c,r=divmod(c-1,26); s=chr(65+r)+s
    return s
FORMULA={8:'fmla_str',9:'fmla_num',10:'fmla_bool',11:'fmla_err'}
CONST={1:'blank',2:'rk',3:'err',4:'bool',5:'real',6:'str',7:'isst'}
def scan(path, rows_wanted=None, max_row=None):
    """yield (row, col, kind) for cells"""
    data=open(path,'rb').read()
    row=None
    for rid,body in records(data):
        if rid==0x00:
            row=struct.unpack_from('<I',body,0)[0]
            if max_row is not None and row>max_row: return
        elif rid in FORMULA or rid in CONST:
            if row is None: continue
            if rows_wanted is not None and row not in rows_wanted: continue
            col=struct.unpack_from('<I',body,0)[0]
            yield row, col, (FORMULA.get(rid) or CONST.get(rid)), rid, body
if __name__=='__main__':
    # 1. Fund View tracked cells classification
    tracked={}
    for line in open('ctx/formulabank.txt'):
        m=re.match(r'^(\d+) (\[.*\])$',line.rstrip('\n'))
        if m and int(m.group(1))>1:
            r=ast.literal_eval(m.group(2)); tracked[r[0]]=(int(m.group(1)),r)
    fv={}
    for row,col,kind,rid,body in scan('xlsb/xl/worksheets/sheet3.bin'):
        fv[f"{colL(col)}{row+1}"]=kind
    kinds=collections.Counter(fv.get(ref,'missing') for ref in tracked)
    print("Fund View tracked-cell kinds:",kinds)
    nonf=[ref for ref in tracked if not fv.get(ref,'').startswith('fmla')]
    print("tracked cells NOT holding a formula:",len(nonf),nonf[:80])
    # untracked formula cells in AT117:AZ159?
    inrange=[k for k in fv if re.fullmatch(r'A[T-Z]\d+',k) and 117<=int(k[2:])<=159]
    print("cells present in AT117:AZ159:",len(inrange),"tracked:",sum(1 for k in inrange if k in tracked),"kinds of untracked:",collections.Counter(fv[k] for k in inrange if k not in tracked))
    print("C6 kind:",fv.get('C6'),"C3:",fv.get('C3'),"AT114:",fv.get('AT114'),"G120:",fv.get('G120'),"E2:",fv.get('E2'),"C2:",fv.get('C2'))
    # 2. source sheets: row 2 & 3 classification (which columns are formulas)
    srcs={'fund':'sheet5','position':'sheet6','company':'sheet7','position_company':'sheet8','company_nav':'sheet9','cashflow':'sheet10','valuation_history':'sheet11','company_breakdown':'sheet12'}
    for name,f in srcs.items():
        cells=list(scan(f'xlsb/xl/worksheets/{f}.bin', rows_wanted={1,2}, max_row=2))
        fcols=sorted({colL(c) for r,c,k,_,_ in cells if k.startswith('fmla')}, key=lambda s:(len(s),s))
        ccols=sorted({colL(c) for r,c,k,_,_ in cells if not k.startswith('fmla')}, key=lambda s:(len(s),s))
        print(f"{name}: rows2-3 formula cols={fcols} const cols={ccols}")
