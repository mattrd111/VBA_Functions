"""Checks the formula-parsing algorithm used by addin/modAuditFormula.bas.

VBA cannot be run outside Excel, so this is a line-by-line mirror of
AnalyseFormula / IsSuspectConstant / ConstantSeverity in Python, exercised
against formulas of the kind that turn up in real models.

It verifies the ALGORITHM, not the VBA implementation - the two have to be
kept in step by hand. If you change the walker, the function lists or the
severity rules in modAuditFormula.bas, change them here too and re-run:

    python build/test_formula_parser.py

It has already earned its place: it caught the function-name stack being
pushed one level too low, which let VLOOKUP's column index through as a
hardcoded number.
"""

STRUCTURAL = {"INDEX","MATCH","XMATCH","VLOOKUP","HLOOKUP","XLOOKUP","LOOKUP","OFFSET","INDIRECT",
"CHOOSE","COLUMN","COLUMNS","ROW","ROWS","ROUND","ROUNDUP","ROUNDDOWN","MROUND","CEILING","FLOOR",
"TRUNC","LEFT","RIGHT","MID","FIND","SEARCH","SUBSTITUTE","REPT","TEXT","SUBTOTAL","AGGREGATE",
"LARGE","SMALL","RANK","PERCENTILE","QUARTILE","DATE","TIME","WEEKDAY","WORKDAY","NETWORKDAYS",
"EOMONTH","EDATE","YEARFRAC","DAYS360","CELL","INFO","ERROR.TYPE","SUMIF","SUMIFS","COUNTIF",
"COUNTIFS","AVERAGEIF","AVERAGEIFS","MAXIFS","MINIFS","TEXTJOIN","SORT","SORTBY","TAKE","DROP","WRAPROWS"}
VOLATILE = {"NOW","TODAY","RAND","RANDBETWEEN","RANDARRAY","OFFSET","INDIRECT","CELL","INFO"}
COMMON = {0,1,2,12,100,360,365,366,1000,10000,1000000,1000000000}

def is_digit(c): return len(c)==1 and '0'<=c<='9'
def ident_start(c): return c.isalpha() or c in '_$\\' or ord(c)>127
def ident_char(c): return c.isalnum() or c in '_.$!:\\?' or ord(c)>127

def at(s,i): return s[i-1] if 1<=i<=len(s) else ''
def next_real(s,i):
    while i<=len(s):
        if at(s,i)!=' ': return at(s,i)
        i+=1
    return ''
def prev_real(s,i):
    i-=1
    while i>=1:
        if at(s,i)!=' ': return at(s,i),i
        i-=1
    return '',0

def analyse(f):
    res={'Hardcodes':[],'Volatiles':[],'WholeRefs':[],'External':0,'IfError':False,'Functions':0}
    n=len(f)
    if n==0 or f[0]!='=': return res
    stack=['']*129; depth=0; i=2; pending=''
    while i<=n:
        ch=at(f,i)
        if ch=='"':
            i+=1
            while i<=n and at(f,i)!='"': i+=1
            i+=1
        elif ch=="'":
            start=i+1; i+=1
            while i<=n and at(f,i)!="'": i+=1
            if '[' in f[start-1:i-1]: res['External']+=1
            i+=1
        elif ch=='[':
            if not (i>1 and ident_char(at(f,i-1))): res['External']+=1
            while i<=n and at(f,i)!=']': i+=1
            i+=1
        elif ident_start(ch):
            start=i
            while i<=n and ident_char(at(f,i)): i+=1
            tok=f[start-1:i-1]
            if next_real(f,i)=='(':
                res['Functions']+=1
                pending=tok.upper()
                if tok.upper() in VOLATILE: res['Volatiles'].append(tok)
                if tok.upper() in ('IFERROR','IFNA'): res['IfError']=True
            else:
                if whole_col(tok): res['WholeRefs'].append(tok)
                if '!' in tok and '[' in tok: res['External']+=1
        elif is_digit(ch) or (ch=='.' and is_digit(at(f,i+1))):
            start=i
            i=consume_number(f,i)
            tok=f[start-1:i-1]
            res['Hardcodes'].append((tok, current_fn(stack,depth), direct_arg(f,start,i)))
        elif ch=='(':
            depth+=1
            if depth<=128: stack[depth]=pending
            pending=''
            i+=1
        elif ch==')':
            if depth>0:
                if depth<=128: stack[depth]=''
                depth-=1
            i+=1
        else: i+=1
    return res

def direct_arg(f,start,after):
    """The literal is the whole argument, not part of an expression."""
    c,pos = prev_real(f,start)
    if c in '+-':                       # a signed argument is still an argument
        c,pos = prev_real(f,pos)
    if c not in '(,': return False
    return next_real(f,after) in ',)'

def consume_number(f,start):
    i=start; seen_dot=False; n=len(f)
    while i<=n:
        ch=at(f,i)
        if is_digit(ch): i+=1
        elif ch=='.' and not seen_dot: seen_dot=True; i+=1
        elif ch in 'Ee' and is_exponent(f,i):
            i+=2
            if is_digit(at(f,i)): i+=1
            while i<=n and is_digit(at(f,i)): i+=1
            break
        else: break
    return i

def is_exponent(f,pos):
    nx=at(f,pos+1)
    if is_digit(nx): return True
    if nx in '+-': return is_digit(at(f,pos+2))
    return False

def current_fn(stack,depth):
    for lvl in range(depth,0,-1):
        if stack[lvl]: return stack[lvl]
    return ''

def whole_col(tok):
    if ':' not in tok: return False
    bare=tok.split('!')[-1]
    parts=bare.split(':')
    if len(parts)!=2: return False
    return all(col_only(p) for p in parts)

def col_only(t):
    t=t.replace('$','')
    return 1<=len(t)<=3 and all('A'<=c.upper()<='Z' for c in t)

def suspect(lit,fn,direct):
    try: v=float(lit)
    except ValueError: return False
    if fn in STRUCTURAL and direct: return False
    return v not in COMMON

def severity(lit):
    if '.' in lit: return 'High'
    try: v=abs(float(lit))
    except ValueError: return 'Medium'
    return 'High' if v>=1000 else 'Medium'

CASES = [
 ("=B5*1.03",                                   ["1.03 High"], 0, [], 0),
 ("=VLOOKUP(A1,Data,3,FALSE)",                  [],            0, [], 0),
 ("=SUM(B5:B20)+250000",                        ["250000 High"],0,[], 0),
 ("=IF(C5>0,C5*0.15,0)",                        ["0.15 High"], 0, [], 0),
 ("=OFFSET(A1,1,2)",                            [],            1, [], 0),
 ('=SUMIF(A:A,"x",B:B)',                        [],            0, ["A:A","B:B"], 0),
 ("=INDEX(Sheet2!A1:D10,MATCH($A5,Sheet2!$A$1:$A$10,0),4)", [],0,[], 0),
 ("='C:\\work\\[Model.xlsx]Sheet1'!A1*2",       [],            0, [], 1),
 ("=Table1[Amount]*2",                          [],            0, [], 0),
 ("=ROUND(B5*1.2,2)",                           ["1.2 High"],  0, [], 0),
 ("=EOMONTH(A1,-1)",                            [],            0, [], 0),
 ("=B5*13",                                     ["13 Medium"], 0, [], 0),
 ("=250000",                                    ["250000 High"],0,[], 0),
 ("=B5*3%",                                     ["3 Medium"],  0, [], 0),
 ('=IFERROR(VLOOKUP(A1,D:F,3,0),"n/a")',        [],            0, ["D:F"], 0),
 ("=Revenue*(1+GrowthRate)",                    [],            0, [], 0),
 ("=NPV(0.08,B5:B20)",                          ["0.08 High"], 0, [], 0),
 ("=TODAY()-A1",                                [],            1, [], 0),
 ('=TEXT(A1,"0.0%")&" of "&B1',                 [],            0, [], 0),
 ("=SUM(Jan:Dec!B5)",                           [],            0, [], 0),
 ("=A1*1E5",                                     ["1E5 High"],  0, [], 0),
 ("=ROUND(VLOOKUP(A1,D:F,3,0)*1.05,2)",          ["1.05 High"], 0, ["D:F"], 0),
 ('=IF(A1>0,"yes (a,b)",7)',                     ["7 Medium"],  0, [], 0),
 ('=INDIRECT("A"&B5)',                           [],            1, [], 0),
 ("=SUM(A:A)/12",                                [],            0, ["A:A"], 0),
 ("=MAX(0,MIN(1,B5*4.5))",                       ["4.5 High"],  0, [], 0),
 ("=B5*(1+2.5%)",                                ["2.5 High"],  0, [], 0),
 ("=CHOOSE(3,A1,A2,A3)",                         [],            0, [], 0),
 ("=SUM(B5:B20)*1000000",                        [],            0, [], 0),
]

fails=0
for f,exp_hc,exp_vol,exp_wr,exp_ext in CASES:
    r=analyse(f)
    hc=[f"{l} {severity(l)}" for l,fn,d in r['Hardcodes'] if suspect(l,fn,d)]
    ok = hc==exp_hc and len(r['Volatiles'])==exp_vol and r['WholeRefs']==exp_wr and r['External']==exp_ext
    if not ok:
        fails+=1
        print(f"FAIL {f}")
        print(f"   hardcodes {hc} expected {exp_hc}")
        print(f"   volatiles {len(r['Volatiles'])} expected {exp_vol}")
        print(f"   wholerefs {r['WholeRefs']} expected {exp_wr}")
        print(f"   external  {r['External']} expected {exp_ext}")
    else:
        print(f"ok   {f:52s} -> {hc if hc else 'clean'}")
print()
print(f"{len(CASES)-fails}/{len(CASES)} passed")
