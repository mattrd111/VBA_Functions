#!/usr/bin/env python3
"""BIFF12 (XLSB) formula (rgce/rgcb) decoder + FormulaBank comparison for HP_Flat_File_Test_v3.3.xlsb.

Usage:  python3 decode.py            (writes compare.txt, summary.json next to this file)

Decodes every formula on "Fund View" (xl/worksheets/sheet3.bin), resolves shared formulas (BrtShrFmla 0x1AB)
and array formulas (BrtArrFmla 0x1AA) referenced through PtgExp, renders Excel-style A1 text and compares the
1729 FormulaBank-tracked cells with FormulaBank column B after normalisation.
"""
import struct, sys, os, re, ast, json, collections, difflib

HERE = os.path.dirname(os.path.abspath(__file__))
SCR = os.path.abspath(os.path.join(HERE, '..', '..'))
sys.path.insert(0, SCR)
from parse_biff import records, wstr

# ----------------------------------------------------------------------------------------------
# Function table (Ftab) - BIFF8 indices 0..378 plus the BIFF12 extension 380..484 ([MS-XLSB] Ftab).
# value = (name, fixed_arg_count_or_None)
# ----------------------------------------------------------------------------------------------
_FT = {
0:('COUNT',None),1:('IF',None),2:('ISNA',1),3:('ISERROR',1),4:('SUM',None),5:('AVERAGE',None),6:('MIN',None),7:('MAX',None),
8:('ROW',None),9:('COLUMN',None),10:('NA',0),11:('NPV',None),12:('STDEV',None),13:('DOLLAR',None),14:('FIXED',None),15:('SIN',1),
16:('COS',1),17:('TAN',1),18:('ATAN',1),19:('PI',0),20:('SQRT',1),21:('EXP',1),22:('LN',1),23:('LOG10',1),24:('ABS',1),25:('INT',1),
26:('SIGN',1),27:('ROUND',2),28:('LOOKUP',None),29:('INDEX',None),30:('REPT',2),31:('MID',3),32:('LEN',1),33:('VALUE',1),34:('TRUE',0),
35:('FALSE',0),36:('AND',None),37:('OR',None),38:('NOT',1),39:('MOD',2),40:('DCOUNT',3),41:('DSUM',3),42:('DAVERAGE',3),43:('DMIN',3),
44:('DMAX',3),45:('DSTDEV',3),46:('VAR',None),47:('DVAR',3),48:('TEXT',2),49:('LINEST',None),50:('TREND',None),51:('LOGEST',None),
52:('GROWTH',None),53:('GOTO',1),54:('HALT',None),55:('RETURN',None),56:('PV',None),57:('FV',None),58:('NPER',None),59:('PMT',None),
60:('RATE',None),61:('MIRR',3),62:('IRR',None),63:('RAND',0),64:('MATCH',None),65:('DATE',3),66:('TIME',3),67:('DAY',1),68:('MONTH',1),
69:('YEAR',1),70:('WEEKDAY',None),71:('HOUR',1),72:('MINUTE',1),73:('SECOND',1),74:('NOW',0),75:('AREAS',1),76:('ROWS',1),77:('COLUMNS',1),
78:('OFFSET',None),79:('ABSREF',2),80:('RELREF',2),81:('ARGUMENT',None),82:('SEARCH',None),83:('TRANSPOSE',1),84:('ERROR',None),
85:('STEP',0),86:('TYPE',1),87:('ECHO',None),88:('SET.NAME',None),89:('CALLER',0),90:('DEREF',1),91:('WINDOWS',None),92:('SERIES',None),
93:('DOCUMENTS',None),94:('ACTIVE.CELL',0),95:('SELECTION',0),96:('RESULT',None),97:('ATAN2',2),98:('ASIN',1),99:('ACOS',1),
100:('CHOOSE',None),101:('HLOOKUP',None),102:('VLOOKUP',None),103:('LINKS',None),104:('INPUT',None),105:('ISREF',1),106:('GET.FORMULA',1),
107:('GET.NAME',None),108:('SET.VALUE',2),109:('LOG',None),110:('EXEC',None),111:('CHAR',1),112:('LOWER',1),113:('UPPER',1),114:('PROPER',1),
115:('LEFT',None),116:('RIGHT',None),117:('EXACT',2),118:('TRIM',1),119:('REPLACE',4),120:('SUBSTITUTE',None),121:('CODE',1),
122:('NAMES',None),123:('DIRECTORY',None),124:('FIND',None),125:('CELL',None),126:('ISERR',1),127:('ISTEXT',1),128:('ISNUMBER',1),
129:('ISBLANK',1),130:('T',1),131:('N',1),132:('FOPEN',None),133:('FCLOSE',1),134:('FSIZE',1),135:('FREADLN',1),136:('FREAD',2),
137:('FWRITELN',2),138:('FWRITE',2),139:('FPOS',None),140:('DATEVALUE',1),141:('TIMEVALUE',1),142:('SLN',3),143:('SYD',4),144:('DDB',None),
145:('GET.DEF',None),146:('REFTEXT',None),147:('TEXTREF',None),148:('INDIRECT',None),149:('REGISTER',None),150:('CALL',None),
151:('ADD.BAR',None),152:('ADD.MENU',None),153:('ADD.COMMAND',None),154:('ENABLE.COMMAND',None),155:('CHECK.COMMAND',None),
156:('RENAME.COMMAND',None),157:('SHOW.BAR',None),158:('DELETE.MENU',None),159:('DELETE.COMMAND',None),160:('GET.CHART.ITEM',None),
161:('DIALOG.BOX',1),162:('CLEAN',1),163:('MDETERM',1),164:('MINVERSE',1),165:('MMULT',2),166:('FILES',None),167:('IPMT',None),
168:('PPMT',None),169:('COUNTA',None),170:('CANCEL.KEY',None),171:('INITIATE',2),172:('REQUEST',2),173:('POKE',3),174:('EXECUTE',2),
175:('TERMINATE',1),176:('RESTART',None),177:('HELP',None),178:('GET.BAR',None),179:('PRODUCT',None),180:('FACT',1),181:('GET.CELL',None),
182:('GET.WORKSPACE',1),183:('GET.WINDOW',None),184:('GET.DOCUMENT',None),185:('DPRODUCT',3),186:('ISNONTEXT',1),187:('GET.NOTE',None),
188:('NOTE',None),189:('STDEVP',None),190:('VARP',None),191:('DSTDEVP',3),192:('DVARP',3),193:('TRUNC',None),194:('ISLOGICAL',1),
195:('DCOUNTA',3),196:('DELETE.BAR',1),197:('UNREGISTER',1),200:('USDOLLAR',None),201:('FINDB',None),202:('SEARCHB',None),
203:('REPLACEB',4),204:('LEFTB',None),205:('RIGHTB',None),206:('MIDB',3),207:('LENB',1),208:('ROUNDUP',2),209:('ROUNDDOWN',2),
210:('ASC',1),211:('DBCS',1),212:('RANK',None),216:('ADDRESS',None),217:('DAYS360',None),218:('TODAY',0),219:('VDB',None),
220:('ELSE',0),221:('ELSE.IF',1),222:('END.IF',0),223:('FOR.CELL',None),227:('MEDIAN',None),228:('SUMPRODUCT',None),229:('SINH',1),
230:('COSH',1),231:('TANH',1),232:('ASINH',1),233:('ACOSH',1),234:('ATANH',1),235:('DGET',3),236:('CREATE.OBJECT',None),
237:('VOLATILE',None),238:('LAST.ERROR',0),239:('CUSTOM.UNDO',None),240:('CUSTOM.REPEAT',None),241:('FORMULA.CONVERT',None),
242:('GET.LINK.INFO',None),243:('TEXT.BOX',None),244:('INFO',1),245:('GROUP',0),246:('GET.OBJECT',None),247:('DB',None),248:('PAUSE',None),
251:('RESUME',None),252:('FREQUENCY',2),253:('ADD.TOOLBAR',None),254:('DELETE.TOOLBAR',None),255:('USER',None),256:('RESET.TOOLBAR',1),
257:('EVALUATE',1),258:('GET.TOOLBAR',None),259:('GET.TOOL',None),260:('SPELLING.CHECK',None),261:('ERROR.TYPE',1),262:('APP.TITLE',None),
263:('WINDOW.TITLE',None),264:('SAVE.TOOLBAR',None),265:('ENABLE.TOOL',3),266:('PRESS.TOOL',3),267:('REGISTER.ID',None),
268:('GET.WORKBOOK',None),269:('AVEDEV',None),270:('BETADIST',None),271:('GAMMALN',1),272:('BETAINV',None),273:('BINOMDIST',4),
274:('CHIDIST',2),275:('CHIINV',2),276:('COMBIN',2),277:('CONFIDENCE',3),278:('CRITBINOM',3),279:('EVEN',1),280:('EXPONDIST',3),
281:('FDIST',3),282:('FINV',3),283:('FISHER',1),284:('FISHERINV',1),285:('FLOOR',2),286:('GAMMADIST',4),287:('GAMMAINV',3),288:('CEILING',2),
289:('HYPGEOMDIST',4),290:('LOGNORMDIST',3),291:('LOGINV',3),292:('NEGBINOMDIST',3),293:('NORMDIST',4),294:('NORMSDIST',1),
295:('NORMINV',3),296:('NORMSINV',1),297:('STANDARDIZE',3),298:('ODD',1),299:('PERMUT',2),300:('POISSON',3),301:('TDIST',3),
302:('WEIBULL',4),303:('SUMXMY2',2),304:('SUMX2MY2',2),305:('SUMX2PY2',2),306:('CHITEST',2),307:('CORREL',2),308:('COVAR',2),
309:('FORECAST',3),310:('FTEST',2),311:('INTERCEPT',2),312:('PEARSON',2),313:('RSQ',2),314:('STEYX',2),315:('SLOPE',2),316:('TTEST',4),
317:('PROB',None),318:('DEVSQ',None),319:('GEOMEAN',None),320:('HARMEAN',None),321:('SUMSQ',None),322:('KURT',None),323:('SKEW',None),
324:('ZTEST',None),325:('LARGE',2),326:('SMALL',2),327:('QUARTILE',2),328:('PERCENTILE',2),329:('PERCENTRANK',None),330:('MODE',None),
331:('TRIMMEAN',2),332:('TINV',2),334:('MOVIE.COMMAND',None),335:('GET.MOVIE',None),336:('CONCATENATE',None),337:('POWER',2),
338:('PIVOT.ADD.DATA',None),339:('GET.PIVOT.TABLE',None),340:('GET.PIVOT.FIELD',None),341:('GET.PIVOT.ITEM',None),342:('RADIANS',1),
343:('DEGREES',1),344:('SUBTOTAL',None),345:('SUMIF',None),346:('COUNTIF',2),347:('COUNTBLANK',1),348:('SCENARIO.GET',None),
349:('OPTIONS.LISTS.GET',1),350:('ISPMT',4),351:('DATEDIF',3),352:('DATESTRING',1),353:('NUMBERSTRING',2),354:('ROMAN',None),
355:('OPEN.DIALOG',None),356:('SAVE.DIALOG',None),357:('VIEW.GET',None),358:('GETPIVOTDATA',None),359:('HYPERLINK',None),360:('PHONETIC',1),
361:('AVERAGEA',None),362:('MAXA',None),363:('MINA',None),364:('STDEVPA',None),365:('VARPA',None),366:('STDEVA',None),367:('VARA',None),
368:('BAHTTEXT',1),369:('THAIDAYOFWEEK',1),370:('THAIDIGIT',1),371:('THAIMONTHOFYEAR',1),372:('THAINUMSOUND',1),373:('THAINUMSTRING',1),
374:('THAISTRINGLENGTH',1),375:('ISTHAIDIGIT',1),376:('ROUNDBAHTDOWN',1),377:('ROUNDBAHTUP',1),378:('THAIYEAR',1),379:('RTD',None),
# BIFF12 extension
380:('CUBEVALUE',None),381:('CUBEMEMBER',None),382:('CUBEMEMBERPROPERTY',3),383:('CUBERANKEDMEMBER',None),384:('HEX2BIN',None),
385:('HEX2DEC',1),386:('HEX2OCT',None),387:('DEC2BIN',None),388:('DEC2HEX',None),389:('DEC2OCT',None),390:('OCT2BIN',None),
391:('OCT2HEX',None),392:('OCT2DEC',1),393:('BIN2DEC',1),394:('BIN2OCT',None),395:('BIN2HEX',None),396:('IMSUB',2),397:('IMDIV',2),
398:('IMPOWER',2),399:('IMABS',1),400:('IMSQRT',1),401:('IMLN',1),402:('IMLOG2',1),403:('IMLOG10',1),404:('IMSIN',1),405:('IMCOS',1),
406:('IMEXP',1),407:('IMARGUMENT',1),408:('IMCONJUGATE',1),409:('IMAGINARY',1),410:('IMREAL',1),411:('COMPLEX',None),412:('IMSUM',None),
413:('IMPRODUCT',None),414:('SERIESSUM',4),415:('FACTDOUBLE',1),416:('SQRTPI',1),417:('QUOTIENT',2),418:('DELTA',None),419:('GESTEP',None),
420:('ISEVEN',1),421:('ISODD',1),422:('MROUND',2),423:('ERF',None),424:('ERFC',1),425:('BESSELJ',2),426:('BESSELK',2),427:('BESSELY',2),
428:('BESSELI',2),429:('XIRR',None),430:('XNPV',3),431:('PRICEMAT',None),432:('YIELDMAT',None),433:('INTRATE',None),434:('RECEIVED',None),
435:('DISC',None),436:('PRICEDISC',None),437:('YIELDDISC',None),438:('TBILLEQ',3),439:('TBILLPRICE',3),440:('TBILLYIELD',3),
441:('PRICE',None),442:('YIELD',None),443:('DOLLARDE',2),444:('DOLLARFR',2),445:('NOMINAL',2),446:('EFFECT',2),447:('CUMPRINC',6),
448:('CUMIPMT',6),449:('EDATE',2),450:('EOMONTH',2),451:('YEARFRAC',None),452:('COUPDAYBS',None),453:('COUPDAYS',None),454:('COUPDAYSNC',None),
455:('COUPNCD',None),456:('COUPNUM',None),457:('COUPPCD',None),458:('DURATION',None),459:('MDURATION',None),460:('ODDLPRICE',None),
461:('ODDLYIELD',None),462:('ODDFPRICE',None),463:('ODDFYIELD',None),464:('RANDBETWEEN',2),465:('WEEKNUM',None),466:('AMORDEGRC',None),
467:('AMORLINC',None),468:('CONVERT',3),469:('ACCRINT',None),470:('ACCRINTM',None),471:('WORKDAY',None),472:('NETWORKDAYS',None),
473:('GCD',None),474:('MULTINOMIAL',None),475:('LCM',None),476:('FVSCHEDULE',2),477:('CUBEKPIMEMBER',None),478:('CUBESET',None),
479:('CUBESETCOUNT',1),480:('IFERROR',2),481:('COUNTIFS',None),482:('SUMIFS',None),483:('AVERAGEIF',None),484:('AVERAGEIFS',None),
}
ERRS = {0x00:'#NULL!',0x07:'#DIV/0!',0x0F:'#VALUE!',0x17:'#REF!',0x1D:'#NAME?',0x24:'#NUM!',0x2A:'#N/A',0x2B:'#GETTING_DATA'}
MAXROW = 0xFFFFF   # 1048575 (0-based)
MAXCOL = 0x3FFF    # 16383

class DecodeError(Exception): pass

def colL(c):
    s=''; c+=1
    while c: c,r=divmod(c-1,26); s=chr(65+r)+s
    return s

def fmt_num(x):
    if x==int(x) and abs(x)<1e15: return str(int(x))
    s=repr(x)
    return s

_plain_sheet = re.compile(r'^[A-Za-z_\\][A-Za-z0-9_.]*$')
def sheet_q(name):
    """Quote a sheet name the way Excel does."""
    if _plain_sheet.match(name) and not re.match(r'^[A-Za-z]{1,3}\d+$',name) and not re.match(r'^R\d*C\d*$',name,re.I):
        return name
    return "'"+name.replace("'","''")+"'"

# ----------------------------------------------------------------------------------------------
class WorkbookCtx:
    def __init__(self, xlsb_dir):
        wb=open(os.path.join(xlsb_dir,'xl','workbook.bin'),'rb').read()
        self.sheets=[]; self.names=[]; self.sups=[]; self.xti=[]
        for rid,body in records(wb):
            if rid==0x9C:
                hs,itab=struct.unpack_from('<II',body,0); rel,off=wstr(body,8); name,off=wstr(body,off)
                self.sheets.append(name)
            elif rid==0x27:
                flags,=struct.unpack_from('<I',body,0); itab,=struct.unpack_from('<i',body,5); name,off=wstr(body,9)
                cce,=struct.unpack_from('<I',body,off); rgce=body[off+4:off+4+cce]; off2=off+4+cce
                cb,=struct.unpack_from('<I',body,off2); rgcb=body[off2+4:off2+4+cb]
                self.names.append(dict(name=name,flags=flags,itab=itab,rgce=rgce,rgcb=rgcb))
            elif rid in (0x163,0x165,0x166,0x167,0x168):   # BrtSupBookSrc / Self / Same / Tabs / Addin
                self.sups.append(rid)
            elif rid==0x16A:
                cxti,=struct.unpack_from('<I',body,0)
                self.xti=[struct.unpack_from('<iii',body,4+12*i) for i in range(cxti)]
    def sheet_prefix(self, ixti):
        ext,first,last=self.xti[ixti]
        if self.sups[ext]==0x165:   # self
            if first==-1 and last==-1: return '#REF!!'
            if first==-2: return ''        # workbook-level (no sheet)
            a=sheet_q(self.sheets[first]) if first<len(self.sheets) else f'#SHEET{first}'
            if first!=last:
                b=self.sheets[last] if last<len(self.sheets) else f'#SHEET{last}'
                inner=(self.sheets[first]+':'+b)
                return "'"+inner.replace("'","''")+"'!" if not _plain_sheet.match(inner.replace(':','')) else inner+'!'
            return a+'!'
        return f'[{ext}]#EXT({first},{last})!'
    def name_text(self, idx):   # 1-based
        if 1<=idx<=len(self.names): return self.names[idx-1]['name']
        return f'#NAME{idx}'

# ----------------------------------------------------------------------------------------------
class RgceDecoder:
    """Decode one rgce/rgcb pair into an A1-style formula string.
    base_row/base_col: the cell the formula belongs to (used for PtgRefN/PtgAreaN relative offsets)."""
    def __init__(self, ctx, base_row=0, base_col=0):
        self.ctx=ctx; self.br=base_row; self.bc=base_col
        self.tokens=[]   # list of (ptg, description) for diagnostics

    # --- reference helpers
    def ref(self, row, colfield, relative_encoding):
        col=colfield&0x3FFF; fcol=bool(colfield&0x4000); frow=bool(colfield&0x8000)
        if relative_encoding:
            if frow: row=(row+self.br)&MAXROW
            if fcol:
                if col&0x2000: col-=0x4000
                col=(col+self.bc)&MAXCOL
        return ('' if fcol else '$')+colL(col)+('' if frow else '$')+str(row+1)
    def area(self, r1, r2, c1f, c2f, relative_encoding):
        c1=c1f&0x3FFF; c2=c2f&0x3FFF
        fc1=bool(c1f&0x4000); fr1=bool(c1f&0x8000); fc2=bool(c2f&0x4000); fr2=bool(c2f&0x8000)
        if relative_encoding:
            if fr1: r1=(r1+self.br)&MAXROW
            if fr2: r2=(r2+self.br)&MAXROW
            if fc1:
                if c1&0x2000: c1-=0x4000
                c1=(c1+self.bc)&MAXCOL
            if fc2:
                if c2&0x2000: c2-=0x4000
                c2=(c2+self.bc)&MAXCOL
        if r1==0 and r2==MAXROW:      # whole column(s)
            return ('' if fc1 else '$')+colL(c1)+':'+('' if fc2 else '$')+colL(c2)
        if c1==0 and c2==MAXCOL:      # whole row(s)
            return ('' if fr1 else '$')+str(r1+1)+':'+('' if fr2 else '$')+str(r2+1)
        a=('' if fc1 else '$')+colL(c1)+('' if fr1 else '$')+str(r1+1)
        b=('' if fc2 else '$')+colL(c2)+('' if fr2 else '$')+str(r2+1)
        return a+':'+b

    def decode(self, rgce, rgcb):
        st=[]; i=0; j=0; n=len(rgce); exp=None
        BIN={0x03:'+',0x04:'-',0x05:'*',0x06:'/',0x07:'^',0x08:'&',0x09:'<',0x0A:'<=',0x0B:'=',0x0C:'>=',0x0D:'>',0x0E:'<>',0x0F:' ',0x10:',',0x11:':'}
        def pop():
            if not st: raise DecodeError('stack underflow')
            return st.pop()
        while i<n:
            ptg=rgce[i]; i+=1
            base=ptg&0x1F; cls=ptg&0x60
            self.tokens.append(ptg)
            if ptg==0x01:   # PtgExp
                row,=struct.unpack_from('<I',rgce,i); i+=4
                col,=struct.unpack_from('<I',rgcb,j); j+=4      # PtgExtraCol
                exp=(row,col); st.append(f'#EXP({row},{col})')
            elif ptg==0x02:  # PtgTbl
                i+=8; st.append('#TBL')
            elif ptg in BIN:
                b=pop(); a=pop(); st.append(a+BIN[ptg]+b)
            elif ptg==0x12: st.append('+'+pop())
            elif ptg==0x13: st.append('-'+pop())
            elif ptg==0x14: st.append(pop()+'%')
            elif ptg==0x15: st.append('('+pop()+')')
            elif ptg==0x16: st.append('')
            elif ptg==0x17:   # PtgStr: cch is 2 bytes in BIFF12
                cch,=struct.unpack_from('<H',rgce,i); s=rgce[i+2:i+2+2*cch].decode('utf-16le'); i+=2+2*cch
                st.append('"'+s.replace('"','""')+'"')
            elif ptg==0x18:  # PtgElf*/PtgList (structured refs etc.)
                raise DecodeError('PtgList/Elf 0x18 unsupported')
            elif ptg==0x19:  # PtgAttr
                sub=rgce[i]; i+=1
                if sub==0x04:   # AttrChoose
                    c,=struct.unpack_from('<H',rgce,i); i+=2+2*(c+1)
                elif sub==0x10:  # AttrSum
                    i+=2; st.append('SUM('+pop()+')')
                elif sub in (0x40,0x41):  # AttrSpace
                    typ=rgce[i]; cch=rgce[i+1]; i+=2
                    # spaces are irrelevant after normalisation; keep a marker for leading spaces only
                else:            # semi(0x01), if(0x02), goto(0x08), baxcel(0x20), iferror(0x80)
                    i+=2
            elif ptg==0x1C: st.append(ERRS.get(rgce[i],f'#ERR{rgce[i]}')); i+=1
            elif ptg==0x1D: st.append('TRUE' if rgce[i] else 'FALSE'); i+=1
            elif ptg==0x1E: st.append(str(struct.unpack_from('<H',rgce,i)[0])); i+=2
            elif ptg==0x1F: st.append(fmt_num(struct.unpack_from('<d',rgce,i)[0])); i+=8
            elif base==0x00 and cls:   # PtgArray 0x20/0x40/0x60
                i+=14
                cols,rows=struct.unpack_from('<II',rgcb,j); j+=8
                vals=[]
                for k in range(rows*cols):
                    t=rgcb[j]; j+=1
                    if t==0x00: j+=8; vals.append('')
                    elif t==0x01: vals.append(fmt_num(struct.unpack_from('<d',rgcb,j)[0])); j+=8
                    elif t==0x02: s,j=wstr(rgcb,j); vals.append('"'+s.replace('"','""')+'"')
                    elif t==0x04: vals.append('TRUE' if rgcb[j] else 'FALSE'); j+=1
                    elif t==0x10: vals.append(ERRS.get(rgcb[j],'#ERR')); j+=1
                    else: raise DecodeError(f'SerAr type {t}')
                st.append('{'+';'.join(','.join(vals[r*cols:(r+1)*cols]) for r in range(rows))+'}')
            elif base==0x01 and cls:   # PtgFunc
                iftab,=struct.unpack_from('<H',rgce,i); i+=2
                name,na=_FT.get(iftab,(f'FUNC{iftab}',None))
                if na is None: raise DecodeError(f'PtgFunc {iftab} ({name}) has no fixed arity')
                args=[pop() for _ in range(na)][::-1]
                st.append(name+'('+','.join(args)+')')
            elif base==0x02 and cls:   # PtgFuncVar
                cp=rgce[i]; tab,=struct.unpack_from('<H',rgce,i+1); i+=3
                iftab=tab&0x7FFF
                args=[pop() for _ in range(cp)][::-1]
                if iftab==255:
                    fname=args[0]; args=args[1:]
                    fname=re.sub(r'^_xlfn\.(_xlws\.)?','',fname)
                else:
                    fname=_FT.get(iftab,(f'FUNC{iftab}',None))[0]
                st.append(fname+'('+','.join(args)+')')
            elif base==0x03 and cls:   # PtgName
                idx,=struct.unpack_from('<I',rgce,i); i+=4
                nm=self.ctx.name_text(idx)
                st.append(re.sub(r'^_xlpm\.','',nm))
            elif base==0x04 and cls:   # PtgRef
                row,=struct.unpack_from('<I',rgce,i); cf,=struct.unpack_from('<H',rgce,i+4); i+=6
                st.append(self.ref(row,cf,False))
            elif base==0x05 and cls:   # PtgArea
                r1,r2=struct.unpack_from('<II',rgce,i); c1,c2=struct.unpack_from('<HH',rgce,i+8); i+=12
                st.append(self.area(r1,r2,c1,c2,False))
            elif base==0x06 and cls:   # PtgMemArea
                i+=6
                cnt,=struct.unpack_from('<I',rgcb,j); j+=4+16*cnt
            elif base==0x07 and cls: i+=6      # PtgMemErr
            elif base==0x08 and cls: i+=6      # PtgMemNoMem
            elif base==0x09 and cls: i+=2      # PtgMemFunc
            elif base==0x0A and cls: i+=6; st.append('#REF!')     # PtgRefErr
            elif base==0x0B and cls: i+=12; st.append('#REF!')    # PtgAreaErr
            elif base==0x0C and cls:   # PtgRefN
                row,=struct.unpack_from('<I',rgce,i); cf,=struct.unpack_from('<H',rgce,i+4); i+=6
                st.append(self.ref(row,cf,True))
            elif base==0x0D and cls:   # PtgAreaN
                r1,r2=struct.unpack_from('<II',rgce,i); c1,c2=struct.unpack_from('<HH',rgce,i+8); i+=12
                st.append(self.area(r1,r2,c1,c2,True))
            elif base==0x19 and cls:   # PtgNameX
                ixti,=struct.unpack_from('<H',rgce,i); idx,=struct.unpack_from('<I',rgce,i+2); i+=6
                nm=self.ctx.name_text(idx)
                pre=self.ctx.sheet_prefix(ixti)
                st.append((pre if pre and not pre.startswith('#REF') else '')+re.sub(r'^_xlpm\.','',nm))
            elif base==0x1A and cls:   # PtgRef3d
                ixti,=struct.unpack_from('<H',rgce,i); row,=struct.unpack_from('<I',rgce,i+2); cf,=struct.unpack_from('<H',rgce,i+6); i+=8
                st.append(self.ctx.sheet_prefix(ixti)+self.ref(row,cf,False))
            elif base==0x1B and cls:   # PtgArea3d
                ixti,=struct.unpack_from('<H',rgce,i); r1,r2=struct.unpack_from('<II',rgce,i+2); c1,c2=struct.unpack_from('<HH',rgce,i+10); i+=14
                st.append(self.ctx.sheet_prefix(ixti)+self.area(r1,r2,c1,c2,False))
            elif base==0x1C and cls:   # PtgRefErr3d
                ixti,=struct.unpack_from('<H',rgce,i); i+=8; st.append(self.ctx.sheet_prefix(ixti)+'#REF!')
            elif base==0x1D and cls:   # PtgAreaErr3d
                ixti,=struct.unpack_from('<H',rgce,i); i+=14; st.append(self.ctx.sheet_prefix(ixti)+'#REF!')
            else:
                raise DecodeError(f'unknown ptg 0x{ptg:02x} at {i-1}')
        if len(st)!=1: raise DecodeError(f'stack has {len(st)} items at end: {st}')
        return st[0], exp

# ----------------------------------------------------------------------------------------------
class SheetFormulas:
    """All formula cells of one worksheet part plus BrtShrFmla/BrtArrFmla/BrtCellMeta bookkeeping."""
    FORMULA={8:'str',9:'num',10:'bool',11:'err'}
    def __init__(self, path):
        data=open(path,'rb').read()
        self.cells={}; self.shared={}; self.arrays={}; self.meta_before=set(); self.dim=None
        row=None; last=None; pending_meta=False
        for rid,body in records(data):
            if rid==0x00:
                row=struct.unpack_from('<I',body,0)[0]
            elif rid==0x94:   # BrtWsDim
                self.dim=struct.unpack_from('<IIII',body,0)
            elif rid==0x31:   # BrtCellMeta (icmb) applies to the next cell
                pending_meta=True
            elif rid in self.FORMULA:
                col=struct.unpack_from('<I',body,0)[0]; off=8
                if rid==8: v,off=wstr(body,off)
                elif rid==9: v=struct.unpack_from('<d',body,off)[0]; off+=8
                else: v=body[off]; off+=1
                g,=struct.unpack_from('<H',body,off); off+=2
                cce,=struct.unpack_from('<I',body,off); rgce=body[off+4:off+4+cce]; off2=off+4+cce
                cb,=struct.unpack_from('<I',body,off2); rgcb=body[off2+4:off2+4+cb]
                self.cells[(row,col)]=dict(kind=self.FORMULA[rid],value=v,grbit=g,rgce=rgce,rgcb=rgcb,meta=pending_meta)
                if pending_meta: self.meta_before.add((row,col))
                last=(row,col); pending_meta=False
            elif rid in (1,2,3,4,5,6,7):
                pending_meta=False
            elif rid in (0x1AA,0x1AB):
                rf=struct.unpack_from('<IIII',body,0)   # rowFirst,rowLast,colFirst,colLast
                if rid==0x1AA:
                    flags=body[16]; off=17
                else:
                    flags=None; off=16
                cce,=struct.unpack_from('<I',body,off); rgce=body[off+4:off+4+cce]; off2=off+4+cce
                cb,=struct.unpack_from('<I',body,off2); rgcb=body[off2+4:off2+4+cb]
                rec=dict(rfx=rf,flags=flags,rgce=rgce,rgcb=rgcb,after=last)
                (self.arrays if rid==0x1AA else self.shared)[(rf[0],rf[2])]=rec

    def decode_cell(self, ctx, row, col):
        """Return (text, info) where info describes shared/array resolution. Raises DecodeError."""
        c=self.cells[(row,col)]
        d=RgceDecoder(ctx,row,col)
        text,exp=d.decode(c['rgce'],c['rgcb'])
        info=dict(kind=c['kind'],grbit=c['grbit'],meta=c['meta'],tokens=d.tokens,storage='cell')
        if exp is not None:
            key=exp
            if key in self.arrays and self._in(self.arrays[key]['rfx'],row,col):
                rec=self.arrays[key]; info['storage']='array'; info['rfx']=rec['rfx']; info['arr_flags']=rec['flags']
            elif key in self.shared and self._in(self.shared[key]['rfx'],row,col):
                rec=self.shared[key]; info['storage']='shared'; info['rfx']=rec['rfx']
            else:
                raise DecodeError(f'PtgExp -> no BrtArrFmla/BrtShrFmla master at {key}')
            d2=RgceDecoder(ctx,row,col)
            text,exp2=d2.decode(rec['rgce'],rec['rgcb'])
            info['tokens']=d2.tokens
            if exp2 is not None: raise DecodeError('nested PtgExp')
        return '='+text, info
    @staticmethod
    def _in(rfx,row,col): return rfx[0]<=row<=rfx[1] and rfx[2]<=col<=rfx[3]

# ----------------------------------------------------------------------------------------------
def a1_to_rc(ref):
    m=re.fullmatch(r'\$?([A-Z]{1,3})\$?(\d+)',ref.upper())
    if not m: raise ValueError(ref)
    col=0
    for ch in m.group(1): col=col*26+(ord(ch)-64)
    return int(m.group(2))-1, col-1

def normalise(f):
    """Normalise formula text for comparison: strip leading '=', '@', '$', spaces and sheet quotes outside
    string literals; strip _xlfn./_xlws./_xlpm. prefixes; upper-case everything outside string literals."""
    f=f.strip()
    if f.startswith('='): f=f[1:]
    out=[]; i=0; n=len(f)
    while i<n:
        if f[i]=='"':
            j=i+1
            while j<n:
                if f[j]=='"':
                    if j+1<n and f[j+1]=='"': j+=2; continue
                    break
                j+=1
            out.append(f[i:j+1]); i=j+1
        else:
            j=i
            while j<n and f[j]!='"': j+=1
            seg=f[i:j]
            seg=re.sub(r'_xlfn\.|_xlws\.|_xlpm\.','',seg)
            seg=seg.replace('@','').replace('$','').replace(' ','').replace("'",'').upper()
            out.append(seg); i=j
    return ''.join(out)

TOK=re.compile(r'"(?:[^"]|"")*"|[A-Z_][A-Z0-9_.]*!|[A-Z]{1,3}\d+:[A-Z]{1,3}\d+|[A-Z]{1,3}:[A-Z]{1,3}|[A-Z]{1,3}\d+|\d+:\d+|\d+(?:\.\d+)?(?:E[+-]?\d+)?|[A-Z_][A-Z0-9_.]*|<>|<=|>=|.',re.S)
def tokens(s): return TOK.findall(s)
def classify(a,b):
    """Describe how normalised formula a (sheet) differs from b (bank)."""
    ta,tb=tokens(a),tokens(b)
    sm=difflib.SequenceMatcher(None,ta,tb,autojunk=False)
    kinds=set(); pairs=[]
    for op,i1,i2,j1,j2 in sm.get_opcodes():
        if op=='equal': continue
        x=' '.join(ta[i1:i2]); y=' '.join(tb[j1:j2]); pairs.append((x,y))
        for t in ta[i1:i2]+tb[j1:j2]:
            if t.startswith('"'): kinds.add('string-constant')
            elif re.fullmatch(r'[A-Z]{1,3}\d+(:[A-Z]{1,3}\d+)?|[A-Z]{1,3}:[A-Z]{1,3}|\d+:\d+',t): kinds.add('range/ref')
            elif t.endswith('!'): kinds.add('sheet')
            elif re.fullmatch(r'\d+(\.\d+)?(E[+-]?\d+)?',t): kinds.add('number')
            elif re.fullmatch(r'[A-Z_][A-Z0-9_.]*',t): kinds.add('function/name')
            else: kinds.add('operator/punct')
    # refine range differences: same shape, only row bounds differ?
    if kinds=={'range/ref'}:
        only_rows=True
        for x,y in pairs:
            rx=re.findall(r'[A-Z]{1,3}',x); ry=re.findall(r'[A-Z]{1,3}',y)
            if rx!=ry: only_rows=False
        kinds={'range-bound(rows)' if only_rows else 'range/ref'}
    return ','.join(sorted(kinds)), pairs

def load_bank(path):
    rows=[]
    for line in open(path,encoding='utf-8'):
        m=re.match(r'^(\d+) (\[.*\])$',line.rstrip('\n'))
        if m and int(m.group(1))>1:
            r=ast.literal_eval(m.group(2)); rows.append((int(m.group(1)),r))
    return rows

def main():
    ctx=WorkbookCtx(os.path.join(SCR,'xlsb'))
    fv=SheetFormulas(os.path.join(SCR,'xlsb','xl','worksheets','sheet3.bin'))
    bank=load_bank(os.path.join(SCR,'ctx','formulabank.txt'))
    results=[]; tokhist=collections.Counter(); status=collections.Counter()
    for bankrow,r in bank:
        ref=r[0]; bformula=r[1]; isarr=r[2]
        rec=dict(bankrow=bankrow,ref=ref,bank=bformula,isarray=isarr,src=r[4] if len(r)>4 else '')
        try:
            row,col=a1_to_rc(ref)
        except ValueError:
            rec.update(status='BAD_REF'); results.append(rec); status['BAD_REF']+=1; continue
        if (row,col) not in fv.cells:
            rec.update(status='NO_FORMULA_ON_SHEET'); results.append(rec); status['NO_FORMULA_ON_SHEET']+=1; continue
        try:
            text,info=fv.decode_cell(ctx,row,col)
        except DecodeError as e:
            rec.update(status='UNDECODABLE',error=str(e)); results.append(rec); status['UNDECODABLE']+=1; continue
        for t in info['tokens']: tokhist[t]+=1
        rec.update(sheet=text,storage=info['storage'],kind=info['kind'],grbit=info['grbit'],meta=info['meta'],rfx=info.get('rfx'),arr_flags=info.get('arr_flags'))
        na,nb=normalise(text),normalise(bformula)
        rec['norm_sheet']=na; rec['norm_bank']=nb
        if na==nb:
            rec['status']='MATCH'
        else:
            rec['status']='MISMATCH'; rec['difftype'],rec['diffs']=classify(na,nb)
        status[rec['status']]+=1
        results.append(rec)
    # ---- write compare.txt
    with open(os.path.join(HERE,'compare.txt'),'w',encoding='utf-8') as f:
        f.write('# FormulaBank (column B) vs decoded Fund View (sheet3.bin) formulas - one block per tracked cell\n')
        f.write('# status: MATCH | MISMATCH(difftype) | UNDECODABLE | NO_FORMULA_ON_SHEET | BAD_REF\n')
        f.write(f'# totals: {dict(status)}\n\n')
        for rec in results:
            f.write(f"[{rec['bankrow']}] {rec['ref']}  {rec['status']}"+(f"  ({rec['difftype']})" if rec.get('difftype') else '')+
                    f"  storage={rec.get('storage')} kind={rec.get('kind')} grbit={rec.get('grbit')} cellmeta={rec.get('meta')} isarray={rec['isarray']}\n")
            f.write(f"  bank : {rec['bank']}\n")
            if 'sheet' in rec: f.write(f"  sheet: {rec['sheet']}\n")
            if rec.get('error'): f.write(f"  error: {rec['error']}\n")
            for x,y in rec.get('diffs',[]): f.write(f"  diff : sheet[{x}]  bank[{y}]\n")
            f.write('\n')
    summary=dict(status=dict(status),tokens={f'0x{k:02x}':v for k,v in sorted(tokhist.items())},
                 n_formula_cells=len(fv.cells),n_shared=len(fv.shared),n_array=len(fv.arrays),n_cellmeta=len(fv.meta_before),dim=fv.dim)
    json.dump(dict(summary=summary,results=results),open(os.path.join(HERE,'results.json'),'w'),default=str,indent=0)
    print(json.dumps(summary,indent=1))
    return ctx,fv,bank,results

if __name__=='__main__':
    main()
