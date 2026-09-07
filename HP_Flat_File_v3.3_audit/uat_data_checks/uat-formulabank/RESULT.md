# UAT data check - FormulaBank vs live "Fund View" formulas (HP_Flat_File_Test_v3.3.xlsb)

Folder: `scratchpad/agents/uat-formulabank/` - `decode.py` (BIFF12 rgce/rgcb decoder + comparison), `compare.txt` (one block per tracked cell: status, bank text, decoded sheet text, token-level diff), `results.json`, `source_tables.json`, `examples.md`, `wb_dump.txt`.

## Method
- Parsed `xl/workbook.bin` (BrtBundleSh sheet order = sheet index used by XTI; 1688 BrtName records, 1-based index used by PtgName/PtgNameX; BrtExternSheet XTI table with externalLink 0 = BrtSupSelf) and `xl/worksheets/sheet3.bin` (= "Fund View": 5657 formula cells, 99 BrtShrFmla, 140 BrtArrFmla, 133 BrtCellMeta).
- Implemented a stack-based decoder for every Ptg present: PtgExp(+PtgExtraCol), binary ops, PtgUminus, PtgParen, PtgStr (2-byte cch), PtgAttr (IfError 0x80, GoTo, If, Choose, Sum, Space, Semi), PtgErr/Bool/Int/Num, PtgArray, PtgFunc (BIFF8/BIFF12 Ftab incl. 480 = IFERROR, 48 = TEXT, 228 = SUMPRODUCT), PtgFuncVar (iftab 255 = user/future function named by the preceding PtgName, e.g. `_xlfn.XMATCH`, `_xlfn.LET`, `_xlpm.n`), PtgName/NameX, PtgRef/Area, PtgRefN/AreaN (shared formulas, relative to base cell), PtgRef3d/Area3d, PtgMem*/RefErr/AreaErr. PtgExp cells are resolved through the BrtArrFmla/BrtShrFmla record whose master cell matches.
- Tokens actually seen in the 1729 tracked cells: 0x05 Mul, 0x08 Concat, 0x0B Eq, 0x13 Uminus, 0x15 Paren, 0x17 Str, 0x19 Attr, 0x1E Int, 0x22/0x42/0x62 FuncVar, 0x23/0x43/0x63 Name, 0x41/0x61 Func, 0x44/0x64 Ref, 0x3B/0x7B Area3d (plus 0x01 Exp in the 7 array cells).
- Normalisation before comparing: strip leading `=`, drop `@`, `$`, spaces and sheet quotes outside string literals, strip `_xlfn.`/`_xlws.`/`_xlpm.` prefixes, upper-case outside string literals. The decoder was validated by the 504 company_breakdown cells (and the position/valuation_history/company_nav cells) decoding to text identical to FormulaBank apart from the row bounds reported below.
- Source-table extents (BrtWsDim + last non-blank key cell), header rows (row 1 via sharedStrings.bin) and Fund View constants (C5, C6, BJ22:BJ24) were read directly from the sheet parts. Excel is not available; nothing was executed live.

## Headline numbers
| Item | Value |
|---|---|
| Tracked cells in FormulaBank | 1729 (rows 2..1730, no duplicate CellRefs, all valid A1 refs inside B1:BL200) |
| Present on Fund View and holding a formula | 1729 / 1729 (1252 string-result, 477 numeric-result) |
| Decoded without error | **1729** (0 undecodable, 0 stack errors) |
| Exact match after normalisation | **504** (all 504 company_breakdown cells) |
| Mismatch | **1225** - every one of them is a *range upper-bound (row) difference only*; no string-constant, function, sheet, operator, `$`-anchoring or `@` differences exist |
| Cells not decodable | 0 |
| AT117:AZ117 stored as array formulas | **Yes** - PtgExp + single-cell BrtArrFmla (0x1AA), no BrtCellMeta (i.e. legacy CSE array, not dynamic array); consistent with IsArray=Y / `.FormulaArray` |

## Range-bound drift per source sheet (the only kind of text difference found)
| Source sheet (cells) | Live sheet formula bound | FormulaBank bound | Actual last data row (BrtWsDim = last non-blank key) | Named KeyRange bound | Effect today / after Reset |
|---|---|---|---|---|---|
| position (27) | `$A$2:$AG$819` | `$AG$820` | 819 | PositionKey A2:A819 | Bank adds 1 blank row - harmless |
| fund (10) | `$A$2:$N$819` | `$N$820` | 819 | FundKey A2:A819 | Bank adds 1 blank row - harmless |
| company (239 + 40 AS cells) | `$A$2:$L$4979` | `$L$4984` | 4979 | CompanyKey D2:D4979 | Bank adds 5 blank rows - harmless |
| company_nav (280) | `$A$2:$J$15134` | `$J$15148` | 15134 | CompanyNavKey D2:D15134 | Bank adds 14 blank rows - harmless (note VBA hard-codes 15131) |
| position_company (286 INDEX/XMATCH cells) | `$A$2:$Q$87982` | `$Q$87989` | 87982 (col E lookup_key non-blank only to row 8322) | PositionCompanyKey E2:E87982 | Bank adds 7 blank rows - harmless |
| position_company LET/SUMPRODUCT array cells AT117:AZ117 (7) | `$A$2:$A$87974` (also $C, $P) | `$A$87981` | 87982 | n/a | **Live formulas miss data rows 87975-87982; bank still misses 87982.** Rows 87975-87982 are `TemplateU Flex` source_row 158/159 components, and row 117 filters on `ROW()=117`, so no numeric effect for any fund today - but the range is stale relative to the table |
| cashflow (250) | `$A$2:$AA$33497` | `$AA$33498` | 33498 | CashflowKey D2:D33498 | **Live formulas cannot see cashflow row 33498 = `Yorktown IX|63`**, which Fund View row 63 cells AE63/AH63/AI63/AK63/AL63/AN63 look up when C6 = Yorktown IX. After a Reset the bank text (33498) makes that row visible: numbers on row 63 change silently for that fund (row currently has blank cashflow_date, so today the change would be blank->row values) |
| valuation_history (126: 121 cells) | `$A$2:$K$32608` | `$K$32609` | 32608 | QuarterKey K2:K32608 | Bank adds 1 blank row - harmless |
| valuation_history (5 cells L23, AA23, L29, AA29, AD29) | `$A$2:$K$32607` | `$K$32609` | 32608 | QuarterKey K2:K32608 | Those 5 cells cannot see row 32608 = `Yorktown IX|2034-12-31`; rows 23/29 hold early quarters so no practical effect, but the 5 cells are out of step with their 121 siblings |
| company_breakdown (504) | `$A$2:$U$1459` | `$U$1459` | 1459 | CompanyBreakdownKey D2:D1459 | identical |

Interpretation: FormulaBank was captured against tables that were larger than they are now (bank bound > live bound everywhere except company_breakdown), then rows were deleted from the tables (Excel shrank the live references, the static bank text did not follow), and in cashflow / valuation_history / position_company rows were later re-added at the bottom beyond the live references. A Reset (RestoreYellowFormulas / RestoreOverriddenCellsOnly, which rewrite from column B) therefore does change the lookup ranges of 1225 cells, in most cases only by re-adding blank rows.

## Checks
| id | name | feature | method | result | details |
|---|---|---|---|---|---|
| FB-01 | Tracked cells exist and hold formulas | FormulaBank col A vs sheet3.bin cell records | scan BrtFmla* records; match every CellRef | PASS | 1729/1729 tracked cells present as formula cells; none overridden in the delivered file; no duplicate or invalid CellRefs |
| FB-02 | Decoder coverage | BIFF12 rgce decoding | decode every tracked cell (and the 7 BrtArrFmla masters) | PASS | 1729 decoded, 0 undecodable; every Ptg in the cells is implemented; cross-validated by 504 exact matches and by named-range rgce (PositionKey, MacroFlag etc.) decoding to the documented targets |
| FB-03 | FormulaBank text == live formula text | column B vs decoded rgce, normalised | string equality after normalisation | FAIL | 504 match, 1225 mismatch. All 1225 differ only in the last row of the source ranges (see drift table). A Reset silently rewrites those ranges |
| FB-04 | Nature of the differences | token-level diff classification | difflib on token lists | INFO | difftype for all 1225 = `range-bound(rows)`; zero differences in string constants (all `MATCH("header")` literals identical), function names, sheet names, operators, `$` anchoring (217 cells use `$C$6`, 1511 use `C6`, identical on both sides) or `@` (bank has `@INDEX` in all 1722 non-array formulas; the sheet stores these as legacy non-dynamic formulas, which Excel displays with `@` - consistent) |
| FB-05 | Live formulas cover the source data | live bound vs BrtWsDim/last key row | per-source comparison | WARN | cashflow live bound 33497 < data 33498 (`Yorktown IX|63`, reachable from Fund View row 63); 5 valuation_history cells bound 32607 < data 32608; AT117:AZ117 bound 87974 < data 87982. Everything else covers the data exactly |
| FB-06 | FormulaBank bounds cover the source data | bank bound vs BrtWsDim | per-source comparison | WARN | Bank covers all current data (it exceeds the tables by 1-14 blank rows) except AT117:AZ117 (`$87981` < data 87982, no numeric effect today). Neither side is future-proof: appended rows beyond the bank bounds will be invisible after a Reset |
| FB-07 | AT117:AZ117 stored as array formulas | BrtArrFmla 0x1AA / PtgExp / grbit | inspect cell records and following records | PASS | Each of the 7 cells holds rgce `01 74000000` (PtgExp row 116) + rgcb PtgExtraCol; followed by BrtArrFmla with rfx = the single cell (e.g. AT117 -> 116,116,45,45), flags byte 0x08, rgce starting `23 2B000000` (PtgName 43 = `_xlfn.LET`) ... `62 01 E400` (SUMPRODUCT array class). Cell grbitFlags = 0 on all 1729 tracked cells (fAlwaysCalc not set). No BrtCellMeta precedes them, so they are classic CSE arrays exactly as `.FormulaArray` would recreate them |
| FB-08 | IsArray flag consistent with storage | FormulaBank col C vs storage | storage type per cell | PASS | 1722 IsArray=N cells are plain cell formulas (no BrtShrFmla/BrtArrFmla involvement); 7 IsArray=Y cells are the array cells. None of the sheet's 133 dynamic-array (BrtCellMeta) cells is tracked |
| FB-09 | AS120 lookup key cell | formula content (bank == sheet) | decoded text + Fund View C5 value | FAIL | Both bank and sheet use `XMATCH(C5&"|"&ROW(),...)` while every other cell uses C6; Fund View C5 = numeric 265 ("Unique identifier"), so the key is `265|120` and never matches a company lookup_key -> AS120 always shows 0 whatever the data. FormulaBank KeyExpr for AS120 says `C6&"|"&ROW()`, so an override typed into AS120 is pushed to the correct company row but the cell can never display it. Not drift (bank matches sheet) but a shared bug that a Reset preserves |
| FB-10 | TargetCol matches the column the formula reads | FormulaBank col H vs `MATCH("header")` / `$BJ$22:24` | header row 1 of each source table | FAIL | 1682/1722 correct. 40 cells AS120:AS159 have TargetCol `M`, but `company` is A:L and `override_value` is column **L** (the formula reads L via MATCH). Sheet730 writes `wsSource.Range(targetCol & targetRow)` = company!M, outside the table: an override typed into AS120:AS159 is stranded in column M and disappears from the cell on Reset |
| FB-11 | SourceSheet / KeyRange / KeyExpr metadata match the formula | FormulaBank cols E,F,G vs decoded formula | regex parse of each formula | PASS (1 exception) | SourceSheet equals the sheet referenced for all 1722 non-array cells; KeyRange name points to the same sheet+column as the XMATCH lookup array for all 1722; KeyExpr equals the XMATCH key for 1721 (exception AS120, see FB-09). BJ22:BJ24 = 7/8/9 = G/H/I matching TargetCol of the valuation_history cells |
| FB-12 | Named KeyRange bounds vs live formula bounds | workbook names vs decoded ranges | compare | INFO | PositionKey/FundKey 819, CompanyKey 4979, CompanyNavKey 15134, PositionCompanyKey 87982, CompanyBreakdownKey 1459 = live formulas. CashflowKey D2:D33498 is one row longer than the live cashflow formulas (33497); QuarterKey K2:K32608 is one row longer than the 5 short valuation_history cells. So for cashflow row 33498 the VBA pushback can find the row but the displayed formula cannot, until a Reset |
| FB-13 | FormulaBank text hygiene for the VBA writers | column B | length / prefix checks | PASS | All 1729 start with `=`; longest formula 232 chars (< 255 limit of `.FormulaArray`); array formulas contain no `@`; non-array formulas contain `@INDEX` (accepted by `.Formula`) |

## Mismatch examples (25 distinct patterns/cells of the 1225; full list in compare.txt)
- **C3** (bank row 2, source `position`) - range-bound(rows): sheet `A2:AG819` vs bank `A2:AG820`; sheet `A2:A819` vs bank `A2:A820`
  - sheet: `=IFERROR(INDEX(position!$A$2:$AG$819,XMATCH(C6,position!$A$2:$A$819,0,1),MATCH("hollyport_id",position!$A$1:$AG$1,0)),"")`
  - bank : `=IFERROR(@INDEX(position!$A$2:$AG$820,XMATCH(C6,position!$A$2:$A$820,0,1),MATCH("hollyport_id",position!$A$1:$AG$1,0)),"")`
- **C8** (bank row 7, source `fund`) - range-bound(rows): sheet `A2:N819` vs bank `A2:N820`; sheet `A2:A819` vs bank `A2:A820`
  - sheet: `=IFERROR(INDEX(fund!$A$2:$N$819,XMATCH(C6,fund!$A$2:$A$819,0,1),MATCH("fund_legal_name",fund!$A$1:$N$1,0)),"")`
  - bank : `=IFERROR(@INDEX(fund!$A$2:$N$820,XMATCH(C6,fund!$A$2:$A$820,0,1),MATCH("fund_legal_name",fund!$A$1:$N$1,0)),"")`
- **L23** (bank row 24, source `valuation_history`) - range-bound(rows): sheet `A2:K32607` vs bank `A2:K32609`; sheet `K2:K32607` vs bank `K2:K32609`
  - sheet: `=IFERROR(INDEX(valuation_history!$A$2:$K$32607,XMATCH($C$6&"|"&TEXT($J23,"yyyy-mm-dd"),valuation_history!$K$2:$K$32607,0,1),$BJ$22),"")`
  - bank : `=IFERROR(@INDEX(valuation_history!$A$2:$K$32609,XMATCH($C$6&"|"&TEXT($J23,"yyyy-mm-dd"),valuation_history!$K$2:$K$32609,0,1),$BJ$22),"")`
- **AD23** (bank row 26, source `valuation_history`) - range-bound(rows): sheet `A2:K32608` vs bank `A2:K32609`; sheet `K2:K32608` vs bank `K2:K32609`
  - sheet: `=IFERROR(INDEX(valuation_history!$A$2:$K$32608,XMATCH($C$6&"|"&TEXT($J23,"yyyy-mm-dd"),valuation_history!$K$2:$K$32608,0,1),$BJ$24),"")`
  - bank : `=IFERROR(@INDEX(valuation_history!$A$2:$K$32609,XMATCH($C$6&"|"&TEXT($J23,"yyyy-mm-dd"),valuation_history!$K$2:$K$32609,0,1),$BJ$24),"")`
- **AE23** (bank row 27, source `cashflow`) - range-bound(rows): sheet `A2:AA33497` vs bank `A2:AA33498`; sheet `D2:D33497` vs bank `D2:D33498`
  - sheet: `=IFERROR(INDEX(cashflow!$A$2:$AA$33497,XMATCH(C6&"|"&ROW(),cashflow!$D$2:$D$33497,0,1),MATCH("cashflow_date",cashflow!$A$1:$AA$1,0)),"")`
  - bank : `=IFERROR(@INDEX(cashflow!$A$2:$AA$33498,XMATCH(C6&"|"&ROW(),cashflow!$D$2:$D$33498,0,1),MATCH("cashflow_date",cashflow!$A$1:$AA$1,0)),"")`
- **O117** (bank row 919, source `position_company`) - range-bound(rows): sheet `A2:Q87982` vs bank `A2:Q87989`; sheet `E2:E87982` vs bank `E2:E87989`
  - sheet: `=IFERROR(INDEX(position_company!$A$2:$Q$87982,XMATCH(C6&"|"&ROW(),position_company!$E$2:$E$87982,0,1),MATCH("ref_date_nav",position_company!$A$1:$Q$1,0)),"")`
  - bank : `=IFERROR(@INDEX(position_company!$A$2:$Q$87989,XMATCH(C6&"|"&ROW(),position_company!$E$2:$E$87989,0,1),MATCH("ref_date_nav",position_company!$A$1:$Q$1,0)),"")`
- **AT117** (bank row 922, array; source position_company) - range-bound(rows): sheet `A2:A87974` vs bank `A2:A87981`; sheet `C2:C87974` vs bank `C2:C87981`; sheet `P2:P87974` vs bank `P2:P87981`
  - sheet: `=LET(n,SUMPRODUCT((position_company!$A$2:$A$87974=$C$6)*(position_company!$C$2:$C$87974=ROW())*(position_company!$P$2:$P$87974=AT$114)*ROW(position_company!$A$2:$A$87974)),v,IF(n=0,"",INDEX(position_company!$Q:$Q,n)),IFERROR(--v,v))`
  - bank : `=LET(n,SUMPRODUCT((position_company!$A$2:$A$87981=$C$6)*(position_company!$C$2:$C$87981=ROW())*(position_company!$P$2:$P$87981=AT$114)*ROW(position_company!$A$2:$A$87981)),v,IF(n=0,"",INDEX(position_company!$Q:$Q,n)),IFERROR(--v,v))`
- **AZ117** (bank row 928, array; source position_company) - same pattern as AT117 with `AZ$114`
  - sheet: `=LET(n,SUMPRODUCT((position_company!$A$2:$A$87974=$C$6)*(position_company!$C$2:$C$87974=ROW())*(position_company!$P$2:$P$87974=AZ$114)*ROW(position_company!$A$2:$A$87974)),v,IF(n=0,"",INDEX(position_company!$Q:$Q,n)),IFERROR(--v,v))`
  - bank : `=LET(n,SUMPRODUCT((position_company!$A$2:$A$87981=$C$6)*(position_company!$C$2:$C$87981=ROW())*(position_company!$P$2:$P$87981=AZ$114)*ROW(position_company!$A$2:$A$87981)),v,IF(n=0,"",INDEX(position_company!$Q:$Q,n)),IFERROR(--v,v))`
- **K120** (bank row 933, source `company`) - range-bound(rows): sheet `A2:L4979` vs bank `A2:L4984`; sheet `D2:D4979` vs bank `D2:D4984`
  - sheet: `=IFERROR(INDEX(company!$A$2:$L$4979,XMATCH(C6&"|"&ROW(),company!$D$2:$D$4979,0,1),MATCH("geography",company!$A$1:$L$1,0)),"")`
  - bank : `=IFERROR(@INDEX(company!$A$2:$L$4984,XMATCH(C6&"|"&ROW(),company!$D$2:$D$4984,0,1),MATCH("geography",company!$A$1:$L$1,0)),"")`
- **AT120** (bank row 943, source `company_nav`) - range-bound(rows): sheet `A2:J15134` vs bank `A2:J15148`; sheet `D2:D15134` vs bank `D2:D15148`
  - sheet: `=IFERROR(INDEX(company_nav!$A$2:$J$15134,XMATCH(C6&"|"&ROW()&"|"&TEXT(AT$114,"yyyy-mm-dd")&"",company_nav!$D$2:$D$15134,0,1),MATCH("nav_value",company_nav!$A$1:$J$1,0)),"")`
  - bank : `=IFERROR(@INDEX(company_nav!$A$2:$J$15148,XMATCH(C6&"|"&ROW()&"|"&TEXT(AT$114,"yyyy-mm-dd")&"",company_nav!$D$2:$D$15148,0,1),MATCH("nav_value",company_nav!$A$1:$J$1,0)),"")`
- **AS120** (bank row 1691, source `company`) - range-bound(rows): sheet `A2:L4979` vs bank `A2:L4984`; sheet `D2:D4979` vs bank `D2:D4984`. NOTE: both texts use `C5` (bank KeyExpr says C6) and TargetCol is `M` although the formula reads column L - see FB-09/FB-10
  - sheet: `=IFERROR(INDEX(company!$A$2:$L$4979,XMATCH(C5&"|"&ROW(),company!$D$2:$D$4979,0,1),MATCH("override_value",company!$A$1:$L$1,0)),0)`
  - bank : `=IFERROR(@INDEX(company!$A$2:$L$4984,XMATCH(C5&"|"&ROW(),company!$D$2:$D$4984,0,1),MATCH("override_value",company!$A$1:$L$1,0)),0)`
- **AY149** (bank row 1499, source `company_nav`) - range-bound(rows): sheet `A2:J15134` vs bank `A2:J15148`; sheet `D2:D15134` vs bank `D2:D15148`
  - sheet: `=IFERROR(INDEX(company_nav!$A$2:$J$15134,XMATCH(C6&"|"&ROW()&"|"&TEXT(AY$114,"yyyy-mm-dd")&"",company_nav!$D$2:$D$15134,0,1),MATCH("nav_value",company_nav!$A$1:$J$1,0)),"")`
  - bank : `=IFERROR(@INDEX(company_nav!$A$2:$J$15148,XMATCH(C6&"|"&ROW()&"|"&TEXT(AY$114,"yyyy-mm-dd")&"",company_nav!$D$2:$D$15148,0,1),MATCH("nav_value",company_nav!$A$1:$J$1,0)),"")`
- **K134** (bank row 1199, source `company`) - range-bound(rows): sheet `A2:L4979` vs bank `A2:L4984`; sheet `D2:D4979` vs bank `D2:D4984`
  - sheet: `=IFERROR(INDEX(company!$A$2:$L$4979,XMATCH(C6&"|"&ROW(),company!$D$2:$D$4979,0,1),MATCH("geography",company!$A$1:$L$1,0)),"")`
  - bank : `=IFERROR(@INDEX(company!$A$2:$L$4984,XMATCH(C6&"|"&ROW(),company!$D$2:$D$4984,0,1),MATCH("geography",company!$A$1:$L$1,0)),"")`
- **C14** (bank row 15, source `fund`) - range-bound(rows): sheet `A2:N819` vs bank `A2:N820`; sheet `A2:A819` vs bank `A2:A820`
  - sheet: `=IFERROR(INDEX(fund!$A$2:$N$819,XMATCH($C$6,fund!$A$2:$A$819,0,1),MATCH("fund_size",fund!$A$1:$N$1,0)),"")`
  - bank : `=IFERROR(@INDEX(fund!$A$2:$N$820,XMATCH($C$6,fund!$A$2:$A$820,0,1),MATCH("fund_size",fund!$A$1:$N$1,0)),"")`
- **F18** (bank row 21, source `position`) - range-bound(rows): sheet `A2:AG819` vs bank `A2:AG820`; sheet `A2:A819` vs bank `A2:A820`
  - sheet: `=IFERROR(INDEX(position!$A$2:$AG$819,XMATCH(C6,position!$A$2:$A$819,0,1),MATCH("manual_override",position!$A$1:$AG$1,0)),"")`
  - bank : `=IFERROR(@INDEX(position!$A$2:$AG$820,XMATCH(C6,position!$A$2:$A$820,0,1),MATCH("manual_override",position!$A$1:$AG$1,0)),"")`
- **AK23** (bank row 30, source `cashflow`) - range-bound(rows): sheet `A2:AA33497` vs bank `A2:AA33498`; sheet `D2:D33497` vs bank `D2:D33498`
  - sheet: `=(IFERROR(INDEX(cashflow!$A$2:$AA$33497,XMATCH(C6&"|"&ROW(),cashflow!$D$2:$D$33497,0,1),MATCH("cash_dist",cashflow!$A$1:$AA$1,0)),""))*1`
  - bank : `=(IFERROR(@INDEX(cashflow!$A$2:$AA$33498,XMATCH(C6&"|"&ROW(),cashflow!$D$2:$D$33498,0,1),MATCH("cash_dist",cashflow!$A$1:$AA$1,0)),""))*1`
- **AN27** (bank row 119, source `cashflow`) - range-bound(rows): sheet `A2:AA33497` vs bank `A2:AA33498`; sheet `D2:D33497` vs bank `D2:D33498`
  - sheet: `=IFERROR(INDEX(cashflow!$A$2:$AA$33497,XMATCH(C6&"|"&ROW(),cashflow!$D$2:$D$33497,0,1),MATCH("withheld",cashflow!$A$1:$AA$1,0)),"")`
  - bank : `=IFERROR(@INDEX(cashflow!$A$2:$AA$33498,XMATCH(C6&"|"&ROW(),cashflow!$D$2:$D$33498,0,1),MATCH("withheld",cashflow!$A$1:$AA$1,0)),"")`
- **G120** (bank row 932, source `position_company`) - range-bound(rows): sheet `A2:Q87982` vs bank `A2:Q87989`; sheet `E2:E87982` vs bank `E2:E87989`
  - sheet: `=IFERROR(INDEX(position_company!$A$2:$Q$87982,XMATCH(C6&"|"&ROW(),position_company!$E$2:$E$87982,0,1),MATCH("company_id",position_company!$A$1:$Q$1,0)),"")`
  - bank : `=IFERROR(@INDEX(position_company!$A$2:$Q$87989,XMATCH(C6&"|"&ROW(),position_company!$E$2:$E$87989,0,1),MATCH("company_id",position_company!$A$1:$Q$1,0)),"")`
- **L45** (bank row 499, source `valuation_history`) - range-bound(rows): sheet `A2:K32608` vs bank `A2:K32609`; sheet `K2:K32608` vs bank `K2:K32609`
  - sheet: `=IFERROR(INDEX(valuation_history!$A$2:$K$32608,XMATCH($C$6&"|"&TEXT($J45,"yyyy-mm-dd"),valuation_history!$K$2:$K$32608,0,1),$BJ$22),"")`
  - bank : `=IFERROR(@INDEX(valuation_history!$A$2:$K$32609,XMATCH($C$6&"|"&TEXT($J45,"yyyy-mm-dd"),valuation_history!$K$2:$K$32609,0,1),$BJ$22),"")`
- **AD29** (bank row 158, source `valuation_history`) - range-bound(rows): sheet `A2:K32607` vs bank `A2:K32609`; sheet `K2:K32607` vs bank `K2:K32609`
  - sheet: `=IFERROR(INDEX(valuation_history!$A$2:$K$32607,XMATCH($C$6&"|"&TEXT($J29,"yyyy-mm-dd"),valuation_history!$K$2:$K$32607,0,1),$BJ$24),"")`
  - bank : `=IFERROR(@INDEX(valuation_history!$A$2:$K$32609,XMATCH($C$6&"|"&TEXT($J29,"yyyy-mm-dd"),valuation_history!$K$2:$K$32609,0,1),$BJ$24),"")`
- **AE63** (bank row 882, source `cashflow`) - range-bound(rows): sheet `A2:AA33497` vs bank `A2:AA33498`; sheet `D2:D33497` vs bank `D2:D33498` (this is the cell that would start seeing `Yorktown IX|63` after a Reset)
  - sheet: `=IFERROR(INDEX(cashflow!$A$2:$AA$33497,XMATCH(C6&"|"&ROW(),cashflow!$D$2:$D$33497,0,1),MATCH("cashflow_date",cashflow!$A$1:$AA$1,0)),"")`
  - bank : `=IFERROR(@INDEX(cashflow!$A$2:$AA$33498,XMATCH(C6&"|"&ROW(),cashflow!$D$2:$D$33498,0,1),MATCH("cashflow_date",cashflow!$A$1:$AA$1,0)),"")`
- **F5** (bank row 4, source `position`) - range-bound(rows): sheet `A2:AG819` vs bank `A2:AG820`; sheet `A2:A819` vs bank `A2:A820`
  - sheet: `=IFERROR(INDEX(position!$A$2:$AG$819,XMATCH(C6,position!$A$2:$A$819,0,1),MATCH("commitment",position!$A$1:$AG$1,0)),"")`
  - bank : `=IFERROR(@INDEX(position!$A$2:$AG$820,XMATCH(C6,position!$A$2:$A$820,0,1),MATCH("commitment",position!$A$1:$AG$1,0)),"")`
- **K16** (bank row 18, source `position`) - range-bound(rows): sheet `A2:AG819` vs bank `A2:AG820`; sheet `A2:A819` vs bank `A2:A820`
  - sheet: `=IFERROR(INDEX(position!$A$2:$AG$819,XMATCH(C6,position!$A$2:$A$819,0,1),MATCH("ix_ownership_pct",position!$A$1:$AG$1,0)),"")`
  - bank : `=IFERROR(@INDEX(position!$A$2:$AG$820,XMATCH(C6,position!$A$2:$A$820,0,1),MATCH("ix_ownership_pct",position!$A$1:$AG$1,0)),"")`
- **AA23** (bank row 25, source `valuation_history`) - range-bound(rows): sheet `A2:K32607` vs bank `A2:K32609`; sheet `K2:K32607` vs bank `K2:K32609`
  - sheet: `=IFERROR(INDEX(valuation_history!$A$2:$K$32607,XMATCH($C$6&"|"&TEXT($J23,"yyyy-mm-dd"),valuation_history!$K$2:$K$32607,0,1),$BJ$23),"")`
  - bank : `=IFERROR(@INDEX(valuation_history!$A$2:$K$32609,XMATCH($C$6&"|"&TEXT($J23,"yyyy-mm-dd"),valuation_history!$K$2:$K$32609,0,1),$BJ$23),"")`
- **AH23** (bank row 28, source `cashflow`) - range-bound(rows): sheet `A2:AA33497` vs bank `A2:AA33498`; sheet `D2:D33497` vs bank `D2:D33498`
  - sheet: `=IFERROR(INDEX(cashflow!$A$2:$AA$33497,XMATCH(C6&"|"&ROW(),cashflow!$D$2:$D$33497,0,1),MATCH("calls_addon",cashflow!$A$1:$AA$1,0)),"")`
  - bank : `=IFERROR(@INDEX(cashflow!$A$2:$AA$33498,XMATCH(C6&"|"&ROW(),cashflow!$D$2:$D$33498,0,1),MATCH("calls_addon",cashflow!$A$1:$AA$1,0)),"")`

## AT117:AZ117 storage detail
- Cell records (BrtFmlaNum / BrtFmlaString, grbitFlags = 0): rgce = `01 74 00 00 00` (PtgExp, row 116), rgcb = PtgExtraCol (`2D..33 00 00 00` = columns 45..51).
- Immediately after each cell: BrtArrFmla (0x1AA), rfx = (116,116,col,col) - a single-cell array, flags byte 0x08, cce 189, rgce = PtgName(43 `_xlfn.LET`) PtgName(80 `_xlpm.n`) PtgArea3d ... PtgFuncVar(228 SUMPRODUCT, array class) ... PtgFuncVar(255 with `_xlfn.LET`, 6 params).
- No BrtCellMeta (0x31) before any of the 7 cells -> not dynamic-array formulas. This matches FormulaBank IsArray=Y and what `Range.FormulaArray = "=LET(...)"` produces, so a Reset preserves the storage type (only the row bounds change, 87974 -> 87981).

## Overall verdict
**FAIL (data drift) with two additional content defects.** FormulaBank does not match the live Fund View for 1225 of 1729 cells; the differences are confined to the last row of every source-table reference (bank 1-14 rows longer than the live formulas; live cashflow, 5 valuation_history cells and the 7 AT117:AZ117 arrays are shorter than their tables). A Reset therefore silently changes lookup ranges and, for cashflow row 63 / fund Yorktown IX, the visible numbers. Independently of drift, AS120 keys on C5 instead of C6 (always 0) and AS120:AS159 carry TargetCol `M` while the company table ends at L (overrides stranded outside the table).

Recommended fixes (for the coordinator): regenerate FormulaBank column B from the live sheet (or, better, rewrite both to whole-column/structured references or the named KeyRanges so that table growth cannot desynchronise them); correct AS120 to `C6`; set TargetCol `L` for AS120:AS159; extend the 5 short valuation_history cells and AT117:AZ117 to the full table. Live Excel UAT still needed to confirm: (a) that `.Formula` round-trips the `@INDEX` text unchanged, (b) values of AE63:AN63 for Yorktown IX before/after RestoreYellowFormulas, (c) that an AS-column override is written to company!M as predicted.
