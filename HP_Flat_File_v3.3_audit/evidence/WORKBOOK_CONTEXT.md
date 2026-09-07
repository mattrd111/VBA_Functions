# Workbook context bundle for the VBA audit of HP_Flat_File_Test_v3.3.xlsb

All paths below are under <scratch>

## VBA modules (extracted with olevba; read these in full)
- vba/modules/ThisWorkbook.cls  (7 lines: Workbook_Open / Workbook_BeforeSave set named cell MacroFlag)
- vba/modules/Sheet730.cls      (523 lines) = the LIVE code-behind of worksheet "Fund View" (codeName Sheet730)
- vba/modules/Main.bas          (432 lines) standard module: OverrideYellowCells, RestoreYellowFormulas, ApplyYellow, BlankChunk, AddNewFund, CheckDuplicateFundIDs, FindDuplicatesInRange, ExportAllFundSnapshots, CountOverriddenTrackedCells, SafeSheetName, RestoreOverriddenCellsOnly
- vba/modules/Sheet731.cls      (295 lines) a Document module that is NOT attached to any worksheet in the workbook (no sheet has codeName Sheet731). It is an older, better-commented copy of the Fund View module without the company-quarter feature. Orphaned/dead.
- All other SheetNN.cls modules are empty stubs.
- vba/olevba_full.txt = raw olevba output.

## Buttons on Fund View (drawing1.xml)
- "Button 16" -> macro AddNewFund
- "Button 17" -> macro RestoreOverriddenCellsOnly  (the "Reset Sheet" button per User Guide)
No button calls OverrideYellowCells, RestoreYellowFormulas, CheckDuplicateFundIDs or ExportAllFundSnapshots (run from the macro dialog).

## Sheets (64). Relevant ones
- "Fund View" (codeName Sheet730). Dimension B1:BL200. C6 = selected fund (position_id), has a data-validation dropdown. C6 currently = "Abenex IV (A)".
  Row 114: AT114:AZ114 hold quarter-end dates (formula cells, number format d-mmm-yy; values 46022,45930,45838,45747,45657,45565,45473 = 2025-12-31 ... 2024-06-30). A note by Matthew Davies sits on AT114.
  Rows 116-119 are rollup rows: B116 "Net NAV"?? (I116 'Net NAV'), I117 'Carry', I118 'Management Fee', I119 'Gross NAV'; G119 holds the header text "Company ID", H119 "Exited?".
  Rows 120-159 are company rows: G = company_id (e.g. G120 = "NC_PQ/15558"), J = company name, AT:AZ = NAV per quarter (from company_nav).
  AT117:AZ117 are single-cell array formulas (LET/SUMPRODUCT over position_company, len 232 chars) tracked in FormulaBank with IsArray=Y and CanPushback=N.
  AT118:AZ118 are BLANK, untracked cells that carry stale legacy notes "Overridden by MatthewDavies on 25/08/2026 09:45 / Previous value: unknown (multi-cell edit/paste)". AT119:AZ119 are untracked formula cells.
  Every one of the 1729 tracked cells currently holds a formula (1252 string-result, 477 numeric-result). No override is active in the delivered file.
  E2 = formula banner "MACROS ENABLED" driven by MacroFlag. C3 hollyport_id lookup.
- "FormulaBank" (hidden). Header row 1: CellRef | Formula | IsArray | CanPushback | SourceSheet | KeyExpr | KeyRange | TargetCol. Rows 2..1730 = 1729 tracked cells. Full dump: ctx/formulabank.txt
  SourceSheet counts: company_breakdown 504, position_company 286, company_nav 280, cashflow 250, company 239, valuation_history 126, position 27, fund 10, (blank) 7 (the AT117:AZ117 array rows).
  KeyExpr patterns: C6&"|"&ROW() (1198), C6&"|"&TEXT(J#,"yyyy-mm-dd") (126, valuation_history), $C$6&"|"&ROW() (81), C6&"|"&ROW()&"|"&TEXT(AT$114,"yyyy-mm-dd")&"" (7x40 company_nav), C6 (37 position/fund). Formulas store "@INDEX(...XMATCH(...))" text; none exceed 255 chars.
  Tracked columns: C,F,G,J,K,L,M,N,O,P,R,S,T,Z,AA,AD,AE,AH,AI,AK,AL,AN,AS,AT..AZ,BA..BF; rows 3..159. All 280 company_nav-sourced cells are inside AT120:AZ159.
- "_Config" (very hidden): A1 label, A2 = MacroFlag (currently TRUE).
- Source tables (row 1 headers, data from row 2, alphabetically sorted by key):
  fund (A2:N819): fund_key, source_tab, fund_id, fund_legal_name, preqin_fund_name, preqin_manager, gp_id, hollyport_fund_type, geography, vintage, carry_terms_raw, fees_terms_raw, fund_currency, fund_size. All constants.
  position (A2:AG819): position_id, source_tab, hollyport_id, unique_identifier, hollyport_deal, vehicle, fund_id, underlying_fund_currency, reference_date, completion_date, nav_date_at_completion, liquidated_date, commitment(M), uncalled_at_reference_date(N), headline_price(O), percentage_of_fund_held(P), ix_ownership_pct(Q), viii_ownership_pct(R), date_of_update(S), hollyport_reference_alt, pre_completion_dist_forecast(U), ..., lp_interest_type(AA), updated_by(AB), liquidated_in_sf, manual_override(AD), notes(AE), position_key(AF), recycling_date(AG). All constants.
  company (A2:L4979): source_tab, source_row, line_type, lookup_key(D), company_id, company_name, geography, sector, strategy, public_private, exited, override_value. lookup_key = source_tab|source_row. All constants.
  position_company (A2:Q87982): source_tab, position_id, source_row, line_type, lookup_key(E), company_id, ref_date_nav, forecast_calls, multiple, gross_nav_input, methodology, information_quality, exit_type, as_of_date, position_company_key, component_label(P), component_value(Q). All constants.
  company_nav (A2:J15134): source_tab, position_id, source_row, lookup_key(D) = source_tab|source_row|yyyy-mm-dd, company_id(E), company_name(F), currency(G), Nav_date(H, numeric serial), nav_value(I), as_of_date(J). All constants. 15133 data rows, last row 15134.
  cashflow (A2:AH33498): source_tab, position_id, source_row, lookup_key(D)=source_tab|source_row, currency, as_of_date, entered_by, cashflow_key, cashflow_date(I), quarter_end(J), fx_rate_at_date, calls_addon(L), calls_fees(M), total_calls, cash_dist(O), in_specie_dist(P), total_gross_dist, withheld(R), total_net_dist, total_cd, cd_val_adj, then FORMULA columns V:AH (net_dist_usd, calls_usd, fees_usd, addon_usd, withheld_usd, total_net_dist_plus_calls, pos_row, comp_date, recyc_date, period, in_recycling, has_quarter, relevant).
  valuation_history (A2:K32608): source_tab, position_id, source_row, lookup_key(D), currency, as_of_date, val_date(G), valuation_date(H), unadjusted_nav(I), quarter_end(J), quarter_key(K) = position_id|yyyy-mm-dd. All constants.
  company_breakdown (A2:U1459): source_tab, position_id, source_row, lookup_key(D), val_type(E), carry_paid(F), carry_recipient_name, carry_recipient_amount, as_of_date, company_1_name(J)...company_5_proceeds(S), then FORMULA columns T:U (proceeds_total, carry_pre_comp).
- Defined names used by VBA: MacroFlag='_Config'!A2; PositionKey=position!A2:A819; FundKey=fund!A2:A819. Others: CompanyNavKey=company_nav!D2:D15134, CashflowKey=cashflow!D2:D33498, CompanyKey=company!D2:D4979, PositionCompanyKey=position_company!E2:E87982, QuarterKey=valuation_history!K2:K32608, CompanyBreakdownKey=company_breakdown!D2:D1459, ValuationHistoryKey=valuation_history!D2:D32608.
  NOTE: Fund View formulas reference position!$A$2:$A$820, company_nav!$D$2:$D$15148 etc. (slightly larger than the named ranges), while Sheet730.PushCompanyQuarterOverride hard-codes company_nav rows 2..15131.

## Data-check results (python over the actual tables) - see ctx/data_checks.txt
- 818 position ids, 818 fund keys, sets identical, no duplicates (exact or case-insensitive), none contain * ? ~, max length 31, none numeric-looking. ONE id has a trailing space: 'IK VII No.4 ' (both tables).
- company_nav: source_tab == position_id on every row. lookup_key always equals source_tab|source_row|yyyy-mm-dd(Nav_date). 485 rows have blank company_id. 43 lookup_keys are already duplicated (e.g. 'Edgewater Growth II|120|2026-03-31', 'EQT VII|120|2025-12-31'). Rows 15132-15134 ('Yorktown IX|130|2026-03-31', 'Yorktown IX|131|2025-06-30', 'Yorktown IX|131|2026-03-31') lie BEYOND the VBA's hard-coded search bound of row 15131. No company_nav rows exist with source_row 116-119.
- Current fund 'Abenex IV (A)' has company_nav rows only for 2024-09-30 and 2025-12-31 (row 120, company NC_PQ/15558) while the grid headers cover 7 quarters, so most AT120:AZ120 cells are blank formulas whose override would go down the 'insert new row' path.

## Documentation inside the workbook (ctx/guide_issuelog.txt)
User Guide says: yellow cells hold lookup formulas; typing over one stores the value on the source table; typing on a blank formula cell in nav or company positions creates a new nav period / company position; red = overridden; cell gets a note with old value/user/time; "Reset Sheet" button = RestoreOverriddenCellsOnly; RestoreYellowFormulas "not for general use"; AddNewFund = "New Fund" button; ExportAllFundSnapshots "not fully tested"; known issues: multi-cell paste notes say unknown; yellow cells can only be overridden with values not formulas; FormulaBank manually maintained (add ' before formula).
Issue Log #13 (High, Open): "cashflow AB:AH - Write-back must preserve the new formula columns (pos_row, comp_date, recyc_date, period, in_recycling, has_quarter, relevant) when it inserts rows. Not yet tested against a live run."

## Environment note
Excel is not available in this Linux container, so nothing can be executed live. Findings must be reasoned from the code + these facts; be explicit about what a live Excel UAT run would be needed to confirm.
