Attribute VB_Name = "modAddFunds"
Option Explicit

' =====================================================================================
' modAddFunds  (v2 - 08/09/2026)
'
' Purpose: Read newly pasted fund rows (Summary sheet) and their company detail
'          (Detailed sheet) from THIS workbook (the IX Onboarder) and append them
'          into the fund / position / company / position_company / company_nav
'          tables in the Flat File workbook, which must already be open.
'
' Entry point: AddNewFundsAndCompanies  (wire this to a button on the Summary sheet)
'
' Behaviour agreed with Matthew (07/09/2026):
'   - Trigger: button on Summary sheet, not automatic on paste.
'   - Duplicates: if ANY pasted fund name already exists in the fund table, the
'     whole run aborts before writing anything, and a message lists the clashing
'     fund names so they can be resolved manually first.
'
' v2 changes (08/09/2026) - rows are now written the way the Flat File's Fund View
' actually reads them (lookup is fund & "|" & ROW() against lookup_key):
'   - Company block uses the Fund View row numbers: 116-119 rollup rows,
'     120 Net Curr, 121-159 companies (max 39 companies per fund).
'   - NAV goes to position_company.ref_date_nav (Fund View col O), not gross_nav_input.
'     gross_nav_input carries the Fund View col I ordinal (1 = Net Curr row) on
'     company rows and the row label (Net NAV / Carry / ...) on rollup rows.
'   - line_type is set: rollup / net_current_assets / company.
'   - Overall Upside (Detailed block KE) goes to position_company.multiple (Fund View
'     col R). Written as-is: the Flat File holds return multiples (1.0 = no upside).
'   - The Detailed block headed "Reinvested" (IP) is written to
'     position_company.forecast_calls. The template formulas in that block sum
'     Known Calls + Assumed Calls + Post Comp. Calls per company - the "Reinvested"
'     header is stale. Non-numeric values (e.g. Yes/No) are logged, not written.
'   - Standard funds (Summary Valuation Methodology <> "Detailed") get the placeholder
'     block every Standard fund in the Flat File carries: rollup rows with NAV 1 and
'     multiple = fund upside, one Net Curr row with NAV 1, and a company_nav row.
'     The upside comes from Summary "NAV plus Calls Upside" (col Z) when it looks like
'     a multiple; otherwise 1 is written and the raw value is logged.
'   - Detailed fund-level values go on the rollup rows: Carry -> row 117 ref_date_nav,
'     Carry Multiple -> row 117 multiple, Fees -> row 118 ref_date_nav,
'     F'Cast Fees Multiplier Value / Fees -> row 118 multiple (legacy fund-tab formula),
'     F'Cast Calls -> row 120 forecast_calls. Rows 116 and 119 are written as the
'     sums the Fund View recomputes (_Agg reads row 116 and row 120 ref_date_nav).
'   - A company_nav row is written per company at the reference date so Fund View
'     "Current Value" (col AA) resolves.
'
' Still logged to the "Import Log" sheet (no home in the Flat File):
'   - Summary fund currency: Summary only has an FX rate (col L). A currency CODE is
'     needed in fund.fund_currency and position.underlying_fund_currency (the Fund
'     View reads the latter into C22 and uses it to pick the FX rate).
'   - Detailed fund-level "Net Curr Reinvested" (PG).
'   - Summary NAV (col Q): the Fund View takes NAV from valuation_history, which this
'     macro does not write.
'
' Column layouts below were read directly off the files on 07-08/09/2026 - re-check
' them if either file's structure changes.
' =====================================================================================

Private Const TARGET_WORKBOOK_NAME As String = "HP Flat File Test v3.6.xlsb"

' --- Detailed sheet block layout (row 2 = block title, row 3 = slot numbers 1-40, data from row 4) ---
Private Const DET_HEADER_ROW As Long = 3
Private Const DET_DATA_START_ROW As Long = 4
Private Const DET_COL_FUND_NAME As Long = 1   ' A
Private Const DET_BLOCK_COMPANIES As Long = 4     ' D  - "Companies"
Private Const DET_BLOCK_GEO As Long = 45    ' AS - "Geo"
Private Const DET_BLOCK_INDUSTRY As Long = 86    ' CH - "Industry"
Private Const DET_BLOCK_CLASSIFICATION As Long = 127   ' DW - "Classification"
Private Const DET_BLOCK_PUBPRIV As Long = 168   ' FL - "Public/Private"
Private Const DET_BLOCK_NAV As Long = 209   ' HA - "NAV"
Private Const DET_BLOCK_FCAST_CALLS As Long = 250   ' IP - headed "Reinvested", holds forecast calls (see header note)
Private Const DET_BLOCK_UPSIDE As Long = 291   ' KE - "Overall Upside"
Private Const DET_BLOCK_VALMETHOD As Long = 332   ' LT - "HPT Val methodology"
Private Const DET_BLOCK_INFOQUALITY As Long = 373   ' NI - "HPT Info Quality"
Private Const DET_SLOT_COUNT As Long = 40
Private Const DET_COL_CURR_ASSETS As Long = 414   ' OX
Private Const DET_COL_TOTAL_LIAB As Long = 415   ' OY
Private Const DET_COL_NET_CURR As Long = 416   ' OZ
Private Const DET_COL_CARRY As Long = 417   ' PA
Private Const DET_COL_CARRY_MULTIPLE As Long = 418   ' PB
Private Const DET_COL_FEES As Long = 419   ' PC
Private Const DET_COL_FCAST_FEES_MULT As Long = 420   ' PD
Private Const DET_COL_FCAST_CALLS As Long = 421   ' PE
Private Const DET_COL_NET_CURR_GEOG As Long = 422   ' PF
Private Const DET_COL_NET_CURR_REINVEST As Long = 423   ' PG

' --- Summary sheet layout (header row 4, data from row 5) ---
Private Const SUM_HEADER_ROW As Long = 4
Private Const SUM_DATA_START_ROW As Long = 5

' --- Fund View row layout in the Flat File (position_company.source_row / lookup_key) ---
Private Const FV_ROW_NET_NAV As Long = 116
Private Const FV_ROW_CARRY As Long = 117
Private Const FV_ROW_MGMT_FEE As Long = 118
Private Const FV_ROW_GROSS_NAV As Long = 119
Private Const FV_ROW_NET_CURR As Long = 120
Private Const FV_ROW_FIRST_COMPANY As Long = 121
Private Const FV_ROW_LAST_COMPANY As Long = 159

' Summary "NAV plus Calls Upside" is accepted as a multiple only inside this range;
' anything else is logged and 1 is written (the Fund View multiplies NAV plus Calls by it).
Private Const MAX_PLAUSIBLE_MULTIPLE As Double = 5#

' --- Module-level state for the current run (set up in AddNewFundsAndCompanies) ---
Private mWsLog As Worksheet
Private mWsCompany As Worksheet, mWsPC As Worksheet, mWsNav As Worksheet
Private mMapCompany As Object, mMapPC As Object, mMapNav As Object
Private mCompaniesAdded As Long, mRollupsAdded As Long, mNavRowsAdded As Long

Sub AddNewFundsAndCompanies()

    Dim wbSrc As Workbook, wbTgt As Workbook
    Dim wsSummary As Worksheet, wsDetailed As Worksheet
    Dim wsFund As Worksheet, wsPosition As Worksheet

    Set wbSrc = ThisWorkbook

    ' --- Find the target workbook among currently open workbooks ---
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name = TARGET_WORKBOOK_NAME Then
            Set wbTgt = wb
            Exit For
        End If
    Next wb

    If wbTgt Is Nothing Then
        MsgBox "Could not find '" & TARGET_WORKBOOK_NAME & "' among the open workbooks." & vbCrLf & _
               "Please open it first, then run this macro again.", vbExclamation, "Flat File not open"
        Exit Sub
    End If

    Set mWsCompany = Nothing: Set mWsPC = Nothing: Set mWsNav = Nothing
    On Error Resume Next
    Set wsFund = wbTgt.Sheets("fund")
    Set wsPosition = wbTgt.Sheets("position")
    Set mWsCompany = wbTgt.Sheets("company")
    Set mWsPC = wbTgt.Sheets("position_company")
    Set mWsNav = wbTgt.Sheets("company_nav")
    On Error GoTo 0

    If wsFund Is Nothing Or wsPosition Is Nothing Or mWsCompany Is Nothing Or mWsPC Is Nothing Then
        MsgBox "One or more expected sheets (fund / position / company / position_company) " & _
               "were not found in '" & TARGET_WORKBOOK_NAME & "'. Aborting - has the Flat File layout changed?", _
               vbCritical, "Sheet not found"
        Exit Sub
    End If

    Set wsSummary = wbSrc.Sheets("Summary")
    Set wsDetailed = wbSrc.Sheets("Detailed")

    ' --- Defensive check that the Detailed block layout hasn't shifted ---
    If wsDetailed.Cells(2, DET_BLOCK_COMPANIES).Value <> "Companies" Or _
       wsDetailed.Cells(2, DET_BLOCK_NAV).Value <> "NAV" Or _
       wsDetailed.Cells(2, DET_BLOCK_UPSIDE).Value <> "Overall Upside" Then
        MsgBox "The Detailed sheet's column layout doesn't match what this macro expects " & _
               "(block headers have moved). Aborting rather than writing to the wrong columns." & vbCrLf & _
               "Ask for the macro to be updated for the new layout.", vbCritical, "Layout check failed"
        Exit Sub
    End If

    Set mWsLog = GetOrCreateLogSheet(wbSrc)

    ' --- Build header maps for the target tables (lookup column index by header name) ---
    Dim mapFund As Object, mapPosition As Object
    Set mapFund = GetHeaderMap(wsFund)
    Set mapPosition = GetHeaderMap(wsPosition)
    Set mMapCompany = GetHeaderMap(mWsCompany)
    Set mMapPC = GetHeaderMap(mWsPC)
    If mWsNav Is Nothing Then
        Set mMapNav = Nothing
    Else
        Set mMapNav = GetHeaderMap(mWsNav)
    End If

    Dim missing As String
    missing = CheckHeaders(mapFund, "fund", "fund_key,source_tab,fund_legal_name") & _
              CheckHeaders(mapPosition, "position", "position_id,source_tab,reference_date") & _
              CheckHeaders(mMapCompany, "company", "source_tab,source_row,line_type,lookup_key,company_id,company_name") & _
              CheckHeaders(mMapPC, "position_company", "source_tab,position_id,source_row,line_type,lookup_key,company_id,ref_date_nav,forecast_calls,multiple,gross_nav_input,methodology,information_quality")
    If missing <> "" Then
        MsgBox "Expected columns are missing from the Flat File tables:" & vbCrLf & missing & vbCrLf & _
               "Aborting - has the Flat File layout changed?", vbCritical, "Column not found"
        Exit Sub
    End If

    ' --- Build a set of existing fund keys in the Flat File (column "fund_key") ---
    Dim existingFunds As Object
    Set existingFunds = CreateObject("Scripting.Dictionary")
    existingFunds.CompareMode = 1 ' vbTextCompare - case-insensitive
    Dim fundKeyCol As Long, lastFundRow As Long, r As Long
    fundKeyCol = mapFund("fund_key")
    lastFundRow = wsFund.Cells(wsFund.Rows.Count, fundKeyCol).End(xlUp).Row
    For r = 2 To lastFundRow
        Dim fk As String
        fk = Trim(CStr(wsFund.Cells(r, fundKeyCol).Value))
        If fk <> "" Then
            If Not existingFunds.Exists(fk) Then existingFunds.Add fk, True
        End If
    Next r

    ' --- Pass 1: read Summary rows, check for clashes before writing anything ---
    Dim lastSumRow As Long
    lastSumRow = wsSummary.Cells(wsSummary.Rows.Count, 1).End(xlUp).Row

    Dim clashList As String
    clashList = ""
    Dim fundNames As Collection
    Set fundNames = New Collection

    For r = SUM_DATA_START_ROW To lastSumRow
        Dim fundName As String
        fundName = SafeStr(wsSummary.Cells(r, 1).Value)
        If fundName <> "" Then
            fundNames.Add fundName
            If existingFunds.Exists(fundName) Then
                clashList = clashList & "  - " & fundName & vbCrLf
            End If
        End If
    Next r

    If fundNames.Count = 0 Then
        MsgBox "No fund names found on the Summary sheet (column A, from row " & SUM_DATA_START_ROW & "). Nothing to do.", vbInformation
        Exit Sub
    End If

    If clashList <> "" Then
        MsgBox "Stopped - nothing has been written." & vbCrLf & vbCrLf & _
               "These funds already exist in the Flat File's fund table:" & vbCrLf & vbCrLf & _
               clashList & vbCrLf & _
               "Resolve these (rename, remove the duplicate row, or handle manually) and run again.", _
               vbExclamation, "Duplicate fund names found"
        Exit Sub
    End If

    If mWsNav Is Nothing Then
        LogSkipped "(all)", "company_nav", "", "No company_nav sheet found in the Flat File - Fund View 'Current Value' (col AA) will be blank for these funds"
    End If

    ' --- Pass 2: no clashes, safe to write. Loop Summary rows and append. ---
    Dim fundsAdded As Long
    fundsAdded = 0: mCompaniesAdded = 0: mRollupsAdded = 0: mNavRowsAdded = 0

    For r = SUM_DATA_START_ROW To lastSumRow
        fundName = SafeStr(wsSummary.Cells(r, 1).Value)
        If fundName = "" Then GoTo NextSummaryRow

        Dim s As Object
        Set s = ReadSummaryRow(wsSummary, r)

        Dim refDate As Variant
        refDate = ToDateOrRaw(s("ValDate"))

        ' --- Append to fund table ---
        Dim newFundRow As Long
        newFundRow = NextEmptyRow(wsFund, mapFund("fund_key"))
        PutVal wsFund, mapFund, newFundRow, "fund_key", fundName
        PutVal wsFund, mapFund, newFundRow, "source_tab", fundName
        PutVal wsFund, mapFund, newFundRow, "fund_legal_name", fundName
        PutVal wsFund, mapFund, newFundRow, "preqin_fund_name", s("PreqinName")
        PutVal wsFund, mapFund, newFundRow, "preqin_manager", s("Manager")
        PutVal wsFund, mapFund, newFundRow, "hollyport_fund_type", s("Strat")
        PutVal wsFund, mapFund, newFundRow, "geography", s("Geog")
        PutVal wsFund, mapFund, newFundRow, "vintage", s("Vint")
        PutVal wsFund, mapFund, newFundRow, "carry_terms_raw", s("Carry")
        PutVal wsFund, mapFund, newFundRow, "fees_terms_raw", s("Fees")
        PutVal wsFund, mapFund, newFundRow, "fund_size", s("SizeM")
        ' fund_currency / underlying_fund_currency left blank - Summary only carries an FX rate, not a currency code
        LogSkipped fundName, "fund_currency", s("FX"), _
            "Summary has an FX rate (col L), not a currency code - enter the code in fund.fund_currency AND position.underlying_fund_currency (Fund View C22 reads the latter)"

        ' --- Append to position table ---
        Dim newPosRow As Long
        newPosRow = NextEmptyRow(wsPosition, mapPosition("position_id"))
        PutVal wsPosition, mapPosition, newPosRow, "position_id", fundName
        PutVal wsPosition, mapPosition, newPosRow, "source_tab", fundName
        PutVal wsPosition, mapPosition, newPosRow, "reference_date", refDate
        PutVal wsPosition, mapPosition, newPosRow, "completion_date", ToDateOrRaw(s("TransferDate"))
        PutVal wsPosition, mapPosition, newPosRow, "commitment", s("Commit")
        PutVal wsPosition, mapPosition, newPosRow, "uncalled_at_reference_date", s("Uncalled")
        PutVal wsPosition, mapPosition, newPosRow, "headline_price", s("HeadlinePrice")
        PutVal wsPosition, mapPosition, newPosRow, "percentage_of_fund_held", s("PctFundHeld")
        PutVal wsPosition, mapPosition, newPosRow, "pre_completion_dist_forecast", s("PreCompDists")
        PutVal wsPosition, mapPosition, newPosRow, "pre_completion_calls_forecast", s("PreCompCalls")
        PutVal wsPosition, mapPosition, newPosRow, "post_completion_calls_forecast", s("PostCompCalls")
        PutVal wsPosition, mapPosition, newPosRow, "recycling_dist_forecast", s("TwelveMonthDist")
        PutVal wsPosition, mapPosition, newPosRow, "effective_nav_forecast", s("EffNAVAcquired")
        PutVal wsPosition, mapPosition, newPosRow, "implied_nav_forecast", s("ImpNAVAcquired")
        PutVal wsPosition, mapPosition, newPosRow, "date_of_update", Now
        PutVal wsPosition, mapPosition, newPosRow, "updated_by", Application.UserName
        PutVal wsPosition, mapPosition, newPosRow, "notes", "Added via Onboarder macro " & Format(Now, "yyyy-mm-dd")
        PutVal wsPosition, mapPosition, newPosRow, "position_key", fundName & "_" & Format(Now, "yyyy-mm-dd")

        If Not IsEmptyValue(s("NAV")) Then
            LogSkipped fundName, "NAV (Summary col Q)", s("NAV"), _
                "Fund View NAV comes from valuation_history, which this macro does not write - add the reference-date NAV there"
        End If

        fundsAdded = fundsAdded + 1

        ' --- Company block: Detailed funds get their companies, Standard funds get the placeholder block ---
        Dim isDetailed As Boolean, detRow As Long
        isDetailed = (UCase(SafeStr(s("ValMethodology"))) = "DETAILED")
        detRow = FindDetailedRow(wsDetailed, fundName)

        If isDetailed And detRow = 0 Then
            LogSkipped fundName, "(companies)", s("ValMethodology"), _
                "Valuation Methodology is Detailed but no matching row was found on the Detailed sheet - Standard placeholder block written instead"
            isDetailed = False
        ElseIf (Not isDetailed) And detRow > 0 Then
            LogSkipped fundName, "(companies)", s("ValMethodology"), _
                "Valuation Methodology is not Detailed, so the row on the Detailed sheet was ignored - Standard placeholder block written"
        End If

        If isDetailed Then
            WriteDetailedBlock wsDetailed, detRow, fundName, s, refDate
        Else
            WriteStandardBlock fundName, s, refDate
        End If

NextSummaryRow:
    Next r

    MsgBox "Done." & vbCrLf & vbCrLf & _
           fundsAdded & " fund(s) added" & vbCrLf & _
           mCompaniesAdded & " company / Net Curr row(s) added" & vbCrLf & _
           mRollupsAdded & " rollup row(s) added" & vbCrLf & _
           mNavRowsAdded & " company_nav row(s) added" & vbCrLf & vbCrLf & _
           "Check the 'Import Log' sheet in this workbook for anything that couldn't be mapped." & vbCrLf & _
           "Nothing has been saved yet - review the Flat File before saving it.", _
           vbInformation, "Import complete"

End Sub

' =====================================================================================
' Standard fund: the placeholder block every Standard fund in the Flat File carries.
' Rows 116/117/119 and the Net Curr row 120 hold NAV 1 and multiple = fund upside, so
' Fund View "Upside" (F35 = R116) resolves to that multiple and Total Return works.
' =====================================================================================
Private Sub WriteStandardBlock(ByVal fundName As String, ByVal s As Object, ByVal refDate As Variant)
    Dim upside As Double
    upside = ResolveFundMultiple(fundName, s("NAVPlusCallsUpside"))

    WriteRollupRow fundName, FV_ROW_NET_NAV, "Net NAV", 1#, 0#, upside
    WriteRollupRow fundName, FV_ROW_CARRY, "Carry", 0#, 0#, upside
    WriteRollupRow fundName, FV_ROW_MGMT_FEE, "Management Fee", 0#, 0#, 0#
    WriteRollupRow fundName, FV_ROW_GROSS_NAV, "Gross NAV", 1#, 0#, upside

    WriteNetCurrRows fundName, SafeStr(s("Geog")), 1#, 0#, upside, refDate
End Sub

' =====================================================================================
' Detailed fund: unpivot the Detailed sheet's company slots onto Fund View rows 121-159,
' Net Curr onto row 120, and the fund-level trailing columns onto the rollup rows.
' =====================================================================================
Private Sub WriteDetailedBlock(ByVal wsDetailed As Worksheet, ByVal detRow As Long, ByVal fundName As String, ByVal s As Object, ByVal refDate As Variant)

    Dim sumNav As Double, sumCalls As Double, sumExpected As Double
    sumNav = 0: sumCalls = 0: sumExpected = 0

    ' --- Net Curr row 120 (fund-level trailing columns) ---
    Dim netCurr As Double, ncCalls As Double
    netCurr = NumericOrDefault(fundName, "Net Curr (OZ)", wsDetailed.Cells(detRow, DET_COL_NET_CURR).Value, 0#)
    ncCalls = NumericOrDefault(fundName, "F'Cast Calls (PE)", wsDetailed.Cells(detRow, DET_COL_FCAST_CALLS).Value, 0#)
    WriteNetCurrRows fundName, SafeStr(wsDetailed.Cells(detRow, DET_COL_NET_CURR_GEOG).Value), netCurr, ncCalls, 1#, refDate
    sumNav = sumNav + netCurr
    sumCalls = sumCalls + ncCalls
    sumExpected = sumExpected + (netCurr + ncCalls) * 1#

    Dim netCurrReinvVal As Variant
    netCurrReinvVal = wsDetailed.Cells(detRow, DET_COL_NET_CURR_REINVEST).Value
    If Not IsEmptyValue(netCurrReinvVal) Then
        LogSkipped fundName, "Net Curr Reinvested", netCurrReinvVal, "No matching column in company or position_company - nothing on the Fund View reads a reinvested flag"
    End If

    ' --- Company rows 121+ ---
    Dim slot As Long, fvRow As Long
    fvRow = FV_ROW_FIRST_COMPANY
    For slot = 1 To DET_SLOT_COUNT
        Dim companyName As String
        companyName = SafeStr(wsDetailed.Cells(detRow, DET_BLOCK_COMPANIES + slot - 1).Value)
        If companyName <> "" Then
            If fvRow > FV_ROW_LAST_COMPANY Then
                LogSkipped fundName, "Company slot " & slot & ": " & companyName, "", _
                    "Fund View company block only runs to row " & FV_ROW_LAST_COMPANY & " (39 companies) - not written"
            Else
                Dim companyId As String, slotLabel As String
                companyId = "NEW-" & SanitizeKey(fundName) & "-" & slot
                slotLabel = " (company slot " & slot & ": " & companyName & ")"

                Dim navVal As Double, callsVal As Double, upsideVal As Variant
                navVal = NumericOrDefault(fundName, "NAV" & slotLabel, wsDetailed.Cells(detRow, DET_BLOCK_NAV + slot - 1).Value, 0#)
                callsVal = NumericOrDefault(fundName, "Forecast calls (block headed 'Reinvested')" & slotLabel, _
                                            wsDetailed.Cells(detRow, DET_BLOCK_FCAST_CALLS + slot - 1).Value, 0#)
                upsideVal = ResolveCompanyMultiple(fundName, "Overall Upside" & slotLabel, wsDetailed.Cells(detRow, DET_BLOCK_UPSIDE + slot - 1).Value)

                WriteCompanyRow fundName, fvRow, "company", companyId, companyName, _
                    SafeStr(wsDetailed.Cells(detRow, DET_BLOCK_GEO + slot - 1).Value), _
                    SafeStr(wsDetailed.Cells(detRow, DET_BLOCK_INDUSTRY + slot - 1).Value), _
                    SafeStr(wsDetailed.Cells(detRow, DET_BLOCK_CLASSIFICATION + slot - 1).Value), _
                    SafeStr(wsDetailed.Cells(detRow, DET_BLOCK_PUBPRIV + slot - 1).Value)

                WritePCRow fundName, fvRow, "company", companyId, navVal, callsVal, upsideVal, _
                    fvRow - FV_ROW_NET_CURR + 1, _
                    SafeStr(wsDetailed.Cells(detRow, DET_BLOCK_VALMETHOD + slot - 1).Value), _
                    SafeStr(wsDetailed.Cells(detRow, DET_BLOCK_INFOQUALITY + slot - 1).Value)

                WriteCompanyNavRow fundName, fvRow, companyId, companyName, refDate, navVal

                sumNav = sumNav + navVal
                sumCalls = sumCalls + callsVal
                If IsNumeric(upsideVal) And Not IsEmpty(upsideVal) Then
                    sumExpected = sumExpected + (navVal + callsVal) * CDbl(upsideVal)
                End If

                mCompaniesAdded = mCompaniesAdded + 1
                fvRow = fvRow + 1
            End If
        End If
    Next slot

    ' --- Rollup rows 116-119 ---
    Dim carryAmt As Double, carryMult As Double, feeAmt As Double, feeMultValue As Double, feeMult As Double
    carryAmt = NumericOrDefault(fundName, "Carry (PA)", wsDetailed.Cells(detRow, DET_COL_CARRY).Value, 0#)
    carryMult = NumericOrDefault(fundName, "Carry Multiple (PB)", wsDetailed.Cells(detRow, DET_COL_CARRY_MULTIPLE).Value, 1#)
    feeAmt = NumericOrDefault(fundName, "Fees (PC)", wsDetailed.Cells(detRow, DET_COL_FEES).Value, 0#)
    feeMultValue = NumericOrDefault(fundName, "F'Cast Fees Multiplier Value (PD)", wsDetailed.Cells(detRow, DET_COL_FCAST_FEES_MULT).Value, 0#)

    ' Legacy fund-tab formula for the Management Fee multiple (left in Detailed!A1):
    ' IF(FeesMultiplierValue = 0, 1, FeesMultiplierValue / FeeNAV)
    If feeMultValue = 0 Then
        feeMult = 1#
    ElseIf feeAmt = 0 Then
        feeMult = 1#
        LogSkipped fundName, "F'Cast Fees Multiplier Value (PD)", feeMultValue, _
            "Fees (PC) is zero so the fee multiple (PD / PC) cannot be derived - wrote 1 on row " & FV_ROW_MGMT_FEE
    Else
        feeMult = feeMultValue / feeAmt
    End If

    Dim grossMult As Double, netNav As Double, netMult As Double
    grossMult = SafeRatio(sumExpected, sumNav + sumCalls, 1#)
    netNav = sumNav + carryAmt + feeAmt
    netMult = SafeRatio(sumExpected + carryAmt * carryMult + feeAmt * feeMult, netNav + sumCalls, 1#)

    WriteRollupRow fundName, FV_ROW_NET_NAV, "Net NAV", netNav, sumCalls, netMult
    WriteRollupRow fundName, FV_ROW_CARRY, "Carry", carryAmt, 0#, carryMult
    WriteRollupRow fundName, FV_ROW_MGMT_FEE, "Management Fee", feeAmt, 0#, feeMult
    WriteRollupRow fundName, FV_ROW_GROSS_NAV, "Gross NAV", sumNav, sumCalls, grossMult
End Sub

' =====================================================================================
' Row writers
' =====================================================================================

' Rollup row on position_company (Fund View rows 116-119). gross_nav_input carries the row label.
Private Sub WriteRollupRow(ByVal fundName As String, ByVal fvRow As Long, ByVal label As String, ByVal refNav As Double, ByVal fcalls As Double, ByVal mult As Double)
    WritePCRow fundName, fvRow, "rollup", "", refNav, fcalls, mult, label, "", ""
    mRollupsAdded = mRollupsAdded + 1
End Sub

' Net Curr line: company row + position_company row (Fund View row 120) + company_nav row.
Private Sub WriteNetCurrRows(ByVal fundName As String, ByVal geo As String, ByVal refNav As Double, ByVal fcalls As Double, ByVal mult As Double, ByVal refDate As Variant)
    Dim ncId As String, ncName As String
    ncId = "NC_" & SanitizeKey(fundName)
    ncName = "Net Curr (" & fundName & ")"
    WriteCompanyRow fundName, FV_ROW_NET_CURR, "net_current_assets", ncId, ncName, geo, "Net Curr", "Net Curr", "Net Curr"
    WritePCRow fundName, FV_ROW_NET_CURR, "net_current_assets", ncId, refNav, fcalls, mult, 1#, "Net Curr", "Net Curr"
    WriteCompanyNavRow fundName, FV_ROW_NET_CURR, ncId, ncName, refDate, refNav
    mCompaniesAdded = mCompaniesAdded + 1
End Sub

Private Sub WriteCompanyRow(ByVal fundName As String, ByVal fvRow As Long, ByVal lineType As String, _
                            ByVal companyId As String, ByVal companyName As String, _
                            ByVal geo As String, ByVal sector As String, ByVal strat As String, ByVal pubPriv As String)
    Dim r As Long
    r = NextEmptyRow(mWsCompany, mMapCompany("source_tab"))
    PutVal mWsCompany, mMapCompany, r, "source_tab", fundName
    PutVal mWsCompany, mMapCompany, r, "source_row", fvRow
    PutVal mWsCompany, mMapCompany, r, "line_type", lineType
    PutVal mWsCompany, mMapCompany, r, "lookup_key", fundName & "|" & fvRow
    PutVal mWsCompany, mMapCompany, r, "company_id", companyId
    PutVal mWsCompany, mMapCompany, r, "company_name", companyName
    PutVal mWsCompany, mMapCompany, r, "geography", geo
    PutVal mWsCompany, mMapCompany, r, "sector", sector
    PutVal mWsCompany, mMapCompany, r, "strategy", strat
    PutVal mWsCompany, mMapCompany, r, "public_private", pubPriv
    PutVal mWsCompany, mMapCompany, r, "exited", "N"
End Sub

' gnavInput: Fund View col I ordinal (1 = Net Curr row) on company rows, the row label on rollup rows.
Private Sub WritePCRow(ByVal fundName As String, ByVal fvRow As Long, ByVal lineType As String, ByVal companyId As String, _
                       ByVal refNav As Variant, ByVal fcalls As Variant, ByVal mult As Variant, ByVal gnavInput As Variant, _
                       ByVal method As String, ByVal quality As String)
    Dim r As Long
    r = NextEmptyRow(mWsPC, mMapPC("source_tab"))
    PutVal mWsPC, mMapPC, r, "source_tab", fundName
    PutVal mWsPC, mMapPC, r, "position_id", fundName
    PutVal mWsPC, mMapPC, r, "source_row", fvRow
    PutVal mWsPC, mMapPC, r, "line_type", lineType
    PutVal mWsPC, mMapPC, r, "lookup_key", fundName & "|" & fvRow
    If companyId <> "" Then PutVal mWsPC, mMapPC, r, "company_id", companyId
    PutVal mWsPC, mMapPC, r, "ref_date_nav", refNav
    PutVal mWsPC, mMapPC, r, "forecast_calls", fcalls
    PutVal mWsPC, mMapPC, r, "multiple", mult
    PutVal mWsPC, mMapPC, r, "gross_nav_input", gnavInput
    If method <> "" Then PutVal mWsPC, mMapPC, r, "methodology", method
    If quality <> "" Then PutVal mWsPC, mMapPC, r, "information_quality", quality
End Sub

' One NAV point per company at the reference date, keyed fund|row|yyyy-mm-dd (Fund View cols AT onward).
Private Sub WriteCompanyNavRow(ByVal fundName As String, ByVal fvRow As Long, ByVal companyId As String, _
                               ByVal companyName As String, ByVal refDate As Variant, ByVal navValue As Double)
    If mWsNav Is Nothing Then Exit Sub
    If Not IsDate(refDate) Then
        LogSkipped fundName, "company_nav" & " (row " & fvRow & ": " & companyName & ")", navValue, _
            "Summary Val Date (col M) is not a date - company_nav row not written"
        Exit Sub
    End If
    Dim r As Long
    r = NextEmptyRow(mWsNav, mMapNav("source_tab"))
    PutVal mWsNav, mMapNav, r, "source_tab", fundName
    PutVal mWsNav, mMapNav, r, "position_id", fundName
    PutVal mWsNav, mMapNav, r, "source_row", fvRow
    PutVal mWsNav, mMapNav, r, "lookup_key", fundName & "|" & fvRow & "|" & Format(CDate(refDate), "yyyy-mm-dd")
    PutVal mWsNav, mMapNav, r, "company_id", companyId
    PutVal mWsNav, mMapNav, r, "company_name", companyName
    PutVal mWsNav, mMapNav, r, "Nav_date", CDate(refDate)
    PutVal mWsNav, mMapNav, r, "nav_value", navValue
    mNavRowsAdded = mNavRowsAdded + 1
End Sub

' =====================================================================================
' Value resolution helpers
' =====================================================================================

' Summary "NAV plus Calls Upside" (col Z) as the fund multiple. Accepted only when it
' looks like a multiple; otherwise 1 is written and the raw value logged for manual entry.
Private Function ResolveFundMultiple(ByVal fundName As String, ByVal v As Variant) As Double
    If IsEmptyValue(v) Then
        LogSkipped fundName, "NAV plus Calls Upside (Summary col Z)", "", _
            "Blank - wrote multiple 1 on rows " & FV_ROW_NET_NAV & "/" & FV_ROW_CARRY & "/" & FV_ROW_GROSS_NAV & "/" & FV_ROW_NET_CURR & " - set manually"
        ResolveFundMultiple = 1#
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 And CDbl(v) <= MAX_PLAUSIBLE_MULTIPLE Then
            ResolveFundMultiple = CDbl(v)
        Else
            LogSkipped fundName, "NAV plus Calls Upside (Summary col Z)", v, _
                "Does not look like a return multiple (expected 0 < x <= " & MAX_PLAUSIBLE_MULTIPLE & ") - wrote multiple 1, set position_company.multiple on rows " & _
                FV_ROW_NET_NAV & "/" & FV_ROW_CARRY & "/" & FV_ROW_GROSS_NAV & "/" & FV_ROW_NET_CURR & " manually"
            ResolveFundMultiple = 1#
        End If
    Else
        LogSkipped fundName, "NAV plus Calls Upside (Summary col Z)", v, "Not numeric - wrote multiple 1 - set manually"
        ResolveFundMultiple = 1#
    End If
End Function

' Detailed "Overall Upside" -> position_company.multiple. Numeric is written as-is,
' "nm" becomes 0 (the legacy fund-tab formula did the same), blank stays blank, anything else is logged.
Private Function ResolveCompanyMultiple(ByVal fundName As String, ByVal fieldName As String, ByVal v As Variant) As Variant
    If IsEmptyValue(v) Then
        LogSkipped fundName, fieldName, "", "Blank - position_company.multiple left empty (Fund View Expected Proceeds will be 0)"
        ResolveCompanyMultiple = Empty
    ElseIf IsNumeric(v) Then
        ResolveCompanyMultiple = CDbl(v)
    ElseIf UCase(SafeStr(v)) = "NM" Then
        ResolveCompanyMultiple = 0#
    Else
        LogSkipped fundName, fieldName, v, "Not numeric - position_company.multiple left empty"
        ResolveCompanyMultiple = Empty
    End If
End Function

' Numeric value or a default; non-blank non-numeric values are logged.
Private Function NumericOrDefault(ByVal fundName As String, ByVal fieldName As String, ByVal v As Variant, ByVal dflt As Double) As Double
    If IsEmptyValue(v) Then
        NumericOrDefault = dflt
    ElseIf IsNumeric(v) Then
        NumericOrDefault = CDbl(v)
    Else
        LogSkipped fundName, fieldName, v, "Not numeric - wrote " & dflt & " instead"
        NumericOrDefault = dflt
    End If
End Function

Private Function SafeRatio(ByVal num As Double, ByVal den As Double, ByVal dflt As Double) As Double
    If den = 0 Then
        SafeRatio = dflt
    Else
        SafeRatio = num / den
    End If
End Function

' Returns a real Date when the cell holds a date or a parseable date string, else the raw value.
Private Function ToDateOrRaw(ByVal v As Variant) As Variant
    If IsEmptyValue(v) Then
        ToDateOrRaw = Empty
    ElseIf IsDate(v) Then
        ToDateOrRaw = CDate(v)
    Else
        ToDateOrRaw = v
    End If
End Function

Private Function IsEmptyValue(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        IsEmptyValue = True
    ElseIf VarType(v) = vbString Then
        IsEmptyValue = (Trim(CStr(v)) = "")
    Else
        IsEmptyValue = False
    End If
End Function

' --- Converts a cell value to a trimmed string, treating #REF!/#N/A/etc as blank instead of raising a type mismatch ---
Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        SafeStr = ""
    Else
        SafeStr = Trim(CStr(v))
    End If
End Function

' --- Reads the Summary sheet fields for one row into a Dictionary keyed by short name ---
Private Function ReadSummaryRow(ws As Worksheet, r As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "FundName", ws.Cells(r, 1).Value
    d.Add "PreqinName", ws.Cells(r, 2).Value
    d.Add "ValMethodology", ws.Cells(r, 3).Value
    d.Add "DetailedTabName", ws.Cells(r, 4).Value
    d.Add "Manager", ws.Cells(r, 5).Value
    d.Add "Strat", ws.Cells(r, 6).Value
    d.Add "Geog", ws.Cells(r, 7).Value
    d.Add "Vint", ws.Cells(r, 8).Value
    d.Add "SizeM", ws.Cells(r, 9).Value
    d.Add "Carry", ws.Cells(r, 10).Value
    d.Add "Fees", ws.Cells(r, 11).Value
    d.Add "FX", ws.Cells(r, 12).Value
    d.Add "ValDate", ws.Cells(r, 13).Value
    d.Add "TransferDate", ws.Cells(r, 14).Value
    d.Add "Commit", ws.Cells(r, 15).Value
    d.Add "Uncalled", ws.Cells(r, 16).Value
    d.Add "NAV", ws.Cells(r, 17).Value
    d.Add "EffNAVAcquired", ws.Cells(r, 18).Value
    d.Add "ImpNAVAcquired", ws.Cells(r, 19).Value
    d.Add "PctFundHeld", ws.Cells(r, 20).Value
    d.Add "HeadlinePrice", ws.Cells(r, 21).Value
    d.Add "PreCompDists", ws.Cells(r, 22).Value
    d.Add "PreCompCalls", ws.Cells(r, 23).Value
    d.Add "PostCompCalls", ws.Cells(r, 24).Value
    d.Add "TwelveMonthDist", ws.Cells(r, 25).Value
    d.Add "NAVPlusCallsUpside", ws.Cells(r, 26).Value
    Set ReadSummaryRow = d
End Function

' --- Finds the Detailed sheet row for a given fund name (column A, from DET_DATA_START_ROW). Returns 0 if not found. ---
Private Function FindDetailedRow(ByVal ws As Worksheet, ByVal fundName As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, DET_COL_FUND_NAME).End(xlUp).Row
    For r = DET_DATA_START_ROW To lastRow
        If SafeStr(ws.Cells(r, DET_COL_FUND_NAME).Value) = fundName Then
            FindDetailedRow = r
            Exit Function
        End If
    Next r
    FindDetailedRow = 0
End Function

' =====================================================================================
' Table plumbing
' =====================================================================================

' --- Builds a Dictionary of header name -> column index for row 1 of a target-table sheet (case-insensitive) ---
Private Function GetHeaderMap(ws As Worksheet) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' vbTextCompare
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        Dim h As String
        h = Trim(CStr(ws.Cells(1, c).Value))
        If h <> "" And Not d.Exists(h) Then d.Add h, c
    Next c
    Set GetHeaderMap = d
End Function

' --- Returns a message line per required header missing from the map, "" if all present ---
Private Function CheckHeaders(ByVal map As Object, ByVal tableName As String, ByVal requiredCsv As String) As String
    Dim parts() As String, i As Long, msg As String
    parts = Split(requiredCsv, ",")
    For i = LBound(parts) To UBound(parts)
        If Not map.Exists(Trim(parts(i))) Then msg = msg & "  - " & tableName & "." & Trim(parts(i)) & vbCrLf
    Next i
    CheckHeaders = msg
End Function

' --- Writes a value to (row, header) if the header exists on the sheet; silently skips unknown headers ---
Private Sub PutVal(ByVal ws As Worksheet, ByVal map As Object, ByVal r As Long, ByVal header As String, ByVal v As Variant)
    If map.Exists(header) Then ws.Cells(r, map(header)).Value = v
End Sub

' --- First empty row below the header, based on a column that should never be blank for a real row ---
Private Function NextEmptyRow(ByVal ws As Worksheet, ByVal anchorCol As Long) As Long
    NextEmptyRow = ws.Cells(ws.Rows.Count, anchorCol).End(xlUp).Row + 1
End Function

' --- Makes a fund name safe to use inside a generated key (strip characters that are awkward in keys) ---
Private Function SanitizeKey(ByVal s As String) As String
    Dim result As String, i As Long, ch As String
    result = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch Like "[A-Za-z0-9]" Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next i
    SanitizeKey = result
End Function

Private Function GetOrCreateLogSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("Import Log")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = "Import Log"
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Fund"
        ws.Cells(1, 3).Value = "Field"
        ws.Cells(1, 4).Value = "Value"
        ws.Cells(1, 5).Value = "Reason not written to Flat File"
        ws.Rows(1).Font.Bold = True
    End If
    Set GetOrCreateLogSheet = ws
End Function

Private Sub LogSkipped(ByVal fundName As String, ByVal fieldName As String, ByVal val As Variant, ByVal reason As String)
    Dim r As Long
    r = mWsLog.Cells(mWsLog.Rows.Count, 1).End(xlUp).Row + 1
    mWsLog.Cells(r, 1).Value = Now
    mWsLog.Cells(r, 2).Value = fundName
    mWsLog.Cells(r, 3).Value = fieldName
    mWsLog.Cells(r, 4).Value = SafeStr(val)
    mWsLog.Cells(r, 5).Value = reason
End Sub
