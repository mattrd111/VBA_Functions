Attribute VB_Name = "modFundHelper"
'==============================================================================
' modFundHelper - a worked example of the fund maths
'------------------------------------------------------------------------------
' Functions nobody knows about are functions nobody uses. This builds a sheet
' with a sample fund on it and every function from modFinance and modWaterfall
' written out as a live formula, so the syntax and the answer are both in front
' of you. Change the sample numbers and everything moves.
'
' Depends on: modDoctorCommon, modFinance, modWaterfall, modWorkbook, modRange
'==============================================================================
Option Explicit

Private Const FIRST_FLOW As Long = 6
Private Const LAST_FLOW As Long = 13
Private Const NAV_ROW As Long = 15
Private Const VALUED_ROW As Long = 16
Private Const HURDLE_ROW As Long = 31
Private Const CAPITAL_ROW As Long = 32
Private Const PREF_ROW As Long = 33
Private Const DISTRIBUTABLE_ROW As Long = 36
Private Const CARRY_ROW As Long = 37
Private Const CATCHUP_ROW As Long = 38
Private Const LP_TOTAL_ROW As Long = 47
Private Const GP_TOTAL_ROW As Long = 48

'------------------------------------------------------------------------------
' ShowFundMathsReference   (menu action)
'------------------------------------------------------------------------------
Public Sub ShowFundMathsReference()
    Dim book As Workbook
    Dim ws As Worksheet
    Dim data As Variant
    Dim flows As String, dates As String, index As String

    flows = "$B$" & FIRST_FLOW & ":$B$" & LAST_FLOW
    dates = "$A$" & FIRST_FLOW & ":$A$" & LAST_FLOW
    index = "$C$" & FIRST_FLOW & ":$C$" & LAST_FLOW

    ReDim data(1 To 50, 1 To 3)

    SetRow data, 1, "Fund maths reference", "", ""
    SetRow data, 2, "Every formula below is live. Change the sample fund and watch them move.", "", ""

    SetRow data, 4, "A sample fund", "", ""
    SetRow data, 5, "Date", "Cash flow", "Index level"
    SetRow data, 6, DateSerial(2019, 1, 15), -100, 100
    SetRow data, 7, DateSerial(2019, 7, 1), -150, 106
    SetRow data, 8, DateSerial(2020, 3, 31), -80, 88
    SetRow data, 9, DateSerial(2021, 6, 30), 60, 128
    SetRow data, 10, DateSerial(2022, 9, 30), 120, 121
    SetRow data, 11, DateSerial(2023, 12, 31), 90, 140
    SetRow data, 12, DateSerial(2024, 6, 28), 75, 152
    SetRow data, 13, DateSerial(2025, 6, 30), 40, 165

    SetRow data, NAV_ROW, "Residual NAV", 120, "What is still in the ground"
    SetRow data, VALUED_ROW, "Valued at", DateSerial(2025, 12, 31), ""

    SetRow data, 18, "Returns and multiples", "", ""
    SetRow data, 19, "Measure", "Result", "The formula"
    ShowFormula data, 20, "IRR on cash alone", "=FundXIRR(" & flows & "," & dates & ")"
    ShowFormula data, 21, "IRR including NAV", "=FundIRR(" & flows & "," & dates & ",$B$" & NAV_ROW & ",$B$" & VALUED_ROW & ")"
    ShowFormula data, 22, "Paid in", "=PaidIn(" & flows & ")"
    ShowFormula data, 23, "Distributed", "=Distributed(" & flows & ")"
    ShowFormula data, 24, "DPI", "=FundDPI(" & flows & ")"
    ShowFormula data, 25, "RVPI", "=RVPI(" & flows & ",$B$" & NAV_ROW & ")"
    ShowFormula data, 26, "TVPI  (MOIC is the same number)", "=TVPI(" & flows & ",$B$" & NAV_ROW & ")"
    ShowFormula data, 27, "KS-PME against the index", "=KSPME(" & flows & "," & dates & "," & index & ",$B$" & NAV_ROW & ")"
    ShowFormula data, 28, "Value of it at 8%", "=FundXNPV(0.08," & flows & "," & dates & ")"

    SetRow data, 30, "Preferred return", "", ""
    SetRow data, HURDLE_ROW, "Hurdle rate", 0.08, "Change this and everything below follows"
    ShowFormula data, CAPITAL_ROW, "Unreturned capital", "=UnreturnedCapital(" & flows & "," & dates & ",$B$" & HURDLE_ROW & ",$B$" & VALUED_ROW & ")"
    ShowFormula data, PREF_ROW, "Accrued preferred", "=AccruedPref(" & flows & "," & dates & ",$B$" & HURDLE_ROW & ",$B$" & VALUED_ROW & ")"

    SetRow data, 35, "Distribution waterfall", "", ""
    SetRow data, DISTRIBUTABLE_ROW, "To distribute now", 300, "Try changing this"
    SetRow data, CARRY_ROW, "Carry", 0.2, ""
    SetRow data, CATCHUP_ROW, "Catch-up", 1#, "100% catch-up. Try 50%"

    SetRow data, 40, "Tier", "Amount", "The formula"
    ShowFormula data, 41, "LP  return of capital", WaterfallCall("LPCapital")
    ShowFormula data, 42, "LP  preferred", WaterfallCall("LPPref")
    ShowFormula data, 43, "GP  catch-up", WaterfallCall("GPCatchUp")
    ShowFormula data, 44, "LP  share of the catch-up tier", WaterfallCall("LPCatchUp")
    ShowFormula data, 45, "LP  residual", WaterfallCall("LPResidual")
    ShowFormula data, 46, "GP  carry", WaterfallCall("GPCarry")
    ShowFormula data, LP_TOTAL_ROW, "LP total", WaterfallCall("LPTotal")
    ShowFormula data, GP_TOTAL_ROW, "GP total", WaterfallCall("GPTotal")
    ShowFormula data, 49, "LP + GP  (ties back to what was distributed)", _
                "=$B$" & LP_TOTAL_ROW & "+$B$" & GP_TOTAL_ROW
    ShowFormula data, 50, "What the LP keeps, per pound distributed", _
                "=NetToGross($B$" & DISTRIBUTABLE_ROW & ",$B$" & CAPITAL_ROW & ",$B$" & PREF_ROW & _
                ",$B$" & CARRY_ROW & ",$B$" & CATCHUP_ROW & ")"

    Set book = Application.Workbooks.Add
    Set ws = DumpToSheet(data, "Fund maths", book, False)
    If ws Is Nothing Then Exit Sub

    FormatReference ws

    MsgBox "A worked example of every fund function, with live formulas." & vbNewLine & vbNewLine & _
           "The functions come from the add-in, so they work in any workbook while it is " & _
           "installed. If you send this sheet to someone without the add-in they will see " & _
           "#NAME? - convert it to values first.", vbInformation, DOCTOR_NAME
End Sub

'==============================================================================
' Private helpers
'==============================================================================
Private Function WaterfallCall(ByVal part As String) As String
    WaterfallCall = "=Waterfall($B$" & DISTRIBUTABLE_ROW & ",$B$" & CAPITAL_ROW & ",$B$" & PREF_ROW & _
                    ",$B$" & CARRY_ROW & ",$B$" & CATCHUP_ROW & ",""" & part & """)"
End Function

Private Sub SetRow(ByRef data As Variant, ByVal r As Long, ByVal a As Variant, _
                ByVal b As Variant, ByVal c As Variant)
    data(r, 1) = a
    data(r, 2) = b
    data(r, 3) = c
End Sub

' Label in column A, the live formula in column B, and the same formula as text
' in column C so it can be read and copied.
Private Sub ShowFormula(ByRef data As Variant, ByVal r As Long, ByVal label As String, _
                 ByVal formulaText As String)
    data(r, 1) = label
    data(r, 2) = formulaText
    data(r, 3) = "'" & formulaText
End Sub

Private Sub FormatReference(ByVal ws As Worksheet)
    On Error Resume Next

    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Font.Italic = True

    BoldRow ws, 4
    BoldRow ws, 18
    BoldRow ws, 30
    BoldRow ws, 35

    ws.Range("A5:C5").Font.Bold = True
    ws.Range("A19:C19").Font.Bold = True
    ws.Range("A40:C40").Font.Bold = True

    ws.Range("A" & FIRST_FLOW & ":A" & LAST_FLOW).NumberFormat = "dd mmm yyyy"
    ws.Range("B" & VALUED_ROW).NumberFormat = "dd mmm yyyy"
    ws.Range("B" & FIRST_FLOW & ":B" & LAST_FLOW).NumberFormat = "#,##0.0;[Red](#,##0.0)"
    ws.Range("B20:B21").NumberFormat = "0.0%"
    ws.Range("B22:B23").NumberFormat = "#,##0.0"
    ws.Range("B24:B28").NumberFormat = "0.00"
    ws.Range("B" & HURDLE_ROW).NumberFormat = "0.0%"
    ws.Range("B" & CAPITAL_ROW & ":B" & PREF_ROW).NumberFormat = "#,##0.0"
    ws.Range("B" & CARRY_ROW & ":B" & CATCHUP_ROW).NumberFormat = "0%"
    ws.Range("B" & DISTRIBUTABLE_ROW).NumberFormat = "#,##0.0"
    ws.Range("B41:B49").NumberFormat = "#,##0.0"
    ws.Range("B50").NumberFormat = "0.0%"

    ws.Range("C:C").Font.Name = "Consolas"
    ws.Range("C:C").Font.Size = 9
    ws.Range("C:C").Font.Color = RGB(120, 120, 120)

    ws.Range("A49:B49").Interior.Color = RGB(242, 242, 242)

    ws.Columns("A").ColumnWidth = 38
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 62
    ws.Range("A1").Select

    On Error GoTo 0
End Sub

Private Sub BoldRow(ByVal ws As Worksheet, ByVal r As Long)
    With ws.Range("A" & r)
        .Font.Bold = True
        .Font.Size = 11
    End With
End Sub
