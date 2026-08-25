Attribute VB_Name = "modExamples"
'==============================================================================
' modExamples - runnable examples
'------------------------------------------------------------------------------
' Import this alongside the modules in \src and run any of the subs below from
' the VBE (F5) or Alt+F8. Each one is written to explain itself and to fail
' politely when the data it wants is not there.
'
' This module DOES depend on the others - it is the one place in the repository
' that does.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' 1. The shape every long-running macro should have: state saved, progress
'    shown, timer running, and Excel handed back in one piece even if the code
'    blows up in the middle.
'------------------------------------------------------------------------------
Public Sub Example_LongJobPattern()
    Dim ws As Worksheet
    Dim lastRowNumber As Long
    Dim r As Long
    Dim blanks As Long

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    lastRowNumber = LastRow(ws, 1)
    If lastRowNumber < 2 Then
        MsgBox "Put a header and some values in column A first.", vbInformation
        Exit Sub
    End If

    StartTimer
    FastMode True
    On Error GoTo Cleanup

    For r = 2 To lastRowNumber
        If r Mod 250 = 0 Then ShowProgress r, lastRowNumber, "Scanning column A"
        If IsBlank(ws.Cells(r, 1).value) Then blanks = blanks + 1
    Next r

Cleanup:
    FastMode False
    ClearStatusBar

    If Err.Number <> 0 Then
        MsgBox ErrorInfo("Example_LongJobPattern"), vbExclamation
    Else
        MsgBox "Checked " & Format$(lastRowNumber - 1, "#,##0") & " rows in " & ElapsedText & "." & _
               vbNewLine & Format$(blanks, "#,##0") & " were blank.", vbInformation
    End If
End Sub

'------------------------------------------------------------------------------
' 2. Group a column by another and write the summary to its own sheet - the
'    SUMIF-per-row pattern replaced by a single pass over the data.
'    Wants columns headed Region and Amount on the active sheet.
'------------------------------------------------------------------------------
Public Sub Example_GroupAndReport()
    Dim ws As Worksheet
    Dim totals As Object
    Dim summary As Variant
    Dim regionColumn As Long, amountColumn As Long

    Set ws = ActiveSheet
    regionColumn = FindColumnByHeader(ws, "Region")
    amountColumn = FindColumnByHeader(ws, "Amount")

    If regionColumn = 0 Or amountColumn = 0 Then
        MsgBox "This example wants columns headed 'Region' and 'Amount' " & _
               "on the active sheet.", vbInformation
        Exit Sub
    End If

    Set totals = GroupSum(DataRange(ws), regionColumn, amountColumn, True)
    If totals.count = 0 Then MsgBox "No data to summarise.", vbInformation: Exit Sub

    summary = WithHeader(DictToArray(totals, True), "Region", "Total")
    DumpToSheet summary, "Summary by Region", ws.Parent

    MsgBox totals.count & " regions summarised.", vbInformation
End Sub

'------------------------------------------------------------------------------
' 3. One index instead of 100,000 VLOOKUPs. Reads a lookup table from a sheet
'    called Prices (code in column A, price in column B) and fills in the price
'    column of the active sheet.
'------------------------------------------------------------------------------
Public Sub Example_FastLookup()
    Dim ws As Worksheet, prices As Worksheet
    Dim index As Object
    Dim codes As Variant
    Dim results() As Variant
    Dim codeColumn As Long, priceColumn As Long
    Dim firstRow As Long, lastRowNumber As Long
    Dim r As Long, misses As Long
    Dim found As Variant

    Set ws = ActiveSheet
    Set prices = GetSheet("Prices", ws.Parent)
    If prices Is Nothing Then
        MsgBox "This example wants a sheet called 'Prices' with codes in " & _
               "column A and prices in column B.", vbInformation
        Exit Sub
    End If

    codeColumn = FindColumnByHeader(ws, "Code")
    priceColumn = FindColumnByHeader(ws, "Price")
    If codeColumn = 0 Or priceColumn = 0 Then
        MsgBox "This example wants columns headed 'Code' and 'Price' on the " & _
               "active sheet.", vbInformation
        Exit Sub
    End If

    Set index = BuildLookup(DataRange(prices), 1, 2, True)

    firstRow = 2
    lastRowNumber = LastRow(ws, codeColumn)
    If lastRowNumber < firstRow Then Exit Sub

    codes = RangeToArray(ws.Range(ws.Cells(firstRow, codeColumn), ws.Cells(lastRowNumber, codeColumn)))
    ReDim results(1 To UBound(codes, 1), 1 To 1)

    FastMode True
    On Error GoTo Cleanup

    For r = 1 To UBound(codes, 1)
        found = LookupValue(index, codes(r, 1), Empty)
        If IsEmpty(found) Then
            results(r, 1) = "not found"
            misses = misses + 1
        Else
            results(r, 1) = found
        End If
    Next r

    WriteArray results, ws.Cells(firstRow, priceColumn)

Cleanup:
    FastMode False
    If Err.Number <> 0 Then
        MsgBox ErrorInfo("Example_FastLookup"), vbExclamation
    Else
        MsgBox Format$(UBound(codes, 1), "#,##0") & " rows filled, " & _
               misses & " code(s) not in the price list.", vbInformation
    End If
End Sub

'------------------------------------------------------------------------------
' 4. Working days, month ends and fiscal periods.
'------------------------------------------------------------------------------
Public Sub Example_WorkingDays()
    Dim holidays As Variant
    Dim message As String

    holidays = Array(DateSerial(2024, 12, 25), DateSerial(2024, 12, 26), DateSerial(2025, 1, 1))

    message = "December 2024" & vbNewLine & String$(40, "-") & vbNewLine
    message = message & "Working days:            " & _
              WorkdaysBetween(DateSerial(2024, 12, 1), DateSerial(2024, 12, 31), holidays) & vbNewLine
    message = message & "10 working days from 20th: " & _
              Format$(AddWorkdays(DateSerial(2024, 12, 20), 10, holidays), "ddd dd mmm yyyy") & vbNewLine
    message = message & "Last working day:        " & _
              Format$(LastWorkdayOfMonth(DateSerial(2024, 12, 1), holidays), "ddd dd mmm yyyy") & vbNewLine
    message = message & "Quarter (April year end): " & _
              PeriodLabel(DateSerial(2024, 12, 1), 4) & vbNewLine
    message = message & "Prior month end:         " & _
              Format$(EndOfMonth(DateSerial(2024, 12, 1), -1), "dd mmm yyyy")

    MsgBox message, vbInformation, "modDate"
End Sub

'------------------------------------------------------------------------------
' 5. Matching names that nearly agree - the everyday reconciliation problem.
'------------------------------------------------------------------------------
Public Sub Example_FuzzyMatch()
    Dim candidates As Variant
    Dim target As String
    Dim message As String
    Dim i As Long
    Dim score As Double, bestScore As Double
    Dim best As String

    target = "Barclays Bank PLC"
    candidates = Array("BARCLAYS BANK P.L.C.", "Barclay's Bank Plc", _
                       "HSBC Bank plc", "Lloyds Bank plc")

    For i = LBound(candidates) To UBound(candidates)
        score = SimilarityPercent(CleanText(target), CleanText(candidates(i)))
        message = message & PadRight(CStr(candidates(i)), 24) & Format$(score, "0.0") & "%" & vbNewLine
        If score > bestScore Then
            bestScore = score
            best = CStr(candidates(i))
        End If
    Next i

    MsgBox "Closest match to """ & target & """:" & vbNewLine & vbNewLine & _
           message & vbNewLine & "Best: " & best, vbInformation, "modString"
End Sub

'------------------------------------------------------------------------------
' 6. Inventory a folder tree on a new sheet, then write it out as CSV.
'------------------------------------------------------------------------------
Public Sub Example_FolderInventory()
    Dim folderPath As String
    Dim found As Collection
    Dim listing() As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim csvPath As String

    folderPath = PickFolder("Pick a folder to inventory")
    If Len(folderPath) = 0 Then Exit Sub

    Set found = ListFiles(folderPath, "*.*", True)
    If found.count = 0 Then MsgBox "No files found.", vbInformation: Exit Sub

    ReDim listing(1 To found.count + 1, 1 To 4)
    listing(1, 1) = "File"
    listing(1, 2) = "Folder"
    listing(1, 3) = "Size (KB)"
    listing(1, 4) = "Modified"

    For i = 1 To found.count
        listing(i + 1, 1) = FileNameOnly(found(i))
        listing(i + 1, 2) = ParentFolder(found(i))
        listing(i + 1, 3) = Round(FileSizeBytes(found(i)) / 1024, 1)
        listing(i + 1, 4) = FileModified(found(i))
    Next i

    Set ws = DumpToSheet(listing, "Folder inventory")
    FreezeHeader ws, 1
    AutoFitColumns ws

    csvPath = UniqueFilePath(JoinPath(folderPath, "inventory.csv"))
    If ExportRangeToCsv(DataRange(ws), csvPath) Then
        MsgBox found.count & " files listed." & vbNewLine & "CSV written to " & csvPath, vbInformation
    Else
        MsgBox found.count & " files listed. The CSV could not be written to " & csvPath, vbExclamation
    End If
End Sub

'==============================================================================
' Private helpers used by the examples
'==============================================================================
' Puts a header row on top of a two column array.
Private Function WithHeader(ByVal data As Variant, ByVal header1 As String, _
                            ByVal header2 As String) As Variant
    Dim out As Variant
    Dim r As Long, c As Long

    If Not IsArray(data) Then Exit Function

    ReDim out(1 To UBound(data, 1) + 1, 1 To UBound(data, 2))
    out(1, 1) = header1
    If UBound(data, 2) >= 2 Then out(1, 2) = header2

    For r = 1 To UBound(data, 1)
        For c = 1 To UBound(data, 2)
            out(r + 1, c) = data(r, c)
        Next c
    Next r

    WithHeader = out
End Function
