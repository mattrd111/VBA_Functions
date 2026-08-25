Attribute VB_Name = "modDoctorNames"
'==============================================================================
' modDoctorNames - defined name clean-up
'------------------------------------------------------------------------------
' Classifies every defined name in a workbook, reports what it found, and
' deletes the categories you approve.
'
' Categories
'   Reserved    Print_Area, Print_Titles, _FilterDatabase and friends - never
'               touched, because Excel needs them.
'   In use      Something refers to it. Never deleted.
'   Broken      Refers to #REF! and nothing refers to it. Safe to delete.
'   Broken (referenced)  Refers to #REF! but a formula still uses it. Reported
'               only - deleting it would turn #REF! into #NAME?, so fix the
'               formula instead.
'   External    Points at another workbook and nothing refers to it. These are
'               what a copied sheet leaves behind.
'   Hidden      Marked hidden and unreferenced - almost always junk.
'   Unused      Visible, local, and nothing appears to refer to it. The riskiest
'               category, so deleting it is a separate opt-in.
'
' Depends on: modDoctorCommon, modDoctorScan, modDictionary, modApp (from \src)
'==============================================================================
Option Explicit

Public Const KIND_RESERVED As String = "Reserved"
Public Const KIND_IN_USE As String = "In use"
Public Const KIND_BROKEN As String = "Broken"
Public Const KIND_BROKEN_USED As String = "Broken (referenced)"
Public Const KIND_EXTERNAL As String = "External"
Public Const KIND_HIDDEN As String = "Hidden"
Public Const KIND_UNUSED As String = "Unused"

'------------------------------------------------------------------------------
' ShowNamesReport   (menu action)
' Read-only: lists every name with its verdict and changes nothing.
'------------------------------------------------------------------------------
Public Sub ShowNamesReport()
    Dim wb As Workbook
    Dim names As Collection
    Dim summary As Object

    Set wb = TargetWorkbook(False)
    If wb Is Nothing Then Exit Sub

    StartTimer
    FastMode True
    Set names = AnalyseNames(wb, summary)
    FastMode False

    If names.count = 0 Then
        MsgBox "'" & wb.Name & "' has no defined names at all.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    ShowReport BuildNamesReport(wb, names, summary, "Defined names"), "Defined names"
End Sub

'------------------------------------------------------------------------------
' CleanNames   (menu action)
' Analyses, asks, deletes, then reports exactly what went.
'------------------------------------------------------------------------------
Public Sub CleanNames()
    Dim wb As Workbook
    Dim names As Collection
    Dim summary As Object
    Dim safeCount As Long, unusedCount As Long
    Dim deleteUnused As Boolean
    Dim deleted As Long, failed As Long
    Dim rows As Collection

    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub

    StartTimer
    FastMode True
    Set names = AnalyseNames(wb, summary)
    FastMode False

    If names.count = 0 Then
        MsgBox "'" & wb.Name & "' has no defined names at all.", vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    safeCount = DictGet(summary, KIND_BROKEN, 0) + DictGet(summary, KIND_EXTERNAL, 0) + _
                DictGet(summary, KIND_HIDDEN, 0)
    unusedCount = DictGet(summary, KIND_UNUSED, 0)

    If safeCount + unusedCount = 0 Then
        MsgBox "Nothing to clean - all " & names.count & " names are in use or reserved." & _
               vbNewLine & vbNewLine & SummaryText(summary), vbInformation, DOCTOR_NAME
        Exit Sub
    End If

    If Not Confirm(SummaryText(summary) & vbNewLine & vbNewLine & _
                   "Delete the " & safeCount & " broken, external and hidden name(s)?") Then
        ShowReport BuildNamesReport(wb, names, summary, "Defined names"), "Defined names"
        Exit Sub
    End If

    If Not OfferBackup(wb) Then Exit Sub

    If unusedCount > 0 Then
        deleteUnused = (MsgBox("Also delete the " & unusedCount & " name(s) that nothing appears " & _
                               "to refer to?" & vbNewLine & vbNewLine & _
                               "These are found by scanning formulas, conditional formatting, " & _
                               "validation, charts and other names. A name used only from VBA, " & _
                               "or built up in a formula from text, cannot be seen by that scan." & _
                               vbNewLine & vbNewLine & "Choose No to keep them and review the report.", _
                               vbYesNo + vbQuestion + vbDefaultButton2, DOCTOR_NAME) = vbYes)
    End If

    FastMode True
    Set rows = DeleteNames(names, deleteUnused, deleted, failed)
    FastMode False

    ShowReport rows, "Names deleted"

    MsgBox deleted & " name(s) deleted in " & ElapsedText & "." & _
           IIf(failed > 0, vbNewLine & failed & " could not be deleted - see the report.", "") & _
           vbNewLine & vbNewLine & "Save the workbook to keep the change.", vbInformation, DOCTOR_NAME
End Sub

'------------------------------------------------------------------------------
' AnalyseNames
' Every name with its verdict. Each entry is
'   Array(kind, localName, scope, refersTo, visible, nameObject)
' summary comes back as a dictionary of kind -> count.
'------------------------------------------------------------------------------
Public Function AnalyseNames(ByVal wb As Workbook, ByRef summary As Object) As Collection
    Dim result As Collection
    Dim usage As Object
    Dim nm As Name
    Dim localName As String, refersTo As String, scope As String
    Dim kind As String
    Dim visible As Boolean
    Dim referenced As Boolean
    Dim done As Long

    Set result = New Collection
    Set summary = NewDictionary(True)
    Set AnalyseNames = result
    If wb Is Nothing Then Exit Function
    If wb.names.count = 0 Then Exit Function

    Set usage = BuildUsageIndex(wb)

    For Each nm In wb.names
        done = done + 1
        If done Mod 50 = 0 Then ShowProgress done, wb.names.count, "Checking names"

        localName = ""
        refersTo = ""
        visible = True

        On Error Resume Next
        localName = LocalNameOf(nm.Name)
        refersTo = nm.refersTo
        visible = nm.visible
        On Error GoTo 0

        scope = ScopeOf(nm)
        referenced = usage.Exists(UCase$(localName))
        kind = Classify(localName, refersTo, visible, referenced)

        DictIncrement summary, kind, 1
        result.Add Array(kind, localName, scope, refersTo, visible, nm)
    Next nm

    ClearStatusBar
End Function

'------------------------------------------------------------------------------
' DeleteNames
' Deletes the approved categories and returns a report of what happened.
'------------------------------------------------------------------------------
Public Function DeleteNames(ByVal names As Collection, ByVal includeUnused As Boolean, _
                            ByRef deleted As Long, ByRef failed As Long) As Collection
    Dim rows As Collection
    Dim item As Variant
    Dim nm As Name
    Dim kind As String
    Dim i As Long

    Set rows = NewReport()
    ReportHeading rows, "Names deleted"
    ReportRow rows, "Result", "Name", "Scope", "Category", "Referred to"

    For i = 1 To names.count
        item = names(i)
        kind = CStr(item(0))

        If ShouldDelete(kind, includeUnused) Then
            Set nm = item(5)
            On Error Resume Next
            nm.Delete
            If Err.Number = 0 Then
                deleted = deleted + 1
                ReportRow rows, "Deleted", item(1), item(2), kind, item(3)
            Else
                failed = failed + 1
                ReportRow rows, "Failed: " & Err.Description, item(1), item(2), kind, item(3)
                Err.Clear
            End If
            On Error GoTo 0
        End If

        If i Mod 25 = 0 Then ShowProgress i, names.count, "Deleting names"
    Next i

    ClearStatusBar
    If deleted + failed = 0 Then ReportRow rows, "Nothing matched the categories chosen."

    Set DeleteNames = rows
End Function

'------------------------------------------------------------------------------
' BuildNamesReport
' The full listing, worst first.
'------------------------------------------------------------------------------
Public Function BuildNamesReport(ByVal wb As Workbook, ByVal names As Collection, _
                                 ByVal summary As Object, ByVal reportTitle As String) As Collection
    Dim rows As Collection
    Dim order As Variant
    Dim item As Variant
    Dim k As Long
    Dim kind As String

    Set rows = NewReport()
    ReportRow rows, DOCTOR_NAME & " - " & reportTitle
    ReportRow rows, "Workbook", wb.Name
    ReportRow rows, "Checked", Format$(Now, "dd mmm yyyy hh:nn")
    ReportRow rows, "Names", names.count

    ReportHeading rows, "Summary"
    ReportRow rows, "Category", "Count", "What it means"
    ReportRow rows, KIND_BROKEN, DictGet(summary, KIND_BROKEN, 0), "Refers to #REF! and unreferenced - safe to delete"
    ReportRow rows, KIND_BROKEN_USED, DictGet(summary, KIND_BROKEN_USED, 0), "Refers to #REF! but a formula still uses it - fix the formula"
    ReportRow rows, KIND_EXTERNAL, DictGet(summary, KIND_EXTERNAL, 0), "Points at another workbook and unreferenced"
    ReportRow rows, KIND_HIDDEN, DictGet(summary, KIND_HIDDEN, 0), "Hidden and unreferenced - usually left by copied sheets"
    ReportRow rows, KIND_UNUSED, DictGet(summary, KIND_UNUSED, 0), "Nothing appears to refer to it - review before deleting"
    ReportRow rows, KIND_IN_USE, DictGet(summary, KIND_IN_USE, 0), "Referenced somewhere - kept"
    ReportRow rows, KIND_RESERVED, DictGet(summary, KIND_RESERVED, 0), "Excel needs these - never deleted"

    order = Array(KIND_BROKEN, KIND_BROKEN_USED, KIND_EXTERNAL, KIND_HIDDEN, _
                  KIND_UNUSED, KIND_IN_USE, KIND_RESERVED)

    For k = LBound(order) To UBound(order)
        kind = CStr(order(k))
        If DictGet(summary, kind, 0) > 0 Then
            ReportHeading rows, kind & " (" & DictGet(summary, kind, 0) & ")"
            ReportRow rows, "Name", "Scope", "Visible", "Refers to"
            For Each item In names
                If StrComp(CStr(item(0)), kind, vbTextCompare) = 0 Then
                    ReportRow rows, item(1), item(2), IIf(item(4), "Yes", "No"), item(3)
                End If
            Next item
        End If
    Next k

    Set BuildNamesReport = rows
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Function Classify(ByVal localName As String, ByVal refersTo As String, _
                          ByVal visible As Boolean, ByVal referenced As Boolean) As String
    If IsReservedName(localName) Then Classify = KIND_RESERVED: Exit Function

    If InStr(1, refersTo, "#REF!", vbTextCompare) > 0 Then
        If referenced Then Classify = KIND_BROKEN_USED Else Classify = KIND_BROKEN
        Exit Function
    End If

    If referenced Then Classify = KIND_IN_USE: Exit Function
    If PointsOutside(refersTo) Then Classify = KIND_EXTERNAL: Exit Function
    If Not visible Then Classify = KIND_HIDDEN: Exit Function

    Classify = KIND_UNUSED
End Function

' True when the reference is to another workbook rather than this one.
Private Function PointsOutside(ByVal refersTo As String) As Boolean
    If InStr(refersTo, "[") = 0 Then Exit Function

    If InStr(1, refersTo, ".xls", vbTextCompare) > 0 Then PointsOutside = True
    If InStr(1, refersTo, ".xlsx", vbTextCompare) > 0 Then PointsOutside = True
    If InStr(1, refersTo, ".xlsm", vbTextCompare) > 0 Then PointsOutside = True
    If InStr(1, refersTo, ".xlsb", vbTextCompare) > 0 Then PointsOutside = True
    If InStr(refersTo, "\") > 0 Then PointsOutside = True
    If InStr(refersTo, "//") > 0 Then PointsOutside = True
End Function

Private Function ShouldDelete(ByVal kind As String, ByVal includeUnused As Boolean) As Boolean
    Select Case kind
        Case KIND_BROKEN, KIND_EXTERNAL, KIND_HIDDEN
            ShouldDelete = True
        Case KIND_UNUSED
            ShouldDelete = includeUnused
    End Select
End Function

Private Function SummaryText(ByVal summary As Object) As String
    Dim text As String

    text = "Defined names found:" & vbNewLine
    text = text & vbNewLine & "  Broken (#REF!)      " & DictGet(summary, KIND_BROKEN, 0)
    text = text & vbNewLine & "  Broken but in use   " & DictGet(summary, KIND_BROKEN_USED, 0)
    text = text & vbNewLine & "  External links      " & DictGet(summary, KIND_EXTERNAL, 0)
    text = text & vbNewLine & "  Hidden              " & DictGet(summary, KIND_HIDDEN, 0)
    text = text & vbNewLine & "  Unused              " & DictGet(summary, KIND_UNUSED, 0)
    text = text & vbNewLine & "  In use              " & DictGet(summary, KIND_IN_USE, 0)
    text = text & vbNewLine & "  Reserved            " & DictGet(summary, KIND_RESERVED, 0)

    SummaryText = text
End Function
