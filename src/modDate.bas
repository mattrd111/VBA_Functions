Attribute VB_Name = "modDate"
'==============================================================================
' modDate - date and period helpers
'------------------------------------------------------------------------------
' Month and quarter ends, fiscal periods, working day maths and forgiving date
' parsing.
'
' Standalone: no dependency on the other modules in this repository.
' Working day routines treat Saturday and Sunday as the weekend.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' StartOfMonth / EndOfMonth / DaysInMonth
' monthsToAdd shifts the result: EndOfMonth(d, -1) is last month end, which is
' the same as the worksheet function EOMONTH.
'------------------------------------------------------------------------------
Public Function StartOfMonth(ByVal aDate As Date, Optional ByVal monthsToAdd As Long = 0) As Date
    StartOfMonth = DateSerial(Year(aDate), Month(aDate) + monthsToAdd, 1)
End Function

Public Function EndOfMonth(ByVal aDate As Date, Optional ByVal monthsToAdd As Long = 0) As Date
    EndOfMonth = DateSerial(Year(aDate), Month(aDate) + monthsToAdd + 1, 0)
End Function

Public Function DaysInMonth(ByVal aDate As Date) As Long
    DaysInMonth = Day(EndOfMonth(aDate))
End Function

'------------------------------------------------------------------------------
' StartOfYear / EndOfYear / StartOfWeek
'------------------------------------------------------------------------------
Public Function StartOfYear(ByVal aDate As Date, Optional ByVal yearsToAdd As Long = 0) As Date
    StartOfYear = DateSerial(Year(aDate) + yearsToAdd, 1, 1)
End Function

Public Function EndOfYear(ByVal aDate As Date, Optional ByVal yearsToAdd As Long = 0) As Date
    EndOfYear = DateSerial(Year(aDate) + yearsToAdd, 12, 31)
End Function

Public Function StartOfWeek(ByVal aDate As Date, _
                            Optional ByVal firstDayOfWeek As VbDayOfWeek = vbMonday) As Date
    StartOfWeek = Int(aDate) - ((Weekday(aDate, firstDayOfWeek) - 1))
End Function

'------------------------------------------------------------------------------
' Quarter / QuarterStart / QuarterEnd
' Calendar quarters unless a fiscal year start month is supplied.
'
'   Quarter(#15-Jun-2024#)      ->  2
'   Quarter(#15-Jun-2024#, 4)   ->  1   ' April year start
'------------------------------------------------------------------------------
Public Function Quarter(ByVal aDate As Date, Optional ByVal fiscalYearStartMonth As Long = 1) As Long
    Dim offset As Long

    If fiscalYearStartMonth < 1 Or fiscalYearStartMonth > 12 Then fiscalYearStartMonth = 1
    offset = (Month(aDate) - fiscalYearStartMonth + 12) Mod 12
    Quarter = offset \ 3 + 1
End Function

Public Function QuarterStart(ByVal aDate As Date, Optional ByVal fiscalYearStartMonth As Long = 1) As Date
    Dim offset As Long

    If fiscalYearStartMonth < 1 Or fiscalYearStartMonth > 12 Then fiscalYearStartMonth = 1
    offset = (Month(aDate) - fiscalYearStartMonth + 12) Mod 12
    QuarterStart = DateSerial(Year(aDate), Month(aDate) - (offset Mod 3), 1)
End Function

Public Function QuarterEnd(ByVal aDate As Date, Optional ByVal fiscalYearStartMonth As Long = 1) As Date
    QuarterEnd = EndOfMonth(QuarterStart(aDate, fiscalYearStartMonth), 2)
End Function

'------------------------------------------------------------------------------
' FiscalYear / FiscalYearLabel / PeriodLabel
' The fiscal year is labelled by the calendar year it ends in, so with an April
' start 15-Jun-2024 falls in FY2025.
'------------------------------------------------------------------------------
Public Function FiscalYear(ByVal aDate As Date, Optional ByVal fiscalYearStartMonth As Long = 1) As Long
    If fiscalYearStartMonth < 1 Or fiscalYearStartMonth > 12 Then fiscalYearStartMonth = 1

    If fiscalYearStartMonth = 1 Then
        FiscalYear = Year(aDate)
    ElseIf Month(aDate) >= fiscalYearStartMonth Then
        FiscalYear = Year(aDate) + 1
    Else
        FiscalYear = Year(aDate)
    End If
End Function

Public Function FiscalYearLabel(ByVal aDate As Date, Optional ByVal fiscalYearStartMonth As Long = 1) As String
    If fiscalYearStartMonth = 1 Then
        FiscalYearLabel = CStr(Year(aDate))
    Else
        FiscalYearLabel = "FY" & CStr(FiscalYear(aDate, fiscalYearStartMonth))
    End If
End Function

'   PeriodLabel(#15-Jun-2024#, 4)  ->  "Q1 FY2025"
Public Function PeriodLabel(ByVal aDate As Date, Optional ByVal fiscalYearStartMonth As Long = 1) As String
    PeriodLabel = "Q" & Quarter(aDate, fiscalYearStartMonth) & " " & _
                  FiscalYearLabel(aDate, fiscalYearStartMonth)
End Function

'------------------------------------------------------------------------------
' IsWeekend / IsWorkday
' The two weekend days can be changed for markets that do not rest at the
' weekend, for example IsWeekend(d, vbFriday, vbSaturday).
'------------------------------------------------------------------------------
Public Function IsWeekend(ByVal aDate As Date, _
                          Optional ByVal weekendDay1 As VbDayOfWeek = vbSaturday, _
                          Optional ByVal weekendDay2 As VbDayOfWeek = vbSunday) As Boolean
    Dim dow As Long

    dow = Weekday(aDate, vbSunday)
    IsWeekend = (dow = weekendDay1 Or dow = weekendDay2)
End Function

Public Function IsWorkday(ByVal aDate As Date, Optional ByVal holidays As Variant) As Boolean
    Dim hol As Object

    If IsWeekend(aDate) Then Exit Function
    Set hol = HolidaySet(holidays)
    IsWorkday = Not hol.Exists(CLng(Int(aDate)))
End Function

'------------------------------------------------------------------------------
' WorkdaysBetween
' Working days from start to end inclusive of both, like NETWORKDAYS. Negative
' when the end date is before the start date. holidays may be a range, an array
' or a single date, and may be left out.
'
'   WorkdaysBetween(#01-Jul-2024#, #31-Jul-2024#, ws.Range("Holidays"))
'------------------------------------------------------------------------------
Public Function WorkdaysBetween(ByVal startDate As Date, ByVal endDate As Date, _
                                Optional ByVal holidays As Variant) As Long
    Dim hol As Object
    Dim firstDay As Date, lastDay As Date, cursor As Date, holiday As Date
    Dim swap As Date
    Dim sign As Long
    Dim totalDays As Long, fullWeeks As Long, extraDays As Long
    Dim i As Long, total As Long
    Dim key As Variant

    firstDay = Int(startDate)
    lastDay = Int(endDate)
    sign = 1
    If firstDay > lastDay Then
        swap = firstDay
        firstDay = lastDay
        lastDay = swap
        sign = -1
    End If

    totalDays = DateDiff("d", firstDay, lastDay) + 1
    fullWeeks = totalDays \ 7
    extraDays = totalDays Mod 7
    total = fullWeeks * 5

    cursor = DateAdd("d", fullWeeks * 7, firstDay)
    For i = 1 To extraDays
        If Not IsWeekend(cursor) Then total = total + 1
        cursor = cursor + 1
    Next i

    Set hol = HolidaySet(holidays)
    For Each key In hol.Keys
        holiday = CDate(CLng(key))
        If holiday >= firstDay And holiday <= lastDay Then
            If Not IsWeekend(holiday) Then total = total - 1
        End If
    Next key

    WorkdaysBetween = total * sign
End Function

'------------------------------------------------------------------------------
' AddWorkdays
' The date a given number of working days away, like WORKDAY. Negative counts
' go backwards.
'
'   AddWorkdays(Date, 10, holidayRange)
'------------------------------------------------------------------------------
Public Function AddWorkdays(ByVal startDate As Date, ByVal workdaysToAdd As Long, _
                            Optional ByVal holidays As Variant) As Date
    Dim hol As Object
    Dim cursor As Date
    Dim direction As Long
    Dim remaining As Long

    cursor = Int(startDate)
    AddWorkdays = cursor
    If workdaysToAdd = 0 Then Exit Function

    Set hol = HolidaySet(holidays)
    direction = IIf(workdaysToAdd > 0, 1, -1)
    remaining = Abs(workdaysToAdd)

    Do While remaining > 0
        cursor = cursor + direction
        If Not IsWeekend(cursor) Then
            If Not hol.Exists(CLng(cursor)) Then remaining = remaining - 1
        End If
    Loop

    AddWorkdays = cursor
End Function

'------------------------------------------------------------------------------
' LastWorkdayOfMonth
' The last working day of the month containing aDate - month end reporting
' deadlines, in one call.
'------------------------------------------------------------------------------
Public Function LastWorkdayOfMonth(ByVal aDate As Date, Optional ByVal holidays As Variant) As Date
    Dim hol As Object
    Dim cursor As Date

    Set hol = HolidaySet(holidays)
    cursor = EndOfMonth(aDate)
    Do While IsWeekend(cursor) Or hol.Exists(CLng(cursor))
        cursor = cursor - 1
    Loop
    LastWorkdayOfMonth = cursor
End Function

'------------------------------------------------------------------------------
' MonthsBetween / YearsBetween
' Whole periods completed between two dates. A month end counts as a full
' month, so 31-Jan to 28-Feb is 1 month, while 31-Jan to 27-Feb is 0.
'------------------------------------------------------------------------------
Public Function MonthsBetween(ByVal startDate As Date, ByVal endDate As Date) As Long
    Dim months As Long

    months = (Year(endDate) - Year(startDate)) * 12 + (Month(endDate) - Month(startDate))
    If Day(endDate) < Day(startDate) Then
        If Day(endDate) <> Day(EndOfMonth(endDate)) Then months = months - 1
    End If
    MonthsBetween = months
End Function

Public Function YearsBetween(ByVal startDate As Date, ByVal endDate As Date) As Long
    YearsBetween = MonthsBetween(startDate, endDate) \ 12
End Function

'------------------------------------------------------------------------------
' ParseDateSafe
' Best effort date parsing that returns 0 instead of raising. Handles real
' dates, Excel serial numbers, ISO text (2024-03-31 and 20240331) and the
' usual separators. dayFirst decides how 03/04/2024 is read.
'------------------------------------------------------------------------------
Public Function ParseDateSafe(ByVal value As Variant, Optional ByVal dayFirst As Boolean = True) As Date
    Dim s As String
    Dim parts() As String
    Dim result As Date

    If IsObject(value) Then Exit Function
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    If IsArray(value) Then Exit Function

    If VarType(value) = vbDate Then ParseDateSafe = CDate(value): Exit Function

    If IsNumeric(value) Then
        If CDbl(value) >= 1 And CDbl(value) <= 2958465 Then
            ParseDateSafe = CDate(CDbl(value))            ' an Excel serial number
        ElseIf CDbl(value) >= 10000101 And CDbl(value) <= 99991231 Then
            ParseDateSafe = FromDateKey(CLng(value))      ' yyyymmdd held as a number
        End If
        Exit Function
    End If

    s = Trim$(CStr(value))
    If Len(s) = 0 Then Exit Function

    If Len(s) = 8 And IsAllDigits(s) Then
        ParseDateSafe = BuildDate(CLng(Left$(s, 4)), CLng(Mid$(s, 5, 2)), CLng(Right$(s, 2)))
        If ParseDateSafe <> 0 Then Exit Function
    End If

    s = Replace$(s, ".", "/")
    s = Replace$(s, "-", "/")
    s = Replace$(s, " ", "/")
    parts = Split(s, "/")
    If UBound(parts) = 2 Then
        If IsAllDigits(parts(0)) And IsAllDigits(parts(1)) And IsAllDigits(parts(2)) Then
            If Len(parts(0)) = 4 Then
                result = BuildDate(CLng(parts(0)), CLng(parts(1)), CLng(parts(2)))
            ElseIf dayFirst Then
                result = BuildDate(FullYear(CLng(parts(2))), CLng(parts(1)), CLng(parts(0)))
            Else
                result = BuildDate(FullYear(CLng(parts(2))), CLng(parts(0)), CLng(parts(1)))
            End If
            If result <> 0 Then ParseDateSafe = result: Exit Function
        End If
    End If

    On Error Resume Next
    If IsDate(value) Then ParseDateSafe = CDate(value)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' DateKey / FromDateKey
' yyyymmdd as a Long - compact, sortable and safe to use as a dictionary key.
'------------------------------------------------------------------------------
Public Function DateKey(ByVal aDate As Date) As Long
    DateKey = Year(aDate) * 10000 + Month(aDate) * 100 + Day(aDate)
End Function

Public Function FromDateKey(ByVal key As Long) As Date
    If key < 10000101 Or key > 99991231 Then Exit Function
    FromDateKey = BuildDate(key \ 10000, (key \ 100) Mod 100, key Mod 100)
End Function

'==============================================================================
' Private helpers
'==============================================================================
' Accepts a range, an array, a collection, a single date or nothing at all and
' returns a dictionary of date serial numbers.
Private Function HolidaySet(ByVal holidays As Variant) As Object
    Dim result As Object
    Dim item As Variant

    Set result = CreateObject("Scripting.Dictionary")
    Set HolidaySet = result

    If IsMissing(holidays) Then Exit Function
    If IsError(holidays) Then Exit Function
    If IsNull(holidays) Or IsEmpty(holidays) Then Exit Function

    If IsObject(holidays) Then
        If holidays Is Nothing Then Exit Function
        For Each item In holidays
            AddHoliday result, item
        Next item
    ElseIf IsArray(holidays) Then
        For Each item In holidays
            AddHoliday result, item
        Next item
    Else
        AddHoliday result, holidays
    End If
End Function

Private Sub AddHoliday(ByRef target As Object, ByVal item As Variant)
    Dim value As Variant
    Dim serial As Long

    If IsObject(item) Then
        On Error Resume Next
        value = item.value                  ' a cell from a range
        On Error GoTo 0
    Else
        value = item
    End If

    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Sub
    If Not IsDate(value) Then Exit Sub

    serial = CLng(Int(CDate(value)))
    If Not target.Exists(serial) Then target.Add serial, True
End Sub

Private Function BuildDate(ByVal yearPart As Long, ByVal monthPart As Long, ByVal dayPart As Long) As Date
    If yearPart < 100 Then yearPart = FullYear(yearPart)
    If yearPart < 1900 Or yearPart > 9999 Then Exit Function
    If monthPart < 1 Or monthPart > 12 Then Exit Function
    If dayPart < 1 Or dayPart > 31 Then Exit Function
    If dayPart > Day(DateSerial(yearPart, monthPart + 1, 0)) Then Exit Function

    BuildDate = DateSerial(yearPart, monthPart, dayPart)
End Function

' Two digit years: 00-29 are this century, 30-99 the last one.
Private Function FullYear(ByVal yearPart As Long) As Long
    If yearPart >= 100 Then FullYear = yearPart: Exit Function
    If yearPart <= 29 Then
        FullYear = 2000 + yearPart
    Else
        FullYear = 1900 + yearPart
    End If
End Function

Private Function IsAllDigits(ByVal text As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function
