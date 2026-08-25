Attribute VB_Name = "modFinance"
'==============================================================================
' modFinance - fund and deal maths
'------------------------------------------------------------------------------
' The numbers a fund reports on itself: money-weighted returns, the multiples,
' the preferred return, and how a fund did against a public index.
'
' Every function here is a worksheet function as well as a VBA one. Install the
' add-in and they are available in any workbook:
'
'   =FundXIRR(C5:C40, B5:B40)
'   =TVPI(C5:C40, NAV)
'   =AccruedPref(C5:C40, B5:B40, 8%, TODAY())
'
' Standalone: no dependency on the other modules in this repository.
' Day count is ACT/365 throughout, which is the convention Excel's own XIRR
' uses. Errors come back as real Excel errors, not zeros, so a broken input is
' visible rather than silently wrong.
'==============================================================================
Option Explicit

Private Const DAYS_PER_YEAR As Double = 365#
Private Const MAX_BISECTIONS As Long = 200

'------------------------------------------------------------------------------
' FundXNPV
' Net present value of dated cash flows, discounted from the first date.
'
'   =FundXNPV(0.08, C5:C40, B5:B40)
'------------------------------------------------------------------------------
Public Function FundXNPV(ByVal rate As Double, ByVal values As Variant, ByVal dates As Variant) As Variant
    Dim v() As Double, d() As Double
    Dim n As Long

    n = PairedSeries(values, dates, v, d)
    If n < 1 Then FundXNPV = CVErr(xlErrValue): Exit Function
    If rate <= -1 Then FundXNPV = CVErr(xlErrNum): Exit Function

    FundXNPV = NetPresentValue(rate, v, d, n)
End Function

'------------------------------------------------------------------------------
' FundXIRR
' Money-weighted return on irregularly dated cash flows.
'
' Excel's XIRR runs Newton's method from a single guess and gives up with #NUM!
' more often than it should, particularly on long fund lives with a big terminal
' value. This brackets the root first and then bisects, so it finds the answer
' whenever one exists between -99% and +100,000%.
'
'   =FundXIRR(C5:C40, B5:B40)
'
' Contributions are negative, distributions positive. Include the closing NAV as
' a final positive flow, or use FundIRR, which does that for you.
'------------------------------------------------------------------------------
Public Function FundXIRR(ByVal values As Variant, ByVal dates As Variant) As Variant
    Dim v() As Double, d() As Double
    Dim n As Long

    n = PairedSeries(values, dates, v, d)
    If n < 2 Then FundXIRR = CVErr(xlErrValue): Exit Function
    If Not ChangesSign(v, n) Then FundXIRR = CVErr(xlErrNum): Exit Function

    FundXIRR = SolveRate(v, d, n)
End Function

'------------------------------------------------------------------------------
' FundIRR
' FundXIRR with the closing NAV added as a final inflow - the since-inception
' IRR as an LP would report it.
'
'   =FundIRR(C5:C40, B5:B40, NAV, TODAY())
'------------------------------------------------------------------------------
Public Function FundIRR(ByVal values As Variant, ByVal dates As Variant, _
                        Optional ByVal residualValue As Double = 0, _
                        Optional ByVal valuationDate As Variant) As Variant
    Dim v() As Double, d() As Double
    Dim n As Long
    Dim asOf As Double

    n = PairedSeries(values, dates, v, d)
    If n < 1 Then FundIRR = CVErr(xlErrValue): Exit Function

    If residualValue <> 0 Then
        If IsMissing(valuationDate) Then
            asOf = d(n)
        Else
            asOf = ToSerialDate(valuationDate)
            If asOf < 0 Then FundIRR = CVErr(xlErrValue): Exit Function
        End If

        n = n + 1
        ReDim Preserve v(1 To n)
        ReDim Preserve d(1 To n)
        v(n) = residualValue
        d(n) = asOf
    End If

    If n < 2 Then FundIRR = CVErr(xlErrValue): Exit Function
    If Not ChangesSign(v, n) Then FundIRR = CVErr(xlErrNum): Exit Function

    FundIRR = SolveRate(v, d, n)
End Function

'==============================================================================
' Multiples
'------------------------------------------------------------------------------
' All take a signed cash-flow series: negative called, positive distributed.
'==============================================================================
Public Function PaidIn(ByVal values As Variant) As Variant
    Dim total As Double
    Dim v() As Double
    Dim n As Long, i As Long

    n = Series(values, v)
    If n < 1 Then PaidIn = CVErr(xlErrValue): Exit Function

    For i = 1 To n
        If v(i) < 0 Then total = total - v(i)
    Next i
    PaidIn = total
End Function

Public Function Distributed(ByVal values As Variant) As Variant
    Dim total As Double
    Dim v() As Double
    Dim n As Long, i As Long

    n = Series(values, v)
    If n < 1 Then Distributed = CVErr(xlErrValue): Exit Function

    For i = 1 To n
        If v(i) > 0 Then total = total + v(i)
    Next i
    Distributed = total
End Function

' Distributions to paid-in - cash actually returned, ignoring what is still in
' the ground. Named FundDPI, not DPI, because DPI is a valid column letter and
' Excel should never have to guess whether you meant a function or a reference.
Public Function FundDPI(ByVal values As Variant) As Variant
    Dim called As Variant

    called = PaidIn(values)
    If IsError(called) Then FundDPI = called: Exit Function
    If called = 0 Then FundDPI = CVErr(xlErrDiv0): Exit Function

    FundDPI = Distributed(values) / called
End Function

' Residual value to paid-in - what is left, as a multiple of what went in.
Public Function RVPI(ByVal values As Variant, ByVal residualValue As Double) As Variant
    Dim called As Variant

    called = PaidIn(values)
    If IsError(called) Then RVPI = called: Exit Function
    If called = 0 Then RVPI = CVErr(xlErrDiv0): Exit Function

    RVPI = residualValue / called
End Function

' Total value to paid-in: cash back plus what is left, over what went in.
Public Function TVPI(ByVal values As Variant, Optional ByVal residualValue As Double = 0) As Variant
    Dim called As Variant

    called = PaidIn(values)
    If IsError(called) Then TVPI = called: Exit Function
    If called = 0 Then TVPI = CVErr(xlErrDiv0): Exit Function

    TVPI = (Distributed(values) + residualValue) / called
End Function

' The same number as TVPI, under the name a deal team would use for it.
Public Function MOIC(ByVal values As Variant, Optional ByVal residualValue As Double = 0) As Variant
    MOIC = TVPI(values, residualValue)
End Function

'==============================================================================
' Preferred return
'------------------------------------------------------------------------------
' Walks the cash flows in date order, accruing the hurdle on unreturned capital
' and applying distributions against it.
'
' capitalFirst decides the order a distribution is applied in, and LPAs differ:
'   True   capital is repaid first, then accrued preferred  (the common case)
'   False  accrued preferred is paid first, then capital
'
' compound decides how the hurdle grows:
'   True   the rate compounds, and unpaid preferred earns it too
'   False  simple interest on unreturned capital
'
' Both are worth checking against the document before anyone relies on the
' answer. They can move the number by a lot.
'==============================================================================
Public Function AccruedPref(ByVal values As Variant, ByVal dates As Variant, _
                            ByVal prefRate As Double, ByVal asOfDate As Variant, _
                            Optional ByVal capitalFirst As Boolean = True, _
                            Optional ByVal compound As Boolean = True) As Variant
    Dim capital As Double, accrued As Double

    If Not PrefState(values, dates, prefRate, asOfDate, capitalFirst, compound, capital, accrued) Then
        AccruedPref = CVErr(xlErrValue)
        Exit Function
    End If
    AccruedPref = accrued
End Function

Public Function UnreturnedCapital(ByVal values As Variant, ByVal dates As Variant, _
                                  ByVal prefRate As Double, ByVal asOfDate As Variant, _
                                  Optional ByVal capitalFirst As Boolean = True, _
                                  Optional ByVal compound As Boolean = True) As Variant
    Dim capital As Double, accrued As Double

    If Not PrefState(values, dates, prefRate, asOfDate, capitalFirst, compound, capital, accrued) Then
        UnreturnedCapital = CVErr(xlErrValue)
        Exit Function
    End If
    UnreturnedCapital = capital
End Function

'==============================================================================
' Public market equivalent
'------------------------------------------------------------------------------
' KSPME
' Kaplan-Schoar: every cash flow is grown to the valuation date at the index's
' return, and the distributions are divided by the contributions.
'
'   above 1  the fund beat the index
'   below 1  the money would have done better in the index
'
'   =KSPME(C5:C40, B5:B40, D5:D40, NAV)
'
' indexLevels is the index level on each cash-flow date - one per cash flow.
'==============================================================================
Public Function KSPME(ByVal values As Variant, ByVal dates As Variant, _
                      ByVal indexLevels As Variant, _
                      Optional ByVal residualValue As Double = 0) As Variant
    Dim v() As Double, d() As Double, x() As Double
    Dim n As Long, i As Long
    Dim endLevel As Double, lastDate As Double
    Dim futureDistributions As Double, futureContributions As Double

    n = TripleSeries(values, dates, indexLevels, v, d, x)
    If n < 1 Then KSPME = CVErr(xlErrValue): Exit Function

    lastDate = d(1)
    endLevel = x(1)
    For i = 1 To n
        If d(i) >= lastDate Then
            lastDate = d(i)
            endLevel = x(i)
        End If
    Next i
    If endLevel <= 0 Then KSPME = CVErr(xlErrNum): Exit Function

    For i = 1 To n
        If x(i) <= 0 Then KSPME = CVErr(xlErrNum): Exit Function
        If v(i) > 0 Then
            futureDistributions = futureDistributions + v(i) * (endLevel / x(i))
        ElseIf v(i) < 0 Then
            futureContributions = futureContributions - v(i) * (endLevel / x(i))
        End If
    Next i

    If futureContributions = 0 Then KSPME = CVErr(xlErrDiv0): Exit Function
    KSPME = (futureDistributions + residualValue) / futureContributions
End Function

'==============================================================================
' Simple return maths
'==============================================================================
' Compound annual growth rate.
Public Function CAGR(ByVal beginValue As Double, ByVal endValue As Double, _
                     ByVal years As Double) As Variant
    If beginValue <= 0 Or years <= 0 Then CAGR = CVErr(xlErrNum): Exit Function
    If endValue < 0 Then CAGR = CVErr(xlErrNum): Exit Function

    CAGR = (endValue / beginValue) ^ (1 / years) - 1
End Function

' Turns a total return over a period into an annual one.
Public Function AnnualisedReturn(ByVal totalReturn As Double, ByVal years As Double) As Variant
    If years <= 0 Then AnnualisedReturn = CVErr(xlErrNum): Exit Function
    If totalReturn <= -1 Then AnnualisedReturn = CVErr(xlErrNum): Exit Function

    AnnualisedReturn = (1 + totalReturn) ^ (1 / years) - 1
End Function

' Time-weighted return: chain-links a series of periodic returns, so the answer
' does not depend on when money went in or out. Spelled out rather than TWR,
' which is a valid column letter.
Public Function TimeWeightedReturn(ByVal periodReturns As Variant) As Variant
    Dim r() As Double
    Dim n As Long, i As Long
    Dim growth As Double

    n = Series(periodReturns, r)
    If n < 1 Then TimeWeightedReturn = CVErr(xlErrValue): Exit Function

    growth = 1
    For i = 1 To n
        If r(i) <= -1 Then TimeWeightedReturn = -1: Exit Function
        growth = growth * (1 + r(i))
    Next i

    TimeWeightedReturn = growth - 1
End Function

'==============================================================================
' Private - the solver
'==============================================================================
Private Function NetPresentValue(ByVal rate As Double, ByRef v() As Double, _
                                 ByRef d() As Double, ByVal n As Long) As Double
    Dim i As Long
    Dim total As Double
    Dim start As Double

    start = d(1)
    For i = 1 To n
        total = total + v(i) / (1 + rate) ^ ((d(i) - start) / DAYS_PER_YEAR)
    Next i
    NetPresentValue = total
End Function

' Bracket the sign change, then bisect. Bisection cannot diverge, which is the
' whole point - a fund with twenty years of flows and a large terminal value is
' exactly where Newton's method wanders off.
Private Function SolveRate(ByRef v() As Double, ByRef d() As Double, ByVal n As Long) As Variant
    Dim low As Double, high As Double, middle As Double
    Dim fLow As Double, fMiddle As Double, fRate As Double
    Dim rate As Double, increment As Double
    Dim i As Long
    Dim bracketed As Boolean

    low = -0.9999
    fLow = NetPresentValue(low, v, d, n)

    rate = -0.99
    increment = 0.01
    Do While rate < 1000
        fRate = NetPresentValue(rate, v, d, n)
        If fRate * fLow <= 0 Then
            high = rate
            bracketed = True
            Exit Do
        End If
        low = rate
        fLow = fRate
        rate = rate + increment
        increment = increment * 1.05        ' coarser as it climbs, so 100,000% is reachable
    Loop

    If Not bracketed Then SolveRate = CVErr(xlErrNum): Exit Function

    For i = 1 To MAX_BISECTIONS
        middle = (low + high) / 2
        fMiddle = NetPresentValue(middle, v, d, n)

        If Abs(fMiddle) < 0.000000001 Then SolveRate = middle: Exit Function
        If (high - low) < 0.000000000001 Then Exit For

        If fMiddle * fLow <= 0 Then
            high = middle
        Else
            low = middle
            fLow = fMiddle
        End If
    Next i

    SolveRate = (low + high) / 2
End Function

'==============================================================================
' Private - the preferred return walk
'==============================================================================
Private Function PrefState(ByVal values As Variant, ByVal dates As Variant, _
                           ByVal prefRate As Double, ByVal asOfDate As Variant, _
                           ByVal capitalFirst As Boolean, ByVal compound As Boolean, _
                           ByRef capital As Double, ByRef accrued As Double) As Boolean
    Dim v() As Double, d() As Double
    Dim order() As Long
    Dim n As Long, i As Long, k As Long
    Dim asOf As Double, last As Double
    Dim remaining As Double, paid As Double

    n = PairedSeries(values, dates, v, d)
    If n < 1 Then Exit Function

    asOf = ToSerialDate(asOfDate)
    If asOf < 0 Then Exit Function

    order = SortedByDate(d, n)
    capital = 0
    accrued = 0
    last = d(order(1))

    For k = 1 To n
        i = order(k)
        If d(i) > asOf Then Exit For

        accrued = accrued + Accrual(capital, accrued, d(i) - last, prefRate, compound)
        last = d(i)

        If v(i) < 0 Then
            capital = capital - v(i)
        ElseIf v(i) > 0 Then
            remaining = v(i)
            If capitalFirst Then
                paid = MinOf(remaining, capital): capital = capital - paid: remaining = remaining - paid
                paid = MinOf(remaining, accrued): accrued = accrued - paid: remaining = remaining - paid
            Else
                paid = MinOf(remaining, accrued): accrued = accrued - paid: remaining = remaining - paid
                paid = MinOf(remaining, capital): capital = capital - paid: remaining = remaining - paid
            End If
        End If
    Next k

    If asOf > last Then
        accrued = accrued + Accrual(capital, accrued, asOf - last, prefRate, compound)
    End If

    PrefState = True
End Function

Private Function Accrual(ByVal capital As Double, ByVal accrued As Double, ByVal days As Double, _
                         ByVal rate As Double, ByVal compound As Boolean) As Double
    If days <= 0 Then Exit Function
    If rate = 0 Then Exit Function

    If compound Then
        If rate <= -1 Then Exit Function
        Accrual = (capital + accrued) * ((1 + rate) ^ (days / DAYS_PER_YEAR) - 1)
    Else
        Accrual = capital * rate * days / DAYS_PER_YEAR
    End If
End Function

Private Function SortedByDate(ByRef d() As Double, ByVal n As Long) As Long()
    Dim order() As Long
    Dim i As Long, j As Long, key As Long

    ReDim order(1 To n)
    For i = 1 To n
        order(i) = i
    Next i

    ' Insertion sort - cash-flow series are short and usually already in order.
    For i = 2 To n
        key = order(i)
        j = i - 1
        Do While j >= 1
            If d(order(j)) <= d(key) Then Exit Do
            order(j + 1) = order(j)
            j = j - 1
        Loop
        order(j + 1) = key
    Next i

    SortedByDate = order
End Function

'==============================================================================
' Private - input handling
'==============================================================================
' A range, a 1D array or a 2D array becomes a 1-based Double array. Blanks and
' anything non-numeric are dropped. Returns the count, or 0.
Private Function Series(ByVal source As Variant, ByRef out() As Double) As Long
    Dim raw As Variant
    Dim n As Long, i As Long

    raw = Flatten(source)
    If Not IsArray(raw) Then Exit Function
    If UBound(raw) < 1 Then Exit Function

    ReDim out(1 To UBound(raw))
    For i = 1 To UBound(raw)
        If IsUsableNumber(raw(i)) Then
            n = n + 1
            out(n) = CDbl(raw(i))
        End If
    Next i

    If n = 0 Then Exit Function
    ReDim Preserve out(1 To n)
    Series = n
End Function

' Values and dates together, keeping them aligned: a pair is dropped only when
' one side of it is missing, so a cash flow never ends up against the wrong date.
Private Function PairedSeries(ByVal values As Variant, ByVal dates As Variant, _
                              ByRef v() As Double, ByRef d() As Double) As Long
    Dim rawValues As Variant, rawDates As Variant
    Dim count As Long, n As Long, i As Long
    Dim serial As Double

    rawValues = Flatten(values)
    rawDates = Flatten(dates)
    If Not IsArray(rawValues) Or Not IsArray(rawDates) Then Exit Function

    count = UBound(rawValues)
    If UBound(rawDates) < count Then count = UBound(rawDates)
    If count < 1 Then Exit Function

    ReDim v(1 To count)
    ReDim d(1 To count)

    For i = 1 To count
        If IsUsableNumber(rawValues(i)) Then
            serial = ToSerialDate(rawDates(i))
            If serial >= 0 Then
                n = n + 1
                v(n) = CDbl(rawValues(i))
                d(n) = serial
            End If
        End If
    Next i

    If n = 0 Then Exit Function
    ReDim Preserve v(1 To n)
    ReDim Preserve d(1 To n)
    PairedSeries = n
End Function

' Cash flows, their dates and the index level on each of those dates, kept in
' step. A row is dropped only when one of the three is missing, so an index
' level can never end up against the wrong cash flow.
Private Function TripleSeries(ByVal values As Variant, ByVal dates As Variant, _
                              ByVal indexLevels As Variant, _
                              ByRef v() As Double, ByRef d() As Double, ByRef x() As Double) As Long
    Dim rawValues As Variant, rawDates As Variant, rawIndex As Variant
    Dim count As Long, n As Long, i As Long
    Dim serial As Double

    rawValues = Flatten(values)
    rawDates = Flatten(dates)
    rawIndex = Flatten(indexLevels)
    If Not IsArray(rawValues) Or Not IsArray(rawDates) Or Not IsArray(rawIndex) Then Exit Function

    count = UBound(rawValues)
    If UBound(rawDates) < count Then count = UBound(rawDates)
    If UBound(rawIndex) < count Then count = UBound(rawIndex)
    If count < 1 Then Exit Function

    ReDim v(1 To count)
    ReDim d(1 To count)
    ReDim x(1 To count)

    For i = 1 To count
        If IsUsableNumber(rawValues(i)) And IsUsableNumber(rawIndex(i)) Then
            serial = ToSerialDate(rawDates(i))
            If serial >= 0 Then
                n = n + 1
                v(n) = CDbl(rawValues(i))
                d(n) = serial
                x(n) = CDbl(rawIndex(i))
            End If
        End If
    Next i

    If n = 0 Then Exit Function
    ReDim Preserve v(1 To n)
    ReDim Preserve d(1 To n)
    ReDim Preserve x(1 To n)
    TripleSeries = n
End Function

Private Function Flatten(ByVal source As Variant) As Variant
    Dim data As Variant
    Dim out() As Variant
    Dim r As Long, c As Long, n As Long, i As Long

    If IsObject(source) Then
        If source Is Nothing Then Exit Function
        On Error Resume Next
        data = source.value
        On Error GoTo 0
    Else
        data = source
    End If

    If IsEmpty(data) Then Exit Function

    If Not IsArray(data) Then
        ReDim out(1 To 1)
        out(1) = data
        Flatten = out
        Exit Function
    End If

    If Dimensions(data) = 2 Then
        n = (UBound(data, 1) - LBound(data, 1) + 1) * (UBound(data, 2) - LBound(data, 2) + 1)
        If n < 1 Then Exit Function
        ReDim out(1 To n)
        n = 0
        For r = LBound(data, 1) To UBound(data, 1)
            For c = LBound(data, 2) To UBound(data, 2)
                n = n + 1
                out(n) = data(r, c)
            Next c
        Next r
    Else
        n = UBound(data) - LBound(data) + 1
        If n < 1 Then Exit Function
        ReDim out(1 To n)
        For i = 1 To n
            out(i) = data(LBound(data) + i - 1)
        Next i
    End If

    Flatten = out
End Function

Private Function IsUsableNumber(ByVal value As Variant) As Boolean
    If IsObject(value) Or IsError(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    If VarType(value) = vbString Then
        If Len(Trim$(CStr(value))) = 0 Then Exit Function
    End If
    If VarType(value) = vbBoolean Then Exit Function

    IsUsableNumber = IsNumeric(value)
End Function

' The date as an Excel serial number, or -1 when it is not a date at all.
Private Function ToSerialDate(ByVal value As Variant) As Double
    ToSerialDate = -1

    If IsObject(value) Or IsError(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    If IsMissing(value) Then Exit Function

    If VarType(value) = vbDate Then ToSerialDate = CDbl(value): Exit Function

    If IsNumeric(value) Then
        If CDbl(value) >= 0 Then ToSerialDate = CDbl(value)
        Exit Function
    End If

    On Error Resume Next
    If IsDate(value) Then ToSerialDate = CDbl(CDate(value))
    On Error GoTo 0
End Function

Private Function ChangesSign(ByRef v() As Double, ByVal n As Long) As Boolean
    Dim i As Long
    Dim positive As Boolean, negative As Boolean

    For i = 1 To n
        If v(i) > 0 Then positive = True
        If v(i) < 0 Then negative = True
        If positive And negative Then ChangesSign = True: Exit Function
    Next i
End Function

Private Function MinOf(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinOf = a Else MinOf = b
End Function

Private Function Dimensions(ByVal arr As Variant) As Long
    Dim d As Long
    Dim test As Long

    If Not IsArray(arr) Then Exit Function
    On Error GoTo Done
    Do
        d = d + 1
        test = LBound(arr, d)
    Loop While d < 60
Done:
    Dimensions = d - 1
End Function
