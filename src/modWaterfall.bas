Attribute VB_Name = "modWaterfall"
'==============================================================================
' modWaterfall - splitting a distribution between LP and GP
'------------------------------------------------------------------------------
' The tiers everyone rebuilds by hand, in one worksheet function:
'
'   1  return of capital      to the LP, until contributed capital is back
'   2  preferred return       to the LP, until the accrued hurdle is paid
'   3  catch-up               to the GP, until it holds its carry percentage
'                             of the profit paid out so far
'   4  the split              whatever is left, at the carry rate
'
' Ask for one part at a time, so it works in any version of Excel and reads
' like what it is:
'
'   =Waterfall(D5, $B$2, $B$3, 20%, 100%, "LPTotal")
'   =Waterfall(D5, $B$2, $B$3, 20%, 100%, "GPCarry")
'
' The parts always add up to the amount distributed. "Total" returns that
' amount, so a model can tie itself out.
'
' Feed it the capital and preferred balances from modFinance:
'
'   capitalDue = UnreturnedCapital(flows, dates, 8%, TODAY())
'   prefDue    = AccruedPref(flows, dates, 8%, TODAY())
'
' Standalone: no dependency on the other modules in this repository.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Waterfall
'
' distributable   the cash to be split now
' capitalDue      unreturned capital contributions
' prefDue         accrued but unpaid preferred return
' carryRate       the GP's share of profit, 20% typically
' catchUpRate     the GP's share of the catch-up tier: 100% for a full catch-up,
'                 50% or 80% for a slower one, 0% for no catch-up tier at all
' part            which line you want:
'
'   LPCapital     tier 1, return of capital
'   LPPref        tier 2, preferred return
'   GPCatchUp     tier 3, the GP's side of the catch-up
'   LPCatchUp     tier 3, the LP's side of it when the catch-up is not 100%
'   LPResidual    tier 4, the LP's share of what is left
'   GPCarry       tier 4, the GP's carry on what is left
'   LPTotal       tiers 1, 2, 3 and 4 to the LP
'   GPTotal       tiers 3 and 4 to the GP
'   Total         the amount distributed - LPTotal + GPTotal always equals this
'------------------------------------------------------------------------------
Public Function Waterfall(ByVal distributable As Double, ByVal capitalDue As Double, _
                          ByVal prefDue As Double, ByVal carryRate As Double, _
                          Optional ByVal catchUpRate As Double = 1, _
                          Optional ByVal part As String = "LPTotal") As Variant
    Dim remaining As Double
    Dim returnOfCapital As Double, preferred As Double
    Dim gpCatchUp As Double, lpCatchUp As Double
    Dim lpResidual As Double, gpCarry As Double
    Dim target As Double, tier As Double

    If carryRate < 0 Or carryRate >= 1 Then Waterfall = CVErr(xlErrNum): Exit Function
    If catchUpRate > 1 Then catchUpRate = 1
    If catchUpRate < 0 Then catchUpRate = 0
    If distributable < 0 Then distributable = 0
    If capitalDue < 0 Then capitalDue = 0
    If prefDue < 0 Then prefDue = 0

    remaining = distributable

    ' Tier 1 - capital back
    returnOfCapital = MinOf(remaining, capitalDue)
    remaining = remaining - returnOfCapital

    ' Tier 2 - the hurdle
    preferred = MinOf(remaining, prefDue)
    remaining = remaining - preferred

    ' Tier 3 - catch-up. The GP is owed the carry percentage of the profit paid
    ' out so far, and the profit so far is the preferred plus the catch-up
    ' itself, which is what puts carryRate / (1 - carryRate) in the sum.
    If catchUpRate > 0 And carryRate > 0 And preferred > 0 Then
        target = carryRate * preferred / (1 - carryRate)
        tier = MinOf(remaining, target / catchUpRate)
        gpCatchUp = tier * catchUpRate
        lpCatchUp = tier * (1 - catchUpRate)
        remaining = remaining - tier
    End If

    ' Tier 4 - the split
    lpResidual = remaining * (1 - carryRate)
    gpCarry = remaining * carryRate

    Select Case UCase$(Trim$(part))
        Case "LPCAPITAL":  Waterfall = returnOfCapital
        Case "LPPREF":     Waterfall = preferred
        Case "GPCATCHUP":  Waterfall = gpCatchUp
        Case "LPCATCHUP":  Waterfall = lpCatchUp
        Case "LPRESIDUAL": Waterfall = lpResidual
        Case "GPCARRY":    Waterfall = gpCarry
        Case "LPTOTAL":    Waterfall = returnOfCapital + preferred + lpCatchUp + lpResidual
        Case "GPTOTAL":    Waterfall = gpCatchUp + gpCarry
        Case "TOTAL":      Waterfall = distributable
        Case Else:         Waterfall = CVErr(xlErrValue)
    End Select
End Function

'------------------------------------------------------------------------------
' CarriedInterest
' The GP's whole take on one distribution - the same as
' Waterfall(..., "GPTotal"), under the name people say out loud.
'------------------------------------------------------------------------------
Public Function CarriedInterest(ByVal distributable As Double, ByVal capitalDue As Double, _
                                ByVal prefDue As Double, ByVal carryRate As Double, _
                                Optional ByVal catchUpRate As Double = 1) As Variant
    CarriedInterest = Waterfall(distributable, capitalDue, prefDue, carryRate, catchUpRate, "GPTotal")
End Function

'------------------------------------------------------------------------------
' NetToGross
' What an LP keeps, as a multiple of the gross. Useful for the one-line sanity
' check on a returns page: gross MOIC 2.5x, net 2.1x, is that plausible?
'------------------------------------------------------------------------------
Public Function NetToGross(ByVal grossProceeds As Double, ByVal capitalDue As Double, _
                           ByVal prefDue As Double, ByVal carryRate As Double, _
                           Optional ByVal catchUpRate As Double = 1) As Variant
    Dim lpShare As Variant

    If grossProceeds <= 0 Then NetToGross = CVErr(xlErrNum): Exit Function

    lpShare = Waterfall(grossProceeds, capitalDue, prefDue, carryRate, catchUpRate, "LPTotal")
    If IsError(lpShare) Then NetToGross = lpShare: Exit Function

    NetToGross = lpShare / grossProceeds
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Function MinOf(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinOf = a Else MinOf = b
End Function
