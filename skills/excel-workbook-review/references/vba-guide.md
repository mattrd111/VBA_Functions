# Fix patterns and mechanics

Read this before fixing. Four fixers working in parallel otherwise invent four
different error-handling styles in one project, which is its own maintenance
problem.

## The house style these fixes assume

```vb
Attribute VB_Name = "modThing"
Option Explicit
```

Both lines are required and must stay - the export/import round-trip breaks
without `Attribute VB_Name`, and without `Option Explicit` a typo becomes a
silent empty `Variant` instead of a compile error.

Typed variables always. `Long`, never `Integer`, for anything counting rows -
`Integer` overflows at 32,767 and Excel has had a million rows since 2007.

## Restoring application state

The failure path is the one that matters: without the handler, an error leaves
the user staring at a frozen Excel with events off, wondering why nothing
works.

```vb
Public Sub Recalculate(ByVal ws As Worksheet)
    Dim previousCalc As XlCalculation
    previousCalc = Application.Calculation

    On Error GoTo Restore
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ... work ...

Restore:
    Application.Calculation = previousCalc
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub
```

Save the previous value rather than assuming `xlCalculationAutomatic` - the
caller may have set manual mode deliberately, and clobbering it turns one slow
macro into a whole session of stale results.

## Scoping On Error Resume Next

It suppresses everything until cancelled, so an unclosed one hides errors from
every line that follows, often in a completely unrelated procedure.

```vb
On Error Resume Next
Set sheet = ThisWorkbook.Worksheets(name)   ' the one statement allowed to fail
On Error GoTo 0
If sheet Is Nothing Then Err.Raise 9, "LoadSheet", "No sheet named " & name
```

## Qualifying references

`Range("A1")` means "the active sheet", which is whatever the user last
clicked. The same macro then works when tested and corrupts data in production.

```vb
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Data")    ' by name, or better by CodeName
total = ws.Range("A1").Value2
```

Sheet indexes (`Sheets(1)`) break when someone reorders tabs. CodeNames survive
renaming and reordering both.

## Reading in bulk

A cell-by-cell loop crosses the COM boundary once per cell; one array read
crosses it once. On tens of thousands of rows the difference is minutes.

```vb
Dim values As Variant
values = ws.Range("A1:C10000").Value2       ' one round trip, 1-based 2-D array
For r = LBound(values, 1) To UBound(values, 1)
    ' ...
Next r
ws.Range("E1:E10000").Value2 = results      ' one write back
```

`.Value2` also skips Currency and Date coercion. Use `.Value` deliberately when
you actually want dates as dates.

## Failing loudly

A function that returns 0 when it cannot compute produces a report full of
plausible zeros. Raise instead, and let the caller decide:

```vb
If target Is Nothing Then Err.Raise 91, "SumRange", "target range is not set"
```

## Writing a regression test

```vb
Public Sub Test_SumRange_RaisesOnNothing()
    Dim result As Variant
    On Error Resume Next
    result = SumRange(Nothing)
    modAssert.AssertRaises 91, "SumRange(Nothing) should raise"
    On Error GoTo 0
End Sub
```

Then register it - `RunTest "Test_SumRange_RaisesOnNothing"` inside
`modTestRunner.RunAllTests`. Unregistered tests never run, and a suite that
silently skips is worse than no suite at all: it reports green either way.

The test must fail against the pre-fix code. If it passes both before and
after, it is documentation, not a regression test, and the approver will say so.

## Getting the code back into Excel

Nothing in this pass modifies the user's workbook - VBA cannot be written back
into an `.xlsm` without Excel. The fixed modules stay text files until the user
imports them:

1. Excel → File → Options → Trust Center → Trust Center Settings → Macro
   Settings → tick **Trust access to the VBA project object model**. Nothing
   below works without it, and the error it produces (1004) does not say so.
2. Open the workbook, Alt+F11, File → Import File… → `assets/vba_io.bas`.
3. Ctrl+G, then `ImportModules "<path to the review workspace>"`.
4. Modules in `src/document-modules/` (`ThisWorkbook`, `Sheet1`, …) cannot be
   imported as components - open each in the VBA editor and paste the code in.
5. Save as `.xlsm`.

`ExportModules` is the reverse trip, for when the user would rather export from
Excel than have the macros extracted from the binary.

## Running the tests (Windows only)

```powershell
pwsh -File scripts/Invoke-VbaTests.ps1 -Workbook Book.xlsm
# no pwsh? powershell -ExecutionPolicy Bypass -File scripts\Invoke-VbaTests.ps1 -Workbook Book.xlsm
```

It copies the workbook to a scratch file, imports `src/` and `tests/`, runs
`modTestRunner.RunAllTests` and reads the JSON the runner writes. Exit 0 green,
1 test failures, **2 could not run** - which is a report of "unverified", not a
pass. Close the workbook in Excel first.
