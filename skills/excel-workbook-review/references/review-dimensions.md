# Review dimensions

One reviewer per dimension. Append the relevant block to the reviewer brief.

## 1. VBA correctness

Wrong results and silent wrongness: off-by-one over ranges and arrays, 1-based
`Range` versus 0-based array assumptions, `Variant` coercion, `Integer`
overflow above 32,767 rows, missing `Set` on object assignment, comparing
against `Empty`/`Null`/`""` interchangeably, `IsNumeric` accepting things the
caller does not expect, functions that swallow a failure and return 0 or an
empty string that reads as real data downstream.

## 2. Error handling and application state

`On Error Resume Next` opened and never closed with `On Error GoTo 0`.
`Application.ScreenUpdating`, `.EnableEvents`, `.Calculation` or
`.DisplayAlerts` set but not restored on the failure path - the happy path is
not enough, since the user is left with a frozen or silent Excel after any
error. Handlers that lose the original `Err.Number`/`Description`. Procedures
that leave the workbook half-modified when they fail partway.

## 3. Workbook structure and formulas

Cached `#REF!`, `#VALUE!`, `#N/A` cells. Volatile functions (`NOW`, `TODAY`,
`RAND`, `OFFSET`, `INDIRECT`) forcing full recalculation. `INDIRECT` and
`OFFSET` that break silently when rows move. External links to files the user
may not have. Hidden sheets holding live logic. Defined names pointing at
nothing. Hardcoded ranges that break the moment a row is inserted. Merged cells
inside data ranges. Protection that is only cosmetic.

`scripts/wb_inventory.py` reports all of these mechanically - use its output as
the starting point and investigate what it flags rather than re-listing it.

## 4. Maintainability and macro safety

Untyped variables, dead code, duplicated logic, magic numbers, sheets
referenced by index (`Sheets(1)`) rather than CodeName, procedures doing five
unrelated things, module-level `Public` state creating order-dependent
behaviour.

Safety: `Shell`, `CreateObject("WScript.Shell")`, `Kill`, `SendKeys`, file
writes to fixed paths, anything running automatically on open, and any of those
reached by unvalidated input. The point is not to accuse the user's own macros
of being malware - it is that these are the lines a corporate security review
will stop on, and unvalidated input reaching them is a genuine hazard.

## Optional: performance

Worth its own reviewer when the user says the workbook is slow. Cell-by-cell
loops that should be a single `.Value2` array read and write, repeated
`Application.Calculate`, `Select`/`Activate` inside loops, string concatenation
in long loops, `Worksheets` lookups repeated inside a loop body.
