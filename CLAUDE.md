# VBA_Functions - working notes for agents

## What the source of truth is

VBA lives as text in `src/*.bas|cls|frm` and `tests/*.bas`. The workbook binary
is a build output - never treat `xl/vbaProject.bin` as editable. Workbooks under
review sit in `workbook/` (gitignored).

## Running the tools

`python3` on macOS/Linux; on Windows use `python` (or `py`) if `python3` is not
on PATH - check with `python --version` before reporting a tool as broken.

```
python3 tools/vba_lint.py src tests --fail-on error     # static gate, exit 1 on errors
python3 tools/wb_inventory.py workbook/Book.xlsm        # structural report
pwsh -File tools/Invoke-VbaTests.ps1 -Workbook workbook/Book.xlsm
```

The last one is Windows + Excel only. `pwsh` may not be installed; fall back to
`powershell -ExecutionPolicy Bypass -File tools/Invoke-VbaTests.ps1 ...`. Exit
code 2 means the behavioural suite **did not run** - report it as unrun, never
as a pass, and never infer runtime behaviour from lint alone.

## Conventions

- Every module starts with `Attribute VB_Name` then `Option Explicit`. Both are
  required for the export/import round-trip - do not strip them.
- Typed variables, `Long` not `Integer` for anything counting rows, `.Value2`
  for bulk reads, explicit worksheet qualification (never `Selection`,
  `ActiveSheet`, or `Sheets(1)`).
- Any procedure that sets `Application.ScreenUpdating`/`EnableEvents`/
  `Calculation`/`DisplayAlerts` restores it on the failure path too.
- New tests go in `tests/Test_*.bas` **and** get registered in
  `modTestRunner.RunAllTests` - VBA cannot enumerate procedures, so an
  unregistered test silently never runs.
- Never skip, delete, loosen, or comment out a test to reach green.

## The review pass

`/workbook-review` runs it. Role separation is the point: reviewers are
read-only, the fixer owns `src/`, approvers are read-only and independent, the
tester owns `tests/` only. See `docs/agent-review-workflow.md`.
