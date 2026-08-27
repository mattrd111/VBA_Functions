# Multi-agent workbook review

How the review-fix-approve-test pass is wired, and how to run it.

## The idea

One agent doing all four jobs marks its own homework: it finds a problem,
patches it, declares it fine, and reports success. Splitting the jobs across
agents with different permissions makes that structurally impossible.

```
                        ┌──────────────────────────────────────────┐
                        │  Inventory  (wb-reviewer, read-only)     │
                        │  sheets · formulas · modules · lint base  │
                        └────────────────────┬─────────────────────┘
                                             │
        ┌────────────────┬───────────────────┼───────────────────┬────────────────┐
        │                │                   │                   │                │
  correctness      error handling     structure/formulas   maintainability   (your own)
   wb-reviewer       wb-reviewer          wb-reviewer         + safety
        └────────────────┴───────────────────┼───────────────────┴────────────────┘
                                             │  dedup by file:line:title
                                             ▼
                              one batch of findings per file
                                             │
                                    ┌────────┴────────┐
                                    │  wb-fixer       │  writes src/ + tests/
                                    │  minimal diff   │
                                    └────────┬────────┘
                                             ▼
                          ┌──────────────────────────────────────┐
                          │  wb-approver  ×2, independent, r/o   │
                          │  reads the real diff, can reject     │
                          └────────┬─────────────────┬───────────┘
                                   │ all approve     │ any block
                                   │                 ▼
                                   │        one rework round → re-vote
                                   │        still blocked → report, don't force
                                   ▼
                          ┌──────────────────────────────────────┐
                          │  wb-tester   owns tests/ only        │
                          │  static → structural → behavioural   │
                          └────────┬─────────────────┬───────────┘
                                   │ green           │ red
                                   ▼                 ▼
                                report        wb-fixer repair round → re-test
```

Permissions are the enforcement, not good intentions:

| Agent | Writes | Never |
|---|---|---|
| `wb-reviewer` | nothing | edits, commits |
| `wb-fixer` | `src/`, `tests/` | approves its own fix, widens scope beyond assigned findings |
| `wb-approver` | nothing | edits; approves what it could not verify |
| `wb-tester` | `tests/` | edits `src/`, weakens or skips a test to reach green |

## Why source lives as text

Excel stores VBA in `xl/vbaProject.bin` - an opaque binary that git cannot diff
and agents cannot read. So the modules are exported to `src/*.bas|cls|frm` and
`tests/*.bas`, and those text files are the source of truth. The workbook is
rehydrated from them.

One-time setup in Excel:

1. File → Options → Trust Center → Trust Center Settings → Macro Settings →
   tick **Trust access to the VBA project object model**.
2. Import `tools/vba_io.bas` into the workbook (VBE → File → Import File).
3. Run `ExportModules "C:\path\to\VBA_Functions"` to write every module out
   (`Test_*`, `modAssert` and `modTestRunner` go to `tests/`, the rest to `src/`).
4. After the agents finish, `ImportModules "C:\path\to\VBA_Functions"` puts the
   fixed modules back into the workbook.

## The gates

| Gate | Command | Runs where |
|---|---|---|
| Static (VBA lint) | `python3 tools/vba_lint.py src tests --fail-on error` | anywhere, incl. Linux/CI |
| Structural (workbook) | `python3 tools/wb_inventory.py Book.xlsm` | anywhere (reads OOXML directly) |
| Behavioural (tests) | `pwsh -File tools/Invoke-VbaTests.ps1 -Workbook Book.xlsm` | Windows + Excel only |

Both Python tools are stdlib-only - no openpyxl, no oletools.

`vba_lint.py` exits non-zero on any `error`-severity finding, so it works as a
gate. Current rules: missing `Option Explicit`, unscoped `On Error Resume Next`,
`Application.*` state left unrestored, selection-dependent code, unqualified
`Range`/`Cells`, sheets by index, untyped variables, `As Integer` overflow,
hardcoded paths, module-level `Public` state, risky calls (`Shell`, `SendKeys`,
`Kill`, `WScript.Shell`, `Auto_Open`), `.Value` vs `.Value2`.

`Invoke-VbaTests.ps1` copies the workbook to a scratch file, imports `src/` and
`tests/` into it, runs `modTestRunner.RunAllTests`, and reads the JSON the
runner writes. Exit codes: `0` green, `1` test failures, `2` could not run.
Exit `2` is reported as **unrun**, never as a pass - on Linux the behavioural
gate simply does not exist, and the pass ends at `green_static_only`.

## Running the pass

```
/workbook-review                      # skill drives the whole thing
```

or drive the pipeline directly:

```
Workflow({ name: 'workbook-review',
           args: { workbook: 'Book.xlsm', approvers: 2, maxFiles: 4 } })
```

`args`:

- `workbook` - path to the `.xlsm`/`.xlsx`; omit to review exported source only.
- `dimensions` - `[{key, prompt}]` to narrow or extend the review lenses.
- `approvers` - independent approvers per fix (default 2; all must approve).
- `maxFiles` - files fixed per run (default 4). Whatever is skipped is logged
  explicitly, so a partial run never reads as full coverage.

Agent count is roughly `1 + dimensions + (files × (1 + approvers)) + 1`.

Without a workflow, the same thing by hand: launch one `wb-reviewer` per
dimension in parallel, then per file one `wb-fixer`, then two `wb-approver`
agents on `git diff`, then one `wb-tester`. Same contract, same order.

## Adding a test

1. Write `Public Sub Test_Thing_DoesX()` in a `tests/Test_*.bas` file, using
   `modAssert.AssertEqual` / `AssertTrue` / `AssertClose` / `AssertRaises`.
2. Register it in `modTestRunner.RunAllTests` - VBA cannot enumerate
   procedures, so an unregistered test silently never runs.
3. A test that does not fail against the pre-fix code is not a regression test.

## Extending the review

- **New lens**: add a `{key, prompt}` to `DIMENSIONS` in
  `.claude/workflows/workbook-review.js`, or pass `dimensions` in `args`.
- **New lint rule**: add it in `lint_file()` in `tools/vba_lint.py`. Anything
  that produces wrong numbers or silent failure is `error`; anything fragile is
  `warn`; taste is `info`.
- **Stricter approval**: raise `approvers`, or give each approver a distinct
  lens (correctness / collateral damage / test quality) rather than the same
  brief three times.
- **CI**: the static and structural gates run on Linux, so they belong in a
  GitHub Action. The behavioural suite needs a Windows runner with Excel.
