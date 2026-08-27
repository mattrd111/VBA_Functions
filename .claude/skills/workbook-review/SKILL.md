---
name: workbook-review
description: Run the multi-agent review-fix-approve-test pass over an Excel workbook and its exported VBA. Use when the user asks to review, audit, clean up, harden, or QA a workbook or its macros, or invokes /workbook-review. Orchestrates wb-reviewer, wb-fixer, wb-approver and wb-tester agents; falls back to a sequential chain when a workflow is not wanted.
---

# Workbook review pass

Four roles, deliberately separated so nobody grades their own homework:

| Agent | Can write? | Job |
|---|---|---|
| `wb-reviewer` | no | finds defects in one dimension, returns structured findings |
| `wb-fixer` | `src/`, `tests/` | fixes assigned findings in one file, minimal diff, adds a regression test |
| `wb-approver` | no | independently verifies each fix diff against the finding; can reject |
| `wb-tester` | `tests/` only | runs the gates, adds missing coverage, reports honestly |

## Before starting

1. **Find the source of truth.** VBA lives in `src/*.bas|cls|frm` as text, not
   inside the workbook binary - that is what git and the agents can read. If the
   user has a workbook whose modules are not exported yet, tell them to run
   `vba_io.ExportModules` once (import `tools/vba_io.bas` into the workbook,
   enable Trust Center > "Trust access to the VBA project object model").
2. **Locate the workbook file** if there is one (`*.xlsm`, `*.xlsx`). It is
   optional - the pass works on exported source alone, with the structural and
   behavioural gates skipped.
3. **Take a baseline** so the tester can compare against it:
   `python3 tools/wb_inventory.py <workbook>` and
   `python3 tools/vba_lint.py src tests --json`.
4. **Confirm the branch is clean** (`git status --short`). Every fix lands as a
   working-tree change the user can inspect and revert per file.

## Running it

Preferred - the deterministic pipeline:

```
Workflow({
  name: 'workbook-review',
  args: { workbook: 'Book.xlsm', approvers: 2, maxFiles: 4 }
})
```

That is ~1 + 4 + (files x 3) + 1 agents. Raise `maxFiles` only when the user
asks for more scope; it caps how many files get fixed in one run and the
workflow logs what it left out.

`dimensions` can be overridden to focus the pass, e.g.
`dimensions: [{ key: 'error-handling', prompt: '...' }]`.

If the user has not opted into multi-agent orchestration and does not want to,
run the same stages sequentially with the `Agent` tool instead: one
`wb-reviewer` per dimension in parallel, then per file a `wb-fixer`, then two
`wb-approver` agents on the diff, then one `wb-tester`. Same contract, same
order, just slower and hand-driven.

## Rules that hold either way

- A fixer never approves its own work; approvers never edit.
- Two approvers must both approve. Any `reject` or `needs_changes` triggers one
  rework round; still blocked after that, leave the change in the working tree,
  report it, and let the user decide - do not force it through.
- The tester owns `tests/`, the fixer owns `src/`. A red gate goes back to a
  fixer; the tester never edits `src/` to get green, and never weakens a test.
- The behavioural suite needs Windows + Excel. On Linux or in CI, `verdict` is
  `green_static_only` - report that as unrun, never as passing.

## Reporting back

Give the user: findings by severity, which files were fixed, which fixes were
reworked or blocked and why, tests added, gate results, and anything deferred
because it needed a decision. Then `git diff --stat` so they can see the scope
at a glance. Do not commit unless they ask.
