---
name: excel-workbook-review
description: Review, fix, verify and test an Excel workbook and its VBA macros using four separated agent roles, so no agent signs off on its own work. Use this whenever someone wants a workbook, spreadsheet, or macro checked, audited, cleaned up, hardened, QA'd, debugged, code-reviewed or made more reliable - including "is there anything wrong with this xlsm", "my macro breaks sometimes", "tidy up this VBA", "can you sanity-check this model's formulas", or when they hand over an .xlsm/.xlsb/.xlsx and ask what could go wrong. Reads macros straight out of the workbook binary without needing Excel installed, so it works on any machine.
---

# Excel workbook review

Reviewing a workbook badly is easy and looks identical to doing it well: skim
the macros, describe some problems, patch them, declare victory. Nobody can
tell the difference from the output. The defence is structural - split the work
into four roles and let each one only do its own job:

| Role | May change | Job |
|---|---|---|
| **Reviewer** | nothing | find defects in one dimension, tied to file and line |
| **Fixer** | one file under `src/`, plus `tests/` | fix its assigned findings, minimal diff, add a regression test |
| **Approver** | nothing | verify the diff against the finding; can reject |
| **Tester** | `tests/` only | run the gates, add missing coverage, report honestly |

The value is in the separation, not the ceremony. A fixer that grades itself
reliably says "fixed". An approver with no stake, reading the actual diff
against the original failure scenario, catches the ones that only look fixed.

## 1. Build the workspace

Everything happens in a folder beside the workbook, so the original file is
never modified and the "before" version survives for diffing:

```
<workbook-stem>-review/
  src/           modules the fixer edits
  src-original/  pristine copies - never edited, used for diffs
  tests/         test harness + regression tests
  report.md      what you hand back
```

## 2. Get the VBA out

```bash
python3 scripts/vba_extract.py <workbook> -o <workspace>/src
```

This reads `xl/vbaProject.bin` directly - the OLE container and its
RLE-compressed module streams - so it needs no Excel and no third-party
packages. Standard modules land as `.bas`, classes as `.cls`, and document
modules (`ThisWorkbook`, `Sheet1`, …) in `src/document-modules/`, which matters
later: those cannot be re-imported as new components, their code has to be
pasted back by hand.

Copy the result to `src-original/` immediately.

Exit code 3 means the project could not be read (damaged, encrypted, or an
unusual build). Do not guess at the bytes - ask the user to export from Excel
with `assets/vba_io.bas` (`ExportModules`) and work from those files. A
workbook with no macros at all is fine; skip to the structural review.

## 3. Take a baseline

```bash
python3 scripts/wb_inventory.py <workbook>        # sheets, formulas, errors, links
python3 scripts/vba_lint.py <workspace>/src --json
```

The inventory reports hidden sheets, formula counts, cached `#REF!`/`#VALUE!`
cells, volatile and fragile functions, external links, defined names and
protection. The linter is a mechanical gate (exit 1 on error-severity
findings). Both are stdlib-only and run anywhere.

Keep both outputs - the tester compares against them at the end, which is how
you notice a "fix" that quietly introduced a new broken reference.

## 4. Review, in parallel, one dimension each

Spawn one reviewer subagent per dimension, all in the same turn. Give each the
reviewer brief from `references/agent-roles.md` plus its dimension from
`references/review-dimensions.md`:

1. VBA correctness
2. Error handling and application state
3. Workbook structure and formulas
4. Maintainability and macro safety

Narrow the set when the user has a specific worry ("it's slow", "it crashes on
some files") - a focused pass beats a broad one nobody reads. Reviewers are
read-only; they return findings as JSON, most severe first.

Then merge: drop duplicates on `file:line:title`, sort by severity, and group
the survivors **by file**. Grouping matters because fixers run in parallel and
two agents editing one file corrupt each other's work.

Show the user the finding list before fixing anything if there are more than a
handful, or if any fix would change behaviour rather than just repair it.

## 5. Fix, one agent per file

Give each fixer the fixer brief, its one file, and only its own findings. The
constraints that make the diff reviewable - minimal change, cause not symptom,
no drive-by refactors, a regression test per behavioural fix - are in the
brief. `references/vba-guide.md` has the idiomatic fixes for the common
findings, which stops four fixers inventing four different error-handling
styles.

## 6. Approve, independently

For each fixed file, spawn **two** approvers in the same turn with the approver
brief. Each reads the real diff:

```bash
diff -u <workspace>/src-original/<file> <workspace>/src/<file>
```

Both must approve. On any `reject` or `needs_changes`, hand the blocking points
back to the fixer for **one** rework round, then re-vote. Still blocked after
that: restore the file from `src-original/`, and report it as a finding the
user has to decide on. Forcing a disputed change through is the one outcome
worse than leaving a known bug in place, because now nobody knows the state.

## 7. Test

Copy `assets/modAssert.bas` and `assets/modTestRunner.bas` into `tests/` if the
workspace has no harness yet, then run the gates in order:

```bash
python3 scripts/vba_lint.py <workspace>/src <workspace>/tests --fail-on error
python3 scripts/wb_inventory.py <workbook>          # compare against the baseline
pwsh -File scripts/Invoke-VbaTests.ps1 -Workbook <workbook>   # Windows + Excel only
```

Every regression test must be registered in `modTestRunner.RunAllTests` - VBA
cannot enumerate procedures, so an unregistered test silently never runs, which
is the worst possible failure mode for a test suite.

The behavioural gate needs Windows and Excel. Anywhere else it simply cannot
run, and the honest verdict is **static checks passed, behaviour unverified**.
Never present that as a pass, and never infer runtime behaviour from lint - the
whole point of the exercise is that the user can trust the report.

If a gate goes red, that is a fixer's job, not the tester's. The tester never
edits `src/` to reach green, and never weakens, skips or deletes a test.

## 8. Hand it back

Write `report.md` and tell the user, in this order:

- what was found, by severity, and what it would have cost them
- what was fixed, and what each fix changed
- what was **not** fixed: blocked by approvers, or deferred because it needed a
  decision only they can make
- gate results, stating plainly whether the behavioural suite ran
- how to get the fixes back into the workbook

The last one matters because nothing has touched their file yet. VBA cannot be
written back into an `.xlsm` without Excel, so the fixed modules are text files
until they import them:

> 1. Excel → File → Options → Trust Center → Trust Center Settings → Macro
>    Settings → tick "Trust access to the VBA project object model".
> 2. Open the workbook, press Alt+F11, then File → Import File… → pick
>    `assets/vba_io.bas`.
> 3. Press Ctrl+G and run: `ImportModules "<path to the review workspace>"`
> 4. Anything in `src/document-modules/` must be pasted in by hand - open
>    ThisWorkbook or the sheet in the VBA editor and replace the code.
> 5. Save as `.xlsm`.

## Working without subagents

The roles still hold when you run them yourself; only the parallelism is lost.
Do the passes in order and keep the discipline that makes them worth anything:
when you switch to approving, re-derive the original failure scenario against
the new code from scratch rather than checking whether you did what you meant
to do. Write the verdict down before moving on. If you cannot argue the fix is
correct without referring to your own intent, it is not approved.

## Rules that hold throughout

- The user's workbook is never modified in place. All work happens on extracted
  copies; they choose when to import.
- Never weaken a test, an assertion, or a check to reach green. If a finding
  cannot be fixed cleanly, report it - an honest "still broken" is worth more
  than a green run that lies.
- A finding needs a failure scenario. "This could be cleaner" is not a defect;
  "this returns 0 instead of raising when the range is empty, so the report
  silently shows zero revenue" is.
