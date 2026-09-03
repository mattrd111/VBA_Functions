# VBA_Functions

Useful Excel VBA functions, kept as text so they can be diffed, linted and tested.

## Layout

```
src/                    VBA modules - the source of truth (.bas/.cls/.frm)
tests/                  test harness (modAssert, modTestRunner) + Test_*.bas suites
tools/
  vba_extract.py        pull VBA source out of a workbook binary (stdlib only)
  wb_inventory.py       workbook structural report (stdlib only)
  vba_lint.py           static VBA checks / CI gate (stdlib only)
  vba_io.bas            Export/ImportModules - round-trip between workbook and repo
  Invoke-VbaTests.ps1   headless test run via Excel COM (Windows)
.claude/
  agents/               wb-reviewer, wb-fixer, wb-approver, wb-tester
  workflows/            workbook-review.js - the orchestration
  skills/               /workbook-review entry point
docs/agent-review-workflow.md
```

## Quick start

```bash
python3 tools/vba_extract.py Book.xlsm -o src    # macros out of the binary, no Excel needed
python3 tools/wb_inventory.py Book.xlsm          # what's in the workbook
python3 tools/vba_lint.py src tests              # static checks (exit 1 on errors)
pwsh -File tools/Invoke-VbaTests.ps1 -Workbook Book.xlsm   # Windows + Excel only
```

To review and fix a workbook with agents, run `/workbook-review` in Claude Code.
Four separated roles - reviewers find, a fixer patches, two independent
approvers verify the diff, a tester runs the gates - so no agent signs off on
its own work. See [docs/agent-review-workflow.md](docs/agent-review-workflow.md).

The same pass is packaged as a portable skill in `skills/excel-workbook-review/`
for use outside this repo (Cowork, claude.ai, any project) - build the
installable `.skill` file with the skill-creator packager.

Getting VBA out of a workbook and into `src/` without Excel:
`python3 tools/vba_extract.py Book.xlsm -o src`. To export from Excel instead
(the reliable route for damaged or unusual projects): enable Trust Center → Macro
Settings → "Trust access to the VBA project object model", import
`tools/vba_io.bas`, then run `ExportModules` (and `ImportModules` to push fixes
back).
