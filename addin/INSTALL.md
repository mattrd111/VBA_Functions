# Workbook Doctor — build and install

An Excel add-in for workbooks that have grown heavy: unused defined names,
tens of thousands of duplicate cell styles, a used range that runs to row
1,048,576, invisible drawing objects, dead external links.

The `.xlam` itself is not in this repository — an add-in is a binary file, and a
binary in git is a binary nobody can review. You build it from the source in
[`../addin`](.) and [`../src`](../src) in about thirty seconds.

---

## Build it with the script

```powershell
cd build
.\Build-AddIn.ps1
```

That writes `dist\WorkbookDoctor.xlam`. Add `-Install` to have the script copy it
into your Excel add-ins folder and switch it on:

```powershell
.\Build-AddIn.ps1 -Install
```

The script needs one Excel setting first, because it has to write code into a
workbook:

> **File → Options → Trust Center → Trust Center Settings → Macro Settings →
> Trust access to the VBA project object model**

Tick it, build, and untick it afterwards if you would rather not leave it on.
Nothing else needs it.

## Or build it by hand

No PowerShell, no trust setting, five minutes:

1. Open Excel, new blank workbook, `Alt+F11` for the VBA editor.
2. **File → Import File…** and import all eight modules from `src`:
   `modApp`, `modArray`, `modDate`, `modDictionary`, `modFile`, `modRange`,
   `modString`, `modWorkbook`.
3. Import all sixteen modules from `addin`:
   `modDoctorCommon`, `modDoctorScan`, `modDoctorNames`, `modDoctorStyles`,
   `modDoctorSheets`, `modDoctorLinks`, `modDoctorTools`, `modDoctorAudit`,
   `modDoctorRunner`, `modAuditFormula`, `modAuditCore`, `modModelAudit`,
   `modWrangleStack`, `modWrangleShape`, `modWrangleMatch`, `modDoctorMenu`.
4. Open `addin/ThisWorkbook.cls` in a text editor, copy everything from
   `Option Explicit` down, and paste it into the `ThisWorkbook` module of your
   new workbook. (Document modules cannot be imported — they have to be pasted.)
5. `Debug → Compile VBAProject`. It should compile clean.
6. Back in Excel: **File → Save As**, choose **Excel Add-In (\*.xlam)**, name it
   `WorkbookDoctor`. Excel will offer its add-ins folder — accept it.
7. **File → Options → Add-ins → Manage: Excel Add-ins → Go**, tick
   **Workbook Doctor**.

## Using it

The menu appears on the **Add-ins** tab of the ribbon as **Workbook Doctor**.

Start with **Audit workbook**. It changes nothing, and its first section is a
list of what is actually worth doing to the workbook in front of you.

The menu is grouped by what each tool does to your workbook.

**Audit** — reads, reports, changes nothing.

| | |
| --- | --- |
| Audit workbook (size and bloat) | Health check: findings first, then the detail per sheet |
| Audit model (formulas) | Formula integrity: broken rows, numbers typed over formulas, hardcoded assumptions, volatile functions, error cells |
| Formula map (this sheet) | One character per cell — the shape of a sheet in one screen |
| Select flagged cells (this sheet) | Selects the suspects so you can see them in context |
| List defined names / List cell styles | Report only |
| External links | Every linked workbook, whether the file is still there, and the cells that depend on it |

**Clean** — changes the workbook. Each one confirms and offers a backup.

| | |
| --- | --- |
| Clean everything (safe) | The low-risk half of the tools below, in one pass |
| Clean names | Broken, external-link and hidden defined names; unused ones are opt-in |
| Clean styles | Duplicate styles (`Normal 2`), or every custom style |
| Reset used range | Trims each sheet back to its real data |
| Remove invisible objects | Hidden and zero-size shapes left by bad pastes |
| Delete empty sheets | Sheets with no cells, shapes, tables or pivots |
| Clean conditional formatting | Removes rules that apply only to empty cells |
| Break external links | Replaces linked formulas with their values |

**Data** — turning an extract into something you can pivot.

| | |
| --- | --- |
| Stack selected sheets | Ctrl+click the tabs, then run it. Columns are matched by header name, not position |
| Stack files in a folder | The same, across every workbook in a folder |
| Unpivot a cross-tab | Months across the top become rows down the side |
| Fill blanks down | The group label that appears once and is blank for the next forty rows |
| Fuzzy match two lists | Two lists that are the same list, where none of the strings match |

**Cells** — acts on the selection.

| | |
| --- | --- |
| Trim and clean selection | Strips non-breaking spaces, control characters, doubled spaces |
| Selection to values | Freezes formulas |
| Remove hyperlinks from selection | Keeps the text |

And loose: **Unhide all sheets** (including very hidden ones) and **Backup this
workbook** (timestamped copy beside the original).

## Stacking and matching

**Stack selected sheets** takes the sheets you have Ctrl+clicked in the tab bar,
so there is no list to drive. Both stackers match columns by header *name*: a
source with its columns in a different order lands in the right place, a source
missing a column leaves blanks, and the stack report names which source was
missing what. That report is usually the finding, not the footnote.

**Fuzzy match two lists** works in three passes: exact once both sides are
tidied, exact once company decoration is stripped (`THE Acme Group Ltd.` and
`Acme Group Limited` are the same company), then closest match by edit distance.
It gives you a score, a verdict, and the *next best* match with its score —
because when the best and second-best are close, the match is a coin toss
however high the score looks. The `Your call` column is left empty on purpose.

## Reading a model audit

**Audit model** never changes the workbook. It reports:

| Finding | What it means |
| --- | --- |
| **Number typed into a calculated row** | Someone replaced a formula with a value. This is the one that costs money |
| **Formula differs from the row** | The other cells in that run share one formula and this one does not. At the first or last cell of a run it drops to Medium, because opening balances and closing columns often differ on purpose |
| **Number inside a formula** | An assumption buried where nobody will find it. Indexes and counts standing alone as arguments to lookup, rounding and date functions are ignored, so `VLOOKUP(x,y,3,0)` is quiet while `ROUND(B5*1.2,2)` is not |
| **Error showing** | `#REF!`, `#VALUE!` and friends. `#N/A` is reported separately at Low, since a lookup is often meant to miss |
| **Volatile function** / **Whole-column reference** | Why the model takes ten seconds to recalculate |
| **Reads another workbook** | A dependency on a file that may not be there tomorrow |

Where it deliberately says nothing:

- Runs shorter than four cells are not judged — "inconsistent" means nothing
  across three cells.
- A run with no dominant formula is left alone. If a row is a set of one-offs,
  there is no pattern to deviate from.
- Array and spilled formulas are skipped. Their R1C1 shifts per cell by design
  and would otherwise light up the whole report.
- Rows only, not columns. Financial models run periods across the page, and
  scanning both directions doubles the false positives without finding much.
- No check reports more than 150 findings per sheet.

The algorithm behind the hardcode check has a test —
`python build/test_formula_parser.py` — which mirrors the VBA in Python so the
rules can be exercised outside Excel.

## Before you run anything destructive

- **Ctrl+Z will not undo any of it.** Every destructive tool offers to take a
  backup first — take it.
- **Deleting a cell style changes how cells using it look.** They fall back to
  Normal. Formatting applied directly to the cell is untouched. If in doubt use
  the duplicates-only option, which only removes styles named after a built-in
  with a number stuck on the end.
- **"Unused" names are a judgement, not a fact.** The scan reads formulas,
  conditional formatting, validation, chart series, pivot sources and other
  names. A name used only from VBA, or assembled in a formula from text, cannot
  be seen — which is why deleting those is a separate opt-in, and why the report
  lists them before you decide.
- **Reset used range deletes formatting below your data.** That is the point,
  but a sheet with deliberate formatting waiting for next year's rows will lose
  it.
- **Save, close and reopen** after a clean-up. The last cell and the file size
  do not settle until Excel rewrites the file.

## Uninstalling

**File → Options → Add-ins → Manage: Excel Add-ins → Go**, untick it. The menu
goes with it. Delete the `.xlam` if you want it gone for good.

## Troubleshooting

**The build script says Excel would not let it touch the VBA project.**
Trust access to the VBA project object model is off — see above.

**The add-in is ticked but there is no menu.**
Look on the **Add-ins** tab, not Home. If the tab is missing entirely, macros
are probably disabled: Trust Center → Macro Settings, or put the `.xlam` in a
trusted location.

**"Cannot run the macro" when clicking a menu item.**
The menu is left over from a previous session. Untick and re-tick the add-in.

**Clean styles is taking minutes.**
That is normal on a workbook with tens of thousands of styles — every delete is
a separate operation. Press **Esc** to stop early; what has already been deleted
stays deleted.
