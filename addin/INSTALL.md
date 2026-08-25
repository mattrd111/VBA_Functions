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

## If PowerShell is blocked

On a managed machine you will most likely get:

> running scripts is disabled on this system

Try the bypass first, which changes no machine setting:

```powershell
powershell -ExecutionPolicy Bypass -File .\Build-AddIn.ps1 -Install
```

If Group Policy is enforcing the policy that will fail too. `Get-ExecutionPolicy -List`
tells you: a value against **MachinePolicy** or **UserPolicy** means it is locked
and no flag will get past it.

In that case use the VBA builder instead, which needs no PowerShell at all:

1. Excel: turn on **Trust access to the VBA project object model** (as above).
2. New blank workbook, `Alt+F11`, **File → Import File…**, pick
   [`build/BuildAddIn.bas`](../build/BuildAddIn.bas).
3. Put the cursor in `BuildWorkbookDoctor` and press **F5**.
4. Point it at the folder you unzipped - the one with `src` and `addin` in it.
5. Close the builder workbook without saving.

It imports all 28 modules, merges `ThisWorkbook`, saves the `.xlam` into your
Excel add-ins folder and switches it on - the same job the PowerShell script
does, from inside Excel.

## Or build it by hand

No PowerShell, no trust setting, five minutes:

1. Open Excel, new blank workbook, `Alt+F11` for the VBA editor.
2. **File → Import File…** and import all ten modules from `src`:
   `modApp`, `modArray`, `modDate`, `modDictionary`, `modFile`, `modFinance`,
   `modRange`, `modString`, `modWaterfall`, `modWorkbook`.
3. Import all nineteen modules from `addin`:
   `modDoctorCommon`, `modDoctorScan`, `modDoctorNames`, `modDoctorStyles`,
   `modDoctorSheets`, `modDoctorLinks`, `modDoctorTools`, `modDoctorAudit`,
   `modDoctorRunner`, `modAuditFormula`, `modAuditCore`, `modModelAudit`,
   `modWrangleStack`, `modWrangleShape`, `modWrangleMatch`, `modFundHelper`,
   `modHouseStyle`, `modDoctorMenu`.
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
| Fund maths reference sheet | A worked example of every fund function, with live formulas |

**Format** — the Alpha FMC house style.

| | |
| --- | --- |
| Cycle number / percent / date format | Each press steps the selection along a ladder and reports where it landed. `0dp → 1dp → 2dp → thousands → millions → £m` |
| Style as Alpha table | Header fill, house type, no borders, banded rows |
| Style as header row / total row | The pieces on their own |
| House type on this sheet | Quire Sans throughout, gridlines off — what makes a sheet screenshot cleanly into a slide |
| Colour inputs and formulas | Blue typed, black calculated, green from another sheet, red from another workbook |
| Alpha chart colours | Recolours the selected chart in the template's series order |
| Alpha palette reference | A swatch sheet: every colour, its hex, and what it is for |

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

## The house style

Read out of the Alpha FMC 2026 PowerPoint theme:

| | |
| --- | --- |
| Ink | `#2A2723` |
| Primary violet | `#503AF5`, tints to `#DCD7FC` |
| Table header fill | `#EDEBFD` |
| Panel / banding | `#F0F5EB`, sage `#BBC1B2` |
| Type | Quire Sans — 10pt body, 9pt headers, headers not bold |
| Chart series | `#1E2999 · #6085DC · #64B0E2 · #4BA379 · #79C792` |

Two things to know:

- The template's tables have **no borders** and their headers are **not bold**.
  The formatter follows that. If you want a house rule different from the
  template, `ApplyHeaderStyle` in `modHouseStyle` is four lines.
- The **number format ladders are not from the template**, which does not
  specify any. They are the usual consulting conventions, sitting in
  `NumberLadder`, `PercentLadder` and `DateLadder` at the top of the module so
  they can be changed in one place.

If Quire Sans is not installed, Excel substitutes a font and everything else
still applies.

## The fund functions

`modFinance` and `modWaterfall` are worksheet functions, not menu commands.
While the add-in is installed they work in any workbook:

```
=FundXIRR(C5:C40, B5:B40)                       money-weighted return
=FundIRR(C5:C40, B5:B40, NAV, TODAY())          ... including the closing NAV
=TVPI(C5:C40, NAV)                              total value to paid-in
=AccruedPref(C5:C40, B5:B40, 8%, TODAY())       the compounded hurdle
=KSPME(C5:C40, B5:B40, D5:D40, NAV)             against a public index
=Waterfall(D5, capital, pref, 20%, 100%, "GPCarry")
```

**Data → Fund maths reference sheet** builds a worked example of all of them
with live formulas, which is the quickest way to see the syntax.

Two things worth knowing:

- The preferred return has two switches that differ by LPA and move the answer
  a long way: whether a distribution repays capital or the hurdle first
  (`capitalFirst`), and whether unpaid preferred earns the rate itself
  (`compound`). Check both against the document before anyone relies on the
  number.
- A sheet using these functions shows `#NAME?` for anyone without the add-in
  installed. Convert to values before sending it out.

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
