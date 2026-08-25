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
3. Import all ten modules from `addin`:
   `modDoctorCommon`, `modDoctorScan`, `modDoctorNames`, `modDoctorStyles`,
   `modDoctorSheets`, `modDoctorLinks`, `modDoctorTools`, `modDoctorAudit`,
   `modDoctorRunner`, `modDoctorMenu`.
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

| Menu item | What it does |
| --- | --- |
| Audit workbook | Read-only health check: findings first, then the detail per sheet |
| Clean everything (safe) | Runs the low-risk half of every tool in one pass |
| Clean names | Deletes broken, external-link and hidden defined names; unused ones are opt-in |
| Clean styles | Deletes duplicate styles (`Normal 2`), or every custom style |
| Reset used range | Trims each sheet back to its real data |
| Remove invisible objects | Hidden and zero-size shapes left by bad pastes |
| Delete empty sheets | Sheets with no cells, shapes, tables or pivots |
| Clean conditional formatting | Removes rules that apply only to empty cells |
| List defined names / List cell styles | Report only, changes nothing |
| External links / Break external links | Lists them, or replaces linked formulas with values |
| Unhide all sheets | Including very hidden ones |
| Trim and clean selection | Strips non-breaking spaces, control characters, doubled spaces |
| Selection to values | Freezes formulas |
| Remove hyperlinks from selection | Keeps the text |
| Backup this workbook | Timestamped copy beside the original |

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
