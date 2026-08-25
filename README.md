# VBA_Functions

Useful Excel VBA functions — a small standard library for the things every
Excel macro ends up needing: finding the edges of your data, moving whole
blocks in and out of arrays, fast lookups, working-day maths, tidying text,
handling files, and leaving Excel exactly as you found it.

Nothing here needs a project reference, a class module or an add-in. Import a
`.bas` file and the functions are available.

It also ships **Workbook Doctor**, an Excel add-in built on top of the library
that cleans out unused defined names, duplicate cell styles, runaway used ranges
and the rest of the junk that makes a workbook heavy —
[build and install it](addin/INSTALL.md).

---

## What's in it

| Module | What it covers |
| --- | --- |
| [`modApp`](src/modApp.bas) | Speed-up switches, progress bar, stopwatch, error text |
| [`modArray`](src/modArray.bas) | Sorting, de-duplicating and reshaping arrays, including the 2D arrays from `Range.Value` |
| [`modDate`](src/modDate.bas) | Month and quarter ends, fiscal periods, working days, forgiving date parsing |
| [`modDictionary`](src/modDictionary.bas) | Dictionary-based replacements for VLOOKUP, SUMIF and COUNTIF loops |
| [`modFile`](src/modFile.bas) | Paths, folder listings, text and CSV files, file and folder pickers |
| [`modRange`](src/modRange.bas) | Last row and column, header lookups, bulk read/write, tidy-ups |
| [`modString`](src/modString.bas) | Cleaning, slicing, padding, regular expressions, fuzzy matching |
| [`modWorkbook`](src/modWorkbook.bas) | Creating, finding and clearing sheets, opening workbooks, PDF export |
| [`modExamples`](examples/modExamples.bas) | Runnable examples of the patterns above |
| [`addin/`](addin) | **Workbook Doctor** — the clean-up add-in built on the modules above |

Every module in `src` is standalone — import one without the others and it
still compiles. Only `modExamples` depends on the rest.

---

## Getting started

1. Download the `.bas` files you want from [`src`](src).
2. In Excel press `Alt+F11` to open the VBA editor.
3. **File → Import File…**, or just drag the `.bas` file onto the Project
   window.
4. Call the functions from your own code. `Alt+F8` runs anything in
   `modExamples`.

To keep the macros in one place, import them into `PERSONAL.XLSB` (the
personal macro workbook) and they are available in every workbook you open.

---

## The four things worth reading first

**Stop hard-coding column letters.** Headers move; header text usually does
not.

```vba
Dim amountColumn As Long
amountColumn = FindColumnByHeader(ws, "Net Amount")   ' 0 when it is not there
```

**Read and write in blocks, not cells.** A loop over 50,000 cells takes
minutes; the same work in an array takes under a second.

```vba
Dim data As Variant
data = RangeToArray(BodyRange(ws))        ' every value, one hit
' ... work on data(r, c) ...
WriteArray data, ws.Range("A2")           ' and back again, one hit
```

**Index once instead of looking up per row.** This is the difference between a
macro that runs while you fetch a coffee and one that runs before you stand up.

```vba
Dim prices As Object
Set prices = BuildLookup(pricesSheet.Range("A1:B20000"), 1, 2)
For r = 2 To lastRow
    result(r, 1) = LookupValue(prices, codes(r, 1), "not found")
Next r
```

**Wrap long jobs so Excel always comes back.** `FastMode` remembers how it
found things and nests safely, so an inner routine cannot switch the screen
back on halfway through an outer one.

```vba
StartTimer
FastMode True
On Error GoTo Cleanup

' ... the slow part, calling ShowProgress r, lastRow, "Rebuilding" now and then

Cleanup:
    FastMode False
    ClearStatusBar
    If Err.Number <> 0 Then MsgBox ErrorInfo("RebuildReport"), vbExclamation
```

If a macro ever dies inside `FastMode` and leaves the screen frozen, run
`ResetApp`.

---

## Function index

### modApp — application state, progress and timing

| Function | Notes |
| --- | --- |
| `FastMode([enable])` | Screen updating, events, alerts and calculation off; restores what it found. Nested calls are counted |
| `ResetApp()` | Panic button — puts Excel back to normal whatever happened |
| `ShowProgress(current, total, [message], [barWidth])` | Progress bar in the status bar |
| `ClearStatusBar()` | Hands the status bar back to Excel |
| `StartTimer()` / `ElapsedSeconds()` / `ElapsedText()` | Stopwatch; `ElapsedText` gives `"1m 12.4s"` |
| `DebugLog(part, part, …)` | Timestamped line in the Immediate window |
| `LogToFile(path, message)` | Timestamped line appended to a text file |
| `ErrorInfo([procedureName])` | One-line description of the current error |

### modArray — arrays

| Function | Notes |
| --- | --- |
| `IsArrayAllocated(arr)` | True only when the array has elements |
| `ArrayDimensions(arr)` / `ArrayLength(arr, [dimension])` | Shape without the error handling |
| `ArrayIndexOf(arr, value, [caseSensitive])` / `ArrayContains(…)` | Search a 1D array |
| `ArrayUnique(arr, [caseSensitive], [keepBlanks])` | Distinct values in first-seen order |
| `ArrayReverse(arr)` / `SortArray(arr, [descending])` | Copies, sorted numerically then alphabetically |
| `QuickSort(arr, [low], [high])` | Sorts a 1D array in place |
| `Sort2D(arr, keyColumn, [descending])` | Sorts the rows of a 2D array by one column |
| `GetColumn(arr, columnIndex)` / `GetRow(arr, rowIndex)` | One column or row as a 1D array |
| `TransposeArray(arr)` | No `Application.Transpose` row limit |
| `To2D(arr, [asRow])` | 1D array to a shape a range will accept |
| `RemoveBlanks(arr)` / `ConcatArrays(first, second)` | |
| `ArrayToString(arr, [delimiter])` | Readable dump for the Immediate window |

### modDate — dates and periods

| Function | Notes |
| --- | --- |
| `StartOfMonth(d, [monthsToAdd])` / `EndOfMonth(d, [monthsToAdd])` | `EndOfMonth(d, -1)` is prior month end |
| `DaysInMonth(d)` / `StartOfYear(d, …)` / `EndOfYear(d, …)` / `StartOfWeek(d, [firstDay])` | |
| `Quarter(d, [fyStartMonth])` / `QuarterStart(…)` / `QuarterEnd(…)` | Calendar or fiscal |
| `FiscalYear(d, [fyStartMonth])` / `FiscalYearLabel(…)` / `PeriodLabel(…)` | `PeriodLabel(d, 4)` → `"Q1 FY2025"` |
| `IsWeekend(d, [day1], [day2])` / `IsWorkday(d, [holidays])` | Weekend days are configurable |
| `WorkdaysBetween(start, end, [holidays])` | Inclusive of both ends, like `NETWORKDAYS` |
| `AddWorkdays(start, count, [holidays])` | Like `WORKDAY`; negative counts go backwards |
| `LastWorkdayOfMonth(d, [holidays])` | Month-end deadlines in one call |
| `MonthsBetween(start, end)` / `YearsBetween(start, end)` | Whole periods completed |
| `ParseDateSafe(value, [dayFirst])` | Text, serial numbers and ISO forms; returns 0 rather than raising |
| `DateKey(d)` / `FromDateKey(key)` | `yyyymmdd` as a Long |

`holidays` accepts a range, an array, a single date, or nothing at all.

### modDictionary — lookups, grouping and counting

| Function | Notes |
| --- | --- |
| `NewDictionary([ignoreCase])` | Late-bound, so no reference needed |
| `DictGet(dict, key, [default])` | No accidental key creation |
| `DictIncrement(dict, key, [amount])` / `DictAppend(dict, key, value)` | Totals, or a Collection per key |
| `SortedKeys(dict, [descending])` / `DictToArray(dict, [sortByKey])` | Ready to write to a sheet |
| `BuildLookup(source, [keyCol], [returnCol], [skipHeader], [keepFirst])` | Index a range or array once |
| `LookupValue(index, key, [default])` | Reads an index back, matching keys the same way |
| `GroupSum(source, keyCol, valueCol, [skipHeader])` | Every `SUMIF` in one pass |
| `CountBy(source, keyCol, [skipHeader])` | Every `COUNTIF` in one pass |
| `UniqueValues(source, [ignoreCase])` / `DuplicateKeys(source, keyCol, …)` | |

`BuildLookup` treats numbers stored as text as numbers, but leaves text that
merely looks numeric alone — so `"007"` and `7` stay apart.

### modFile — files and folders

| Function | Notes |
| --- | --- |
| `FileExists(path)` / `FolderExists(path)` | A folder never passes as a file |
| `EnsureFolder(path)` | Creates every missing parent too |
| `JoinPath(part, part, …)` | Single backslashes however many you supply |
| `FileNameOnly(p)` / `BaseName(p)` / `FileExtension(p)` / `ParentFolder(p)` | Pure string work |
| `UniqueFilePath(path)` | Inserts `(1)`, `(2)` … so nothing is overwritten |
| `ListFiles(folder, [pattern], [includeSubfolders])` | Full paths, as a Collection |
| `ReadTextFile(path, [charset])` / `WriteTextFile(path, content, [append])` | |
| `WriteUtf8File(path, content, [includeBOM])` | With or without the byte order mark |
| `ExportRangeToCsv(rng, path, [delimiter], [includeBOM])` | Proper quoting, ISO dates, no `SaveAs` |
| `PickFile([title], [filterName], [pattern])` / `PickFolder([title])` | `""` when cancelled |
| `DeleteFileSafe(path)` / `FileSizeBytes(path)` / `FileModified(path)` | |

### modRange — ranges

| Function | Notes |
| --- | --- |
| `LastRow(ws, [column])` / `LastColumn(ws, [row])` | 0 when empty; column may be `2` or `"B"` |
| `DataRange(ws, [headerRow])` / `BodyRange(ws, [headerRow])` | The real block, not `UsedRange` |
| `RangeToArray(rng)` | Always 2D, even for one cell |
| `WriteArray(values, topLeftCell, [asRow])` | 1D or 2D, in a single write |
| `HeaderMap(ws, [headerRow])` | Header text → column number |
| `FindColumnByHeader(ws, headerText, [headerRow])` | 0 when the header is missing |
| `FindCell(searchRange, what, [wholeCell], [searchValues])` | `Find` that returns `Nothing` instead of raising |
| `ColumnLetter(n)` / `ColumnNumber(letters)` | `28` ↔ `"AB"` |
| `IsRangeEmpty(rng)` / `CopyValues(source, destination)` | Values only, no clipboard |
| `DeleteEmptyRows(ws, [checkColumn], [firstRow])` | One delete for the lot |
| `AutoFitColumns(ws, [maxWidth], [minWidth])` | AutoFit with an upper bound |
| `FreezeHeader(ws, [rows], [columns])` | Restores the sheet that was active |

### modString — text

| Function | Notes |
| --- | --- |
| `CleanText(value)` | Non-breaking spaces, control characters, doubled spaces |
| `IsBlank(value)` | Empty, Null, an error, or whitespace only |
| `StartsWith` / `EndsWith` / `ContainsText` | Case insensitive by default |
| `TextBefore` / `TextAfter` / `TextBeforeLast` / `TextAfterLast` | Slice around a delimiter |
| `PadLeft` / `PadRight` / `Repeat` | |
| `JoinNonBlank(values, [delimiter])` | Skips the empty ones |
| `KeepDigits(text)` / `ExtractNumber(text, [default])` | `"Fee: 1,250.75 GBP"` → `1250.75` |
| `TitleCase(text)` | Copes with `O'Neill` and `Smith-Jones` |
| `RegExTest` / `RegExReplace` / `RegExExtract` / `RegExMatches` | Late bound, no reference needed |
| `Levenshtein(a, b, [ignoreCase])` / `SimilarityPercent(a, b, …)` | Fuzzy matching for reconciliations |

### modWorkbook — workbooks and sheets

| Function | Notes |
| --- | --- |
| `SheetExists(name, [wb])` / `GetSheet(name, [wb])` | `GetSheet` returns `Nothing`, never raises |
| `GetOrCreateSheet(name, [wb], [clearIfExists])` | |
| `ClearSheet(ws, [keepFormats])` | Contents, formats, shapes, filters, merges |
| `DeleteSheet(name, [wb])` | No prompt; refuses to delete the last visible sheet |
| `SafeSheetName(name, [replacement])` / `UniqueSheetName(base, [wb])` | 31 characters, no `: \ / ? * [ ]` |
| `SheetNames([wb], [visibleOnly])` | |
| `WorkbookIsOpen(nameOrPath)` / `GetOpenWorkbook(nameOrPath)` | Name or full path |
| `OpenWorkbook(path, [readOnly], [updateLinks])` | Hands back the copy that is already open |
| `DumpToSheet(data, sheetName, [wb], [hasHeaderRow])` | Array to a formatted sheet in one go |
| `ExportToPdf(target, path, [openAfter])` | Workbook, sheet or range |
| `RefreshAllAndWait([wb])` | Unlike `RefreshAll`, does not return early |
| `ProtectAllSheets([wb], [password])` / `UnprotectAllSheets(…)` | |

---

## Workbook Doctor (the add-in)

A workbook that has had sheets copied into it for a few years quietly collects
tens of thousands of cell styles, hundreds of broken defined names, a used range
that runs to row 1,048,576 and a pile of invisible drawing objects. It opens
slowly, saves slowly, and one day says *"Too many different cell formats"*.

Workbook Doctor finds all of that and offers to remove it — and reviews the
formulas while it is in there. It lives on the **Add-ins** tab of the ribbon.

**Start with Audit workbook.** It changes nothing and its first section is a
prioritised list of what is actually worth doing to the workbook in front of you:

```
What is worth doing
Priority   Tool                       Why
High       Clean styles               8,412 duplicate styles such as 'Normal 2'
High       Reset used range           1,048,300 empty rows are still counted as used
Medium     Clean names                37 broken, external or hidden name(s)
High       Fix by hand                2 name(s) refer to #REF! and are still used by formulas
Low        Remove invisible objects   14 hidden or zero-size drawing object(s)
```

Then run the tools it points at, or **Clean everything (safe)** for the low-risk
half of all of them in one pass.

**Audit model** is the other half, and the one to reach for when management
hands over a model and you have two days to decide whether to trust it. It
changes nothing, and it finds what a reviewer scrolling through cells will not:

```
High    Forecast   D14   Number typed into a calculated row
                         The rest of this row calculates; this cell holds a typed number.
High    Forecast   H22   Number inside a formula
                         The value 1.03 is typed into the formula. An assumption belongs
                         in a cell of its own.  (this formula appears 48 times)
Medium  Opex       B9    Formula differs from the row
                         The other 47 cells in this run use =RC[-1]*(1+R4C2)
```

The rule it runs on is the one every model audit tool is built on: in a working
model a row holds one formula copied across, so read the block in R1C1 and the
odd one out falls straight out of the comparison. Then it reads each distinct
formula once to find assumptions typed into the middle of a calculation —
skipping the ones that are really indexes, so `VLOOKUP(x,y,3,0)` stays quiet
while `ROUND(B5*1.2,2)` does not.

**Formula map** draws the sheet one character per cell — inputs, formulas, and
the cells that break the pattern — so its shape is visible in a single screen.
**Select flagged cells** picks out the suspects on the active sheet so you can
look at them where they live.

### The data tools

The other half of the job: thirty-six monthly extracts that need to become one
table before anything can be analysed.

**Stack selected sheets** — Ctrl+click the tabs, run it. Columns are matched by
header *name*, not position, so a source with its columns in a different order
lands in the right place and one missing a column leaves blanks. The stack
report names which source was missing what, which is usually the finding rather
than the footnote. **Stack files in a folder** does the same across a directory.

**Unpivot a cross-tab** turns months running across the top into rows down the
side. **Fill blanks down** fixes the group label that appears once and is blank
for the next forty rows. Between them they are what stands between a management
extract and a pivot table.

**Fuzzy match two lists** runs three passes — exact once both sides are tidied,
exact once company decoration is stripped (`THE Acme Group Ltd.` and `Acme Group
Limited` are the same company), then closest match by edit distance. It reports
the best match, its score, the verdict, and the *next best* match with its score,
because when the top two are close the match is a coin toss however high the
number looks. The `Your call` column is left empty on purpose.

| | |
| --- | --- |
| **Names** | Classifies every name as reserved, in use, broken, external, hidden or unused, and deletes only the categories you approve. "Unused" is worked out by reading every formula, conditional format, validation rule, chart series, pivot source and other name — so it is a good guess, not a fact, which is why deleting those is a separate opt-in |
| **Styles** | Duplicates-only mode removes `Normal 2`, `Comma 3 4` and friends; the aggressive mode removes every custom style. Esc stops a long run |
| **Used range** | Trims each sheet back to its real data, protecting anything anchored under a shape or chart |
| **Sheets** | Deletes genuinely empty sheets, unhides very hidden ones, removes invisible and zero-size drawing objects |
| **Conditional formatting** | Counts rules per sheet and removes the ones that apply only to empty cells |
| **Links** | Lists every linked workbook, says which files have gone missing, shows the cells that depend on them, and breaks them on request |
| **Selection tools** | Trim and clean text, freeze formulas to values, strip hyperlinks |
| **Model integrity** | Row-by-row formula consistency, numbers typed over formulas, hardcoded assumptions, volatile functions, whole-column references, error cells — all read-only |
| **Data** | Stack many extracts into one table matched by header name, unpivot a cross-tab, fill blanks down, fuzzy-match two lists that nearly agree |
| **Backup** | Timestamped copy beside the original, offered automatically before anything destructive |

Nothing it does can be undone with Ctrl+Z, so every destructive action confirms
first and offers a backup. Read
[addin/INSTALL.md](addin/INSTALL.md) for the build, the full menu and the
caveats worth knowing before you point it at a workbook that matters.

### Building it

The `.xlam` is not committed — a binary in git is a binary nobody can review.
Build it from source in about thirty seconds:

```powershell
cd build
.\Build-AddIn.ps1 -Install
```

or import the modules by hand: [addin/INSTALL.md](addin/INSTALL.md).

---

## Repository layout

```
src/        the library - eight standalone modules, no dependencies
addin/      Workbook Doctor - the add-in, built on src
build/      Build-AddIn.ps1, which assembles the .xlam
            test_formula_parser.py, which checks the audit's parsing rules
examples/   runnable examples of the library patterns
```

The formula parser behind the hardcode check is the one piece of logic here
with real edge cases, so it has a test that runs outside Excel:

```
python build/test_formula_parser.py     # 29/29 passed
```

---

## Conventions

- `Option Explicit` everywhere.
- Functions return a safe empty value — `0`, `""`, `Nothing`, `Array()` —
  rather than raising, so callers can test the result instead of wrapping
  every call in an error handler.
- Arrays returned by these functions are 1-based, matching what Excel hands
  back from `Range.Value`.
- Optional arguments come last and default to the common case.
- `Scripting.Dictionary`, `Scripting.FileSystemObject`, `VBScript.RegExp` and
  `ADODB.Stream` are all late bound, so there are no references to add and
  nothing to break on another machine.

## Compatibility

Written for Excel 2010 and later on Windows, 32 or 64 bit — there are no API
declarations, so nothing needs `PtrSafe`. On Mac, everything except the
`Scripting`, `VBScript.RegExp` and `ADODB` based routines will work; that rules
out the regular expression helpers, most of `modFile`, and anything returning a
dictionary.

## Licence

MIT — see [LICENSE](LICENSE).
