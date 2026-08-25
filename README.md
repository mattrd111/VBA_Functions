# VBA_Functions

Useful Excel VBA functions — a small standard library for the things every
Excel macro ends up needing: finding the edges of your data, moving whole
blocks in and out of arrays, fast lookups, working-day maths, tidying text,
handling files, and leaving Excel exactly as you found it.

Nothing here needs a project reference, a class module or an add-in. Import a
`.bas` file and the functions are available.

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
