# Tools used for the audit

The workbook is an `.xlsb` (binary) file, which LibreOffice could not open, so the audit
worked directly on the package parts.

## Extracting the VBA
```bash
pip install oletools pyxlsb openpyxl
unzip -o HP_Flat_File_Test_v3.3.xlsb -d xlsb
olevba --decode xlsb/xl/vbaProject.bin > olevba_full.txt      # all module sources
```
The per-module files in `../original/` were split out of that output.

## Reading the workbook without Excel
- `parse_biff.py` - minimal MS-XLSB record reader; maps sheet tab names to VBA code names
  (BrtBundleSh + BrtWsProp) and decodes the simple defined names (BrtName, PtgRef3d/PtgArea3d).
- `cellscan.py` - classifies every cell of a sheet part as formula vs constant and lists which
  columns of the source tables hold formulas (used to prove which columns an inserted row lacks).
- `pyxlsb` was used for cell values (FormulaBank dump, headers, key columns, data checks).

## Putting the fixed modules back into the workbook
1. Take a copy of the workbook. Open it, press Alt+F11.
2. In the Project Explorer remove the orphaned module `Sheet731` (right-click > Remove, No to export).
3. `Main`: right-click > Remove, then File > Import File... `fixed/Main.bas`.
4. `Sheet730 (Fund View)`: open it, select all, paste the contents of `fixed/Sheet730.cls`.
5. `ThisWorkbook`: open it, select all, paste the contents of `fixed/ThisWorkbook.cls`.
6. Debug > Compile VBAProject must report no errors before saving as `.xlsb`.
7. Run the UAT script in `../UAT_TEST_PLAN.md` on the copy.
