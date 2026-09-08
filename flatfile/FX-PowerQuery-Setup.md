# FX Input as a Power Query connection

The FX tab's instructions say: open the ECB euro reference rates page, copy the whole
CSV into the FX Input tab, calculate, and the FX tab picks up the new rates. The query
in `FX_Input.pq` does that copy for you. Everything downstream is untouched:

- **FX tab** rows 4 to 16 and the "For FX formula in underlying fund tabs" block look
  rates up from FX Input by currency header and by the dates in row 1.
- **Fund Summary** helper columns CA (fx_row), CB (fx_ccy_col) and CC (fx_at_comp) read
  `'FX Input'!$A$1:$AP$1` by header name and `'FX Input'!$A$2:$A$7039` by date.
- **Fund View** only reads the FX tab (C25 for the model-date rate, K22/K23 for the
  per-date rates), never FX Input.

So the query only has to land data on the FX Input sheet in the same shape: header row
in row 1, "Date" in column A, ECB currency columns in ECB order, newest date first.
Power Query works in .xlsb files, so the Flat File does not need converting.

## Setting it up (once)

1. Open the Flat File. On the **FX Input** sheet, note the top date in A2 and copy
   the FX tab's B4:H16 values somewhere for a before-and-after check.
2. Clear the FX Input sheet (Ctrl+A, Delete). Keep the sheet.
3. **Data > Get Data > From Other Sources > Blank Query**. In the editor choose
   **Home > Advanced Editor**, delete the default text, paste the whole of
   `FX_Input.pq`, click **Done**.
4. In the Query Settings pane rename the query to `FX_Input`.
5. If Excel asks about privacy levels, set the ECB URL to **Public**.
6. **Home > Close & Load To...** and pick **Table**, **Existing worksheet**,
   `='FX Input'!$A$1`. Excel writes the table with the headers in row 1.
7. Right-click the query in the Queries & Connections pane > **Properties**:
   untick *Enable background refresh*, tick *Refresh data when opening the file*
   if you want it automatic. Leave *Refresh every n minutes* off.
8. Recalculate. FX tab B4:H16 should equal the values you copied in step 1
   (the ECB file is the same source, so only newly published dates can differ).
   The FX tab's `Check` cell at N22 should still show `ok`.

## Refreshing

**Data > Refresh All**, or right-click the table on FX Input > Refresh. The
`RefreshFXInput` macro in `modFX.bas` does the same thing from VBA and reports the newest
date it pulled; import it into the Flat File's VBA project if you want a button.

The roll-forward step on the FX tab (update B1 with the new month-end date, shift the
older dates along C1:H1) is still manual. The formulas below row 1 look the rates up by
those dates, so nothing else needs copying.

## Things to know

- **Proxy.** The query needs outbound HTTPS to `www.ecb.europa.eu`. If the corporate
  proxy blocks it, download `eurofxref-hist.csv` from the ECB page and switch the
  `Raw` step to the `File.Contents` line that is already in the query.
- **N/A becomes blank.** The pasted sheet held the text "N/A" for discontinued
  currencies (CYP, EEK, LTL, LVL, MTL, ROL, SIT, SKK, TRL, HRK, RUB, ISK gaps). The query
  makes those blank so the rate columns are numeric. None of those currencies appear on
  the FX tab or in any fund, so no formula changes value.
- **Row count.** Fund Summary CA fixes its date range at `$A$2:$A$7039`. Because the
  table is newest first, new dates enter at the top and only 1999-era dates fall outside
  the range. If you ever want that formula to grow with the table, change the range to
  `FX_Input[Date]`.
- **Units.** ECB rates are currency per 1 EUR. The FX tab divides each currency by the
  USD column to get currency per 1 USD, which is what the fund sheets use. The query does
  not change that.
- **Timing.** The ECB publishes around 16:00 CET on TARGET business days. Refreshing on
  a weekend or before that time returns the previous business day, which the XMATCH
  "exact or next smaller" lookups already handle.
