# Fund Onboarder: modAddFunds

`modAddFunds.bas` is the import macro that lives in `Fund_Onboarder.xlsm`. It reads the
Summary and Detailed sheets and appends rows to the flat tables in
`HP Flat File Test v3.6.xlsb` (fund, position, company, position_company, company_nav).

## Installing the module

1. Open `Fund_Onboarder.xlsm`, press Alt+F11.
2. In the Project pane, right-click the existing `modAddFunds` module and choose
   **Remove modAddFunds...** (answer No to exporting, or export it first if you want a backup).
3. **File > Import File...** and pick `modAddFunds.bas`.
4. The Summary button still calls `AddNewFundsAndCompanies`, so nothing else changes.

## Where each Onboarder value lands

The Fund View in the flat file looks every company-block cell up by
`fund & "|" & ROW()` against `lookup_key`, so rows have to be keyed by Fund View row number.

| Onboarder source | Flat table column | Fund View cell |
|---|---|---|
| Summary fund fields (name, Preqin name, manager, strategy, geography, vintage, carry, fees, size) | `fund.*` | C8 to C16 |
| Summary position fields (val date, transfer date, commitment, uncalled, headline price, % held, forecasts) | `position.*` | C28 onward, F-column forecasts |
| Detailed company name, geo, industry, classification, public/private | `company.*` on row 121+ | J to N |
| Detailed NAV block (HA) | `position_company.ref_date_nav` on row 121+ | O "Ref Date NAV" |
| Detailed block headed "Reinvested" (IP), which the template formulas fill with Known + Assumed + Post Comp. calls | `position_company.forecast_calls` on row 121+ | P "F'cast Calls" |
| Detailed Overall Upside (KE) | `position_company.multiple` on row 121+ | R "Multiple" |
| Detailed HPT Val methodology (LT), HPT Info Quality (NI) | `position_company.methodology`, `.information_quality` | S, T |
| Detailed Net Curr (OZ), Net Curr Geog (PF), F'Cast Calls (PE) | Net Curr line on row 120 (`ref_date_nav`, `company.geography`, `forecast_calls`) | row 120 |
| Detailed Carry (PA), Carry Multiple (PB) | row 117 `ref_date_nav`, `multiple` | O117, R117 |
| Detailed Fees (PC), F'Cast Fees Multiplier Value (PD) | row 118 `ref_date_nav`, `multiple` = PD / PC | O118, R118 |
| Summary NAV plus Calls Upside (Z), Standard funds only | `multiple` on rows 116, 117, 119, 120 of the placeholder block | F35 "Upside" via R116 |
| NAV per company at the reference date | `company_nav.nav_value` keyed `fund|row|yyyy-mm-dd` | AT onward, feeds AA "Current Value" |

Fund View column I ordinal (1 for Net Curr, 2 for the first company) is stored in
`position_company.gross_nav_input`; on rollup rows that column holds the row label.

## Standard versus Detailed funds

There is no methodology column in the flat file. Standard funds carry a placeholder block:
rollup rows 116 to 119 with NAV 1, one Net Curr row 120 with NAV 1, and multiple equal to the
fund upside on rows 116, 117, 119 and 120. The macro writes that block when Summary column C
is anything other than "Detailed", or when a Detailed fund has no row on the Detailed sheet.

## Still logged, not written

These have no home in the flat file and stay on the Import Log sheet:

- Fund currency. Summary only has an FX rate. A currency code is needed in
  `fund.fund_currency` and `position.underlying_fund_currency` (the Fund View reads the
  latter to pick FX rates).
- Net Curr Reinvested (PG). Nothing on the Fund View reads a reinvested flag.
- Summary NAV (Q). The Fund View takes NAV from `valuation_history`, which the macro does not write.
- NAV plus Calls Upside when it does not look like a multiple (outside 0 to 5). The macro
  writes 1 and logs the raw value.
