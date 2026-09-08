// FX_Input - Power Query (M) for the Flat File's "FX Input" sheet.
//
// Replaces the manual step on the FX tab ("copy and paste the entire .csv from the ECB
// link into FX Input"). Pulls the ECB euro reference rate history straight from the ECB
// and lands it in exactly the shape the workbook already expects:
//   - one row per ECB business day, NEWEST FIRST (Fund Summary CA uses XMATCH with
//     match_mode -1 / search_mode -2, which needs descending dates)
//   - column A "Date", then the ECB currency columns in ECB order (USD, JPY, BGN, ...)
//   - rates are units of currency per 1 EUR, as published; the FX tab divides by USD
//   - "N/A" in the ECB file becomes blank (the old paste left the text "N/A")
//
// Source: the same page the FX tab links to. The page offers eurofxref-hist.zip; the
// uncompressed CSV sits beside it at the URL below and is what this query reads.
let
    // ---- source ---------------------------------------------------------------
    // Web source (default). If the proxy blocks www.ecb.europa.eu, download the CSV
    // manually and swap the next line for the File.Contents line beneath it.
    Raw = Web.Contents("https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist.csv"),
    // Raw = File.Contents("C:\path\to\eurofxref-hist.csv"),

    Csv = Csv.Document(Raw, [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Csv, [PromoteAllScalars = true]),

    // ---- shape ----------------------------------------------------------------
    // Every ECB line ends with a trailing comma, which arrives as an unnamed empty column.
    KeepNamed = Table.SelectColumns(
        Promoted,
        List.Select(Table.ColumnNames(Promoted), each Text.Trim(_) <> "" and not Text.StartsWith(_, "Column"))
    ),
    Trimmed = Table.TransformColumnNames(KeepNamed, Text.Trim),

    // Date column -> real dates.
    Dated = Table.TransformColumnTypes(Trimmed, {{"Date", type date}}, "en-GB"),

    // Rate columns -> decimal numbers; "N/A" (discontinued currencies, gaps) -> null.
    RateCols = List.RemoveItems(Table.ColumnNames(Dated), {"Date"}),
    Cleaned = Table.ReplaceValue(Dated, "N/A", null, Replacer.ReplaceValue, RateCols),
    Typed = Table.TransformColumnTypes(
        Cleaned,
        List.Transform(RateCols, each {_, type number}),
        "en-US"
    ),

    // Newest first, matching the pasted layout and the XMATCH descending search.
    Sorted = Table.Sort(Typed, {{"Date", Order.Descending}}),
    Result = Table.Buffer(Sorted)
in
    Result
