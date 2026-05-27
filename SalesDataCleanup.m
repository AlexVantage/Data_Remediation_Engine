// Sales Data Cleanup — Power Query M pipeline
// Companion to dirty-sales-export.csv and README.md
//
// Setup (once):
//   1. Data → Get Data → Blank Query → rename to "SourceFilePath"
//   2. Advanced Editor → paste the SourceFilePath query below → point path at your CSV
//   3. New Blank Query → rename to "SalesDataCleanup" → paste the main query below
//
// Refresh: Data → Refresh All (or Ctrl+Alt+F5)

// ── Query 1: SourceFilePath (Parameter) ─────────────────────────────────────

let
    Source = "C:\Data\sales\dirty-sales-export.csv" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]
in
    Source


// ── Query 2: SalesDataCleanup (Main pipeline) ─────────────────────────────────

let
    // ── Pillar A: Data Ingestion & Connection ───────────────────────────────
    // Explicit external file reference via parameter — swap the path once, reuse everywhere.
    Source = Csv.Document(
        File.Contents(SourceFilePath),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),

    // ── Pillar B: Text & Structural Normalization ───────────────────────────
    // Trim → Clean (non-printable chars) → Capitalize Each Word (Text.Proper)
    TrimmedCustomerName = Table.TransformColumns(
        PromotedHeaders,
        {{"Customer Name", Text.Trim, type text}}
    ),
    CleanedCustomerName = Table.TransformColumns(
        TrimmedCustomerName,
        {{"Customer Name", Text.Clean, type text}}
    ),
    NormalizedCustomerName = Table.TransformColumns(
        CleanedCustomerName,
        {{"Customer Name", Text.Proper, type text}}
    ),

    // Split compound "City State Zip" field by comma delimiter
    SplitLocation = Table.SplitColumn(
        NormalizedCustomerName,
        "City State Zip",
        Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv),
        {"City", "State", "Zip"}
    ),
    TrimmedLocation = Table.TransformColumns(
        SplitLocation,
        {
            {"City", Text.Trim, type text},
            {"State", Text.Trim, type text},
            {"Zip", Text.Trim, type text}
        }
    ),

    // ── Pillar C: Logical Type Casting & Date Standardization ───────────────
    NullSentinels = {"N/A", "", " "},

    ReplacedOrderDate = Table.ReplaceValue(
        TrimmedLocation,
        each [Order Date],
        each if List.Contains(NullSentinels, Text.Trim(Text.From(_))) then null else _,
        Replacer.ReplaceValue,
        {"Order Date"}
    ),

    // Change Type with Locale: DD/MM/YYYY dates on a US-default Excel install
    TypedDates = Table.TransformColumnTypes(
        ReplacedOrderDate,
        {{"Order Date", type date}},
        "en-GB"
    ),

    // Extract currency symbol & text cleanup on Revenue before numeric cast
    CleanedRevenueText = Table.TransformColumns(
        TypedDates,
        {
            {
                "Revenue",
                each
                    let
                        raw = Text.Trim(Text.From(_)),
                        noCurrency = Text.Remove(raw, {"$"}),
                        noSpaces = Text.Remove(noCurrency, {" "}),
                        noThousands = Text.Remove(noSpaces, {","})
                    in
                        noThousands,
                type text
            }
        }
    ),

    // Replace N/A and blank strings with true null (Region + Revenue)
    ReplacedRegion = Table.ReplaceValue(
        CleanedRevenueText,
        each [Region],
        each if List.Contains(NullSentinels, Text.Trim(Text.From(_))) then null else _,
        Replacer.ReplaceValue,
        {"Region"}
    ),
    ReplacedRevenue = Table.ReplaceValue(
        ReplacedRegion,
        each [Revenue],
        each if List.Contains(NullSentinels, Text.Trim(Text.From(_))) then null else _,
        Replacer.ReplaceValue,
        {"Revenue"}
    ),

    TypedRevenue = Table.TransformColumnTypes(
        ReplacedRevenue,
        {{"Revenue", Currency.Type}}
    ),

    // ── Pillar D: Row/Column Deduplication and Schema Control ───────────────
    RemovedDuplicates = Table.Distinct(TypedRevenue, {"Order ID"}),

    FinalSchema = Table.SelectColumns(
        RemovedDuplicates,
        {"Order ID", "Customer Name", "City", "State", "Zip", "Order Date", "Revenue", "Region"}
    )
in
    FinalSchema
