// Shopify Data Cleanup — Power Query M pipeline
// Companion to shopify_data_raw.csv and README.md
//
// Setup (once):
//   1. Data → Get Data → Blank Query → rename to "SourceFilePath"
//   2. Advanced Editor → paste the SourceFilePath query below → point path at your CSV
//   3. New Blank Query → rename to "ShopifyDataCleanup" → paste the main query below
//
// Refresh: Data → Refresh All (or Ctrl+Alt+F5)

// ── Query 1: SourceFilePath (Parameter) ─────────────────────────────────────

let
    Source = "C:\Data\shopify\shopify_data_raw.csv" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]
in
    Source


// ── Query 2: ShopifyDataCleanup (Main pipeline) ─────────────────────────────

let
    // ── Pillar A: Data Ingestion & Connection ───────────────────────────────
    // Explicit external file reference via parameter — swap the path once, reuse everywhere.
    Source = Csv.Document(
        File.Contents(SourceFilePath),
        [Delimiter = ",", Columns = 11, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),

    // ── Pillar B: Text & Structural Normalization ───────────────────────────
    // Customer_Name: Trim → Clean (non-printable chars) → Capitalize Each Word
    TrimmedName = Table.TransformColumns(
        PromotedHeaders,
        {{"Customer_Name", Text.Trim, type text}}
    ),
    CleanedName = Table.TransformColumns(
        TrimmedName,
        {{"Customer_Name", Text.Clean, type text}}
    ),
    NormalizedName = Table.TransformColumns(
        CleanedName,
        {{"Customer_Name", Text.Proper, type text}}
    ),

    // Customer_Email: collapse mixed case to lowercase (emails are case-insensitive)
    LowercasedEmail = Table.TransformColumns(
        NormalizedName,
        {{"Customer_Email", Text.Lower, type text}}
    ),

    // Phone_Number: strip all non-digits, then re-format to a single standard (XXX) XXX-XXXX
    NormalizedPhone = Table.TransformColumns(
        LowercasedEmail,
        {
            {
                "Phone_Number",
                each
                    let
                        digits = Text.Select(Text.From(_), {"0".."9"}),
                        last10 = if Text.Length(digits) >= 10 then Text.End(digits, 10) else digits,
                        formatted =
                            if Text.Length(last10) = 10 then
                                "(" & Text.Start(last10, 3) & ") "
                                    & Text.Middle(last10, 3, 3) & "-"
                                    & Text.End(last10, 4)
                            else
                                last10
                    in
                        formatted,
                type text
            }
        }
    ),

    // Split SKU (PROD-CATEGORY-COLOR-SIZE) by hyphen into 4 columns
    SplitSKU = Table.SplitColumn(
        NormalizedPhone,
        "SKU",
        Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv),
        {"SKU_Prefix", "Category", "Color", "Size"}
    ),

    // ── Pillar C: Logical Type Casting & Date Standardization ───────────────
    // Order_Date arrives as a mix of real dates and US-format text dates (MM/DD/YYYY).
    // No locale override needed — the macro writes US-format, matching a US Excel install.
    TypedColumns = Table.TransformColumnTypes(
        SplitSKU,
        {
            {"Transaction_ID", type text},
            {"Customer_Name", type text},
            {"Customer_Email", type text},
            {"Phone_Number", type text},
            {"Order_Date", type date},
            {"SKU_Prefix", type text},
            {"Category", type text},
            {"Color", type text},
            {"Size", type text},
            {"Region", type text},
            {"Sales_Channel", type text},
            {"Fulfillment_Status", type text},
            {"Revenue_USD", Currency.Type},
            {"Quantity", Int64.Type}
        }
    ),

    // Replace blank/whitespace Region cells with true null
    ReplacedRegion = Table.ReplaceValue(
        TypedColumns,
        each [Region],
        each if [Region] = null or Text.Trim(Text.From([Region])) = "" then null else [Region],
        Replacer.ReplaceValue,
        {"Region"}
    ),

    // Numeric nulls: leave Revenue_USD / Quantity as null (cleaner for analysis than 0,
    // which would distort averages). Flip to 0 here only if your downstream model requires it.
    // ReplacedNulls = Table.ReplaceValue(ReplacedRegion, null, 0, Replacer.ReplaceValue, {"Revenue_USD", "Quantity"}),

    // ── Pillar D: Row Deduplication & Schema Control ────────────────────────
    // Full-row duplicates were introduced via Transaction_ID; dedupe on the primary key.
    RemovedDuplicates = Table.Distinct(ReplacedRegion, {"Transaction_ID"}),

    FinalSchema = Table.SelectColumns(
        RemovedDuplicates,
        {
            "Transaction_ID", "Customer_Name", "Customer_Email", "Phone_Number",
            "Order_Date", "Category", "Color", "Size",
            "Region", "Sales_Channel", "Fulfillment_Status", "Revenue_USD", "Quantity"
        }
    )
in
    FinalSchema