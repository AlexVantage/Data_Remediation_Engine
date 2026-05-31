# Shopify Data Cleanup

Portfolio sample for [AlexVantage](https://alexvantage.com) — a deliberately messy 90-day Shopify-style retail transaction export and a reusable Power Query pipeline organized around four data-quality pillars.

**Case study on the site:** [alexvantage.com/projects/sales-data-cleanup](https://alexvantage.com/projects/sales-data-cleanup)

## Files

| File | Purpose |
|------|---------|
| `shopify_dump_dirty.csv` | Raw export with intentional quality issues |
| `defect_macro.vba` | VBA macro that generates the defects (setup, not solution) |
| `ShopifyDataCleanup.m` | Full M-language pipeline (copy into Advanced Editor) |
| `README.md` | Step-by-step walkthrough (this document) |

---

## Dirty data inventory

The sample export contains ~750 rows of retail transactions with defects introduced via `defect_macro.vba`. Each pillar maps to specific issues:

| Pillar | Column(s) | Issues baked in |
|--------|-----------|-----------------|
| **A — Ingestion** | (file-level) | External CSV meant to be connected, not pasted into a sheet |
| **B — Text & structure** | `Customer_Name`, `Customer_Email`, `Phone_Number`, `SKU` | Leading/trailing whitespace, inconsistent casing, mixed-case emails, four inconsistent phone formats, compound SKU field |
| **C — Types & nulls** | `Order_Date`, `Revenue_USD`, `Quantity`, `Region` | Some dates stored as text, blank cells in numeric and region fields |
| **D — Dedupe & schema** | `Transaction_ID` | Duplicate transaction rows, raw SKU column dropped after split |

---

## Before you start

1. Copy `shopify_dump_dirty.csv` to a stable folder, e.g. `C:\Data\shopify\`.
2. Open Excel → **Blank workbook**.
3. The dataset uses **US-format dates (MM/DD/YYYY)**, so a default US Excel install parses them correctly — no locale override required.

---

## Pillar A — Data Ingestion & Connection

**Goal:** Connect to the external CSV by reference, not a one-time import. A parameter makes the path reusable across refreshes and environments.

### Steps (Excel UI)

1. **Data → Get Data → From Other Sources → Blank Query.**
2. **Home → Advanced Editor.** Replace contents with the `SourceFilePath` query from `ShopifyDataCleanup.m` (top block).
3. Update the path string to your copy of `shopify_dump_dirty.csv`.
4. Marking `IsParameterQuery = true` in the `meta` record (as in the `.m` file) exposes it under **Queries & Connections → Parameters**.
5. Rename the query **`SourceFilePath`**.

### Why this matters

Pasting CSV data into a worksheet creates a static snapshot. A parameterized `File.Contents` connection lets the team drop a new weekly export in the same folder, refresh, and get the full clean pipeline — no rework.

---

## Pillar B — Text & Structural Normalization

**Goal:** Standardize names, emails, and phone numbers, and split the compound SKU field.

### Steps (Excel UI)

1. **Data → Get Data → From Other Sources → Blank Query** → rename **`ShopifyDataCleanup`**.
2. In Advanced Editor, start from the `Source` and `PromotedHeaders` steps in `ShopifyDataCleanup.m`, referencing `SourceFilePath`.
3. Select **`Customer_Name`** → **Transform → Format → Trim**, then **Clean**, then **Capitalize Each Word**.
   - M equivalents: `Text.Trim` → `Text.Clean` → `Text.Proper`
4. Select **`Customer_Email`** → **Transform → Format → lowercase**.
   - M equivalent: `Text.Lower` (emails are case-insensitive; normalizing prevents duplicate-looking records)
5. **`Phone_Number`** — the macro produces four formats: `(555) 123-4567`, `555-123-4567`, `5551234567`, `555.123.4567`. Strip to digits and re-apply one standard. See the `NormalizedPhone` step in the `.m` file.
6. Select **`SKU`** → **Transform → Split Column → By Delimiter** → hyphen → split into **4 columns**.
7. Rename the new columns **`SKU_Prefix`**, **`Category`**, **`Color`**, **`Size`**.

### Spot-check after Pillar B

| Before | After |
|--------|-------|
| `  alex johnson  ` | `Alex Johnson` |
| `ALEX@VANTAGE.COM` | `alex@vantage.com` |
| `555.123.4567` | `(555) 123-4567` |
| `PROD-SHIRT-RED-LRG` | SHIRT / RED / LRG (prefix dropped at schema step) |

---

## Pillar C — Logical Type Casting & Null Handling

**Goal:** Coerce text-formatted dates back to real dates, cast revenue to currency, and convert blank region cells to true nulls.

### Step 1 — Dates

1. Select **`Order_Date`** → **Transform → Data Type → Date**.
2. Because the macro writes US-format text dates (`MM/DD/YYYY`), a US Excel install parses them without a locale override.
   - M equivalent: `{"Order_Date", type date}` in `Table.TransformColumnTypes`

### Step 2 — Revenue & Quantity

1. Select **`Revenue_USD`** → **Transform → Data Type → Currency**.
2. Select **`Quantity`** → **Transform → Data Type → Whole Number**.

### Step 3 — Null handler

1. Select **`Region`** → **Transform → Replace Values** → replace blank/whitespace with empty → mark as null.
   - M approach: see the `ReplacedRegion` step in `ShopifyDataCleanup.m`.
2. **Design choice:** blank `Revenue_USD` and `Quantity` cells are left as **null**, not `0`. Zeros would distort averages and totals. If your downstream model needs zeros, uncomment the `ReplacedNulls` step in the `.m` file.

---

## Pillar D — Row Deduplication & Schema Control

**Goal:** One row per transaction and a lean schema for downstream reporting.

### Steps (Excel UI)

1. Select **`Transaction_ID`** → **Home → Remove Rows → Remove Duplicates** (keeps first occurrence).
   - M equivalent: `Table.Distinct(..., {"Transaction_ID"})`
2. **Home → Choose Columns** → drop the leftover **`SKU_Prefix`** column (always `PROD`, adds nothing).
   - M equivalent: `Table.SelectColumns(...)`
3. Final column order: Transaction_ID → Customer_Name → Customer_Email → Phone_Number → Order_Date → Category → Color → Size → Region → Sales_Channel → Fulfillment_Status → Revenue_USD → Quantity.

### Result

The 15 duplicate rows introduced by the macro collapse back to the unique transaction set, and the redundant `SKU_Prefix` column is excluded from the output.

---

## Load & refresh

1. **Home → Close & Load To… → Table** (or **Only Create Connection** if feeding a data model).
2. Each week: replace the CSV in the source folder → **Data → Refresh All**.

---

## How the dirty data was generated

The defects are introduced programmatically by `defect_macro.vba` (run on a clean export), so the dataset is fully reproducible. The macro introduces defects only — it is **not** the cleaning solution. The cleaning is the `ShopifyDataCleanup.m` pipeline above.

---

## Certification notes

- **PL-300 / DP-600:** Parameterized sources, type casting, text normalization, and schema shaping before load.
- **Portfolio:** Linked from the [Sales Data Cleanup](https://alexvantage.com/projects/sales-data-cleanup) case study on alexvantage.com.

---

## Related

- [AlexVantage](https://alexvantage.com) — data operations consultancy
- [Spreadsheet cleanup service](https://alexvantage.com/services)