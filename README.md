# Data Remediation Engine

**An Excel-native data validation and cleansing engine that takes a messy retail transaction export and produces an audited, analysis-ready dataset — with a self-testing QA harness that proves the logic works.**

Part of the [AlexVantage](https://alexvantage.com) Excel Asset Suite · Data Operations Consultancy

---

## Problem

Raw operational exports (Shopify dumps, CRM extracts, manual staff entry) arrive full of structural defects: non-printable characters, malformed IDs, mixed-case emails with broken domains, inconsistent phone formats, out-of-vocabulary categories, negative revenue, and missing fields. Pushed downstream untouched, this corrupts reports, breaks pivots, and erodes trust in every number that follows.

The usual "fix" is someone eyeballing 1,000 rows by hand. That doesn't scale, isn't repeatable, and leaves no record of *what* was wrong.

## What It Does

This workbook runs a three-stage pipeline plus a validation layer that **classifies** defects instead of silently overwriting them — so every correction is auditable.

- **Validates** more than 1,000 transaction records against field-level rules using Excel REGEX functions
- **Classifies** 8 distinct defect types with explicit `ERROR:` codes (e.g. `ERROR: Invalid TLD`, `ERROR: Invalid Currency`) rather than blanking bad data
- **Checks email domains** against a 1,441-entry IANA top-level-domain reference list — not just a "contains an @" check
- **Normalizes** clean values: title-casing names, lowercasing emails, standardizing phone formats, coercing dates
- **Self-tests** the entire ruleset with a 52-case adversarial QA matrix that compares each formula's output to its expected output — currently **52/52 passing**

## Architecture

A two-pass pipeline keeps raw data untouched and every transformation traceable:


| Sheet             | Role                                                                                                                        |
| ----------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `03_RAW_DATA`     | Zero-modification intake layer. Raw export lands here and is never edited.                                                  |
| `02_STAGING_DATA` | Validation engine. Each cell is tested with REGEX rules; failures emit a structured `ERROR:` code naming the exact defect.  |
| `01_CLEAN_DATA`   | Final analysis-ready output. Validated values pass through; flagged errors are suppressed.                                  |
| `QA_TEST_MATRIX`  | Self-testing harness. 52 adversarial inputs with expected outputs, each scored 🟢 PASS / 🔴 FAIL against the live formulas. |
| `TLD_REF`         | Reference table of valid IANA top-level domains used by the email validation rule.                                          |


**Why this matters:** the raw layer is preserved as evidence, the staging layer documents *why* each row was flagged, and the QA matrix means the cleaning logic is provably correct — not "looks right to me."

## Defect Coverage


| Field              | Rule                                       | Example flag                                     |
| ------------------ | ------------------------------------------ | ------------------------------------------------ |
| Transaction ID     | Format + length pattern                    | `ERROR: Invalid Format`, `ERROR: Invalid Length` |
| Customer Name      | Required, title-cased, whitespace-stripped | `ERROR: Missing First Name`                      |
| Customer Email     | Structure + TLD checked against IANA list  | `ERROR: Invalid TLD`                             |
| Region             | Allowed-value control list                 | `ERROR: Invalid Region`                          |
| Sales Channel      | Allowed-value control list                 | `ERROR: Invalid Channel`                         |
| Fulfillment Status | Allowed-value control list                 | `ERROR: Invalid Status`                          |
| Revenue            | Numeric, non-negative                      | `ERROR: Invalid Currency`                        |
| (Completeness)     | Missing critical fields                    | `ERROR: Missing Data`                            |


## Tech Stack

- Microsoft Excel (365 — dynamic arrays required)
- REGEX functions: `REGEXTEST`, `REGEXREPLACE`
- Spilling/array formulas: `LET`, `INDEX`, `SEQUENCE`, `ROWS`
- Text + safety: `SUBSTITUTE`, `TRIM`, `CLEAN`, `PROPER`, `LOWER`, `IFERROR`, `IF`, `COUNTIF`

## How to Use

1. Open `AlexVantage_Data_Remediation_Engine.xlsx`.
2. Paste a raw export into `03_RAW_DATA` (match the column order in row 1). Leave it untouched after that.
3. Open `02_STAGING_DATA` — every row is validated automatically. Scan for `ERROR:` codes to triage what needs upstream fixing.
4. Read `01_CLEAN_DATA` for the analysis-ready output.
5. Open `QA_TEST_MATRIX` to confirm the ruleset is sound (all rows should read 🟢 PASS).

## Screenshots

Raw data — defects visible
Staging layer — ERROR codes firing
QA matrix — 52/52 passing

## Files

- `AlexVantage_Data_Remediation_Engine.xlsx` — the engine
- `/sample_data/shopify_dump_dirty.csv` — example dirty input for testing
- `/screenshots` — visual walkthrough
- `case_study.md` — full problem → approach → impact writeup