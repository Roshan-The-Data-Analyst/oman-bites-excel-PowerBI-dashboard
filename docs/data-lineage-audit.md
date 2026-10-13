# Oman Bites — Data Quality and Lineage Audit

## Executive summary

Two workbooks were inspected:

- `data/raw/Raw_Data.xlsx`, worksheet `Orders_Raw`
- `data/processed/Oman_Bites_Cleaned_Data.xlsx`, including `Cleaned Table`, `Orders`, `Pover Pivot Q&A`, and `Power Pivot Measure`

The most important finding is **an unresolved source-to-cleaned lineage gap**. The cleaned table has 550 rows, but its Order IDs overlap only partially with the raw file. Treat this as a material caveat until the source version or filtering rule is explained.

## Raw workbook quality profile

| Check | Result |
|---|---:|
| Worksheet rows below header | 1,184 |
| Fully blank rows | 6 |
| Repeated-header rows | 3 |
| Candidate rows after blank/header removal | 1,175 |
| Unique candidate Order IDs | 1,150 |
| Repeated Order ID occurrences | 25 |
| Exact duplicate rows including repeated headers | 27 |
| Negative-quantity rows | 31 |
| Missing cells | 56 |
| Blank Customer Name values | 14 |
| Blank Phone values | 10 |
| Blank Rating (1–5) values | 32 |
| Branch label variants after case normalization | 12 |
| Date strings not matching the four identified patterns | 158 |

The raw date field mixes ISO strings, slash-formatted dates, dash-formatted dates, Excel serials, and an additional group requiring review. Do not parse ambiguous dates by guessing a locale; choose a format policy and flag ambiguous values.

The 12 raw branch label variants represent six intended branches mixed with casing, whitespace, and spelling variants. Category, payment-method, order-type, item, customer, and phone fields also need standardization as described in the retained Power Query guide.

## Cleaned workbook profile

| Check | Result |
|---|---:|
| Cleaned Table rows | 550 |
| Blank rows | 0 |
| Missing cells | 0 |
| Exact duplicate rows | 0 |
| Unique Order IDs | 550 |
| Rows in separate Orders staging sheet | 550 |
| Negative quantities in Orders staging sheet | 15 |
| Missing cells in Orders staging sheet | 90 |

The cleaned table has a cleaner shape than the source/staging sheet. A clean-looking final table alone, however, does not demonstrate that all source orders were retained.

## Critical lineage comparison

| Comparison | Result |
|---|---:|
| Candidate rows from raw workbook | 1,175 |
| Unique raw Order IDs | 1,150 |
| Cleaned table rows / unique IDs | 550 |
| Order IDs shared by raw and cleaned table | 519 |
| Unique raw IDs absent from cleaned table | 631 |
| Cleaned IDs absent from raw candidate data | 31 |

The cleaned table and the separate `Orders` sheet have the same 550 Order IDs. This suggests the cleaned table is a cleaned version of that staging sheet, but that staging sheet itself does not align one-for-one with the current raw workbook.

### Required resolution

1. Determine whether `Orders` is intended to be a separate 550-order sample or a copy of the raw workbook after documented filters.
2. For every source order ID, document whether it was retained, deduplicated, rejected, or intentionally excluded.
3. Investigate the 31 cleaned-only IDs and the 631 raw-only unique IDs.
4. Record a before/after row-count reconciliation in the Power Query guide and project README.
5. Do not state that all raw data issues were fixed until the lineage gap is resolved.

## Existing Power Pivot summary

The cleaned workbook's summary reports:

| Measure | Workbook value |
|---|---:|
| Total Orders | 550 |
| Total Revenue | OMR 3,346.403 |
| Average Order Value | OMR 6.084369 |
| Average Rating | 3.672727 |

Branch revenue in `Pover Pivot Q&A`:

| Branch | Revenue (OMR) |
|---|---:|
| Sohar | 620.380 |
| Muscat | 601.003 |
| Salalah | 556.423 |
| Nizwa | 547.257 |
| Sur | 526.228 |
| Duqm | 495.112 |
| **Total** | **3,346.403** |

The row-level audit script was executed with GitHub Actions. It calculated `SUM(Quantity × Unit Price)` across the 550 cleaned records as **OMR 3,346.403**, matching the cached Power Pivot total (difference **OMR 0.000** at 3-decimal reporting precision). The six branch calculations also match their cached pivot values at OMR 0.001 precision:

| Branch | Recomputed revenue (OMR) | Stored branch pivot (OMR) | Difference |
|---|---:|---:|---:|
| Sohar | 620.380 | 620.380 | 0.000 |
| Muscat | 601.003 | 601.003 | 0.000 |
| Salalah | 556.423 | 556.423 | 0.000 |
| Nizwa | 547.257 | 547.257 | 0.000 |
| Sur | 526.228 | 526.228 | 0.000 |
| Duqm | 495.112 | 495.112 | 0.000 |
| **Total** | **3,346.403** | **3,346.403** | **0.000** |

This verifies the cached dashboard sales calculations against the current cleaned sheet. It does **not** resolve the 631 raw-only IDs and 31 cleaned-only IDs; the current cleaned file may be a selected sample or may derive from a different source version.

## Audit method

Run `python scripts/audit_oman_bites.py` to regenerate:

- `reports/data-lineage-audit.json`
- `reports/data-quality-summary.csv`
- `reports/branch-revenue-reconciliation.csv`

The script uses Python's standard library to read XLSX XML directly, count missing values and duplicates, compare Order IDs, calculate row-level revenue by branch, and compare those figures with the stored pivot and KPI summary. It does not edit the source workbooks.
