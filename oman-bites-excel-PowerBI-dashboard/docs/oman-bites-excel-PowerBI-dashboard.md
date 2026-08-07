# Oman Bites Restaurant Chain — Messy Dataset: Power Query Cleaning Guide

This guide accompanies `Oman_Bites_Messy_Dataset.xlsx` (sheet: `Orders_Raw`, 1,184 order rows across
six branches: Muscat, Salalah, Sohar, Nizwa, Sur and Duqm). The file was deliberately "dirtied" for
practicing real-world cleaning with Power Query in Excel, then loading the cleaned table into Power BI
for analysis.

Every issue below is intentional. Work through them in order — each section names the exact column(s)
affected, why it matters, and the specific Power Query steps to fix it.

## Column Overview

| Column | What it should be | Messy in this file? |
|---|---|---|
| Order ID | Text, format OB-##### | Yes — some rows are plain numbers |
| Order Date | Date | Yes — 5 different formats mixed |
| Branch | One of 6 branch names | Yes — casing, spaces, typos |
| Item | Menu item name | Yes — a few rows in ALL CAPS |
| Category | One of 7 categories | Yes — casing + wording variants |
| Quantity | Whole number | Yes — a few negatives / text values |
| Unit Price | Decimal number (OMR) | Yes — currency prefix, spacing, rounding |
| Payment Method | Cash / Card / Mobile Wallet | Yes — casing + synonyms |
| Order Type | Dine-in / Takeaway / Delivery | Yes — casing + spelling variants |
| Customer Name | Text | Yes — some blank / placeholder values |
| Phone | Omani mobile number | Yes — 4 different formats |
| Rating (1–5) | Whole number 1–5 | Yes — many blanks (genuinely optional) |

## Step 0 — Load the file into Power Query

1. Open Excel → **Data** tab → **Get Data** → **From File** → **From Excel Workbook** → select `Oman_Bites_Messy_Dataset.xlsx`.
2. In the Navigator window, select the `Orders_Raw` sheet and click **Transform Data** (not Load) — this opens the Power Query Editor so every step is recorded.
3. Keep the Power Query Editor open for all the steps below. Each fix becomes a new "Applied Step" you can see, reorder, or undo on the right-hand panel.

## Issue 1 — Messy header row (stray spaces & inconsistent casing)

**Why it matters:** The header row has issues: " Order Date" and "Unit Price " have leading/trailing spaces, and "payment method" is lowercase. This carries through your whole model.

**How to fix:**
1. Confirm the **Promoted Headers** step exists in Applied Steps (Power Query does this automatically on load).
2. Double-click each affected column header (Order Date, Unit Price, Payment Method) and retype the clean name, or right-click → **Rename**.
3. Alternative for many columns at once: **Home** → **Advanced Editor** and edit the `#"Renamed Columns"` step directly.

## Issue 2 — Duplicate rows (~25 exact duplicates)

**Why it matters:** Some orders appear twice with identical values in every column, inflating revenue, item counts and branch totals.

**How to fix:**
1. Select all columns (click the top-left corner of the table, or Ctrl+A on the column headers).
2. **Home** → **Remove Rows** → **Remove Duplicates**.
3. Check the row count before/after in the status bar to confirm.

> Tip: Do this AFTER fixing text casing/spacing issues below, or near-duplicates (e.g. "Muscat" vs " Muscat") won't be caught as exact matches yet.

## Issue 3 — Blank rows and rows that repeat the header

**Why it matters:** A handful of fully blank rows, and a few rows where every cell just contains the column name, are mixed into the data.

**How to fix:**
1. **Home** → **Remove Rows** → **Remove Blank Rows** to drop fully empty rows.
2. For rows that repeat the header text (e.g. a row where Branch = "Branch"), use the filter arrow on the Order ID column and filter out rows where Branch equals "Branch".
3. Alternative: filter the Order ID column to keep only rows where the value does not equal the text "Order ID" (**Text Filters** → **Does Not Equal**).

## Issue 4 — Order ID: mixed text and number formats

**Why it matters:** Most Order IDs look like `OB-10041`, but some rows only contain a plain number (e.g. `10041`) with no prefix.

**How to fix:**
1. Click the data type icon on the Order ID column header and set it to **Text** for all rows first.
2. Add a custom column (**Add Column** → **Custom Column**):

