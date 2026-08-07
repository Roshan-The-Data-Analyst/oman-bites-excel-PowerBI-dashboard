3. For the `dd/mm/yyyy` vs `mm-dd-yyyy` ambiguity, use locale-aware parsing, e.g. `Date.FromText([Order Date], [Format="d/M/yyyy"])` on slash-formatted rows, with a separate logic branch for dash-formatted rows. (Power Query's regional settings under **File** → **Options** → **Regional Settings** also affect this — set to match your source, e.g. English (Oman) or English (United Kingdom).)
4. Once the custom column returns clean dates for every row, delete the original Order Date column and set the new column's type to **Date**.

> Tip: It's fine to do this in stages — fix the easy ISO (yyyy-mm-dd) and serial-number rows first, check the error count, then handle the remaining slash/dash rows.

## Issue 6 — Branch: inconsistent casing, stray spaces, and typos

**Why it matters:** The same six branches appear as "Muscat", "muscat", "MUSCAT", " Muscat", "Muscat ", and misspelled as "Muscatt"/"Musact".

**How to fix:**
1. Right-click Branch → **Transform** → **Trim**.
2. Right-click Branch → **Transform** → **Capitalize Each Word**.
3. **Transform** → **Replace Values** to fix remaining typos: "Muscatt"/"Musact" → "Muscat"; "Salalh" → "Salalah"; "Suhar" → "Sohar"; "Nizwaa" → "Nizwa"; "Duqum" → "Duqm".
4. Sanity check: filter dropdown on Branch should show exactly 6 distinct values.

## Issue 7 — Item: a few rows in ALL CAPS

**Why it matters:** Menu item names are mostly title case ("Fattoush Salad") but a small number are in ALL CAPS ("FATTOUSH SALAD"), fragmenting your item list.

**How to fix:**
1. Right-click Item → **Transform** → **Capitalize Each Word**, or **Lowercase** followed by **Capitalize Each Word** for a fully consistent result.

## Issue 8 — Category: casing and wording variants

**Why it matters:** "Rice & Biryani" also appears as "rice & biryani", "Rice and Biryani" and "RICE & BIRYANI".

**How to fix:**
1. Right-click Category → **Transform** → **Trim**, then **Transform** → **Capitalize Each Word**.
2. **Transform** → **Replace Values**: "Rice and Biryani"/"Rice And Biryani" → "Rice & Biryani"; "Salads and Starters" → "Salads & Starters".
3. Confirm the filter dropdown shows exactly 7 distinct values.

## Issue 9 — Quantity: negative values and text-formatted numbers

**Why it matters:** A small number of rows have a negative quantity (data-entry error) or a quantity stored as text with a trailing space (e.g. "3 "), blocking `SUM()` and aggregate calculations.

**How to fix:**
1. Set the column type to **Whole Number** (surfaces rows Power Query can't convert as errors).
2. Filter for errors using **Home** → **Keep Rows** → **Keep Errors**, fix the underlying text (Trim before re-typing as Whole Number).
3. For negative quantities, add a filter step (**Quantity** → **Number Filters** → "is greater than or equal to 0"), or use **Transform** → **Standard** → **Absolute Value** if negatives should be treated as sign errors rather than removed — decide based on how many rows are affected.

## Issue 10 — Unit Price: inconsistent currency formatting

**Why it matters:** Prices mix plain numbers ("1.500"), currency-prefixed text ("OMR 1.500"), inconsistent decimal places, and stray spaces. As text, this column can't be summed or charted.

**How to fix:**
1. Right-click Unit Price → **Transform** → **Trim**.
2. **Transform** → **Replace Values**: replace "OMR " with an empty string.
3. Change the column type to **Decimal Number**.
4. Optional: apply a display format later in Excel/Power BI (e.g. `0.000 "OMR"`) rather than storing the currency symbol in the data.

## Issue 11 — Payment Method & Order Type: casing and wording variants

**Why it matters:** Both columns mix casing ("Cash"/"cash"/"CASH") and synonyms ("Card" vs "Credit Card" vs "Debit Card"; "Mobile Wallet" vs "eWallet"; "Takeaway" vs "Take Away" vs "Take-away"; "Dine-in" vs "Dine In" vs "Dinein").

**How to fix:**
1. For each column: right-click → **Transform** → **Trim**, then **Transform** → **Capitalize Each Word**.
2. **Transform** → **Replace Values** to merge synonyms — Payment Method: "Credit Card"/"Debit Card" → "Card"; "Ewallet"/"E-wallet" → "Mobile Wallet". Order Type: "Take Away"/"Take-away"/"Takeaway " → "Takeaway"; "Dine In"/"Dinein" → "Dine-in".
3. Confirm each column's filter dropdown shows exactly 3 distinct values.

## Issue 12 — Customer Name & Phone: blanks and placeholder text

**Why it matters:** A small percentage of rows have blank names/phone numbers, or placeholder text such as "N/A", "NULL", "unknown", or "-".

**How to fix:**
1. **Transform** → **Replace Values** on both columns to replace "N/A", "NULL", "null", "n/a", "unknown" and "-" with null (leave the replace-with box empty, use Advanced Options → "Match entire cell contents").
2. Leave true blanks as null rather than deleting the row — Quantity and Unit Price are still valid.
3. Optional: add a `Customer Name Filled` column with **Transform** → **Fill Down** only if you need a placeholder like "Guest" — not required for this exercise.

## Issue 13 — Phone: four different formats

**Why it matters:** Phone numbers appear as `+968 XXXXXXXX`, `968XXXXXXXX`, `0XXXXXXXX`, and plain 8-digit numbers with no prefix, preventing reliable lookups.

**How to fix:**
1. Add a Custom Column that strips everything except digits: