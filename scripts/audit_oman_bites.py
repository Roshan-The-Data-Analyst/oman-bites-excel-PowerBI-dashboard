#!/usr/bin/env python3
"""Audit Oman Bites XLSX row counts, lineage, and revenue reconciliation (stdlib only).

Run from repository root:
  python scripts/audit_oman_bites.py
Optional paths:
  python scripts/audit_oman_bites.py --raw data/raw/Raw_Data.xlsx --processed data/processed/Oman_Bites_Cleaned_Data.xlsx --output-dir reports
This script reads the source workbooks; it does not modify them.
"""
from __future__ import annotations
import argparse
import csv
import json
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

D = Decimal
ZERO = D("0")
CENT = D("0.01")
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
DOC_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def col_index(cell_ref: str) -> int:
    result = 0
    for ch in cell_ref:
        if not ch.isalpha():
            break
        result = result * 26 + ord(ch.upper()) - ord("A") + 1
    return result - 1


def parse_cell(cell: ET.Element, shared: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(MAIN + "t"))
    value = cell.find(MAIN + "v")
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        return shared[int(value.text)]
    return value.text


def read_sheet(path: Path, wanted_sheet: str) -> tuple[list[str], list[dict[str, str]]]:
    with zipfile.ZipFile(path) as archive:
        wb = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rels = {r.attrib["Id"]: r.attrib["Target"] for r in rels_xml.findall(PKG_REL + "Relationship")}
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            ss_xml = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared = ["".join(t.text or "" for t in si.iter(MAIN + "t")) for si in ss_xml.findall(MAIN + "si")]

        sheet_target = None
        sheets = wb.find(MAIN + "sheets")
        if sheets is None:
            raise ValueError(f"No sheets in {path}")
        for sheet in sheets.findall(MAIN + "sheet"):
            if sheet.attrib.get("name") == wanted_sheet:
                sheet_target = rels.get(sheet.attrib.get(DOC_REL + "id", ""))
                break
        if not sheet_target:
            raise ValueError(f"Worksheet {wanted_sheet!r} not found in {path}")
        target = sheet_target.lstrip("/")
        sheet_path = target if target.startswith("xl/") else posixpath.normpath(posixpath.join("xl", target))
        root = ET.fromstring(archive.read(sheet_path))

        matrix: list[list[str]] = []
        for row in root.iter(MAIN + "row"):
            cells: list[str] = []
            for cell in row.findall(MAIN + "c"):
                index = col_index(cell.attrib.get("r", ""))
                while len(cells) <= index:
                    cells.append("")
                cells[index] = parse_cell(cell, shared)
            matrix.append(cells)

    if not matrix:
        raise ValueError(f"Worksheet {wanted_sheet!r} is empty")
    headers = [str(v or "").strip() for v in matrix[0]]
    records = []
    for raw_row in matrix[1:]:
        values = raw_row + [""] * max(0, len(headers) - len(raw_row))
        row = {headers[i]: str(values[i] or "").strip() for i in range(len(headers))}
        records.append(row)
    return headers, records


def nonblank(row: dict[str, str], headers: list[str]) -> bool:
    return any((row.get(key) or "").strip() for key in headers)


def decimal_value(value: str) -> Decimal:
    cleaned = (value or "").strip().replace("OMR", "").replace(",", "").strip()
    if not cleaned:
        raise ValueError("blank numeric value")
    return D(cleaned)


def money(value: Decimal) -> Decimal:
    return value.quantize(CENT, rounding=ROUND_HALF_UP)


def fmt(value: Decimal) -> str:
    return f"{value.quantize(CENT, rounding=ROUND_HALF_UP):.2f}"


def fmt3(value: Decimal) -> str:
    return f"{value.quantize(D('0.001'), rounding=ROUND_HALF_UP):.3f}"


def write_csv(path: Path, fields: list[str], rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description="Audit Oman Bites workbooks and revenue lineage.")
    parser.add_argument("--raw", type=Path, default=root / "data" / "raw" / "Raw_Data.xlsx")
    parser.add_argument("--processed", type=Path, default=root / "data" / "processed" / "Oman_Bites_Cleaned_Data.xlsx")
    parser.add_argument("--output-dir", type=Path, default=root / "reports")
    args = parser.parse_args()

    raw_headers, raw_records_all = read_sheet(args.raw, "Orders_Raw")
    clean_headers, clean_records_all = read_sheet(args.processed, "Cleaned Table")
    orders_headers, orders_records_all = read_sheet(args.processed, "Orders")
    pivot_headers, pivot_records = read_sheet(args.processed, "Pover Pivot Q&A")
    measure_headers, measure_records = read_sheet(args.processed, "Power Pivot Measure")

    raw_nonblank = [r for r in raw_records_all if nonblank(r, raw_headers)]
    raw_blank_rows = len(raw_records_all) - len(raw_nonblank)
    repeated_headers = [r for r in raw_nonblank if (r.get("Order ID", "").casefold() == "order id" or r.get("Branch", "").casefold() == "branch")]
    raw_candidates = [r for r in raw_nonblank if r not in repeated_headers]
    clean_rows = [r for r in clean_records_all if nonblank(r, clean_headers)]
    orders_rows = [r for r in orders_records_all if nonblank(r, orders_headers)]

    raw_signatures = Counter(tuple(r.get(h, "") for h in raw_headers) for r in raw_candidates)
    clean_signatures = Counter(tuple(r.get(h, "") for h in clean_headers) for r in clean_rows)
    raw_ids = [r.get("Order ID", "").strip() for r in raw_candidates if r.get("Order ID", "").strip()]
    clean_ids = [r.get("Order ID", "").strip() for r in clean_rows if r.get("Order ID", "").strip()]
    orders_ids = [r.get("Order ID", "").strip() for r in orders_rows if r.get("Order ID", "").strip()]
    raw_id_set, clean_id_set, orders_id_set = set(raw_ids), set(clean_ids), set(orders_ids)

    raw_missing = {h: sum(not r.get(h, "").strip() for r in raw_nonblank) for h in raw_headers}
    clean_missing = {h: sum(not r.get(h, "").strip() for r in clean_rows) for h in clean_headers}
    orders_missing = {h: sum(not r.get(h, "").strip() for r in orders_rows) for h in orders_headers}
    raw_negative_qty = sum(1 for r in raw_candidates if safe_decimal(r.get("Quantity", "")) is not None and safe_decimal(r.get("Quantity", "")) < ZERO)
    orders_negative_qty = sum(1 for r in orders_rows if safe_decimal(r.get("Quantity", "")) is not None and safe_decimal(r.get("Quantity", "")) < ZERO)
    clean_negative_qty = sum(1 for r in clean_rows if safe_decimal(r.get("Quantity", "")) is not None and safe_decimal(r.get("Quantity", "")) < ZERO)

    date_formats = Counter()
    for r in raw_candidates:
        value = r.get("Order Date", "").strip()
        if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", value):
            date_formats["ISO-like"] += 1
        elif re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", value):
            date_formats["slash-formatted"] += 1
        elif re.fullmatch(r"\d{1,2}-\d{1,2}-\d{4}", value):
            date_formats["dash-formatted"] += 1
        elif re.fullmatch(r"\d{5}", value):
            date_formats["Excel-serial-like"] += 1
        else:
            date_formats["other-needs-review"] += 1

    branch_revenue: dict[str, Decimal] = defaultdict(lambda: ZERO)
    total_revenue = ZERO
    total_quantity = ZERO
    ratings: list[Decimal] = []
    row_revenue_errors = 0
    numeric_errors = Counter()
    for row in clean_rows:
        try:
            quantity = decimal_value(row.get("Quantity", ""))
            unit_price = decimal_value(row.get("Unit Price", ""))
            if quantity < ZERO or quantity != quantity.to_integral_value():
                numeric_errors["invalid_quantity"] += 1
            if unit_price <= ZERO:
                numeric_errors["nonpositive_unit_price"] += 1
            revenue = quantity * unit_price
            total_revenue += revenue
            total_quantity += quantity
            branch_revenue[row.get("Branch", "").strip()] += revenue
        except Exception:
            row_revenue_errors += 1
        try:
            rating = decimal_value(row.get("Rating (1-5)", ""))
            if D("1") <= rating <= D("5"):
                ratings.append(rating)
            else:
                numeric_errors["rating_out_of_range"] += 1
        except Exception:
            numeric_errors["blank_or_invalid_rating"] += 1

    stored_branch: dict[str, Decimal] = {}
    valid_branches = {
        "muscat": "Muscat", "salalah": "Salalah", "sohar": "Sohar",
        "nizwa": "Nizwa", "sur": "Sur", "duqm": "Duqm",
    }
    for r in pivot_records:
        label = (r.get("Row Labels") or "").strip()
        value = (r.get("Total Revenue") or "").strip()
        canonical = valid_branches.get(label.casefold())
        # The worksheet contains multiple pivot tables side by side; only keep branch labels.
        if canonical and value:
            try:
                stored_branch[canonical] = D(value)
            except Exception:
                continue

    branch_rows = []
    for branch_name in sorted(set(branch_revenue) | set(stored_branch)):
        computed = branch_revenue.get(branch_name, ZERO)
        stored = stored_branch.get(branch_name)
        delta = computed - stored if stored is not None else None
        branch_rows.append({
            "branch": branch_name,
            "row_level_revenue_omr": fmt3(computed),
            "stored_pivot_revenue_omr": fmt3(stored) if stored is not None else "",
            "difference_omr": fmt3(delta) if delta is not None else "",
            "status": "MATCH_AT_0.001_OMR" if delta is not None and abs(delta) < D("0.001") else ("CHECK" if stored is not None else "NO_STORED_PIVOT"),
        })

    measures = {}
    if measure_records:
        row = measure_records[0]
        for header in measure_headers:
            if row.get(header, ""):
                measures[header] = row[header]
    stored_total = D(measures["Total Revenue"]) if measures.get("Total Revenue") else None
    stored_orders = D(measures["Total Orders"]) if measures.get("Total Orders") else None
    stored_aov = D(measures["Average Order Value"]) if measures.get("Average Order Value") else None
    stored_rating = D(measures["Average Rating"]) if measures.get("Average Rating") else None
    unique_raw = len(raw_id_set)
    unique_clean = len(clean_id_set)
    lineage = {
        "raw_candidate_rows": len(raw_candidates),
        "raw_unique_order_ids": unique_raw,
        "cleaned_rows": len(clean_rows),
        "cleaned_unique_order_ids": unique_clean,
        "shared_order_ids": len(raw_id_set & clean_id_set),
        "raw_ids_missing_from_cleaned": len(raw_id_set - clean_id_set),
        "cleaned_ids_absent_from_raw": len(clean_id_set - raw_id_set),
        "cleaned_and_orders_sheet_id_sets_match": clean_id_set == orders_id_set,
    }
    audit = {
        "raw_workbook": str(args.raw),
        "processed_workbook": str(args.processed),
        "raw_worksheet": "Orders_Raw",
        "cleaned_worksheet": "Cleaned Table",
        "raw": {
            "rows_below_header": len(raw_records_all),
            "blank_rows": raw_blank_rows,
            "repeated_header_rows": len(repeated_headers),
            "candidate_rows_after_blank_and_header_removal": len(raw_candidates),
            "exact_duplicate_rows_after_header_removal": sum(n - 1 for n in raw_signatures.values() if n > 1),
            "duplicate_order_id_occurrences": len(raw_ids) - unique_raw,
            "missing_values_by_column_in_nonblank_rows": raw_missing,
            "total_missing_cells_in_nonblank_rows": sum(raw_missing.values()),
            "negative_quantity_rows": raw_negative_qty,
            "branch_label_variants_casefolded": len({r.get("Branch", "").strip().casefold() for r in raw_candidates}),
            "date_format_distribution": dict(date_formats),
        },
        "cleaned": {
            "rows": len(clean_rows),
            "exact_duplicate_rows": sum(n - 1 for n in clean_signatures.values() if n > 1),
            "missing_values_by_column": clean_missing,
            "total_missing_cells": sum(clean_missing.values()),
            "negative_quantity_rows": clean_negative_qty,
            "invalid_numeric_rows": dict(numeric_errors),
            "quantity_total": fmt(total_quantity),
            "row_level_revenue_omr": fmt3(total_revenue),
            "average_rating": str((sum(ratings, ZERO) / D(len(ratings))).quantize(D("0.000001"), rounding=ROUND_HALF_UP)) if ratings else None,
            "rows_with_revenue_calculation_errors": row_revenue_errors,
        },
        "orders_staging_sheet": {
            "rows": len(orders_rows),
            "missing_values_by_column": orders_missing,
            "total_missing_cells": sum(orders_missing.values()),
            "negative_quantity_rows": orders_negative_qty,
        },
        "lineage": lineage,
        "stored_power_pivot_measures": {
            "total_orders": str(stored_orders) if stored_orders is not None else None,
            "total_revenue_omr": str(stored_total) if stored_total is not None else None,
            "average_order_value_omr": str(stored_aov) if stored_aov is not None else None,
            "average_rating": str(stored_rating) if stored_rating is not None else None,
            "row_level_revenue_difference_omr": fmt3(total_revenue - stored_total) if stored_total is not None else None,
            "row_level_orders_match": int(stored_orders) == len(clean_rows) if stored_orders is not None else None,
        },
        "branch_revenue_reconciliation": branch_rows,
        "limitations": [
            "A matching pivot total does not resolve the order-ID lineage gap.",
            "Confirm which order records are intentionally excluded before claiming the cleaned table represents all raw data.",
            "Revenue is calculated here as Quantity × Unit Price; validate the intended business definition and discount treatment.",
            "The workbook describes an Oman restaurant scenario and uses OMR in pricing and example visuals; currency labels should remain explicit.",
        ],
    }

    args.output_dir.mkdir(parents=True, exist_ok=True)
    with (args.output_dir / "data-lineage-audit.json").open("w", encoding="utf-8") as f:
        json.dump(audit, f, indent=2, ensure_ascii=False)
        f.write("\n")
    write_csv(args.output_dir / "branch-revenue-reconciliation.csv",
        ["branch", "row_level_revenue_omr", "stored_pivot_revenue_omr", "difference_omr", "status"], branch_rows)
    quality_rows = [
        {"area": "raw", "check": "rows_below_header", "value": len(raw_records_all), "interpretation": "Includes blank rows and repeated headers."},
        {"area": "raw", "check": "blank_rows", "value": raw_blank_rows, "interpretation": "Fully blank worksheet rows."},
        {"area": "raw", "check": "repeated_header_rows", "value": len(repeated_headers), "interpretation": "Rows that repeat the headers."},
        {"area": "raw", "check": "candidate_transaction_rows", "value": len(raw_candidates), "interpretation": "After blank/header removal, before deduplication."},
        {"area": "raw", "check": "unique_order_ids", "value": unique_raw, "interpretation": "Unique candidate Order IDs."},
        {"area": "raw", "check": "negative_quantity_rows", "value": raw_negative_qty, "interpretation": "Requires business-rule review."},
        {"area": "raw", "check": "missing_cells", "value": sum(raw_missing.values()), "interpretation": "Counted over nonblank source rows."},
        {"area": "cleaned", "check": "rows", "value": len(clean_rows), "interpretation": "Rows in Cleaned Table."},
        {"area": "cleaned", "check": "missing_cells", "value": sum(clean_missing.values()), "interpretation": "Counted over cleaned records."},
        {"area": "cleaned", "check": "negative_quantity_rows", "value": clean_negative_qty, "interpretation": "Must be zero unless negative orders are legitimate."},
        {"area": "lineage", "check": "shared_order_ids", "value": len(raw_id_set & clean_id_set), "interpretation": "Raw and cleaned intersection."},
        {"area": "lineage", "check": "raw_ids_missing_from_cleaned", "value": len(raw_id_set - clean_id_set), "interpretation": "Explain each exclusion or filter."},
        {"area": "lineage", "check": "cleaned_ids_absent_from_raw", "value": len(clean_id_set - raw_id_set), "interpretation": "Check source versions or unmatched IDs."},
    ]
    write_csv(args.output_dir / "data-quality-summary.csv", ["area", "check", "value", "interpretation"], quality_rows)
    print(f"Raw candidate rows: {len(raw_candidates):,}; cleaned rows: {len(clean_rows):,}")
    print(f"Shared order IDs: {len(raw_id_set & clean_id_set):,}; raw-only: {len(raw_id_set-clean_id_set):,}; cleaned-only: {len(clean_id_set-raw_id_set):,}")
    print(f"Cleaned row-level revenue (OMR): {fmt3(total_revenue)}; stored KPI total: {fmt3(stored_total) if stored_total is not None else 'not found'}")
    print(f"Reports written to {args.output_dir.resolve()}")


def safe_decimal(value: str) -> Decimal | None:
    try:
        return decimal_value(value)
    except Exception:
        return None


if __name__ == "__main__":
    main()
