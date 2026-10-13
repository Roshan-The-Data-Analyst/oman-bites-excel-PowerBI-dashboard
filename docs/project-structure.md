# Oman Bites Repository Structure

```text
oman-bites-excel-PowerBI-dashboard/
├── README.md
├── .gitignore
├── data/
│   ├── raw/Raw_Data.xlsx
│   └── processed/Oman_Bites_Cleaned_Data.xlsx
├── powerbi/Oman_Bites_Dashboard.pbix
├── docs/
│   ├── project-structure.md
│   └── guides/
│       ├── power-query-cleaning-notes-part-1.md
│       └── power-query-cleaning-notes-part-2.md
└── screenshots/
    ├── dashboard/
    ├── data-cleaning/
    ├── data-model/measures/
    └── insights/
```

## Folder responsibilities

- **data/raw/** — source workbook; keep it unchanged.
- **data/processed/** — cleaned Excel workbook created during preparation.
- **powerbi/** — Power BI dashboard.
- **docs/guides/** — original Power Query cleaning notes retained for reference.
- **screenshots/dashboard/** — dashboard views, filters, and mobile layout.
- **screenshots/data-cleaning/** — evidence of data-cleaning steps.
- **screenshots/data-model/measures/** — DAX/measure-building screenshots.
- **screenshots/insights/** — charts and evidence for branch and restaurant insights.

## Maintenance checklist

- Keep raw and processed files clearly separated.
- Update README image links whenever screenshots move.
- Reconcile row counts and totals across raw workbook, cleaned workbook, and Power BI before quoting definitive metrics.
- Document each measure's numerator, denominator, filter context, and business definition.
- Preserve source evidence until the cleaned output and dashboard have been validated.
- Avoid committing Office lock files and temporary editor artifacts.
