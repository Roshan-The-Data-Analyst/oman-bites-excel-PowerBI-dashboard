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
│   ├── data-lineage-audit.md
│   └── guides/
│       ├── power-query-cleaning-notes-part-1.md
│       └── power-query-cleaning-notes-part-2.md
├── scripts/audit_oman_bites.py
├── reports/
│   ├── data-lineage-audit.json
│   ├── data-quality-summary.csv
│   └── branch-revenue-reconciliation.csv
└── screenshots/
    ├── dashboard/
    ├── data-cleaning/
    ├── data-model/measures/
    └── insights/
```

## Data discipline

- Keep `data/raw/` immutable.
- Save cleaned output under `data/processed/`.
- Record every row-exclusion rule and its row count.
- Reconcile source Order IDs before claiming lineage.
- Check measure definitions and reconcile row-level values to Power Pivot output before presenting KPI figures as verified.
