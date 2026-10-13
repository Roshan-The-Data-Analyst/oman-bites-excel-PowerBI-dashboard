# Oman Bites — Excel & Power BI Sales Analytics Dashboard

A portfolio project demonstrating an end-to-end BI workflow for a simulated restaurant business in Oman: clean operational data, create measures, analyse branch and order performance, and present results through a Power BI dashboard.

![Oman Bites dashboard overview](screenshots/dashboard/dashboard-overview.png)

## Project overview

- **Business scenario:** Multi-branch restaurant operations in Oman
- **Tools:** Microsoft Excel, Power Query, Power Pivot, DAX, and Power BI
- **Data assets:** Raw source workbook and cleaned workbook
- **Analysis areas:** Revenue, branch comparisons, order type, payment method, product/category performance, ratings, quantity, and interactive slicers
- **Power BI report:** `powerbi/Oman_Bites_Dashboard.pbix`

## Repository structure

| Folder | Purpose |
|---|---|
| `data/raw/` | Original source workbook |
| `data/processed/` | Cleaned Excel workbook |
| `powerbi/` | Main Power BI report |
| `docs/guides/` | Retained Power Query cleaning notes |
| `screenshots/dashboard/` | Dashboard overview, detail, filters, and mobile view |
| `screenshots/data-cleaning/` | Evidence of raw data and data preparation |
| `screenshots/data-model/measures/` | DAX and measure-building screenshots |
| `screenshots/insights/` | Branch and restaurant operations insights |

## Workflow

1. **Inspect the source data:** review field names, data types, duplicates, missing values, and inconsistent values.
2. **Clean with Power Query:** standardize branches, categories, dates, quantities, unit prices, payment methods, and order types.
3. **Model and calculate:** prepare a consistent model and measures for sales, quantity, ratings, and branch/product comparisons.
4. **Build the dashboard:** present KPIs and filters that allow users to explore branch and order performance.
5. **Validate before publication:** reconcile raw/cleaned row counts and compare Excel totals with Power BI measures.

## Dashboard preview

![Dashboard detail](screenshots/dashboard/dashboard-detail.png)

Explore [slicers and filters](screenshots/dashboard/slicers-and-filters.png), [revenue by branch](screenshots/insights/revenue-by-branch.png), [orders by branch](screenshots/insights/orders-by-branch.png), and [highest-rated branch](screenshots/insights/highest-rated-branch.png).

## Data notes and limitations

The repository contains both raw and cleaned workbooks. Treat the raw workbook as source evidence and the processed workbook as a transformation output. The earlier README lists 3,045 rows, while the cleaning notes describe a 1,184-row practice sheet; reconcile those counts against the actual workbooks before quoting a definitive record count.

The project scenario is simulated and should be presented as a portfolio case study, not actual company results. Existing assets are preserved by this organization pass; dashboard calculations have not been independently re-audited here.

## Getting started

1. Clone or download the repository.
2. Open `powerbi/Oman_Bites_Dashboard.pbix` in Power BI Desktop.
3. Open the raw and processed workbooks to review source data and transformations.
4. Use [the project structure guide](docs/project-structure.md) to navigate the repository.
5. Refer to the [Power Query cleaning notes](docs/guides/) for transformation context.

## Skills demonstrated

Excel, Power Query, Power Pivot, DAX, Power BI, data cleaning, data modeling, KPI design, interactive dashboarding, and business reporting.
