# Order-Level Sales Data Cleaning and Business Review

## Problem Summary

A retail company exported order-level sales data from multiple internal systems. The raw extract contained inconsistent text values, duplicate records, missing values, date-quality issues, invalid discounts, calculation mismatches, and conflicting order/payment statuses.

This project converts that extract into a clean, validated, analysis-ready Excel dataset and delivers business-review summaries. The original source workbook was preserved and was not overwritten.

## Dataset Description

| Item | Description |
|---|---|
| Source file | `raw_orders.xlsx` |
| Raw records | 932 order-level rows |
| Cleaned file | `cleaned_orders.xlsx` |
| Retained after exact de-duplication | 912 rows |
| Primary business grain | One order-level source record; conflicting repeated order IDs are retained for audit rather than silently collapsed |
| Reporting period | January 2024 through October 2025, based on cleaned order dates |

Key data areas include customer and geography attributes, product category/sub-category, order and ship dates, ship mode, quantities, unit price, cost, discounts, raw sales/profit, order status, and payment status.

## Deliverables

| File | Purpose |
|---|---|
| `raw_orders.xlsx` | Preserved source dataset; not modified |
| `cleaned_orders.xlsx` | Cleaned, standardized, formula-enabled order dataset with record-level quality flags |
| `outputs/data_quality_report.xlsx` | Evaluator-friendly summary of missing values, duplicates, date issues, discount issues, status issues, and calculation mismatches |
| `outputs/pivot_summary.xlsx` | Business review summaries for sales, profit, customer segments, shipping, exceptions, and monthly trend |
| `outputs/cleaning_log.md` | Detailed audit log of all cleaning and rule decisions |

## Tools Used

- Excel-compatible `.xlsx` workbooks with tables, formulas, filters, conditional formatting, and charts.
- Excel-style cleaning techniques: `TRIM`, `SUBSTITUTE`, controlled Find/Replace mappings, standardized date formats, and calculated columns.
- Formula-driven validation for calculated sales, profit, profit margin, shipping delay, and quality classification.
- Programmatic workbook generation and verification using `artifact_tool` for repeatable reporting output.
- Markdown documentation for the cleaning log and this README.

## Cleaning Steps Performed

1. **Preserved the source file** and created a separate `cleaned_orders.xlsx` output.
2. **Standardized text fields**: customer name, segment, region, state, city, category, sub-category, ship mode, payment status, and order status.
   - Removed leading/trailing and repeated spaces.
   - Standardized capitalization/case.
   - Normalized known category and sub-category variants.
   - Avoided destructive character removal from names and product text.
3. **Cleaned dates**:
   - Converted `order_date` and `ship_date` to real Excel dates displayed as `yyyy-mm-dd`.
   - Created `shipping_delay_days` as `ship_date - order_date`.
   - Flagged invalid date sequences rather than changing source evidence.
4. **Handled duplicates**:
   - Removed only later occurrences of exact full-row duplicates.
   - Retained repeated `order_id` records with conflicting information and flagged them for reconciliation.
5. **Applied business-rule validations** for missing region/ship mode, discount validity, order/payment status, date sequence, and calculation variance.
6. **Added calculated and audit columns** including `cleaned_discount`, `calculated_sales`, `calculated_profit`, `profit_margin`, `order_month`, `order_year`, `data_quality_flag`, sales/profit variance, and quality-detail fields.
7. **Built business reports** for data quality and pivot-style sales/profit review.

## Business Rules Applied

| Rule area | Decision |
|---|---|
| Missing `region` | Replaced with `Unknown` and retained as a warning/audit adjustment. |
| Missing `ship_mode` | Replaced with `Unknown` and retained as a warning/audit adjustment. |
| Missing `discount` | Set to `0` only when other sales fields were present and internally valid. |
| Negative discount | Treated as invalid; retained for audit and excluded from eligible completed sales. |
| Discount above allowed range | Treated as invalid above the adopted 50% cap; retained for audit and excluded from eligible completed sales. |
| Cancelled orders | Retained for audit but excluded from the final completed-sales summary. |
| Failed payments | Retained for audit but excluded from the final completed-sales summary. |
| Refunded orders | Retained and summarized separately using recorded raw sales. |
| Ship date before order date | Treated as an invalid shipping record and excluded from completed-sales eligibility. |
| Completed-sales eligibility | Requires `Completed` + `Paid`, valid discount/date sequence, no calculation mismatch, and no unresolved duplicate conflict or review-level issue. |

## Summary of Data Quality Issues Found

| Issue | Count | Treatment |
|---|---:|---|
| Exact duplicate rows | 20 | Removed; only exact full-row duplicates were removed. |
| Duplicate order IDs | 31 | Identified for audit. |
| Conflicting duplicate order IDs | 12 IDs / 24 retained records | Retained and flagged for manual reconciliation. |
| Missing region | 25 retained rows | Filled as `Unknown`; warning retained. |
| Missing ship mode | 21 retained rows | Filled as `Unknown`; warning retained. |
| Missing discount | 18 retained rows | Imputed to `0` after sales-field validation. |
| Invalid discount | 30 retained rows | 15 negative; 15 above 50%; marked invalid. |
| Ship date before order date | 21 retained rows | Marked invalid shipping records. |
| Sales/profit mismatch | 53 unique retained rows | Flagged where variance exceeded 0.01. |
| Completed but unpaid | 2 retained rows | Marked invalid and excluded from completed-sales eligibility. |
| Cancelled orders | 145 retained rows | Excluded from completed-sales summary. |
| Failed-payment orders | 69 retained rows | Excluded from completed-sales summary. |
| Refunded orders | 71 retained rows | Reported separately. |

### Final Quality Classification

| Classification | Records | Share of retained rows |
|---|---:|---:|
| Clean | 765 | 83.88% |
| Warning | 45 | 4.93% |
| Invalid | 102 | 11.18% |
| **Flagged total** | **147** | **16.12%** |
| **Retained total** | **912** | **100.00%** |

> Individual issue counts overlap. Use the mutually exclusive `Clean`, `Warning`, and `Invalid` counts for the final quality population.

## Summary of Final Pivot Reports

`outputs/pivot_summary.xlsx` contains the following review-ready worksheets:

| Pivot-style summary | Scope and purpose |
|---|---|
| Sales & Profit by Region | Eligible completed-sales records only; sales, profit, margin, and sales share by region. Filtered and sorted from highest to lowest calculated sales. |
| Sales & Profit by Category and Sub-Category | Eligible completed-sales records only; category/sub-category performance. Filtered and sorted within category by calculated sales. |
| Order Count by Ship Mode | All 912 retained records for operational volume, with eligible-sales context. Sorted by record count. |
| Profit Margin by Customer Segment | Eligible completed-sales records only; segment sales, profit, margin, and share. Sorted from highest to lowest margin. |
| Refunded/Cancelled/Failed Orders by Region | All retained records; exception counts, unique impacted IDs, and sales exposure by region. |
| Monthly Sales Trend | Eligible completed-sales records only; month, sales, profit, margin, and month-on-month sales growth. |
| Dashboard | Headline KPIs, sales-by-region chart, and monthly sales trend chart. |

## Key Business Insights

1. **Validated completed business:** 537 records qualify for completed-sales reporting, producing **$5,329,453.93** in calculated sales and **$1,549,060.04** in calculated profit at a **29.1%** overall margin.
2. **Regional concentration:** South leads eligible sales at **$1,427,391.84** (26.8% share), followed by West at **$1,360,140.47** (25.5%). South and West together contribute just over half of validated sales.
3. **Category opportunity:** Copiers are the largest sub-category by eligible sales at **$636,659.33** (11.9% share). Machines have the highest listed sub-category margin at **34.1%**, making them a strong profitability focus.
4. **Customer segment view:** Home Office has the highest segment profit margin at **30.2%**, while Small Business generates the largest sales share at **28.2%**.
5. **Shipping operations:** Standard Class has the greatest retained order volume (242 records) and the highest eligible sales among ship modes at **$1,503,391.36**.
6. **Trend signal:** Eligible monthly sales peak in **February 2025** at **$387,223.35**. October 2025 is materially lower at **$121,434.81** and should be interpreted carefully because it may represent an incomplete reporting month.
7. **Exception exposure:** North has the largest exception population, with **61 unique impacted order IDs** across refunded/cancelled/failed conditions and **$640,361.88** of raw sales exposure. It should be prioritized for root-cause review.
8. **Data quality priority:** Calculation mismatches (53 unique records), invalid discounts (30), and date-sequence failures (21) are the most important controls to correct at the source-system level.

## Assumptions and Limitations

### Assumptions

- The allowed discount range is **0% to 50%**, inclusive. The source requirements did not define a maximum discount, so 50% was adopted as a documented operational assumption.
- Slash-formatted numeric dates are interpreted as `MM/DD/YYYY`; hyphen-formatted numeric dates are interpreted as `DD-MM-YYYY`.
- Sales/profit variance greater than **0.01** is a mismatch; smaller variances are treated as rounding variation.
- Missing discounts are not inferred from sales values; they are set to 0 only after other sales fields validate.
- Refunded sales are summarized using `sales_raw` to preserve the recorded transaction amount.

### Limitations

- No external customer, product, geographic, or shipping master data was available to validate standardized text values.
- Conflicting duplicate order IDs remain unresolved until an owner identifies the authoritative transaction/source system.
- Ambiguous date parsing depends on the documented locale assumptions.
- The 50% discount cap should be confirmed against company policy before production deployment.
- Raw sales and profit values were preserved; mismatches are flagged rather than overwritten.
- Orders excluded from completed-sales reporting remain in the cleaned file for audit and operational analysis.

## Screenshots Included

### Data Quality Report Overview

![Data Quality Report Overview](outputs/screenshots/data_quality_report_overview.png)

### Pivot Summary Dashboard

![Pivot Summary Dashboard](outputs/screenshots/pivot_summary_dashboard.png)

### Calculated Columns and Quality Flags

![Calculated Columns Preview](outputs/screenshots/calculated_columns_preview.png)

## Audit Reference

For full implementation details, review `outputs/cleaning_log.md`. For evaluator-focused validation totals and detailed exception lists, use `outputs/data_quality_report.xlsx`.
