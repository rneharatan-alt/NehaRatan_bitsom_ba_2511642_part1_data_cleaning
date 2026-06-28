# Cleaning Log — Order-Level Sales Dataset

**Run date:** June 28, 2026  
**Source preserved:** `raw_orders.xlsx` — 932 source records; not overwritten or modified.  
**Primary cleaned output:** `cleaned_orders.xlsx` — 912 retained records after exact de-duplication.  
**Supporting outputs:** `outputs/data_quality_report.xlsx` and `outputs/pivot_summary.xlsx`.

## 1. Scope

This log documents the cleaning, validation, duplicate handling, business rules, calculated fields, and reporting decisions applied to the order-level sales extract. It is intended to provide an auditable explanation of how raw source records were converted into the analysis-ready cleaned dataset.

## 2. Issues Found

| Issue area | Finding | Records / occurrences | Final treatment |
|---|---:|---:|---|
| Text formatting | Leading/trailing spaces, repeated spaces, case inconsistencies, and known naming variants occurred in requested text fields. | 106 source occurrences | Standardized and recorded as a cleaning action. |
| Missing region | `region` was blank or whitespace-only. | 26 raw; 25 retained after de-duplication | Set to `Unknown`; retained as a warning/audit adjustment. |
| Missing ship mode | `ship_mode` was blank or whitespace-only. | 22 raw; 21 retained after de-duplication | Set to `Unknown`; retained as a warning/audit adjustment. |
| Missing discount | `discount` was missing. | 18 retained | Set to `0` only after other sales fields passed validation. |
| Invalid discount | Negative discounts or discounts above the allowed limit. | 30 retained (15 negative, 15 above 50%) | Flagged invalid; excluded from calculated sales/profit and completed-sales eligibility. |
| Exact duplicates | Entire raw rows matched exactly. | 20 | Later duplicate occurrences removed; the first occurrence was retained. |
| Duplicate order IDs with conflicts | Same `order_id` had non-identical information. | 12 IDs; 24 unique records retained | Retained and flagged for manual reconciliation; never silently deleted. |
| Date sequence | `ship_date` occurred before `order_date`. | 21 retained | Preserved, flagged as invalid shipping records, and excluded from completed-sales eligibility. |
| Sales calculation mismatch | Raw sales differed from formula-calculated sales by more than 0.01. | 34 records | Flagged invalid/review-level. |
| Profit calculation mismatch | Raw profit differed from formula-calculated profit by more than 0.01. | 34 records | Flagged invalid/review-level. |
| Any sales or profit mismatch | At least one of the sales/profit checks failed. | 53 unique records | Flagged invalid/review-level. |
| Status/payment inconsistency | Completed orders that were not paid. | 2 records | Flagged invalid and excluded from completed-sales eligibility. |
| Cancelled orders | `order_status = Cancelled`. | 145 retained | Retained for audit; excluded from completed-sales summary. |
| Failed payments | `payment_status = Failed`. | 69 retained | Retained for audit; excluded from completed-sales summary. |
| Refunded payments | `payment_status = Refunded`. | 71 retained | Retained and reported separately; excluded from completed-sales summary. |

> Issue categories can overlap. The mutually exclusive final quality classification is reported in Section 7.

## 3. Cleaning Actions Performed

### 3.1 Text fields standardized

The following fields were standardized: `customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, and `order_status`.

Actions applied:

- Trimmed leading and trailing spaces.
- Collapsed repeated internal spaces.
- Standardized capitalization and case variants.
- Normalized known equivalent category and sub-category names using controlled mappings.
- Removed or normalized unwanted formatting characters only where they prevented consistent matching; no broad destructive character stripping was applied to names or product-related text.
- Filled missing `region` and `ship_mode` values with `Unknown` rather than guessing a geographic or shipping value.

The cleaning approach corresponds to Excel-style techniques such as `TRIM`, `SUBSTITUTE`, controlled Find/Replace mappings, and standardized output fields. Raw values remain available in the preserved source workbook/sheets for traceability.

### 3.2 Dates standardized and validated

- `order_date` and `ship_date` were converted to real Excel dates and displayed as `yyyy-mm-dd`.
- Numeric slash dates were interpreted as `MM/DD/YYYY`.
- Numeric hyphen dates were interpreted as `DD-MM-YYYY`.
- Recognized text/ISO date values were converted to real dates.
- No retained records had missing or unrecognized date text after cleaning.
- `shipping_delay_days` was calculated as `ship_date - order_date`.
- Negative shipping delays were retained as evidence and flagged instead of being changed without source confirmation.

### 3.3 Duplicates handled

- Exact duplicates were defined as records identical across every raw source field.
- The first occurrence of each exact duplicate set was retained; later identical occurrences were removed.
- Repeated `order_id` values with conflicting field values were **not** collapsed or deleted. All unique conflicting records were retained and marked for review.

### 3.4 Calculation and quality fields added

The cleaned workbook includes the following analysis fields:

| Field | Logic |
|---|---|
| `cleaned_discount` | Valid source discount, or `0` only for eligible missing-discount records; blank for invalid discounts. |
| `calculated_sales` | `quantity × unit_price × (1 − cleaned_discount)`. |
| `calculated_profit` | `calculated_sales − cost_raw`. |
| `profit_margin` | `calculated_profit ÷ calculated_sales`, where calculated sales is valid and non-zero. |
| `shipping_delay_days` | `ship_date − order_date`. |
| `order_month` / `order_year` | Extracted from the cleaned `order_date`. |
| `sales_calculation_variance` | `sales_raw − calculated_sales`. |
| `profit_calculation_variance` | `profit_raw − calculated_profit`. |
| `data_quality_flag` | Mutually exclusive `Clean`, `Warning`, or `Invalid` classification. |
| `data_quality_details` | Human-readable explanation of record-level flags. |

## 4. Business Rules Applied

| Rule area | Decision applied | Result / reporting treatment |
|---|---|---|
| Missing region | Set blank/whitespace values to `Unknown`. | Retained and identified as a warning/audit adjustment. |
| Missing ship mode | Set blank/whitespace values to `Unknown`. | Retained and identified as a warning/audit adjustment. |
| Missing discount | Treat as `0` only when quantity, unit price, raw sales, raw cost, and raw profit are present and internally valid. | 18 records imputed to `0`; no discount was inferred from sales values. |
| Negative discount | Invalid. | Retained for audit; excluded from validated sales/profit calculations and completed-sales eligibility. |
| Discount above allowed range | Invalid when above 50%. | Retained for audit; excluded from validated sales/profit calculations and completed-sales eligibility. |
| Cancelled orders | Do not include in final completed-sales summary. | Retained for audit and operational reporting. |
| Failed payments | Do not include in final completed-sales summary. | Retained for audit and operational reporting. |
| Refunded orders | Do not include in completed-sales summary. | Summarized separately using recorded `sales_raw`. |
| Ship date before order date | Invalid shipping record. | Retained and flagged; excluded from completed-sales eligibility. |
| Completed-sales eligibility | Include only `Completed` + `Paid` records with valid discount, usable dates, no calculation mismatch, no duplicate conflict, and no unresolved review-level issue. | 537 records qualify for final completed-sales reporting. |

## 5. Assumptions Made

1. **Allowed discount range:** The source requirements did not provide a maximum discount. A maximum of **50%** was adopted; valid discounts must be between 0% and 50%, inclusive.
2. **Date parsing:** Slash-formatted numeric dates were treated as `MM/DD/YYYY`, while hyphen-formatted numeric dates were treated as `DD-MM-YYYY`.
3. **Calculation tolerance:** A sales or profit difference greater than **0.01** is treated as a mismatch. Differences at or below 0.01 are treated as rounding variation.
4. **Missing discount imputation:** A missing discount is treated as 0 only when the rest of the transaction fields are present and internally reconcilable; it is never inferred solely from the raw sales value.
5. **Duplicate logic:** Only truly identical full-row records are deleted. Repeated `order_id` values with different content represent possible source-system conflicts and are retained.
6. **Refund amount basis:** Refunded sales are summarized using `sales_raw` to preserve the source transaction amount rather than a recalculated completed-sale value.
7. **Unknown text values:** Missing `region` and `ship_mode` are classified as `Unknown`; no geographic or operational value is guessed from adjacent fields.

## 6. Records Removed and Retained for Review

### Removed

| Item | Count | Removal logic |
|---|---:|---|
| Exact duplicate rows | 20 | Removed only where every raw field matched a previous row exactly. |

### Retained and flagged

| Flag category | Count | Action |
|---|---:|---|
| Warning records | 45 | Retained after a remediated source issue or audit warning, such as `Unknown` region/ship mode. |
| Invalid records | 102 | Retained with a rule breach or calculation mismatch. |
| Total final flagged records | 147 | Warning + Invalid; mutually exclusive classification across 912 retained records. |
| Conflicting duplicate records | 24 | Retained for manual source-system reconciliation. |
| Invalid-discount records | 30 | Retained; not used for calculated sales/profit or final completed sales. |
| Negative-shipping-delay records | 21 | Retained; not auto-corrected. |
| Any sales/profit mismatch records | 53 | Retained for financial/data-owner review. |

## 7. Final Record Quality Outcome

| Final classification | Count | Share of 912 retained records | Meaning |
|---|---:|---:|---|
| Clean | 765 | 83.88% | All validation checks passed. |
| Warning | 45 | 4.93% | Source issue was remediated or retained as an audit warning. |
| Invalid | 102 | 11.18% | At least one invalid rule condition or calculation mismatch exists. |
| **Flagged total** | **147** | **16.12%** | Warning + Invalid. |
| **Retained total** | **912** | **100.00%** | After removal of 20 exact duplicates. |

## 8. Sales Summary Treatment

- **Eligible completed-sales records:** 537
- **Validated completed sales:** 5,329,453.93
- **Validated completed profit:** 1,549,060.04
- **Refunded orders separately summarized:** 71
- **Refunded raw sales value:** 677,073.75
- **Cancelled, failed-payment, refunded, invalid, and unresolved conflict records are not included in the final completed-sales summary.**

## 9. Limitations

1. **No external master-data verification:** Customer, geography, category, and ship-mode corrections were standardized from the supplied data; they were not verified against an external customer, product, or logistics master.
2. **No source-system reconciliation:** Conflicting duplicate `order_id` records remain unresolved until a data owner identifies the authoritative source transaction.
3. **Date interpretation depends on stated assumptions:** Ambiguous slash/hyphen formats were parsed using the documented convention. A different source-system locale could change the intended date for ambiguous values.
4. **Discount cap is an operational assumption:** The 50% maximum should be confirmed with the company’s discount policy before production use.
5. **Calculated values do not overwrite source values:** Raw sales/profit values are preserved. Calculation mismatches are flagged rather than automatically corrected.
6. **Eligibility is a reporting rule, not deletion logic:** Orders excluded from completed-sales reporting remain in the cleaned workbook for audit, operational analysis, and reconciliation.
7. **Issue categories overlap:** Counts by individual issue should not be added together. Use the final mutually exclusive `Clean` / `Warning` / `Invalid` classification for the total flagged population.

## 10. Audit Locations

- `cleaned_orders.xlsx` → `Cleaned_Orders`, `Issue_Log`, `Date_Validation`, `Business_Rules_Summary`, and `Refunded_Sales_Summary`.
- `outputs/data_quality_report.xlsx` → `Overview`, `Missing_Value_Summary`, `Duplicate_Summary`, `Invalid_Discount_Summary`, `Date_Issue_Summary`, `Order_Status_Issues`, `Sales_Profit_Mismatches`, and `Final_Clean_vs_Flagged`.
- `outputs/pivot_summary.xlsx` → business-oriented summaries by region, category, ship mode, segment, status, and month.
