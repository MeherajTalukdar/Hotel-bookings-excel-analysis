# Hotel Bookings — Excel Data Analysis & Business Recommendation

A complete, hands-on Excel data analysis project using a real-world hotel bookings dataset — taken all the way from raw formulas to a business recommendation memo a revenue manager could act on.

## Project Overview

This project applies everything from basic formulas through to Power Query, Power Pivot, and DAX on a real hotel industry dataset with 119,390 booking records.


## Repository Structure

```
hotel-bookings-excel-analysis/
│
├── README.md
│
├── questions/
│   └── business_questions.md            ← All 20 business questions
│
├── solutions/
│   ├── 01_beginner_formulas.md          ← Q1–Q5  : Formulas & Functions
│   ├── 02_intermediate_pivots.md        ← Q6–Q12 : PivotTables & Charts
│   ├── 03_power_query_cleaning.md       ← Q13–Q15: Power Query
│   └── 04_power_pivot_dax.md            ← Q16–Q20: Power Pivot, DAX & Dashboard
│
├── reports/
│   └── Hotel_Portfolio_Business_Recommendation.docx   ← Business recommendation memo (findings → action plan)
│
├── Dashboard & Other Calculations.xlsx  ← Dashboard, Calculations, What-If Analysis
├── hotel_bookings.xlsx                  ← Raw dataset
│
└── assets/
    └── dashboard_preview.png            ← Dashboard screenshot
```

---

## Dataset

**Source:** [Hotel Booking Demand — Kaggle](https://www.kaggle.com/datasets/jessemostipak/hotel-booking-demand)
**Original paper:** Antonio, Almeida & Nunes (2019), *Hotel booking demand datasets*, Data in Brief

| Property    | Value                                            |
| ----------- | ------------------------------------------------ |
| Rows        | 119,390 bookings                                 |
| Columns     | 32 original + 4 engineered                       |
| Period      | July 2015 – August 2017                          |
| Hotels      | City Hotel & Resort Hotel                        |
| File format | CSV (semicolon-delimited) → converted to `.xlsx` |

## Tools & Skills Used

| Tool                 | Skills Applied                                                              |
| -------------------- | ---------------------------------------------------------------------------- |
| **Excel Formulas**   | `COUNTIF`, `AVERAGEIF`, `SUMIF`, `IFS`, `LARGE`, `INDEX/MATCH`                |
| **PivotTables**      | Grouping, calculated fields, top N filters, % of total                       |
| **PivotCharts**      | Line, bar, column, combo charts with slicers                                 |
| **Power Query**      | Import, clean nulls, build date column, append, merge                        |
| **Power Pivot**      | Data Model, relationships, DAX measures                                      |
| **DAX**              | `COUNTROWS`, `CALCULATE`, `FILTER`, `SUMX`, `DIVIDE`, `AVERAGE`, `TOTALYTD`  |
| **What-If Analysis** | Two-variable Data Table                                                      |
| **Dashboard Design** | Slicers, `GETPIVOTDATA`, linked cells, layout                                |
| **Data validation**  | Cross-tab reconciliation, outlier detection, pivot-field auditing            |

---

## Business Questions Summary

> Full questions with hints: [`questions/business_questions.md`](questions/business_questions.md)
> Step-by-step solutions: [`solutions/`](solutions)

### Beginner — Formulas & Functions (Q1–Q5)
1. What is the overall cancellation rate across all bookings?
2. Which hotel type has a higher average daily rate?
3. Add a `total_guests` calculated column to the data.
4. Classify each booking's lead time as Last Minute / Short / Medium / Long.
5. How many bookings had a room upgrade?

### Intermediate — PivotTables & Charts (Q6–Q12)
6. Which month has the highest number of arrivals?
7. What is the cancellation rate by market segment?
8. Build a chart: average ADR by month for each hotel type.
9. Which top 10 countries do most guests come from?
10. What is the average length of stay by customer type?
11. Do guests who make more special requests cancel less?
12. What is the cancellation rate of repeated vs new guests?

### Advanced — Power Query (Q13–Q15)
13. Build a proper `arrival_date` column from three separate columns.
14. Replace "NULL" text values in `agent` and `company` with proper blanks.
15. Split the dataset by hotel type, then Append back into one master table.

### Advanced — Power Pivot & DAX (Q16–Q20)
16. DAX measure: `Est Revenue = adr × total_nights`. Revenue by hotel and year?
17. DAX measure: Cancellation Rate. Does deposit type affect cancellations?
18. DAX measure: Average waiting list days for canceled vs. completed bookings.
19. **Capstone:** Full interactive Hotel Performance Dashboard with slicers.
20. What-If Analysis: How does revenue change if ADR increases by 5/10/15%?

---

## Key Findings

| Metric                        | Value                                      |
| ------------------------------ | ------------------------------------------- |
| Total bookings                 | 119,390                                     |
| Portfolio mix                  | City Hotel 79,330 (66.4%) · Resort Hotel 40,060 (33.6%) |
| Overall cancellation rate      | 37.0% (44,224 bookings)                     |
| Cancellation rate by hotel     | City Hotel 41.7% · Resort Hotel 27.8%       |
| Average length of stay         | 3.4 nights (Contract 5.3 · Group 2.9)       |
| Largest booking channel        | Online TA — 56,477 bookings (47.3%), 36.7% cancellation |
| Highest-cancellation segment   | Groups — 61.1% (19,811 bookings)            |
| Repeat vs. new guest cancellation | 14.5% vs. 37.8%                           |
| Top guest country              | Portugal (48,590 bookings)                  |
| Best revenue year              | 2016                                        |

---

## Business Recommendations

The 20 technical questions feed directly into a one-page memo (`reports/Hotel_Portfolio_Business_Recommendation.docx`) written as if delivered to hotel-group leadership. Summary:

| # | Area | Finding | Recommendation | Priority |
|---|------|---------|-----------------|----------|
| 1 | Distribution & Cancellation Risk | Groups cancel at 61.1%; Online TA (47.3% of volume) cancels at 36.7% | Tiered deposit/cancellation terms for Group & OTA bookings; rebalance spend toward Direct/Corporate | High |
| 2 | Revenue Management | Resort ADR swings sharply by season; City stays flat | Build a formal seasonal pricing calendar for the Resort property | High |
| 3 | Retention | Repeat guests cancel at 14.5% vs. 37.8% new; cancellation drops as special requests rise | Expand loyalty program; train staff to capture preferences at booking | Medium |
| 4 | Market Diversification | Portugal accounts for far more bookings than the UK, France, Spain and Germany combined | Redirect marketing spend toward under-represented markets | Medium |
| 5 | Data Governance | "Non Refund" bookings show 99.4% cancellation — counter-intuitive | Verify the cancellation-status definition with Operations before using it in policy | High |

---

## Data Quality & Validation

Before trusting a dashboard, an analyst has to be willing to break it. Reconciling this workbook's own numbers against each other surfaced two real issues worth documenting rather than hiding:

- **Inconsistent headline metrics.** The live dashboard reports Average ADR as €2,201.85 and Estimated Revenue as €1.26B (Dashboard Calculations tab) or €901.8M (What-If Analysis tab) — none of which match this README's own €101.83 / €42.7M. Root cause: 30.1% of the raw `adr` column exceeds €1,000/night, 16.8% exceeds €5,000, and one record is negative. The median (€149) is far more representative than the contaminated mean.
- **A mislabeled pivot field.** The "Deposit Type: Cancellation Rate" table on the Dashboard Calculations tab is actually grouped by `market_segment`, not `deposit_type` — confirmed by cross-checking the pivot's category values and rates against the raw data.

Both issues, and the recommended fix, are documented in full in `reports/Hotel_Portfolio_Business_Recommendation.docx`.

---

## What This Project Demonstrates

- End-to-end Excel fluency: formulas → PivotTables → Power Query → Power Pivot/DAX → What-If Analysis → dashboard.
- The habit of validating a dashboard's numbers against its own source data before presenting them — not just building the dashboard.
- Translating a technical analysis into a prioritized, owner-assigned business recommendation, the deliverable format used in real revenue-management and BI reporting roles.
