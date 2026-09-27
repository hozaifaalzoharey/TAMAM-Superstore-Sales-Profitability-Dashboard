# TAMAM | Superstore Sales & Profitability Dashboard

An end-to-end Excel analytics project that turns the classic **Sample – Superstore** retail dataset into a 3-page interactive dashboard, branded as **TAMAM** (a personal e-commerce/fulfillment concept), and a full **Define → Clean → Model → Analyze → Decide** analysis workflow.


## 🧭 Project Overview

The goal of this project was not just to build charts, but to **diagnose why profitability doesn't move in line with sales** across categories, regions, and segments — using a proper analytical workflow rather than surface-level reporting.

| | |
|---|---|
| **Dataset** | Sample – Superstore (US retail orders) |
| **Rows (Fact table)** | 9,994 orders |
| **Time range** | Jan 2, 2014 – Dec 30, 2017 (4 years) |
| **Columns (raw)** | 21, before modeling |
| **Tools** | Excel · Power Query · Power Pivot · DAX · PivotTables/PivotCharts |
| **Output** | 3-page interactive dashboard (`.xlsm`) + documented analysis report |

---

## 🗂️ Repository Structure

```
├── Dashboard.xlsm                                  # Final interactive dashboard (Overview / Sales / Customers)
├── Sample_-_Superstore.xlsb                        # Source dataset
├── Superstore_Sales_Analysis_Documentation.docx    # Full write-up: Define → Clean → Model → Analyze → Decide
├── screenshots/                                    # Dashboard page exports (used in this README)
└── README.md
```

---

## 🔍 1. Define — Understanding the Data

The dataset covers **Orders, Shipping, Customers, Products, and Financials** (Sales, Profit, Discount, Quantity), grouped into 5 logical categories:

- **Order & Shipping:** Row ID, Order ID, Order Date, Ship Date, Ship Mode
- **Customer:** Customer ID, Customer Name, Segment (Consumer / Corporate / Home Office)
- **Location:** Country, City, State, Postal Code, Region
- **Product:** Product ID, Category (3), Sub-Category (17), Product Name
- **Financial:** Sales, Quantity, Discount, Profit

Before any cleaning, a separate **Meta Data** sheet was built (Header Names + sample Unique Values) to catch typos or inconsistent categorical values early — the same discipline applied across all of the author's projects.

There was no obvious ML "target" column here; the objective was **performance diagnostics**, not prediction.

---

## 🧹 2. Clean, Process & Model the Data

### Data type & quality fixes (Power Query)
- **`Order Date` / `Ship Date` inconsistency:** ~5,952 rows (≈60%) were stored as **Text** (e.g. `4/15/2017`) while the remaining ~4,042 rows were stored as **Serial Numbers** for the same field — blocking any reliable sort/filter/group by date. Both were unified into a single `Date` type.
- **Data-quality flag (not silently fixed):** after standardizing the type, ~1,708 rows had `Ship Date` earlier than `Order Date` — most likely a Month/Day vs Day/Month ambiguity in the original text values. Rather than guessing, these rows were **flagged for follow-up** with the source system instead of being "corrected" blindly.
- Reviewed and corrected other data types (e.g. Postal Code).
- Handled **Nulls** per column, based on whether the column was a location or financial field.
- Standardized the **Quarter** format for readable Slicers/PivotTables.
- **Split combined location fields** (Country / City / State) in Power Query so each geographic level is independently filterable and groupable.

### Data Modeling
- Built an independent **`Date` dimension table** (Order Date, Year, Quarter, Month, Month Name, Day, Day Name).
- Related it to the Fact table (`Sample - Superstore`) on `Order Date`, following a proper **star-schema** approach instead of relying on the date column inside the fact table directly.
- Loaded both tables into the **Power Pivot Data Model** to enable DAX measures and shared slicers across all 3 dashboard pages.

---

## 📐 3. Analysis — Power Pivot & DAX

All KPIs were built as **explicit DAX measures** (e.g. `Total Sales = SUM('Sample - Superstore'[Sales])`) rather than Excel's implicit auto-sum fields — for consistent naming and easier maintenance.

### Core KPIs

| Measure | Value |
|---|---|
| Total Sales | $2,297,138.90 |
| Total Profit | $286,397.02 |
| Total Orders (distinct) | 5,009 |
| Total Customers (distinct) | 793 |
| AOV (Sales / Orders) | $458.60 |
| Num of Products | 1,862 |
| Categories / Sub-Categories | 3 / 17 |
| States covered | 49 |

> **Known improvement point (documented, not hidden):** the `Total Discount` measure currently `SUM`s the `Discount` column, which is a *rate* (0–1), not a currency value — the result isn't financially meaningful. The correct fix is `AVERAGE` instead of `SUM`, logged as a next-iteration improvement rather than silently removed.

### Category & Sub-Category performance

| Category | Sales | Profit | Margin % |
|---|---|---|---|
| Technology | $836,154.03 | $145,454.95 | 17.40% |
| Office Supplies | $719,047.03 | $122,490.80 | 17.04% |
| Furniture | $741,937.84 | $18,451.27 | **2.49%** |

Furniture nearly matches Office Supplies in sales but its margin collapses — driven by specific sub-categories:

| Sub-Category | Sales | Profit | Margin % |
|---|---|---|---|
| Tables | $206,965.53 | **-$17,725.48** | -8.56% |
| Bookcases | $114,818.04 | **-$3,472.56** | -3.02% |
| Supplies | $46,673.54 | -$1,189.10 | -2.55% |

### Discount → Profit relationship

- Row-level correlation between Discount and Profit: **-0.22** (mild negative — no strong dataset-wide pattern).
- But it's **not uniform**: `Tables` average discount is **26.1%** and `Bookcases` **21.1%**, both well above the dataset average of **15.6%** — directly explaining their losses.
- Counter-example: `Binders` has a high average discount (**37.2%**) yet still turns a solid profit (**$30,221.76**) — proving discount rate alone isn't a valid blanket explanation and must be assessed **per sub-category**, not as one global rule.

### Region & Segment

| Region | Sales | Profit | Margin % |
|---|---|---|---|
| West | $725,457.82 | $108,418.45 | 14.95% |
| East | $678,781.24 | $91,522.78 | 13.49% |
| South | $391,659.94 | $46,749.43 | 11.94% |
| Central | $501,239.89 | $39,706.36 | **7.92%** |

| Segment | Sales | Profit | Margin % |
|---|---|---|---|
| Consumer | $1,161,339.38 | $134,119.21 | 11.55% |
| Corporate | $706,146.37 | $91,979.13 | 13.03% |
| Home Office | $429,653.15 | $60,298.68 | **14.03%** |

Central outsells South but has the **lowest margin of all regions**; Home Office is the smallest segment by sales but the **most profitable per dollar**.

### Shipping & Seasonality
- `Standard Class` accounts for **~60% of all orders** (2,994 of 5,009) — a clear upsell opportunity toward faster (paid) shipping tiers.
- Orders trend **upward year over year**, with **Q4 consistently peaking** in every year from 2014–2017 (from 385 orders in Q1 2014 to 969 in Q4 2017) — a strong, repeatable holiday-season pattern.

---

## ✅ 4. Decide — Insights & Recommendations

1. **Revisit discount policy on Tables & Bookcases specifically** (not Furniture as a whole) — these two sub-categories alone account for the bulk of the category's margin loss.
2. **Plan inventory & marketing ahead of Q4** — the seasonal spike is consistent across all 4 years and predictable.
3. **Invest further in Home Office** — smallest segment by volume, highest margin; a strong candidate for targeted growth over pure volume-chasing.
4. **Audit cost/pricing in the Central region** — not the lowest in sales, but the weakest in margin among all regions.
5. **Promote faster shipping tiers** — with 60% of orders on Standard Class, there's room to upsell Second/First Class as a paid service.

**Bottom line:** overall sales are healthy and growing — the real story is in **profitability concentration**. Losses are localized to specific sub-categories and margins vary meaningfully by region/segment, so decisions should be made at the **Sub-Category / Region level**, not from Total Sales and Total Profit alone.

---

## 🛠️ Tech Stack

`Excel` · `Power Query (ETL)` · `Power Pivot (Data Model)` · `DAX` · `PivotTables & PivotCharts` · `Slicers`

---

## 👤 Author

**Hozaifa Alzohare** — Freelance Data Analyst (Excel, SQL, Python) | Biotechnology student

- LinkedIn: [linkedin.com/in/hozaifa-alzohare](https://linkedin.com/in/hozaifa-alzohare)
- Email: hozaifaalzohare@gmail.com

---

## 📄 License

This project uses the public **Sample – Superstore** dataset for educational/portfolio purposes. Analysis, dashboard design, and documentation are original work by the author.
