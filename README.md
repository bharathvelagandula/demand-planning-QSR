
# Synchronized QSR Demand Planning, Staffing & Inventory System (`Demand_Planning_.xlsx`)

A practical, end-to-end weekly demand planning workbook built in Microsoft Excel for a Quick-Service Restaurant (QSR) franchise. The core design principle is **one synchronized consensus plan**: once the weekly **Approved Demand Forecast** is set, **shift staffing schedules** and **SKU-level inventory order quantities** update automatically—grounded in Demand Management best practices and a continuous **Plan–Do–Check–Act (PDCA)** cycle, while keeping every formula simple enough for a store manager to run without VBA or macros.

## Design Philosophy: Demand Management + PDCA
- **One Synchronized Number:** `Approved Weekly Forecast → Dynamic Labor Capacity (Staffing) → SKU-Level Replenishment (Inventory)` stay locked in sync across the **Sunday order → Wednesday–Tuesday delivery week** cycle (3-day lead time, 7-day coverage window).
- **Demand Management Loop:** Translates the four core pillars—**Plan Demand, Communicate Demand, Influence Demand, and Manage & Prioritize Demand**—into a repeatable weekly store routine.
- **Multi-View Demand & FVA Governance:** Combines a statistical ETS baseline with a structured **Sunday Decisions, Assumption Log & Forecast Value Added (FVA)** table so weather, holiday, and promotional assumptions are documented and measured against actuals in a continuous **PDCA** loop.
- **Simplicity Over Complexity:** Green cells mark user inputs; white/gray cells hold transparent, standard Excel formulas so store leadership can adopt and maintain the process reliably.

## Key Metrics (Dataset Snapshot)
- **Historical POS Sell-Through:** 359 daily records (`Oct 4, 2025 – Oct 3, 2026`), totaling 95,557 units and $687,503 net sales (winter trough: 199.4 units/day in Dec 2025; summer peak: 325.0 units/day in Jul 2026)
- **Active Delivery-Week Plan (`Oct 7 – Oct 13, 2026`):** 2,115 units, $15,364.90 net sales, and 34 total shifts via `FORECAST.ETS` (7-day seasonality)
- **Unadjusted Baseline Accuracy:** 85.8% volume-weighted unit accuracy and 86.2% volume-weighted sales accuracy on the logged Oct 7–13 cycle—*before* manual FVA overrides, establishing a clean statistical benchmark for assumption-driven FVA gains
- **Replenishment Disaggregation:** 14 product categories and 88 active SKUs (2,013 prior-week plan units, automatically filtering out non-inventory `ORDER TYPE` service rows)
- **Monthly ABC–XYZ Portfolio Segmentation:** 102 SKUs ($59,137 4-week revenue, `Sep 2–29, 2026`) segmented into 31 AX (steady core), 4 AY (high-value variable), 13 BX, 10 BY, 1 BZ, 4 CX, 8 CY, and 31 CZ SKUs

## Workbook Architecture (7 Tabs)
1. **Forecast** – Rolling 7-day `FORECAST.ETS` engine (auto-advances to the next Wednesday–Tuesday delivery week), dynamic 4-tier labor capacity table (`AVERAGE`/`AVERAGEIF` + `LOOKUP`, 3–6 staff), and the Sunday Decisions, Assumption Log & FVA tracker.
2. **Regular sales** – Structured Excel Table (`Sales`: Date, Units, Net sales) that auto-expands into `Forecast`, plus a 12-month average daily seasonality summary.
3. **Item sales actuals** – POS data cleansing zone accepting 10-column "Hide Modifiers" CSV exports; auto-excludes `ORDER TYPE` rows and computes Plan Units and SKU mix shares.
4. **Inventory planning** – Synced live to the approved `Forecast` delivery week; disaggregates store demand into 14 category totals and 88 active SKU order quantities (`INDEX/MATCH`) alongside a 4-ingredient Weekly Stock Log (`Projected Close = Opening + Received − Planned Usage`).
5. **ABC-XYZ analysis** – Monthly 4-week revenue Pareto (A ≤ 80%, B ≤ 95%, C > 95%) and weekly demand variability CV (`STDEVP / AVERAGE`: X ≤ 25%, Y ≤ 50%, Z > 50%) with a 3×3 matrix and Top 6 category breakdown.
6. **Try it** – Interactive 10-day demo dataset (`Oct 4 – Oct 13, 2026`) to test the rolling forecast live.


## Try the Live Rollover Demo (3 Steps)
1. Download and open `Demand_Planning_.xlsx` in Microsoft Excel.
2. Open the **Try it** tab and copy the green demo sales rows (`A7:C16`, Oct 4–13).
3. Paste Values at the bottom of the `Sales` table on **Regular sales** (`A364`), then switch to **Forecast** and **Inventory planning** to watch the dates, 7-day ETS forecast, staffing tiers, and SKU order quantities roll automatically from **Oct 7–13** to **Oct 14–20, 2026** (Sunday Oct 11 order).

## Tools & Methods
- **Excel Modeling:** `FORECAST.ETS`, Structured Tables, `SUMIF`/`SUMIFS`, `INDEX/MATCH`, `LOOKUP`, `AVERAGEIF`, `STDEVP`
- **Supply Chain & Demand Planning:** Time-Series Seasonal Forecasting, Forecast Value Added (FVA), Assumption Logging, PDCA Governance, Volume-Weighted Accuracy (WAPE), POS Data Cleansing, Top-Down SKU Disaggregation, ABC–XYZ Segmentation, Labor Capacity Planning.

## Dashboard Previews
