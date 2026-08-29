# Vantage Vapour Distribution Analytics System

A fully automated FMCG **Numeric & Weighted Distribution** reporting workbook, built entirely in native Excel (formulas + VBA) — no Power BI, no external database. Designed as a portfolio piece demonstrating advanced Excel modelling, VBA automation, and FMCG trade marketing analytics.

 **Note:** All data (brands, SKUs, outlets, sales figures) is synthetic and fictional, generated for demonstration purposes. "Vantage Vapour Distribution" is not a real company.
 
## 📊 Overview
Numeric Distribution and Weighted Distribution are core KPIs in FMCG trade marketing:
- **Numeric Distribution** — the % of audited outlets stocking a given SKU (every outlet counts equally)
- **Weighted Distribution** — the % of total sales value held by outlets stocking that SKU (bigger outlets count more)

These are normally reported through tools like Nielsen or a BI platform. This project replicates that reporting entirely in Excel: a single flat dataset drives every summary, matrix, KPI dashboard, and pivot-style breakdown through live formulas, with VBA macros layered on top for report refresh and audit simulation.

<img width="1857" height="902" alt="Dashboard" src="https://github.com/user-attachments/assets/0afea2ad-65ed-4f99-bac3-b65117bdacc9" />

<img width="1892" height="927" alt="Regional Distribution Matrix" src="https://github.com/user-attachments/assets/c2d73dd9-8d06-4bc0-bf96-4a58c26b9074" />

<img width="1846" height="925" alt="Sales Summary Dashboard" src="https://github.com/user-attachments/assets/2d8f3ba3-9878-483a-ab24-30b59c85cd0a" />


## ✨ Key Features
- **5,400-row dataset** — 100 outlets × 54 SKUs across 12 fictional vaping brands, spanning 3 trade channels and 6 South African provinces grouped into 4 regions
- **Fully formula-driven** — every % and total is calculated live via `SUMPRODUCT` / `COUNTIFS` / `SUMIFS` against the raw data table; nothing is hardcoded
- **KPI Dashboard** — headline metrics, a brand drill-down selector, and channel/brand comparison charts
- **Region → Province drill-down matrix** — a Power BI–style view built with native Excel row grouping, showing both Numeric % and Weighted % side by side, collapsed by default
- **80% distribution benchmark** — automatic red/green conditional formatting flags any region, province, channel, or SKU below target
- **Sales Summary Dashboard** — five PivotTable-style breakdowns (by Province, Brand, Channel, Category, Distribution Status) of Target vs. Actual outlet sales, all driven by a single Region Filter dropdown
- **VBA automation** — macros to refresh the report, simulate a new field audit, flag underperforming SKUs, and toggle the drill-down view
- **Gold & gunmetal theme** — a custom colour system throughout, with functional red/green/cream status colours kept distinct from the branding palette
- **Currency and market context** — all figures in South African Rand (ZAR), using real SA provincial/channel structures

## 🗂️ Workbook Structure
| `Dashboard` | Headline KPIs, brand drill-down, brand/channel comparison charts |
| `Sales Summary Dashboard` | PivotTable-style Target vs. Actual breakdown by Province/Brand/Channel/Category/Status, with a Region Filter |
| `Regional Distribution Matrix` | Region → Province drill-down, Numeric % and Weighted % vs. 80% benchmark |
| `Product Distribution Advanced` | SKU × Channel distribution matrix, Gap vs. Target, Brand Rollup |
| `Distribution Tracker` | The single flat source-of-truth table — one row per outlet × SKU |


The workbook is intentionally deformalized: outlet, product, and sales attributes are inline directly into `Distribution Tracker` rather than split across lookup tables, keeping the file to 7 sheets total with no `VLOOKUP` chains to maintain.

**`Distribution Tracker` columns (A–P):** Outlet ID, Store Name, Channel, Province, Region, SKU Code, Brand, Category, Available Now (Y/N), Available Previous (Y/N), Date Checked, Distribution Status, OOS Flag, Outlet Total Sales (R), Outlet Target Sales (R), Quantity Sold (Units).

🧮 Key Formulas Used
Every distribution % and sales figure in this workbook is calculated live from `Distribution Tracker` — nothing is hardcoded. Column references below match the actual sheet layout (`F`=SKU Code, `G`=Brand, `C`=Channel, `D`=Province, `E`=Region, `I`=Available Now, `L`=Distribution Status, `N`=Outlet Total Sales, `O`=Outlet Target Sales).

**Numeric Distribution % (per SKU, per channel)** — % of audited outlets stocking the SKU:
```excel
=IFERROR(
   COUNTIFS('Distribution Tracker'!$F$5:$F$5404,$A6,
            'Distribution Tracker'!$C$5:$C$5404,"Modern Trade",
            'Distribution Tracker'!$I$5:$I$5404,"Y")
 / COUNTIFS('Distribution Tracker'!$F$5:$F$5404,$A6,
            'Distribution Tracker'!$C$5:$C$5404,"Modern Trade")
,0)


**Weighted Distribution % (per SKU, per channel)** — % of sales value held by stocking outlets:
```excel
=IFERROR(
   SUMPRODUCT(('Distribution Tracker'!$F$5:$F$5404=$A6)*
              ('Distribution Tracker'!$C$5:$C$5404="Modern Trade")*
              ('Distribution Tracker'!$I$5:$I$5404="Y")*
              'Distribution Tracker'!$N$5:$N$5404)
 / SUMPRODUCT(('Distribution Tracker'!$F$5:$F$5404=$A6)*
              ('Distribution Tracker'!$C$5:$C$5404="Modern Trade")*
              'Distribution Tracker'!$N$5:$N$5404)
,0)


**Gap vs. 80% Target**:
```excel
=J6-0.8


**Brand Rollup (Overall Numeric % / Weighted %)** — averages every SKU belonging to a brand:
```excel
=IFERROR(AVERAGEIF($B$6:$B$59,A61,J$6:J$59),0)   ' Numeric roll-up
=IFERROR(AVERAGEIF($B$6:$B$59,A61,K$6:K$59),0)   ' Weighted roll-up


**Distribution Status** — compares this period's stock check to the last one:
```excel
=IF(AND(J5="Y",I5="Y"),"Retained",
 IF(AND(J5="Y",I5="N"),"Lost",
 IF(AND(J5="N",I5="Y"),"New","Never Stocked")))


**OOS (Out of Stock) Flag**:
```excel
=IF(I5="N","OOS","")


**Regional Matrix — Weighted % by Region/Province** — filtered dynamically by whichever geography row it sits on:
```excel
=IFERROR(
   SUMPRODUCT(('Distribution Tracker'!$G$5:$G$5404="Cloud9")*
              ('Distribution Tracker'!$E$5:$E$5404="CENTRAL")*
              ('Distribution Tracker'!$I$5:$I$5404="Y")*
              'Distribution Tracker'!$N$5:$N$5404)
 / SUMPRODUCT(('Distribution Tracker'!$G$5:$G$5404="Cloud9")*
              ('Distribution Tracker'!$E$5:$E$5404="CENTRAL")*
              'Distribution Tracker'!$N$5:$N$5404)
,0)


**Regional Matrix — Numeric % by Region/Province** — same filter pattern with `COUNTIFS`:
```excel
=IFERROR(
   COUNTIFS('Distribution Tracker'!$G$5:$G$5404,"Cloud9",
            'Distribution Tracker'!$E$5:$E$5404,"CENTRAL",
            'Distribution Tracker'!$I$5:$I$5404,"Y")
 / COUNTIFS('Distribution Tracker'!$G$5:$G$5404,"Cloud9",
            'Distribution Tracker'!$E$5:$E$5404,"CENTRAL")
,0)


**Dashboard KPI — Overall Weighted Distribution** (all brands, all channels):
```excel
=IFERROR(
   SUMPRODUCT(('Distribution Tracker'!$I$5:$I$5404="Y")*'Distribution Tracker'!$N$5:$N$5404)
 / SUM('Distribution Tracker'!$N$5:$N$5404)
,0)


**Dashboard KPI — distinct-count shortcuts.** Because the panel is balanced (every outlet has exactly one row per SKU), distinct counts don't need an array formula — they can be derived by simple division:
```excel
SKUs Tracked    =COUNTA('Distribution Tracker'!$F$5:$F$5404)/100   ' 100 = outlet count
Outlets Audited =COUNTA('Distribution Tracker'!$A$5:$A$5404)/54    ' 54 = SKU count


**Sales Summary Dashboard — the "OR filter" pattern.** A single Region Filter cell (`All` or a specific region) drives every table on the sheet. Rather than writing two versions of every formula, each `SUMPRODUCT` includes a term that evaluates to 1 for every row when the filter is `"All"`, and acts as a normal equality filter otherwise:
```excel
=SUMPRODUCT(
   ('Distribution Tracker'!$D$5:$D$5404="Gauteng")*
   (('Distribution Tracker'!$E$5:$E$5404=$C$4)+($C$4="All"))*
   'Distribution Tracker'!$O$5:$O$5404
)
```
The `(Region=Filter)+(Filter="All")` term is the trick: when the filter is a real region name, the second term is always 0 and the first behaves as a normal filter; when the filter is `"All"`, the first term is always 0 (no province is literally named "All") and the second term forces every row to count.

**What it demonstrates:** driving Excel's native outline/grouping feature programmatically, rather than relying on the manual gutter clicks — this is what would sit behind an "Expand All / Collapse All" button.

## 🛠️ Tech / Skills Demonstrated
- Advanced Excel formula design (`SUMPRODUCT`, `COUNTIFS`, `SUMIFS`, `AVERAGEIF`, `IFERROR`)
- Native Excel row outlining for pivot-style drill-down (no PivotTables used — everything is formula-built)
- A single-filter-cell pattern (`(condition)+(filter="All")`) driving multiple summary tables at once
- Conditional formatting as a live benchmarking tool
- VBA / macro development (looping, weighted randomisation, procedural formatting, outline control)
- Flat-table data modelling for a denormalised, formula-friendly dataset
- FMCG trade marketing domain knowledge (Numeric/Weighted Distribution methodology, Target vs. Actual reporting)
- Dashboard design (KPI cards, drill-down selectors, embedded charts)
- Synthetic dataset generation with realistic channel/region/brand distribution patterns
