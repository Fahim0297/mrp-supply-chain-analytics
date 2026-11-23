# Findings

What the dashboard surfaces in the dataset. Every number is computed from the files in [`data/`](../data). The period is Jan 2025 to Aug 2026 unless stated.

> The data is synthetic: a stand in for an SAP MRP extract, with a fictional company and placeholder site names. These findings show what the model surfaces, not facts about any real business.

## Network at a glance
- 8 sites, 16 chemicals, 70 site and material combinations
- Revenue AED 497.5M over 20 months; production cost AED 285.4M
- Fill rate 98.5% overall, but it drops every summer

## 1. Capacity is in the wrong place (the headline)
- UAE Site 1 (AE01) averages 93% utilisation and 95% to 97% every summer. UAE Site 2 (AE02), in the same country, averages 58%, and 67% in summer 2026.
- Compare like for like when quoting growth: network utilisation was 77% in Jan to Aug 2025 and 83% in Jan to Aug 2026 (AE01: 90% to 96%).
- Both KSA sites ran above 95% in summer 2026.
- MRP plans against unlimited capacity. AE01's MRP plan exceeded its capacity in 16 of 20 months, peaking at 1.89 times capacity. Plan attainment at AE01 is 72% (71.5%), against about 97% at sites with spare capacity. That is not poor execution: the plan was never feasible.

## 2. The cost of the capacity gap
- AED 8.61M of revenue at risk from backorders, of which AED 7.83M falls in May to Aug 2026.
- AE01 accounts for AED 6.22M (72%) and SA02 for AED 2.39M. No other site loses revenue.
- 62 emergency transfers moved 114 t between sites at AED 128K of expedited freight. Transfers are reactive and only cover part of the gap.

## 3. Forward risk: the September plan will fail
- Sep to Nov 2026 as a whole, the network is planned at only 95% of capacity. Capacity exists; it is in the wrong place.
- In September alone, MRP asks AE01 for 830 t against 480 t of capacity (173%), SA02 for 171% and SA01 for 128%.
- AE02 has 138 t spare in September, and 683 t of AE01's September plan is on materials AE02 already makes.
- Over the three months: AE01 is 401 t over capacity and AE02 has 473 t spare. SA01 and SA02 are 243 t over combined; the spare capacity in Kuwait, Qatar and Bahrain means shipping across borders.

## 4. Demand shift at KSA Site 2 (SA02)
- SA02 demand for RO Antiscalant, phosphonate based (400101) was 90% higher in Mar to Aug 2026 than a year earlier, consistent with a large RO or data centre project ramping up. The forecast lags because it uses a trailing average.

## 5. Quality problem at the Oman Site (OM01)
- Scrap rate 6.0% against 1.4% for the rest of the network.
- About 156 t of excess scrap, worth AED 1.36M (the dashboard measure, which uses a 1.4% target row by row, shows AED 1.34M).
- Cost per kg 16% to 21% higher than other sites making the same materials, so it is the site, not the product mix.

## 6. Cash tied up in slow stock at the Qatar and Kuwait sites
- Four combinations (Defoamer 400304 and Oxygen Scavenger 400401 at QA01 and KW01) hold AED 1.54M of stock: 32% of network stock value (AED 4.78M) from 6% of combinations.
- They carry 15 to 20 months of cover, against a network median of about a third of a month. AED 0.52M is already older than 180 days, and both products have a 12 month shelf life, so write offs are close.
- Root cause: demand fell after a customer loss in 2025. The forecast was never revised down (2026 forecasts ran at 2.1 to 2.3 times actual; over the full 20 months the bias is +47% to +82%, which is what the dashboard table shows), and a firmed production schedule kept running.

## 7. Forecast accuracy
- Network forecast accuracy is 91.9% (1 minus weighted absolute error).
- The four stale forecasts above are the outliers, all heavily biased upward. The average hides them.

## 8. Data quality in the raw SAP extract
- Trailing spaces in site codes, mixed material number formats, 22 duplicated orders, 15 cancelled orders, Oman site quantities in tonnes, and dates as DD.MM.YYYY text.
- 9 open orders have no settled cost, leaving an AED 947K cost gap against Fact_Production. It is caught and explained before anything is published.
- After cleaning, quantities reconcile exactly to Fact_Production.

## How it fits together
Findings 1 to 3: MRP creates plans the network cannot build, and revenue is lost. Findings 5 and 6: margin and cash leak at the less loaded sites. Finding 7 connects them, because the forecast drives both. In Pareto terms, two sites explain all the lost revenue, one site explains most of the scrap cost, and four products hold a third of the stock value.
