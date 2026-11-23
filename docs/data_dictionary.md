# Data dictionary

Two files in [`data/`](../data):

- **MRP_Dataset.xlsx**: the clean, model ready workbook. Each sheet holds one Excel table that Power BI loads directly.
- **SAP_Production_Orders_raw.csv**: the raw, order level export behind `Fact_Production`, with realistic data quality problems.

## Tables in MRP_Dataset.xlsx

| Table | Contents and grain |
|---|---|
| `Dim_Plant` | One row per plant (SAP WERKS). Nameplate blending capacity. |
| `Dim_Material` | One row per material (SAP MATNR). Standard cost and list price per KG. |
| `Dim_Date` | One row per month. Mark as date table on Month_Start. |
| `Plant_Material` | MRP settings per plant and material (SAP MARC style): lot size, safety stock policy, controller. |
| `Fact_Production` | Plant x material x month. Planned_Qty_KG is the unconstrained MRP plan. Gross, scrap and good output are actuals. Actual_Cost_AED is total production cost. |
| `Fact_Demand` | Plant x material x month. Forecast used by MRP, actual customer demand, fulfilment, backorders and revenue. |
| `Fact_Inventory` | Plant x material x month stock movement: opening + receipts - issues - transfers out = closing. Stock_Over_180_Days_KG uses FIFO ageing. |
| `Fact_Transfers` | Emergency inter plant transfers raised to cover stockouts, with freight cost. |
| `Fact_Capacity` | Available capacity per plant and month in KG, for utilisation. |
| `Fact_MRP_Plan` | Forward MRP run for Sep to Nov 2026: forecast and planned production per plant and material. |

**Relationships:** every fact table joins to `Dim_Plant` on `Plant_Code`, to `Dim_Material` on `Material_No` (except `Fact_Capacity`) and to `Dim_Date` on `Month = Month_Start`. `Fact_Transfers` joins to `Dim_Plant` on `To_Plant`. All relationships are many to one, single direction.

Units: quantities in KG, money in AED. Period: actuals Jan 2025 to Aug 2026, forward MRP plan Sep to Nov 2026.

## Fact table columns

### Fact_Production

| Column | Meaning |
|---|---|
| `Planned_Qty_KG` | Quantity MRP planned for the month. MRP plans against unlimited capacity, so this can exceed what the plant can make. |
| `Gross_Produced_KG` | Total quantity produced, including scrap. |
| `Scrap_KG` | Off specification product written off. |
| `Good_Produced_KG` | Gross minus scrap. Posted to stock as a goods receipt. |
| `Actual_Cost_AED` | Total production cost for the month (materials, conversion, overhead). |

### Fact_Demand

| Column | Meaning |
|---|---|
| `Forecast_Demand_KG` | Demand forecast used by MRP for the month. |
| `Actual_Demand_KG` | Customer demand received. |
| `Delivered_From_Stock_KG` | Demand met from the plant's own stock. |
| `Delivered_Via_Transfer_KG` | Demand met using an emergency transfer from another plant. |
| `Backorder_KG` | Demand not met in the month. |
| `Net_Revenue_AED` | Revenue on delivered quantity. |
| `Revenue_At_Risk_AED` | Revenue value of backorders. |

### Fact_Inventory

| Column | Meaning |
|---|---|
| `Opening_Stock_KG` | Stock at start of month. |
| `Receipts_KG` | Good production received into stock. |
| `Issues_KG` | Stock issued to customers. |
| `Transfer_Out_KG` | Stock sent to other plants. |
| `Closing_Stock_KG` | Opening + receipts - issues - transfer out. |
| `Safety_Stock_KG` | Target safety stock (half a month of forward forecast). |
| `Stock_Over_180_Days_KG` | Portion of closing stock older than six months on a first in, first out basis. |

### Fact_Transfers

| Column | Meaning |
|---|---|
| `From_Plant`, `To_Plant` | Sending and receiving plant. Join To_Plant to Dim_Plant. |
| `Freight_Cost_AED` | Expedited freight cost of the transfer. |

### Fact_Capacity

| Column | Meaning |
|---|---|
| `Available_Capacity_KG` | Blending capacity available in the month. |

### Fact_MRP_Plan

| Column | Meaning |
|---|---|
| `MRP_Planned_Qty_KG` | Forward MRP planned production, unconstrained. |

## SAP_Production_Orders_raw.csv

Order level extract (SAP table AFKO/AFPO style). One row per production order.

| Column | Meaning |
|---|---|
| `AUFNR` | Production order number. Unique key. |
| `AUART` | Order type. |
| `WERKS` | Plant. |
| `MATNR` | Material number. SAP pads to 18 digits with leading zeros. |
| `MAKTX` | Material description. |
| `GSTRP / GLTRP` | Basic start and finish dates, format DD.MM.YYYY. Assign orders to months on GSTRP. |
| `PSMNG` | Order target quantity. |
| `WEMNG` | Goods receipt quantity (good output). |
| `XMNGA` | Scrap quantity. |
| `MEINS` | Unit of measure: KG, or TO for tonnes. |
| `STAT` | TECO technically complete, PDLV partially delivered, DLFL flagged for deletion. |
| `ISTKOSTEN` | Actual cost settled to the order, AED. |
| `ERDAT` | Order creation date. |

## Cleaning steps (Power Query)

1. Trim WERKS (some codes have trailing spaces).
2. Convert MATNR to a whole number to remove leading zeros (mixed formats in the extract).
3. Remove orders with STAT = DLFL (cancelled).
4. Remove duplicate orders on AUFNR (a double extract created repeats).
5. Convert quantities where MEINS = TO to KG (multiply by 1,000). The Oman site (OM01) reports in tonnes.
6. Parse GSTRP as a date using locale English (United Kingdom) or DD.MM.YYYY, then derive the month.
7. Flag orders with blank ISTKOSTEN: cost not yet settled. These explain the cost gap to Fact_Production.

**Check:** After steps 1 to 6, quantities by plant, material and month match Fact_Production exactly.
