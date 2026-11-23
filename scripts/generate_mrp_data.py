"""
Generate a synthetic SAP style MRP dataset for a fictional network of water
treatment chemical blending sites in the GCC.

The simulation runs month by month:

  demand  ->  forecast (trailing average)  ->  MRP net requirements
          ->  capacity constrained production  ->  deliveries, backorders
          ->  emergency transfers between sites  ->  closing stock

It deliberately builds in the patterns the dashboard is designed to surface:

  1. Capacity in the wrong place: UAE Site 1 runs near full while UAE Site 2
     has spare capacity, and MRP plans against unlimited capacity.
  2. A demand ramp at KSA Site 2 that a trailing average forecast lags behind.
  3. A quality problem: the Oman site scraps about four times the network rate.
  4. Stale forecasts: two products at the Qatar and Kuwait sites lose a
     customer, the forecast is never revised, and stock piles up.
  5. A raw order level export with typical SAP extract problems.

Outputs (default folder: ./output):
  MRP_Dataset.xlsx               clean, model ready workbook (one Excel table per sheet)
  SAP_Production_Orders_raw.csv  raw production order export with data quality issues

Usage:
  python generate_mrp_data.py [--out OUTPUT_DIR] [--seed SEED]
"""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# --------------------------------------------------------------------------
# Master data
# --------------------------------------------------------------------------

PLANTS = pd.DataFrame(
    [
        # code, name, country, capacity t/month, cost factor, scrap rate, target load
        ("AE01", "UAE Site 1", "UAE", 480, 1.00, 0.012, 0.94),
        ("AE02", "UAE Site 2", "UAE", 380, 1.03, 0.014, 0.55),
        ("SA01", "KSA Site 1", "Saudi Arabia", 410, 0.98, 0.013, 0.85),
        ("SA02", "KSA Site 2", "Saudi Arabia", 280, 1.01, 0.016, 0.78),
        ("OM01", "Oman Site", "Oman", 210, 1.16, 0.058, 0.74),
        ("QA01", "Qatar Site", "Qatar", 160, 1.05, 0.015, 0.68),
        ("KW01", "Kuwait Site", "Kuwait", 150, 1.06, 0.014, 0.60),
        ("BH01", "Bahrain Site", "Bahrain", 105, 1.04, 0.015, 0.70),
    ],
    columns=["Plant_Code", "Plant_Name", "Country", "Capacity_T_per_Month",
             "cost_factor", "scrap_rate", "target_load"],
)

MATERIALS = pd.DataFrame(
    [
        # number, description, family, std cost, list price (AED/kg), demand weight, lot kg, seasonal, shelf life
        (400101, "RO Antiscalant, phosphonate based", "RO and Membrane", 9.8, 17.5, 10.0, 5000, False, 24),
        (400102, "RO Antiscalant, phosphonate free", "RO and Membrane", 13.2, 23.0, 2.4, 2000, False, 24),
        (400110, "Membrane Cleaner, acidic", "RO and Membrane", 6.4, 11.8, 3.2, 2000, False, 24),
        (400111, "Membrane Cleaner, alkaline", "RO and Membrane", 7.1, 12.9, 2.9, 2000, False, 24),
        (400120, "Dechlorinator, bisulphite blend", "RO and Membrane", 3.9, 7.2, 2.0, 2000, False, 12),
        (400201, "Cooling Tower Scale Inhibitor", "Cooling Water", 8.6, 15.4, 4.4, 3000, True, 24),
        (400202, "Cooling Corrosion Inhibitor", "Cooling Water", 10.4, 18.6, 2.6, 2000, True, 24),
        (400203, "Non oxidising Biocide", "Cooling Water", 18.5, 33.0, 2.2, 1000, True, 12),
        (400204, "Oxidising Biocide, stabilised bromine", "Cooling Water", 12.7, 22.4, 1.5, 1000, True, 12),
        (400301, "Coagulant, polyaluminium chloride", "Clarification and Wastewater", 2.9, 5.6, 8.5, 10000, False, 12),
        (400302, "Organic Coagulant, polyDADMAC", "Clarification and Wastewater", 7.8, 14.1, 1.8, 2000, False, 18),
        (400303, "Anionic Flocculant, polyacrylamide", "Clarification and Wastewater", 11.9, 21.0, 1.4, 1000, False, 24),
        (400304, "Defoamer, silicone based", "Clarification and Wastewater", 14.8, 25.5, 0.9, 1000, False, 12),
        (400401, "Oxygen Scavenger", "Boiler Water", 9.1, 16.3, 1.1, 1000, False, 12),
        (400402, "Boiler Internal Treatment, phosphate", "Boiler Water", 8.2, 14.6, 0.8, 1000, False, 24),
        (400403, "Condensate Treatment, neutralising amine", "Boiler Water", 16.9, 29.8, 0.6, 1000, False, 18),
    ],
    columns=["Material_No", "Material_Desc", "Product_Family", "Std_Cost_AED_per_KG",
             "List_Price_AED_per_KG", "weight", "lot_kg", "seasonal", "Shelf_Life_Months"],
)

# Which materials each site blends
ASSORTMENT = {
    "AE01": [400101, 400102, 400110, 400111, 400120, 400201, 400202, 400203, 400204, 400301, 400302, 400303],
    "AE02": [400101, 400110, 400111, 400201, 400202, 400203, 400301, 400303, 400401, 400402],
    "SA01": [400101, 400102, 400110, 400111, 400120, 400201, 400202, 400203, 400301, 400302, 400303],
    "SA02": [400101, 400110, 400111, 400120, 400201, 400203, 400301, 400303],
    "OM01": [400101, 400110, 400201, 400202, 400301, 400303, 400401],
    "QA01": [400101, 400201, 400203, 400301, 400304, 400401, 400402, 400403],
    "KW01": [400110, 400201, 400202, 400301, 400304, 400401, 400402, 400403],
    "BH01": [400101, 400201, 400204, 400301, 400304, 400402],
}

# Monthly seasonality (cooling water demand peaks in the Gulf summer)
SEASON = {1: 0.82, 2: 0.84, 3: 0.92, 4: 1.00, 5: 1.12, 6: 1.25,
          7: 1.33, 8: 1.32, 9: 1.18, 10: 1.02, 11: 0.90, 12: 0.84}

WARMUP = pd.period_range("2024-07", "2024-12", freq="M")   # history used to seed forecasts
ACTUALS = pd.period_range("2025-01", "2026-08", freq="M")
FORWARD = pd.period_range("2026-09", "2026-11", freq="M")

# Planted scenarios
DEMAND_RAMPS = {  # (plant, material): (start month, uplift once fully ramped)
    ("SA02", 400101): (pd.Period("2026-01", "M"), 0.95),
    ("SA02", 400110): (pd.Period("2026-02", "M"), 0.45),
    ("AE01", 400101): (pd.Period("2026-01", "M"), 0.30),
    ("SA01", 400101): (pd.Period("2026-02", "M"), 0.20),
}
CUSTOMER_LOSS = {  # (plant, material): month the customer is lost; forecast is never revised
    ("QA01", 400304): pd.Period("2025-05", "M"),
    ("KW01", 400304): pd.Period("2025-06", "M"),
    ("QA01", 400401): pd.Period("2025-07", "M"),
    ("KW01", 400401): pd.Period("2025-04", "M"),
}

SAFETY_STOCK_MONTHS = 0.5
KG_PER_WEIGHT_UNIT = 118_000 / 4.2


# --------------------------------------------------------------------------
# Demand and forecast
# --------------------------------------------------------------------------

def season_factor(material: pd.Series, month: int) -> float:
    """Full seasonality for cooling water products, a quarter of it for the rest."""
    s = SEASON[month]
    return s if material.seasonal else 1 + 0.25 * (s - 1)


def underlying_demand(plant: str, mat: pd.Series, period: pd.Period, scale: float) -> float:
    months_in = (period - WARMUP[0]).n
    level = KG_PER_WEIGHT_UNIT * mat.weight * scale
    level *= 1 + 0.006 * months_in                      # about 7% underlying growth a year
    level *= season_factor(mat, period.month)

    key = (plant, mat.Material_No)
    if key in DEMAND_RAMPS:
        start, uplift = DEMAND_RAMPS[key]
        k = (period - start).n
        if k >= 0:
            level *= 1 + uplift * min(1.0, (k + 1) / 5)    # ramps up over five months
    if key in CUSTOMER_LOSS:
        k = (period - CUSTOMER_LOSS[key]).n
        if k >= 0:
            level *= max(0.42, 1 - 0.11 * (k + 1))          # volume falls away, then settles
    return level


def calibrate_plant_scales(mats: pd.DataFrame) -> dict[str, float]:
    """Scale each site's demand so its 2025 load matches its target utilisation."""
    year = pd.period_range("2025-01", "2025-12", freq="M")
    scales = {}
    for p in PLANTS.itertuples():
        avg = sum(underlying_demand(p.Plant_Code, mats.loc[m], t, 1.0)
                  for m in ASSORTMENT[p.Plant_Code] for t in year) / 12
        scales[p.Plant_Code] = p.target_load * p.Capacity_T_per_Month * 1000 / avg
    return scales


def build_forecasts(demand: dict, mats: pd.DataFrame) -> dict:
    """Trailing three month, seasonally adjusted forecast made two months ahead.
    Combinations that lost a customer keep their old run rate (a stale forecast)."""
    rng = np.random.default_rng(7)
    last_actual = ACTUALS[-1]
    forecast = {}
    for (plant, m), series in demand.items():
        mat = mats.loc[m]
        sf = lambda t: season_factor(mat, t.month)
        fc = {}
        for t in ACTUALS.append(FORWARD):
            if (plant, m) in CUSTOMER_LOSS:
                base = np.mean([series[pd.Period(x, "M")] / sf(pd.Period(x, "M"))
                                for x in ("2025-01", "2025-02", "2025-03")])
                fc[t] = base * sf(t) * 1.03
            else:
                lags = [x for x in (t - 2, t - 3, t - 4) if x in series and x <= last_actual]
                lags = lags or [last_actual, last_actual - 1, last_actual - 2]
                base = np.mean([series[x] / sf(x) for x in lags])
                fc[t] = base * sf(t) * rng.normal(1.0, 0.04)
        forecast[(plant, m)] = fc
    return forecast


# --------------------------------------------------------------------------
# Monthly MRP simulation
# --------------------------------------------------------------------------

def mrp_order(fc: dict, t: pd.Period, stock: float, lot: int) -> tuple[float, float]:
    """Net requirement rounded up to the lot size, plus the safety stock target."""
    horizon = [fc[x] for x in (t, t + 1, t + 2) if x in fc]
    safety = SAFETY_STOCK_MONTHS * float(np.mean(horizon))
    net = fc[t] + safety - stock
    return (0.0 if net <= 0 else float(np.ceil(net / lot) * lot)), safety


def simulate(seed: int) -> dict[str, pd.DataFrame]:
    rng = np.random.default_rng(seed)
    mats = MATERIALS.set_index("Material_No", drop=False)
    plants = PLANTS.set_index("Plant_Code", drop=False)
    combos = [(p, m) for p, ms in ASSORTMENT.items() for m in ms]

    scales = calibrate_plant_scales(mats)
    demand = {c: {t: max(0.0, underlying_demand(c[0], mats.loc[c[1]], t, scales[c[0]]) * rng.normal(1, 0.07))
                  for t in WARMUP.append(ACTUALS).append(FORWARD)} for c in combos}
    forecast = build_forecasts(demand, mats)

    stock = {c: SAFETY_STOCK_MONTHS * np.mean([demand[c][t] for t in WARMUP]) for c in combos}
    good_history = {c: [] for c in combos}
    freight_rate = {(a, b): (0.45 if plants.loc[a, "Country"] == plants.loc[b, "Country"] else 1.35)
                    for a in plants.index for b in plants.index if a != b}

    prod, dem, inv, trf = [], [], [], []
    for t in ACTUALS:
        planned, safety = {}, {}
        for c in combos:
            planned[c], safety[c] = mrp_order(forecast[c], t, stock[c], int(mats.loc[c[1], "lot_kg"]))
            if c in CUSTOMER_LOSS and t >= CUSTOMER_LOSS[c]:
                # a firmed schedule tied to the old contract keeps running
                firm = forecast[c][t] * (1.0 if t < pd.Period("2026-06", "M") else 0.55)
                lot = int(mats.loc[c[1], "lot_kg"])
                planned[c] = float(np.ceil(firm / lot) * lot)

        # Production is limited by blending capacity; MRP never checked it
        gross, scrap = {}, {}
        for p in plants.itertuples():
            cs = [c for c in combos if c[0] == p.Plant_Code]
            cap = p.Capacity_T_per_Month * 1000 * rng.uniform(0.95, 0.99)
            total = sum(planned[c] for c in cs)
            squeeze = min(1.0, cap / total) if total else 1.0
            for c in cs:
                attain = rng.uniform(0.95, 1.0) if planned[c] else 1.0
                g = planned[c] * (min(attain, squeeze * rng.uniform(0.98, 1.0)) if squeeze < 1 else attain)
                gross[c] = round(g)
                scrap[c] = round(g * p.scrap_rate * rng.uniform(0.8, 1.2))

        available, from_stock, short = {}, {}, {}
        for c in combos:
            available[c] = stock[c] + gross[c] - scrap[c]
            d = round(demand[c][t])
            from_stock[c] = min(d, available[c])
            short[c] = d - from_stock[c]

        # Emergency transfers: cheapest donor first, only part of a gap can be expedited
        t_in = {c: 0.0 for c in combos}
        t_out = {c: 0.0 for c in combos}
        for c in sorted(combos, key=lambda x: -short[x]):
            if short[c] <= 0:
                continue
            need = short[c] * 0.65
            donors = sorted(((q, c[1]) for q in plants.index if q != c[0] and (q, c[1]) in available),
                            key=lambda d: freight_rate[(d[0], c[0])])
            for d in donors:
                if need <= 0:
                    break
                spare = available[d] - from_stock[d] - t_out[d] - safety[d]
                qty = float(np.floor(min(spare, need) / 500) * 500)
                if qty <= 0:
                    continue
                t_out[d] += qty
                t_in[c] += qty
                need -= qty
                trf.append(dict(Month=t, From_Plant=d[0], To_Plant=c[0], Material_No=c[1],
                                Transfer_Type="Emergency", Qty_KG=qty,
                                Freight_Cost_AED=round(qty * freight_rate[(d[0], c[0])] * rng.uniform(0.9, 1.15))))

        for c in combos:
            p, m = c
            d = round(demand[c][t])
            good = gross[c] - scrap[c]
            opening = stock[c]
            closing = available[c] - from_stock[c] - t_out[c]
            stock[c] = closing
            good_history[c].append(good)
            aged = max(0.0, closing - sum(good_history[c][-6:]))  # FIFO: older than six months
            cost_kg = mats.loc[m, "Std_Cost_AED_per_KG"] * plants.loc[p, "cost_factor"] * rng.normal(1, 0.025)
            price = mats.loc[m, "List_Price_AED_per_KG"] * rng.normal(1, 0.02)
            backorder = max(0.0, d - from_stock[c] - t_in[c])

            prod.append(dict(Month=t, Plant_Code=p, Material_No=m, Planned_Qty_KG=planned[c],
                             Gross_Produced_KG=gross[c], Scrap_KG=scrap[c], Good_Produced_KG=good,
                             Actual_Cost_AED=round(gross[c] * cost_kg)))
            dem.append(dict(Month=t, Plant_Code=p, Material_No=m,
                            Forecast_Demand_KG=round(forecast[c][t]), Actual_Demand_KG=d,
                            Delivered_From_Stock_KG=from_stock[c], Delivered_Via_Transfer_KG=t_in[c],
                            Backorder_KG=backorder,
                            Net_Revenue_AED=round((from_stock[c] + t_in[c]) * price),
                            Revenue_At_Risk_AED=round(backorder * price)))
            inv.append(dict(Month=t, Plant_Code=p, Material_No=m, Opening_Stock_KG=round(opening),
                            Receipts_KG=good, Issues_KG=from_stock[c], Transfer_Out_KG=t_out[c],
                            Closing_Stock_KG=round(closing), Safety_Stock_KG=round(safety[c]),
                            Stock_Over_180_Days_KG=round(aged)))

    # Forward MRP run, unconstrained, on the stock left at the end of the actuals
    fwd, projected = [], dict(stock)
    for t in FORWARD:
        for c in combos:
            qty, ss = mrp_order(forecast[c], t, projected[c], int(mats.loc[c[1], "lot_kg"]))
            projected[c] = max(0.0, projected[c] + qty - forecast[c][t])
            fwd.append(dict(Month=t, Plant_Code=c[0], Material_No=c[1],
                            Forecast_Demand_KG=round(forecast[c][t]), MRP_Planned_Qty_KG=qty,
                            Safety_Stock_KG=round(ss)))

    plant_material = pd.DataFrame([
        dict(Plant_Code=p, Material_No=m, MRP_Type="PD", MRP_Controller=f"{p[:2]}{rng.integers(1, 4)}",
             Lot_Size_KG=int(mats.loc[m, "lot_kg"]), Planned_Delivery_Days=int(rng.choice([3, 5, 7])),
             Safety_Stock_Policy_Months=SAFETY_STOCK_MONTHS, Procurement_Type="E (in house)")
        for p, m in combos])

    frames = {"Fact_Production": prod, "Fact_Demand": dem, "Fact_Inventory": inv,
              "Fact_Transfers": trf, "Fact_MRP_Plan": fwd}
    out = {}
    for name, rows in frames.items():
        df = pd.DataFrame(rows)
        df["Month"] = df["Month"].dt.to_timestamp()
        for col in df.columns:
            if col.endswith(("_KG", "_AED")):
                df[col] = df[col].round().astype("int64")
        out[name] = df
    out["Plant_Material"] = plant_material
    return out


# --------------------------------------------------------------------------
# Dimensions, workbook and raw extract
# --------------------------------------------------------------------------

def dimensions() -> dict[str, pd.DataFrame]:
    dim_plant = PLANTS[["Plant_Code", "Plant_Name", "Country", "Capacity_T_per_Month"]].copy()
    dim_plant.insert(3, "Region", "Middle East")

    dim_material = MATERIALS.drop(columns=["weight", "lot_kg", "seasonal"]).copy()
    dim_material.insert(3, "Base_UoM", "KG")
    dim_material.insert(6, "Seasonal_Demand", np.where(MATERIALS.seasonal, "Yes", "No"))

    months = pd.date_range(ACTUALS[0].to_timestamp(), FORWARD[-1].to_timestamp(), freq="MS")
    dim_date = pd.DataFrame({"Month_Start": months})
    dim_date["Year"] = dim_date.Month_Start.dt.year
    dim_date["Quarter"] = "Q" + dim_date.Month_Start.dt.quarter.astype(str)
    dim_date["Year_Quarter"] = dim_date.Year.astype(str) + " " + dim_date.Quarter
    dim_date["Month_Num"] = dim_date.Month_Start.dt.month
    dim_date["Month_Name"] = dim_date.Month_Start.dt.strftime("%b")
    dim_date["Year_Month"] = dim_date.Month_Start.dt.strftime("%Y-%m")
    dim_date["Peak_Season"] = np.where(dim_date.Month_Num.between(5, 9), "Summer peak", "Off peak")
    dim_date["Period_Type"] = np.where(dim_date.Month_Start <= ACTUALS[-1].to_timestamp(), "Actual", "MRP plan")

    capacity = pd.DataFrame([dict(Month=m, Plant_Code=p.Plant_Code,
                                  Available_Capacity_KG=p.Capacity_T_per_Month * 1000)
                             for m in months for p in PLANTS.itertuples()])
    return {"Dim_Plant": dim_plant, "Dim_Material": dim_material,
            "Dim_Date": dim_date, "Fact_Capacity": capacity}


def write_workbook(tables: dict[str, pd.DataFrame], path: Path) -> None:
    order = ["Dim_Plant", "Dim_Material", "Dim_Date", "Plant_Material", "Fact_Production",
             "Fact_Demand", "Fact_Inventory", "Fact_Transfers", "Fact_Capacity", "Fact_MRP_Plan"]
    wb = Workbook()
    wb.remove(wb.active)
    for name in order:
        df = tables[name]
        ws = wb.create_sheet(name)
        ws.append(list(df.columns))
        for row in df.itertuples(index=False):
            ws.append([v.to_pydatetime() if isinstance(v, pd.Timestamp)
                       else (v.item() if hasattr(v, "item") else v) for v in row])
        table = Table(displayName=name, ref=f"A1:{get_column_letter(df.shape[1])}{len(df) + 1}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(table)
        for j, col in enumerate(df.columns, 1):
            letter = get_column_letter(j)
            ws.column_dimensions[letter].width = max(12, min(40, len(col) + 3))
            fmt = ("mmm yyyy" if col in ("Month", "Month_Start")
                   else "#,##0" if col.endswith(("_KG", "_AED")) else "0.00" if "per_KG" in col else None)
            if fmt:
                for cell in ws[letter][1:]:
                    cell.number_format = fmt
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
        ws.freeze_panes = "A2"
    wb.properties.title = "Synthetic SAP MRP dataset"
    wb.save(path)


def raw_order_extract(prod: pd.DataFrame, seed: int) -> pd.DataFrame:
    """Split monthly production into SAP production orders, then add the kind of
    problems a real extract has. After cleaning it reconciles to Fact_Production."""
    rng = np.random.default_rng(seed + 1)
    desc = MATERIALS.set_index("Material_No").Material_Desc

    def split(total: int, parts: int) -> np.ndarray:
        w = rng.dirichlet(np.ones(parts) * 4)
        v = np.floor(w * total).astype(int)
        v[0] += int(total) - v.sum()
        return v

    rows, order_no = [], 1_000_345_000
    for r in prod[prod.Planned_Qty_KG > 0].itertuples(index=False):
        n = int(rng.choice([1, 2, 3, 4], p=[0.25, 0.35, 0.25, 0.15]))
        planned = split(r.Planned_Qty_KG, n)
        weights = planned / planned.sum()
        gross = np.floor(weights * r.Gross_Produced_KG).astype(int)
        gross[0] += r.Gross_Produced_KG - gross.sum()
        scrap = np.floor(weights * r.Scrap_KG).astype(int)
        scrap[0] += r.Scrap_KG - scrap.sum()
        cost = np.round(weights * r.Actual_Cost_AED, 2)
        for k, day in enumerate(sorted(rng.choice(range(1, 26), n, replace=False))):
            order_no += int(rng.integers(3, 40))
            start = r.Month + pd.Timedelta(days=int(day) - 1)
            finish = min(start + pd.Timedelta(days=int(rng.integers(1, 4))), r.Month + pd.offsets.MonthEnd(0))
            rows.append(dict(AUFNR=str(order_no), AUART="ZP01", WERKS=r.Plant_Code,
                             MATNR=str(r.Material_No).zfill(18), MAKTX=desc[r.Material_No],
                             GSTRP=start, GLTRP=finish, PSMNG=float(planned[k]),
                             WEMNG=float(gross[k] - scrap[k]), XMNGA=float(scrap[k]), MEINS="KG",
                             STAT="TECO" if gross[k] >= planned[k] * 0.94 else "PDLV",
                             ISTKOSTEN=float(cost[k]),
                             ERDAT=start - pd.Timedelta(days=int(rng.integers(5, 25)))))
    raw = pd.DataFrame(rows)

    # 1. One site reports in tonnes
    tonnes = raw.WERKS == "OM01"
    raw.loc[tonnes, ["PSMNG", "WEMNG", "XMNGA"]] = (raw.loc[tonnes, ["PSMNG", "WEMNG", "XMNGA"]] / 1000).round(3)
    raw.loc[tonnes, "MEINS"] = "TO"
    # 2. Open orders with no settled cost yet
    open_orders = raw[raw.STAT == "PDLV"].sample(9, random_state=seed).index
    raw.loc[open_orders, "ISTKOSTEN"] = np.nan
    # 3. Rows duplicated by a double extract, 4. cancelled orders flagged for deletion
    duplicates = raw.sample(22, random_state=seed + 2)
    cancelled = raw.sample(15, random_state=seed + 3).copy()
    cancelled["AUFNR"] = [str(1_009_000_000 + i * 17) for i in range(len(cancelled))]
    cancelled[["WEMNG", "XMNGA", "ISTKOSTEN"]] = 0.0
    cancelled["STAT"] = "DLFL"
    raw = pd.concat([raw, duplicates, cancelled]).sort_values(["GSTRP", "WERKS", "AUFNR"]).reset_index(drop=True)
    # 5. Trailing spaces on site codes, 6. material numbers with and without leading zeros
    idx = raw.sample(frac=0.08, random_state=seed + 4).index
    raw.loc[idx, "WERKS"] = raw.loc[idx, "WERKS"] + " "
    idx = raw.sample(frac=0.12, random_state=seed + 5).index
    raw.loc[idx, "MATNR"] = raw.loc[idx, "MATNR"].str.lstrip("0")
    # 7. SAP style text dates
    for col in ("GSTRP", "GLTRP", "ERDAT"):
        raw[col] = pd.to_datetime(raw[col]).dt.strftime("%d.%m.%Y")
    return raw


# --------------------------------------------------------------------------
# Checks
# --------------------------------------------------------------------------

def check(tables: dict[str, pd.DataFrame], raw: pd.DataFrame) -> None:
    inv, dem, prod = tables["Fact_Inventory"], tables["Fact_Demand"], tables["Fact_Production"]
    flow = inv.Opening_Stock_KG + inv.Receipts_KG - inv.Issues_KG - inv.Transfer_Out_KG - inv.Closing_Stock_KG
    assert flow.abs().max() <= 2, "stock movements do not balance"
    split = dem.Actual_Demand_KG - dem.Delivered_From_Stock_KG - dem.Delivered_Via_Transfer_KG - dem.Backorder_KG
    assert split.abs().max() <= 2, "demand does not split into deliveries and backorders"

    clean = raw[raw.STAT != "DLFL"].drop_duplicates("AUFNR").copy()
    clean["WERKS"] = clean.WERKS.str.strip()
    clean["MATNR"] = clean.MATNR.astype(int)
    factor = np.where(clean.MEINS == "TO", 1000, 1)
    clean["PSMNG"] = (clean.PSMNG * factor).round()
    clean["Month"] = pd.to_datetime(clean.GSTRP, format="%d.%m.%Y").dt.to_period("M").dt.to_timestamp()
    planned = clean.groupby(["Month", "WERKS", "MATNR"]).PSMNG.sum()
    target = prod[prod.Planned_Qty_KG > 0].set_index(["Month", "Plant_Code", "Material_No"]).Planned_Qty_KG
    assert (planned.reindex(target.index) - target).abs().max() <= 1, "raw extract does not reconcile"


def summary(tables: dict[str, pd.DataFrame]) -> None:
    prod, dem, cap = tables["Fact_Production"], tables["Fact_Demand"], tables["Fact_Capacity"]
    actual_cap = cap[cap.Month <= ACTUALS[-1].to_timestamp()].groupby("Plant_Code").Available_Capacity_KG.sum()
    util = prod.groupby("Plant_Code").Gross_Produced_KG.sum() / actual_cap
    scrap = prod.groupby("Plant_Code").Scrap_KG.sum() / prod.groupby("Plant_Code").Gross_Produced_KG.sum()
    print(f"Revenue AED {dem.Net_Revenue_AED.sum() / 1e6:,.1f}M, "
          f"revenue at risk AED {dem.Revenue_At_Risk_AED.sum() / 1e6:,.2f}M, "
          f"fill rate {1 - dem.Backorder_KG.sum() / dem.Actual_Demand_KG.sum():.1%}")
    print("Utilisation by site:", ", ".join(f"{k} {v:.0%}" for k, v in util.sort_values().items()))
    print("Scrap rate by site: ", ", ".join(f"{k} {v:.1%}" for k, v in scrap.sort_values().items()))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--out", default="output", help="output folder (default: ./output)")
    parser.add_argument("--seed", type=int, default=20260927, help="random seed")
    args = parser.parse_args()

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    tables = {**simulate(args.seed), **dimensions()}
    raw = raw_order_extract(tables["Fact_Production"], args.seed)
    check(tables, raw)

    write_workbook(tables, out / "MRP_Dataset.xlsx")
    raw.to_csv(out / "SAP_Production_Orders_raw.csv", index=False)
    print(f"Wrote {out / 'MRP_Dataset.xlsx'} and {out / 'SAP_Production_Orders_raw.csv'}")
    summary(tables)


if __name__ == "__main__":
    main()
