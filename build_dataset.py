#!/usr/bin/env python3
import json, sys, pandas as pd, numpy as np
from pathlib import Path

EXPORTS = Path("exports")

def find_one(patterns):
    for p in patterns:
        files = sorted(EXPORTS.glob(p), key=lambda x: x.stat().st_mtime, reverse=True)
        if files:
            return files[0]
    return None

def read_table(path, sheet=None):
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    return pd.read_excel(path, sheet_name=sheet, engine="openpyxl")

def parse_atlas(df: pd.DataFrame) -> pd.DataFrame:
    # Required columns fallbacks
    for c in ["Picked Quantity (WHPK)","Loaded Quantity (WHPK)","Shipped Quantity (WHPK)","Order Quantity (WHPK)"]:
        if c not in df.columns: df[c] = 0
    # Dates & keys
    if "Effective Release Date" in df.columns:
        df["Effective Release Date"] = pd.to_datetime(df["Effective Release Date"], errors="coerce")
    if "Item #" in df.columns:
        df["Item #"] = df["Item #"].astype(str)
    # Metrics
    df["picked_loaded_shipped"] = df[["Picked Quantity (WHPK)","Loaded Quantity (WHPK)","Shipped Quantity (WHPK)"]].sum(axis=1)
    df["remaining_release"] = df["Order Quantity (WHPK)"].fillna(0) - df["picked_loaded_shipped"].fillna(0)
    return df

def parse_wm6(df: pd.DataFrame) -> pd.DataFrame:
    # ID + desc
    if "Product" in df.columns:
        df["product_id"] = df["Product"].astype(str)
    else:
        df["product_id"] = None
    df["desc"] = df.get("Product Description")

    # Excel date serial -> datetime (Date of Production -> Date of Production (dt))
    if "Date of Production" in df.columns:
        serial = pd.to_numeric(df["Date of Production"], errors="coerce")
        mask = serial.between(30000, 50000)  # ~1982-2070
        df["Date of Production (dt)"] = pd.NaT
        df.loc[mask, "Date of Production (dt)"] = pd.to_datetime("1899-12-30") + pd.to_timedelta(serial[mask], unit="D")
    return df

def load_rates_map(df: pd.DataFrame) -> dict:
    if df is None or df.empty: return {}
    df = df.dropna(how="all")
    # Expect SKU / Curent (as in your Rates sheet). If your column is named "Current", change it here.
    if "SKU" in df.columns and "Curent" in df.columns:
        return df.groupby("SKU")["Curent"].max().to_dict()
    return {}

def map_rate_key(desc: str | None) -> str | None:
    if not isinstance(desc, str): return None
    d = desc.lower()
    # Gallons
    if "choc" in d and "gal" in d: return "Chocolate Gallon"
    if "skim" in d and "gal" in d: return "Skim Gallon"
    if "1%" in d and "gal" in d:   return "1% Gallon"
    if "2%" in d and "gal" in d:   return "2% Gallon"
    if "whole" in d and "gal" in d:return "Whole Gallon"
    # Half gallons
    if "half" in d and "choc" in d:return "Chocolate Half Gallon"
    if "half" in d and "1%" in d:  return "1% Half Gallon"
    if "half" in d and "2%" in d:  return "2% Half Gallon"
    if "half" in d and "skim" in d:return "Skim Half Gallon"
    if "half" in d and "whole" in d:return "Whole Half Gallon"
    return None

def main():
    EXPORTS.mkdir(exist_ok=True)

    # Find latest exports (xlsx or csv names containing these keywords)
    atlas_path = find_one(["*atlas*release*.xlsx","*atlas*release*.csv","atlas_release*.xlsx","atlas*.xlsx"])
    wm6_path   = find_one(["*load*unit*table*.xlsx","*load*unit*.csv","load_unit_table*.xlsx","wm6*.xlsx"])
    rates_path = find_one(["*rate*.xlsx","*rate*.csv"])

    if not atlas_path or not wm6_path:
        sys.exit("Missing exports. Upload Atlas Release and Load Unit Table into /exports/ and push again.")

    # Read
    atlas = parse_atlas(read_table(atlas_path))
    wm6   = parse_wm6(read_table(wm6_path))
    rates_map = load_rates_map(read_table(rates_path)) if rates_path else {}

    # Aggregate WM6
    now = pd.Timestamp.utcnow().tz_localize(None)
    def _hold24h(idx):
        if "Date of Production (dt)" not in wm6.columns: return 0
        ages = (now - wm6.loc[idx, "Date of Production (dt)"])
        mask = ages.notna() & (ages.dt.total_seconds() < 24*3600)
        qoh = wm6.loc[idx, "Quantity On Hand"]
        return qoh[mask].sum() if hasattr(qoh, "sum") else 0

    agg_inv = (
        wm6.groupby(["product_id","desc"], dropna=False)
           .agg(total_inventory=("Quantity On Hand","sum"),
                quality_hold=("Quantity On Hand", lambda s: s[wm6.loc[s.index,"Material Status"].ne("Default")].sum()
                              if "Material Status" in wm6.columns else 0),
                hold_24h=("Quantity On Hand", lambda s: _hold24h(s.index)),
                available=("Quantity Available","sum"),
                allocated=("Quantity Allocated","sum"),
                reserved=("Quantity Expected Reserving","sum"))
           .reset_index()
    )
    agg_inv["available_for_picking"] = agg_inv["available"].fillna(0)

    # Atlas order summaries
    if "Effective Release Date" in atlas.columns and "Item #" in atlas.columns:
        release_dates = atlas["Effective Release Date"].dropna().sort_values().unique().tolist()[:3]
        orders = (
            atlas.groupby(["Item #","Item Description","Effective Release Date"])
                 .agg(picked_loaded_shipped=("picked_loaded_shipped","sum"),
                      remaining_release=("remaining_release","sum"))
                 .reset_index()
        )
    else:
        release_dates = []
        orders = pd.DataFrame(columns=["Item #","Item Description","Effective Release Date",
                                       "picked_loaded_shipped","remaining_release"])

    # Build rows
    rows = []
    for _, inv in agg_inv.iterrows():
        pid, desc = inv["product_id"], inv["desc"]
        row = {
            "product": pid,
            "description": desc,
            "total_inventory": float(inv["total_inventory"] or 0),
