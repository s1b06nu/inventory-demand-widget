#!/usr/bin/env python3
import sys, json, pandas as pd, numpy as np
from pathlib import Path

# Usage: python build_dataset.py [path_to_excel]
excel_path = Path(sys.argv[1]) if len(sys.argv)>1 else Path('11270 Inventory and Demand.xlsx')

xl = pd.ExcelFile(excel_path, engine='openpyxl')
ars = pd.read_excel(xl, 'Atlas Release', engine='openpyxl')
lu = pd.read_excel(xl, 'Load Unit Table', engine='openpyxl')
# Rates sheet is optional but recommended
rates = pd.read_excel(xl, 'Rates', engine='openpyxl') if 'Rates' in xl.sheet_names else pd.DataFrame(columns=['SKU','Curent'])

# Dates
ars['Effective Release Date'] = pd.to_datetime(ars['Effective Release Date'], errors='coerce')
serial = pd.to_numeric(lu['Date of Production'], errors='coerce')
mask = serial.between(30000, 50000)
lu['Date of Production (dt)'] = pd.NaT
lu.loc[mask, 'Date of Production (dt)'] = pd.to_datetime('1899-12-30') + pd.to_timedelta(serial[mask], unit='D')
now = pd.Timestamp.utcnow().tz_localize(None)

# Rates map
rates = rates.dropna(how='all')
rates_map = rates.dropna().groupby('SKU')['Curent'].max().to_dict()

# Aggregate inventory by product
lu['product_id'] = lu['Product'].astype(str)
lu['desc'] = lu['Product Description']
agg_inv = lu.groupby(['product_id','desc'], dropna=False).agg(
    total_inventory=('Quantity On Hand','sum'),
    quality_hold=('Quantity On Hand', lambda s: s[lu.loc[s.index, 'Material Status'].ne('Default')].sum()),
    hold_24h=('Quantity On Hand', lambda s: s[(now - lu.loc[s.index, 'Date of Production (dt)']).dt.total_seconds() < 24*3600].sum() if 'Date of Production (dt)' in lu.columns else 0),
    available=('Quantity Available','sum'),
    allocated=('Quantity Allocated','sum'),
    reserved=('Quantity Expected Reserving','sum')
).reset_index()
agg_inv['available_for_picking'] = agg_inv['available']

# Release dates (first 3 by date)
release_dates = ars['Effective Release Date'].dropna().sort_values().unique().tolist()[:3]
ars['picked_loaded_shipped'] = ars[['Picked Quantity (WHPK)','Loaded Quantity (WHPK)','Shipped Quantity (WHPK)']].sum(axis=1)
ars['remaining_release'] = ars['Order Quantity (WHPK)'].fillna(0) - ars['picked_loaded_shipped'].fillna(0)
ars['Item #'] = ars['Item #'].astype(str)
orders = ars.groupby(['Item #','Item Description','Effective Release Date']).agg(
    picked_loaded_shipped=('picked_loaded_shipped','sum'),
    remaining_release=('remaining_release','sum')
).reset_index()

# Map description → rates key
RATE_KEYS = [
    'Whole Gallon','2% Gallon','1% Gallon','Skim Gallon','Chocolate Gallon',
    'Whole Half Gallon','2% Half Gallon','1% Half Gallon','Skim Half Gallon','Chocolate Half Gallon'
]

def map_rate_key(desc:str|None):
    if not isinstance(desc, str):
        return None
    d = desc.lower()
    if 'half' in d and 'choc' in d: return 'Chocolate Half Gallon'
    if 'half' in d and '1%' in d: return '1% Half Gallon'
    if 'half' in d and '2%' in d: return '2% Half Gallon'
    if 'half' in d and 'skim' in d: return 'Skim Half Gallon'
    if 'half' in d and 'whole' in d: return 'Whole Half Gallon'
    if 'choc' in d and 'gal' in d: return 'Chocolate Gallon'
    if 'skim' in d and 'gal' in d: return 'Skim Gallon'
    if '1%' in d and 'gal' in d: return '1% Gallon'
    if '2%' in d and 'gal' in d: return '2% Gallon'
    if 'whole' in d and 'gal' in d: return 'Whole Gallon'
    return None

rows = []
for _, inv_row in agg_inv.iterrows():
    pid = inv_row['product_id']
    desc = inv_row['desc']
    row = {
        'product': pid,
        'description': desc,
        'total_inventory': float(inv_row['total_inventory'] or 0),
        'quality_hold': float(inv_row['quality_hold'] or 0),
        'hold_24h': float(inv_row['hold_24h'] or 0),
        'available_inventory': float(inv_row['available'] or 0),
        'allocated_inventory': float(inv_row['allocated'] or 0),
        'reserved_inventory': float(inv_row['reserved'] or 0),
        'available_for_picking': float(inv_row['available_for_picking'] or 0)
    }
    rate_key = map_rate_key(desc)
    rate = rates_map.get(rate_key)
    for i, d in enumerate(release_dates, start=1):
        sub = orders[(orders['Item #']==pid) & (orders['Effective Release Date']==d)]
        pls = float(sub['picked_loaded_shipped'].sum()) if not sub.empty else 0.0
        rem = float(sub['remaining_release'].sum()) if not sub.empty else 0.0
        rem_inv = row['available_for_picking'] - rem
        htp = (rem / rate) if rate and rate>0 else None
        row[f'release_{i}_date'] = d.strftime('%Y-%m-%d')
        row[f'release_{i}_pls'] = pls
        row[f'release_{i}_remaining'] = rem
        row[f'release_{i}_remaining_inventory'] = rem_inv
        row[f'release_{i}_hours_to_produce'] = htp
    rows.append(row)

payload = {
    'generated_utc': pd.Timestamp.utcnow().tz_localize(None).isoformat(),
    'release_dates': [d.strftime('%Y-%m-%d') for d in release_dates],
    'rows': rows
}

Path('widget_dataset.json').write_text(json.dumps(payload, indent=2))
print(f"Wrote widget_dataset.json with {len(rows)} rows and {len(release_dates)} releases.")
