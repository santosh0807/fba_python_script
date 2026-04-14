import pandas as pd
from pathlib import Path

# File paths (update if needed)
input_file = Path("/home/santosh/Documents/dev/python/fba/fbainput/FBA - To Send.xlsx")
output_file = Path("/home/santosh/Documents/dev/python/fba/FBA Report/fba to send - 1.5x.xlsx")
# State → FC mapping (used for sales columns)
state_to_fc = {
    'Maharashtra':     'BOM7',
    'Karnataka':       'BLR8',
    'Telangana':       'HYD3',
    'TAMIL NADU':      'MAA4',
    'Tamil Nadu':      'MAA4',
    'Uttar Pradesh':   'LKO1',
    'West Bengal':     'CCX2',
    # North East mapped to GAX1
    'Assam':           'GAX1',
    'Delhi':           'DEX3',
    'HARYANA':         'DEL4',
    'Haryana':         'DEL4',
    'Madhya Pradesh':  'IDX2',
    'Rajasthan':       'JPX2',
}

# FC / Rack Id → State mapping (for grouping stock)
fc_to_state = {
    # Karnataka
    'BLR5': 'Karnataka',
    'BLR7': 'Karnataka',
    'BLR8': 'Karnataka',
    'SBLY': 'Karnataka',
    'ZSHA': 'Karnataka',
    'FBLJ': 'Karnataka',
    'SBLL': 'Karnataka',

    # Maharashtra
    'BOM5': 'Maharashtra',
    'BOM7': 'Maharashtra',
    'NAX1': 'Maharashtra',
    'PNQ3': 'Maharashtra',
    'FBOD': 'Maharashtra',
    'FBOF': 'Maharashtra',

    # West Bengal
    'CCX1': 'West Bengal',
    'CCX2': 'West Bengal',
    'CCX4': 'West Bengal',

    # Tamil Nadu
    'CJB1': 'Tamil Nadu',
    'MAA4': 'Tamil Nadu',
    'FMAB': 'Tamil Nadu',

    # Haryana
    'DED4': 'Haryana',
    'DEL2': 'Haryana',
    'DEL4': 'Haryana',
    'DEL5': 'Haryana',
    'FDLB': 'Haryana',

    # Telangana
    'HYD3': 'Telangana',
    'HYD8': 'Telangana',
    'FHYE': 'Telangana',

    # Rajasthan
    'JPX2': 'Rajasthan',

    # Others
    'DEX3': 'Delhi',
    'GAX1': 'Assam', 
    'IDX2': 'Madhya Pradesh',
    'LKO1': 'Uttar Pradesh',
    'RNR0': 'Uttar Pradesh',   # ← change to 'Noida' if you want it separate
}
# ── Read sales / past movement sheet ──
df_sales = pd.read_excel(input_file, sheet_name="Product Send")
df_sales.columns = df_sales.columns.str.strip()

# Find state columns (case-insensitive match)
state_cols = []
for col in df_sales.columns:
    for state in state_to_fc:
        if col.strip().upper() == state.upper():
            state_cols.append(col)
            break

print("Detected state columns:", state_cols)

# Core product info - use Main Sku as key for stock matching
result = df_sales[['Main Sku', 'Sku', 'Asin', 'Title']].drop_duplicates(subset=['Main Sku'])

# ── Read Rack data sheet ──
df_rack = pd.read_excel(input_file, sheet_name="Rack data")
df_rack.columns = df_rack.columns.str.strip()

# Columns: SKU (B=1), Rack Id (D=3), Quantity (E=4)
df_rack = df_rack.iloc[:, [1, 3, 4]].copy()
df_rack.columns = ['Main Sku', 'FC', 'Quantity']

df_rack['Quantity'] = pd.to_numeric(df_rack['Quantity'], errors='coerce').fillna(0)
df_rack = df_rack[df_rack['Quantity'] > 0]

# Group by state
df_rack['State'] = df_rack['FC'].map(fc_to_state)
df_rack = df_rack[df_rack['State'].notna()]  # remove unmapped racks

df_rack['Main Sku'] = df_rack['Main Sku'].astype(str).str.strip()

# Pivot stock by STATE
stock_pivot = df_rack.pivot_table(
    index='Main Sku',
    columns='State',
    values='Quantity',
    aggfunc='sum',
    fill_value=0
).reset_index()

stock_pivot.columns = ['Main Sku'] + [f"{col} Stock" for col in stock_pivot.columns[1:]]

# ── Calculate projected sales (1.5x) per FC ──
all_fcs = sorted(set(state_to_fc.values()))

sales_by_fc = pd.DataFrame(0.0, index=result.index, columns=all_fcs)

for idx, row in df_sales.iterrows():
    main_sku = row['Main Sku']
    mask = result['Main Sku'] == main_sku
    if not mask.any():
        continue

    for state_col in state_cols:
        qty = pd.to_numeric(row[state_col], errors='coerce')
        if pd.isna(qty) or qty <= 0:
            continue

        fc = state_to_fc.get(state_col.strip())
        if fc and fc in sales_by_fc.columns:
            projected = qty * 1.5
            sales_by_fc.loc[mask, fc] += projected

sales_by_fc = sales_by_fc.round(0).astype(int)

# ── Merge everything ──
result = result.merge(stock_pivot, on='Main Sku', how='left')
result = pd.concat([result, sales_by_fc], axis=1)

# Fill missing stock with 0
stock_cols = [c for c in result.columns if 'Stock' in c]
result[stock_cols] = result[stock_cols].fillna(0).astype(int)

# Debug info
print("\nStates with stock data:", [c.replace(' Stock', '') for c in stock_cols])
print("FCs with projected sales > 0:", [fc for fc in all_fcs if result[fc].sum() > 0])

# ── Calculate "to send" = FC sales - STATE stock ──
final_cols = ['Main Sku', 'Sku', 'Asin', 'Title']

# Reverse lookup: FC → State
fc_to_state_name = {v: k for k, v in state_to_fc.items()}

for fc in all_fcs:
    sale_col   = fc
    state_name = fc_to_state_name.get(fc)
    stock_col  = f"{state_name} Stock" if state_name else None

    has_sale  = sale_col   in result.columns
    has_stock = stock_col  in result.columns if stock_col else False

    if has_sale or has_stock:
        sales = result[sale_col]   if has_sale  else 0
        stock = result[stock_col]  if has_stock else 0

        result[f"{fc} to send"] = (sales - stock).clip(lower=0).astype(int)

        if has_sale:
            final_cols.append(sale_col)
        if has_stock:
            final_cols.append(stock_col)
        final_cols.append(f"{fc} to send")

# Remove rows with zero total to send
send_cols = [c for c in result.columns if 'to send' in c]
if send_cols:
    total_to_send = result[send_cols].sum(axis=1)
    result = result[total_to_send > 0].copy()
else:
    print("Warning: No send quantities calculated")

# Final selection & sort
final_cols = [col for col in final_cols if col in result.columns]
result = result[final_cols].sort_values('Sku')

# Save
result.to_excel(output_file, index=False)

print("\nOutput saved to:", output_file)
print("Final columns include:", final_cols)
print(f"Rows with positive send quantity: {len(result)}")
print("Done.")