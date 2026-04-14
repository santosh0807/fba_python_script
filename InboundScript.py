import pandas as pd
import sys

# -------------------------------------------------
# CONFIG: set the years you want (or empty list for all years)
# Examples:
#   YEARS = [2025]          # only 2025
#   YEARS = [2026]          # only 2026
#   YEARS = [2025, 2026]    # both 2025 and 2026
#   YEARS = []              # all years available in the data
# -------------------------------------------------
YEARS = [2025, 2026]       # ← change here

# -------------------------------------------------
# 1. MAIN INBOUND SHEET (Planned shipments)
# -------------------------------------------------
url_main = "https://docs.google.com/spreadsheets/d/1KDQJM6ZIVZCkFoylfoXqsbead93zwjzCc61wHGf2GSM/export?format=csv&gid=220119882"
df_main = pd.read_csv(url_main)

# Columns we need: B=Shipment Date, C=Sku, E=Asin, F=Fnsku, H=Qty, I=Shipment id
df_main = df_main.iloc[:, [1, 2, 4, 5, 7, 8]]
df_main.columns = ['Shipment Date', 'Sku', 'Asin', 'Fnsku', 'Qty', 'Shipment id']

# Clean basic columns
df_main['Sku'] = df_main['Sku'].astype(str).str.strip()
df_main['Shipment id'] = df_main['Shipment id'].astype(str).str.strip()
df_main['Qty'] = pd.to_numeric(df_main['Qty'], errors='coerce').fillna(0).astype(int)

# Parse date (DD-MM-YYYY)
df_main['Shipment Date'] = pd.to_datetime(df_main['Shipment Date'],
                                          format='%d-%m-%Y',
                                          errors='coerce')

# Drop rows with invalid dates
df_main = df_main.dropna(subset=['Shipment Date'])

print(f"Total rows after date parsing: {len(df_main)}")
print(f"Date range: {df_main['Shipment Date'].min().date()} → {df_main['Shipment Date'].max().date()}")

# Filter by year(s) if specified
if YEARS:
    before = len(df_main)
    df_main = df_main[df_main['Shipment Date'].dt.year.isin(YEARS)]
    print(f"Rows after filtering to years {YEARS}: {len(df_main)} (from {before})")
else:
    print("No year filter applied → using all available data")

if df_main.empty:
    print("No shipments found for the selected year(s).")
    sys.exit()

# Group planned shipments
main_grouped = (df_main
    .groupby(['Shipment id', 'Sku'], as_index=False)
    .agg({
        'Shipment Date': 'first',
        'Asin': 'first',
        'Fnsku': 'first',
        'Qty': 'sum'
    })
    .rename(columns={'Qty': 'Total Qty'})
)

# -------------------------------------------------
# 2. RECEIVED DATA SHEET
# -------------------------------------------------
url_inv = "https://docs.google.com/spreadsheets/d/1KDQJM6ZIVZCkFoylfoXqsbead93zwjzCc61wHGf2GSM/export?format=csv&gid=152394457"
df_inv = pd.read_csv(url_inv)

df_inv = df_inv.iloc[:, [3, 6, 7, 8]]
df_inv.columns = ['Sku', 'Shipment id', 'Received Qty', 'FC']

df_inv['Sku'] = df_inv['Sku'].astype(str).str.strip()
df_inv['Shipment id'] = df_inv['Shipment id'].astype(str).str.strip()
df_inv['FC'] = df_inv['FC'].astype(str).str.strip()
df_inv['Received Qty'] = pd.to_numeric(df_inv['Received Qty'], errors='coerce').fillna(0).astype(int)

# -------------------------------------------------
# 3–6. Rest of the logic (unchanged)
# -------------------------------------------------
received_summary = (df_inv
    .groupby(['Shipment id', 'Sku'], as_index=False)
    .agg(
        Received_Qty=('Received Qty', 'sum'),
        Received_At_FC=('FC', lambda x: ', '.join(sorted(set(x.dropna()))))
    )
)

final = main_grouped.merge(received_summary, on=['Shipment id', 'Sku'], how='left')

final['Received_Qty'] = final['Received_Qty'].fillna(0).astype(int)
final['Received_At_FC'] = final['Received_At_FC'].fillna('— Not Received')

final['Shortage'] = final['Total Qty'] - final['Received_Qty']
final['Status'] = final['Shortage'].apply(lambda x: 'Complete' if x <= 0 else 'Short')

# Overall SKU match
sku_totals = final.groupby('Sku', as_index=False)[['Total Qty', 'Received_Qty']].sum()
sku_totals['Match'] = sku_totals['Total Qty'] == sku_totals['Received_Qty']
match_dict = dict(zip(sku_totals['Sku'], sku_totals['Match']))
final['Overall SKU Match'] = final['Sku'].map(match_dict).fillna(False)

# Final column order
final = final[[
    'Shipment Date', 'Sku', 'Asin', 'Fnsku',
    'Total Qty', 'Received_Qty', 'Received_At_FC',
    'Shortage', 'Status', 'Shipment id', 'Overall SKU Match'
]]

# Save file
if YEARS:
    year_suffix = "_" + "_".join(map(str, sorted(YEARS)))
else:
    year_suffix = "_All"
output_file = f"/home/santosh/Documents/Python/Scripts/FBA/FBA Report/Inbound{year_suffix}.xlsx"

final.to_excel(output_file, index=False, sheet_name='Inbound')

print("\nSUCCESS! Report generated")
print(f"Years filtered: {YEARS if YEARS else 'All'}")
print(f"Rows           : {len(final)}")
print(f"Unique SKUs    : {final['Sku'].nunique()}")
print(f"Perfect matches: {final['Overall SKU Match'].sum()}")
print(f"File → {output_file}")