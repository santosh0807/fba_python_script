import pandas as pd
from urllib.parse import quote
from datetime import datetime

# ==============================
# GOOGLE SHEET CONFIG
# ==============================
sheet_id = "1KDQJM6ZIVZCkFoylfoXqsbead93zwjzCc61wHGf2GSM"
sheet_name = quote("Inventory Details")

url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"

# ==============================
# LOAD DATA
# ==============================
df = pd.read_csv(url)
df.columns = df.columns.str.strip()

# ==============================
# SELECT REQUIRED COLUMNS
# ==============================
df = df.iloc[:, [0, 2, 3, 5, 7, 8]]
df.columns = ["Date", "ASIN", "MSKU", "Event Type", "Quantity", "Fulfillment Center"]

# ==============================
# FIX DATE (MIXED FORMAT)
# ==============================
df["Date"] = pd.to_datetime(df["Date"], errors='coerce')

mask = df["Date"].isna()
df.loc[mask, "Date"] = pd.to_datetime(
    pd.to_numeric(df.loc[mask, "Date"], errors='coerce'),
    origin='1899-12-30',
    unit='D',
    errors='coerce'
)

# Convert Quantity
df["Quantity"] = pd.to_numeric(df["Quantity"], errors='coerce')

# ==============================
# BASE (ALL SKUs)
# ==============================
df_base = df[["MSKU", "ASIN", "Fulfillment Center"]].drop_duplicates()

# ==============================
# SALES (OUTWARD)
# ==============================
df_sales = df[df["Quantity"] < 0]

df_sales_summary = df_sales.groupby(
    ["MSKU", "ASIN", "Fulfillment Center"],
    as_index=False
)["Quantity"].sum()

df_sales_summary["Total Sold"] = df_sales_summary["Quantity"].abs()
df_sales_summary.drop(columns=["Quantity"], inplace=True)

# ==============================
# INWARD (TOTAL RECEIVED)
# ==============================
df_inward = df[df["Quantity"] > 0]

df_inward_summary = df_inward.groupby(
    ["MSKU", "ASIN", "Fulfillment Center"],
    as_index=False
)["Quantity"].sum()

df_inward_summary.rename(columns={"Quantity": "Total Inward"}, inplace=True)

# ==============================
# LAST INWARD DATE
# ==============================
df_receipts_summary = df_inward.groupby(
    ["MSKU", "ASIN", "Fulfillment Center"],
    as_index=False
)["Date"].max()

df_receipts_summary.rename(columns={"Date": "Last Inward Date"}, inplace=True)

# ==============================
# MERGE ALL
# ==============================
df_final = pd.merge(df_base, df_sales_summary,
                    on=["MSKU", "ASIN", "Fulfillment Center"], how="left")

df_final = pd.merge(df_final, df_inward_summary,
                    on=["MSKU", "ASIN", "Fulfillment Center"], how="left")

df_final = pd.merge(df_final, df_receipts_summary,
                    on=["MSKU", "ASIN", "Fulfillment Center"], how="left")

# Fill missing values
df_final["Total Sold"] = df_final["Total Sold"].fillna(0)
df_final["Total Inward"] = df_final["Total Inward"].fillna(0)

# ==============================
# STOCK LEFT
# ==============================
df_final["Stock Left"] = df_final["Total Inward"] - df_final["Total Sold"]

# ==============================
# DAYS SINCE INWARD
# ==============================
today = pd.to_datetime(datetime.today().date())

df_final["Days Since Inward"] = (today - df_final["Last Inward Date"]).dt.days

# ==============================
# RECALL LOGIC (UPDATED)
# ==============================
def get_decision(row):
    if row["Stock Left"] <= 0:
        return "No Stock"

    if pd.isna(row["Last Inward Date"]):
        return "No Inward Data"

    if row["Stock Left"] > 0 and row["Total Sold"] == 0 and row["Days Since Inward"] > 15:
        return "🔴 RECALL"

    elif row["Stock Left"] > 0 and row["Total Sold"] <= 5 and row["Days Since Inward"] > 30:
        return "🔴 RECALL"

    elif row["Stock Left"] > 0 and row["Total Sold"] <= 5:
        return "🟡 MONITOR"

    else:
        return "🟢 HEALTHY"

df_final["Decision"] = df_final.apply(get_decision, axis=1)

# ==============================
# SORT (IMPORTANT PRIORITY)
# ==============================
df_final = df_final.sort_values(
    by=["Decision", "Stock Left", "Days Since Inward"],
    ascending=[True, False, False]
)

# ==============================
# SAVE
# ==============================
df_final.to_excel("Smart_Recall_With_Stock.xlsx", index=False)

print("✅ Done: Smart_Recall_With_Stock.xlsx created")