import streamlit as st
import pandas as pd

st.set_page_config(page_title="Inventory Planning Dashboard", layout="wide")

st.markdown("<h4>📦 Inventory Planning Dashboard</h4>", unsafe_allow_html=True)

# -------------------------------
# AUTO LOAD FILE (NO UPLOAD)
# -------------------------------
file_path = "PureTree_45_days_needed.xlsx"  # keep file in same folder

@st.cache_data
def load_data():
    df = pd.read_excel(file_path)
    return df

df = load_data()

# -------------------------------
# FILTERS
# -------------------------------
st.subheader("🔍 Filters")

col1, col2, col3 = st.columns(3)

with col1:
    sku_filter = st.text_input("SKU")

with col2:
    asin_filter = st.text_input("ASIN")

with col3:
    title_filter = st.text_input("Title")

filtered_df = df.copy()

if sku_filter:
    filtered_df = filtered_df[
        filtered_df["Sku"].astype(str).str.contains(sku_filter, case=False)
    ]

if asin_filter:
    filtered_df = filtered_df[
        filtered_df["Asin"].astype(str).str.contains(asin_filter, case=False)
    ]

if title_filter:
    filtered_df = filtered_df[
        filtered_df["Title"].astype(str).str.lower().str.contains(title_filter.lower())
    ]

# -------------------------------
# HIGHLIGHT LOGIC
# -------------------------------
def highlight_row(row):
    if row["Final Required Qty"] > 0:
        return ["background-color: #FF4C4C"] * len(row)
    return [""] * len(row)

st.subheader("📊 Inventory Table")

# -------------------------------
# ADD TOTAL ROW
# -------------------------------
total_row = filtered_df.select_dtypes(include='number').sum()
total_row["Sku"] = "TOTAL"
total_row["Asin"] = ""
total_row["Title"] = ""

# Convert to DataFrame
total_df = pd.DataFrame([total_row])

# Combine original + total row
final_df = pd.concat([filtered_df, total_df], ignore_index=True)

st.markdown("""
<style>
thead tr th {
    position: sticky;
    top: 0;
    background-color: white;
    z-index: 2;
}
</style>
""", unsafe_allow_html=True)

st.dataframe(
    final_df.style.apply(highlight_row, axis=1),
    use_container_width=True,
    height=700  # increased height for more rows
)

# -------------------------------
# SUMMARY
# -------------------------------
st.subheader("📈 Summary")

total_sku = len(filtered_df)
total_required = filtered_df["Final Required Qty"].sum()
total_sales = filtered_df["Total Sales"].sum()

c1, c2, c3 = st.columns(3)

c1.metric("📦 Total SKU", total_sku)
c2.metric("🔥 Total Sales", int(total_sales))
c3.metric("⚠️ Required Qty", int(total_required))

# -------------------------------
# DOWNLOAD
# -------------------------------
st.subheader("⬇️ Export")

csv = filtered_df.to_csv(index=False).encode("utf-8")

st.download_button(
    "Download Filtered Data",
    data=csv,
    file_name="inventory_report.csv",
    mime="text/csv"
)