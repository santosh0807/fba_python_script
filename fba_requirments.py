import streamlit as st
import pandas as pd

st.set_page_config(page_title="Inventory Planning Dashboard", layout="wide")

st.markdown("<h4>📦 Inventory Planning Dashboard</h4>", unsafe_allow_html=True)

# -------------------------------
# FILE PATHS
# -------------------------------
inventory_file = "PureTree_45_days_needed.xlsx"
packing_file = "Daily_packingreport.xlsx"
fba_file = "fbastock_report.xlsx"

# -------------------------------
# LOAD DATA
# -------------------------------
@st.cache_data
def load_inventory():
    return pd.read_excel(inventory_file)

@st.cache_data
def load_packing():
    return pd.read_excel(packing_file, sheet_name="Main")

@st.cache_data
def load_fba():
    return pd.read_excel(fba_file)

df = load_inventory()
packing_df = load_packing()
fba_df = load_fba()

# -------------------------------
# 🔐 LOGIN SYSTEM
# -------------------------------
def login():

    st.title("🔐 Login Required")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username == "santosh" and password == "111":
            st.session_state["logged_in"] = True
        else:
            st.error("Invalid Username or Password")

# Session check
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

# Show login if not logged in
if not st.session_state["logged_in"]:
    login()
    st.stop()

# -------------------------------
# TABS
# -------------------------------
tab1, tab2, tab3 = st.tabs(
    ["📦 Inventory", "📊 Daily Packing Report", "🏬 FBA Stock"]
)

# =========================================================
# TAB 1 → INVENTORY
# =========================================================
with tab1:

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

    # Highlight logic
    def highlight_row(row):
        if row["Final Required Qty"] > 0:
            return ["background-color: #FF4C4C"] * len(row)
        return [""] * len(row)

    st.subheader("📊 Inventory Table")

    # Total Row
    total_row = filtered_df.select_dtypes(include='number').sum()
    total_row["Sku"] = "TOTAL"
    total_row["Asin"] = ""
    total_row["Title"] = ""

    total_df = pd.DataFrame([total_row])
    final_df = pd.concat([filtered_df, total_df], ignore_index=True)

    st.dataframe(
        final_df.style.apply(highlight_row, axis=1),
        use_container_width=True,
        height=700
    )

    # Summary
    st.subheader("📈 Summary")

    total_sku = len(filtered_df)
    total_required = filtered_df["Final Required Qty"].sum()
    total_sales = filtered_df["Total Sales"].sum()

    c1, c2, c3 = st.columns(3)

    c1.metric("📦 Total SKU", total_sku)
    c2.metric("🔥 Total Sales", int(total_sales))
    c3.metric("⚠️ Required Qty", int(total_required))

    # Download
    csv = filtered_df.to_csv(index=False).encode("utf-8")

    st.download_button(
        "Download Filtered Data",
        data=csv,
        file_name="inventory_report.csv",
        mime="text/csv"
    )


# =========================================================
# TAB 2 → DAILY PACKING REPORT
# =========================================================
with tab2:

    st.subheader("📊 Daily Packing Report")

    required_cols = ["Date", "FBA", "Other Marketplaces", "Total", "Average"]
    packing_df = packing_df[required_cols]

    # ✅ FIX DATE (NO TIME)
    if "Date" in packing_df.columns:
        packing_df["Date"] = pd.to_datetime(
            packing_df["Date"],
            format="%m/%d/%Y",
            errors="coerce"
        )

        packing_df["Date"] = packing_df["Date"].dt.date

    st.dataframe(
        packing_df,
        use_container_width=True,
        height=600
    )

    # Summary
    st.subheader("📈 Packing Summary")

    total_fba = packing_df["FBA"].sum()
    total_other = packing_df["Other Marketplaces"].sum()
    total_total = packing_df["Total"].sum()

    c1, c2, c3 = st.columns(3)

    c1.metric("📦 FBA", int(total_fba))
    c2.metric("🛒 Other", int(total_other))
    c3.metric("📊 Total", int(total_total))

# =========================================================
# TAB 3 → FBA STOCK REPORT
# =========================================================
with tab3:

    st.subheader("🏬 FBA Stock Report")

    required_cols = [
        "Brand", "SKU", "Asin", "Product name", "Total",
        "Intransit", "Noida", "Karnataka", "Delhi", "Telangana",
        "Maharashtra", "Tamil Nadu", "Haryana", "West Bengal",
        "Uttar Pradesh", "Assam", "Madhya Pradesh", "Rajasthan"
    ]

    available_cols = [col for col in required_cols if col in fba_df.columns]
    fba_df = fba_df[available_cols]

    # -------------------------------
    # 🔍 TYPE SEARCH FILTERS
    # -------------------------------
    st.markdown("### 🔍 Filters")

    filtered_df = fba_df.copy()

    col1, col2 = st.columns(2)

    with col1:
        brand_filter = st.text_input("Brand", key="fba_brand")
        sku_filter = st.text_input("SKU", key="fba_sku")

    with col2:
        asin_filter = st.text_input("Asin", key="fba_asin")
        name_filter = st.text_input("Product name", key="fba_name")

    if brand_filter:
        filtered_df = filtered_df[
            filtered_df["Brand"].astype(str).str.contains(brand_filter, case=False, na=False)
        ]

    if sku_filter:
        filtered_df = filtered_df[
            filtered_df["SKU"].astype(str).str.contains(sku_filter, case=False, na=False)
        ]

    if asin_filter:
        filtered_df = filtered_df[
            filtered_df["Asin"].astype(str).str.contains(asin_filter, case=False, na=False)
        ]

    if name_filter:
        filtered_df = filtered_df[
            filtered_df["Product name"].astype(str).str.contains(name_filter, case=False, na=False)
        ]

    # -------------------------------
    # 🎨 HIGHLIGHT STOCK (RED)
    # -------------------------------
    def highlight_stock(val):
        try:
            if float(val) > 0:
                return "background-color: red; color: white"
        except:
            return ""
        return ""

    stock_cols = [
        col for col in filtered_df.columns
        if col not in ["Brand", "SKU", "Asin", "Product name"]
    ]

    styled_df = filtered_df.style.map(highlight_stock, subset=stock_cols)

    # -------------------------------
    # DISPLAY
    # -------------------------------
    st.dataframe(
        styled_df,
        use_container_width=True,
        height=600
    )

    # -------------------------------
    # SUMMARY
    # -------------------------------
    st.subheader("📈 Stock Summary")

    total_sku = len(filtered_df)
    total_stock = filtered_df["Total"].sum() if "Total" in filtered_df.columns else 0

    c1, c2 = st.columns(2)

    c1.metric("📦 Total SKU", total_sku)
    c2.metric("📊 Total Stock", int(total_stock))