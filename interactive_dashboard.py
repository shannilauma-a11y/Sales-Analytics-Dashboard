import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --------------------------------
# 1. PAGE SETTINGS
# --------------------------------

st.set_page_config(
    page_title="Sales Analytics Dashboard",
    page_icon="📊",
    layout="wide"
)

# --------------------------------
# 2. BLACK AND GOLD DESIGN
# --------------------------------

st.markdown("""
<style>
.stApp {
    background-color: #101010;
    color: white;
}

[data-testid="stSidebar"] {
    background-color: #1C1C1C;
}

h1, h2, h3 {
    color: #D4AF37 !important;
}

div[data-testid="stMetric"] {
    background-color: #222222;
    border: 1px solid #D4AF37;
    padding: 18px;
    border-radius: 12px;
}

div[data-testid="stMetricLabel"] {
    color: #D4AF37;
}

div[data-testid="stMetricValue"] {
    color: white;
}

.stDownloadButton button {
    background-color: #D4AF37;
    color: black;
    font-weight: bold;
    border-radius: 8px;
}
</style>
""", unsafe_allow_html=True)

# --------------------------------
# 3. DASHBOARD TITLE
# --------------------------------

st.title("SALES PERFORMANCE DASHBOARD")
st.write("Interactive Business Sales Analysis")

# --------------------------------
# 4. LOAD EXCEL DATA
# --------------------------------

try:
    df = pd.read_excel(
        "Sales Inventory Analytics.xlsx",
        sheet_name="Orders"
    )
except Exception as error:
    st.error(f"Could not load the sales workbook: {error}")
    st.stop()

# --------------------------------
# 5. SIDEBAR FILTERS
# --------------------------------

st.sidebar.header("DASHBOARD FILTERS")

regions = ["All"] + sorted(
    df["Region"].dropna().unique().tolist()
)

segments = ["All"] + sorted(
    df["Segment"].dropna().unique().tolist()
)

selected_region = st.sidebar.selectbox(
    "Select Region",
    regions
)

selected_segment = st.sidebar.selectbox(
    "Select Customer Segment",
    segments
)

# --------------------------------
# 6. FILTER SALES DATA
# --------------------------------

filtered = df.copy()

if selected_region != "All":
    filtered = filtered[
        filtered["Region"] == selected_region
    ]

if selected_segment != "All":
    filtered = filtered[
        filtered["Segment"] == selected_segment
    ]

if filtered.empty:
    st.warning("No sales records match these filters.")
    st.stop()

# --------------------------------
# 7. CALCULATE BUSINESS KPIs
# --------------------------------

revenue = filtered["Sales"].sum()
profit = filtered["Profit"].sum()
cost = revenue - profit

margin = (
    profit / revenue * 100
    if revenue != 0 else 0
)

orders = filtered["Order ID"].nunique()

# --------------------------------
# 8. DISPLAY KPI CARDS
# --------------------------------

st.subheader("BUSINESS PERFORMANCE")

col1, col2, col3, col4, col5 = st.columns(5)

col1.metric("Revenue", f"{revenue:,.2f}")
col2.metric("Profit", f"{profit:,.2f}")
col3.metric("Cost", f"{cost:,.2f}")
col4.metric("Profit Margin", f"{margin:.2f}%")
col5.metric("Unique Orders", f"{orders:,}")

# --------------------------------
# 9. REVENUE BY REGION
# --------------------------------

st.subheader("Revenue by Region")

region_sales = filtered.groupby("Region")["Sales"].sum()

st.bar_chart(region_sales, color="#D4AF37")

# --------------------------------
# 10. PROFIT BY REGION
# --------------------------------

st.subheader("Profit by Region")

region_profit = filtered.groupby("Region")["Profit"].sum()

st.bar_chart(region_profit, color="#D4AF37")

# --------------------------------
# 11. REVENUE BY CUSTOMER SEGMENT
# --------------------------------

st.subheader("Revenue by Customer Segment")

segment_sales = filtered.groupby("Segment")["Sales"].sum()

st.bar_chart(segment_sales, color="#D4AF37")

# --------------------------------
# 12. MONTHLY SALES TREND
# --------------------------------

st.subheader("Monthly Sales Trend")

filtered["Order Date"] = pd.to_datetime(
    filtered["Order Date"],
    errors="coerce"
)

monthly = filtered.dropna(
    subset=["Order Date"]
).groupby(
    filtered["Order Date"].dt.to_period("M")
)["Sales"].sum()

monthly.index = monthly.index.astype(str)

st.line_chart(monthly, color="#D4AF37")

# --------------------------------
# 13. CREATE FORMATTED EXCEL REPORT
# --------------------------------

output = BytesIO()

with pd.ExcelWriter(output, engine="openpyxl") as writer:

    filtered.to_excel(
        writer,
        index=False,
        sheet_name="Sales Report"
    )

    worksheet = writer.sheets["Sales Report"]

    # Gold column headings
    for cell in worksheet[1]:
        cell.fill = PatternFill(
            start_color="D4AF37",
            end_color="D4AF37",
            fill_type="solid"
        )

        cell.font = Font(
            bold=True,
            color="000000"
        )

        cell.alignment = Alignment(
            horizontal="center"
        )

    # Adjust column widths
    for column in worksheet.columns:
        max_length = max(
            len(str(cell.value or ""))
            for cell in column
        )

        worksheet.column_dimensions[
            column[0].column_letter
        ].width = min(max_length + 3, 40)

    # Add borders
    thin_border = Border(
        bottom=Side(style="thin", color="D4AF37")
    )

    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border

    # Keep headings visible
    worksheet.freeze_panes = "A2"

excel_file = output.getvalue()

# --------------------------------
# 14. DOWNLOAD EXCEL REPORT
# --------------------------------

st.subheader("DOWNLOAD YOUR SALES REPORT")

st.download_button(
    label="Download Excel Sales Report",
    data=excel_file,
    file_name="Filtered_Sales_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# --------------------------------
# 15. FOOTER
# --------------------------------

st.write("---")
st.write("Created using Python, Pandas, OpenPyXL and Streamlit")
st.caption("Sales Analytics Project | Business Intelligence")