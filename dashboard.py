import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import plotly.express as px
from PIL import Image

# =========================
# Branding Config
# =========================
import streamlit as st

col1, col2 = st.columns([3,4])
with col1:
    st.image("kyndryl_logo.png", width=130)  # Adjust size


st.set_page_config(
    page_title="💻 Windows Upgrade Dashboard",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================
# Custom CSS for Corporate Look
# =========================
st.markdown("""
    <style>
    .main {
        background-color: #f9fafc;
        font-family: "Segoe UI", sans-serif;
    }
    h1, h2, h3 {
        color: #005baa;
    }
    [data-testid="stMetricValue"] {
        color: #005baa;
        font-weight: bold;
    }
    .stDataFrame {
        border: 1px solid #ddd;
        border-radius: 10px;
    }
    [data-testid="stSidebar"] {
        background-color: #e6f0fa;
    }
    </style>
""", unsafe_allow_html=True)


# =========================
# Load and Merge Data
# =========================
cmdb_file = "CMDB_Data.xlsx"
os_file = "OS_Version.xlsx"
wsus_file = "WSUS_Data.xlsx"

cmdb_data = pd.read_excel(cmdb_file)
os_data = pd.read_excel(os_file)
wsus_data = pd.read_excel(wsus_file)

# Normalize Asset ID column name across all 3 files
cmdb_data = cmdb_data.rename(columns={"Asset ID": "Asset ID"})
os_data = os_data.rename(columns={"Name": "Asset ID"})
wsus_data = wsus_data.rename(columns={"Computer Name": "Asset ID"})

# Normalize case for Asset IDs across all datasets
cmdb_data["Asset ID"] = cmdb_data["Asset ID"].str.upper()
os_data["Asset ID"] = os_data["Asset ID"].str.upper()
wsus_data["Asset ID"] = wsus_data["Asset ID"].str.upper()

# Merge step
merged = cmdb_data.merge(
    os_data, how="left", on="Asset ID"
).merge(
    wsus_data, how="left", on="Asset ID"
)

# Rename overlapping columns
merged = merged.rename(columns={
    "Status": "WSUS Status",
    "Status(Hardware Status)": "Hardware Status",
    "OperatingSystem": "Operating System"
})

# Replace NaN values
merged = merged.fillna({
    "WSUS Status": "Not Reporting",
    "Hardware Status": "Not Found",
    "Operating System": "Not Found",
    "LastLogonDate": "Not Found"
})

# =========================
# Normalize Operating Systems
# =========================
def normalize_os(os_name):
    if pd.isna(os_name):
        return "Not Found"
    os_name = str(os_name)

    win10_variants = [
        "Windows 10 Enterprise", "Windows 10 Enterprise 2016 LTSB",
        "Windows 10 IoT Enterprise LTSC", "Windows 10 Pro",
        "Windows 10 Pro for Workstations", "Windows 10 Pro N",
        "Windows 10 Pro N for Workstations", "Windows Workstation"
    ]
    win11_variants = [
        "Windows 11 Enterprise", "Windows 11 Pro",
        "Windows 11 Pro for Workstations", "Windows 11 Pro N"
    ]
    win8_variants = ["Windows 8 Pro", "Windows 8.1 Pro"]
    win7_variants = ["Windows 7 Professional"]
    winXP_variants = ["Windows XP Professional"]
    Cisco_variants = ["Cisco Identity Services Engine"]
    winserver_variants = [
        "Windows Server 2003", "Windows Server 2008 R2 Enterprise",
        "Windows Server 2008 R2 Standard", "Windows Server 2012 R2 Datacenter",
        "Windows Server 2012 R2 Standard", "Windows Server 2012 R2 Standard Evaluation",
        "Windows Server 2012 Standard", "Windows Server 2016 Datacenter",
        "Windows Server 2016 Standard", "Windows Server 2016 Standard Evaluation",
        "Windows Server 2019 Datacenter", "Windows Server 2019 Standard",
        "Windows Server 2019 Standard Evaluation", "Windows Server 2022 Datacenter",
        "Windows Server 2022 Datacenter Azure Edition", "Windows Server 2022 Standard",
        "Windows Server 2022 Standard Evaluation", "Windows Server 2025 Datacenter",
        "Windows Server 2025 Standard", "Windows Server® 2008 Enterprise",
        "Windows Server® 2008 Standard", "Windows Storage Server 2008 R2 Standard",
        "Windows Storage Server 2012 Standard", "Windows Storage Server 2016 Standard"
    ]    
    if os_name in win10_variants:
        return "Win10"
    elif os_name in win11_variants:
        return "Win11"
    elif os_name in win8_variants:
        return "Win8"
    elif os_name in win7_variants:
        return "Win7"
    elif os_name in winserver_variants:
        return "Win Server"        
    elif os_name in winXP_variants:
        return "WinXP"        
    elif os_name in Cisco_variants:
        return "Cisco"
    else:
        return os_name

merged["Operating System"] = merged["Operating System"].apply(normalize_os)


# =========================
# Sidebar Filters
# =========================
st.sidebar.header("🔍 Filters")

status_filter = st.sidebar.multiselect(
    "Filter by Hardware Status:",
    options=merged["Hardware Status"].dropna().unique(),
    default=["Live"]
)

ci_type_filter = st.sidebar.multiselect(
    "Filter by CI Type:",
    options=merged["CI Type"].dropna().unique(),
    default=["Workstation", "Desktop", "Laptop"]
)

asset_criteria_options = merged["Asset Criteria"].dropna().unique()
default_asset_criteria = [x for x in asset_criteria_options if x not in ["L1", "L2", "No Data"]]

asset_criteria = st.sidebar.multiselect(
    "Filter by Asset Criteria:",
    options=asset_criteria_options,
    default=default_asset_criteria
)

last_logon_filter = st.sidebar.checkbox(
    "Only show LastLogonDate within last 3 months", 
    value=True
)

company_filter = st.sidebar.multiselect(
    "Filter by Company:",
    merged["Company"].dropna().unique()
)

support_group_filter = st.sidebar.multiselect(
    "Filter by Support Group:",
    merged["Support group"].dropna().unique()
)

# Apply Filters
filtered = merged.copy()

if status_filter:
    filtered = filtered[filtered["Hardware Status"].isin(status_filter)]

if ci_type_filter:
    filtered = filtered[filtered["CI Type"].isin(ci_type_filter)]

if asset_criteria:
    filtered = filtered[filtered["Asset Criteria"].isin(asset_criteria)]

if last_logon_filter:
    cutoff_date = datetime.now() - timedelta(days=90)
    filtered = filtered[pd.to_datetime(filtered["LastLogonDate"], errors="coerce") >= cutoff_date]
    
if company_filter:
    filtered = filtered[filtered["Company"].isin(company_filter)]

if support_group_filter:
    filtered = filtered[filtered["Support group"].isin(support_group_filter)]

# Clean WSUS Status
filtered["WSUS Status"] = (
    filtered["WSUS Status"]
    .astype(str)
    .str.strip()
    .str.title()
)

# =========================
# KPI Function (Card Style)
# =========================
def kpi_card(title, count, percent=None, color="#005baa", accent="#FF6600"):
    percent_html = f"<div style='font-size:18px; color:{accent}; font-weight:500;'>({percent:.1f}%)</div>" if percent is not None else ""
    st.markdown(f"""
        <div style="
            background: white;
            padding: 20px;
            border-radius: 18px;
            box-shadow: 2px 2px 12px rgba(0,0,0,0.08);
            text-align: center;
            border-top: 5px solid {accent};
        ">
            <h4 style="color: #444; margin-bottom: 12px; font-size:16px;">{title}</h4>
            <div style="font-size: 28px; font-weight: 700; color:{color}; line-height:1.2;">
                {count}
            </div>
            {percent_html}
        </div>
    """, unsafe_allow_html=True)


# =========================
# KPIs + Chart Side by Side
# =========================
st.divider()
st.subheader("📊 Operating System Overview")

# KPI calculations
total_assets = len(filtered)
win11_devices = (filtered["Operating System"] == "Win11").sum()
win10_devices = (filtered["Operating System"] == "Win10").sum()
older_devices = filtered["Operating System"].isin(["Win8", "Win7", "WinXP"]).sum()
not_found_devices = (filtered["Operating System"] == "Not Found").sum()
other_os_devices = total_assets - (win11_devices + win10_devices + older_devices + not_found_devices)

win11_percent = (win11_devices / total_assets * 100) if total_assets > 0 else 0
win10_percent = (win10_devices / total_assets * 100) if total_assets > 0 else 0
older_percent = (older_devices / total_assets * 100) if total_assets > 0 else 0
other_os_pct = (other_os_devices / total_assets * 100) if total_assets > 0 else 0

col1, col2 = st.columns([2, 4])

with col2:
    k1, k2, k3, k4 = st.columns(4)

    with k1: 
        kpi_card("Win11 Devices", win11_devices, win11_percent, color="#005baa", accent="#005baa")  # Tata Steel Blue
    with k2: 
        kpi_card("Win10 Devices", win10_devices, win10_percent, color="#FF6600", accent="#FF6600")  # Kyndryl Orange
    with k3: 
        kpi_card("Older OS Devices", older_devices, older_percent, color="red", accent="#FF6600")
    with k4: 
        kpi_card("Other OS Devices", other_os_devices, other_os_pct, color="orange", accent="#FF6600")

    kpi_center = st.columns([1, 2, 1])
    with kpi_center[1]:
        kpi_card("Total Assets", total_assets, None, color="green", accent="#005baa")


with col1:
    os_counts = filtered["Operating System"].value_counts().reset_index()
    os_counts.columns = ["Operating System", "Count"]

    fig_donut = px.pie(
        os_counts,
        names="Operating System",
        values="Count",
        hole=0.3
    )
    st.plotly_chart(fig_donut, use_container_width=True)
st.markdown("---")
# =========================
# Non-Compliance Report
# =========================
st.subheader("⚠️ Win11 Non-Compliance Devices Overview")

# Filter non-compliant devices: all except Win11 and Other OS
non_compliant = filtered[~filtered["Operating System"].isin(["Win11", "Other OS"])]

# KPIs for non-compliance
total_nc = len(non_compliant)
win10_nc = (non_compliant["Operating System"] == "Win10").sum()
older_nc = (
    (non_compliant["Operating System"] == "Win8").sum() +
    (non_compliant["Operating System"] == "Win7").sum() +
    (non_compliant["Operating System"] == "WinXP").sum()
)
winserver_nc = (non_compliant["Operating System"] == "Win Server").sum()


k5, k6, k7, k8 = st.columns(4)
with k5: kpi_card("Total Non-Compliant", total_nc)
with k6: kpi_card("Win10", win10_nc)
with k7: kpi_card("Older OS", older_nc)
with k8: kpi_card("Win Server", winserver_nc)
st.markdown("---")
# =========================
# Charts
# =========================
#st.subheader("📊 Non Complient System Distribution")
#os_counts = filtered["Operating System"].value_counts()
#st.bar_chart(os_counts)

# =========================
# WSUS Status for Non-Compliant Devices
# =========================
st.subheader("⚠️ WSUS Status for Win10 Live Devices")

# Filter WSUS data for non-compliant live devices
wsus_nc = non_compliant[
    (non_compliant["Operating System"] == "Win10") &
    (non_compliant["Hardware Status"] == "Live")
]

col_wsus1, col_wsus2 = st.columns([4, 2])  # Left: KPIs, Right: Donut Chart

with col_wsus1:
    # KPIs for WSUS Status
    installed = (wsus_nc["WSUS Status"] == "Installed").sum()
    not_installed = (wsus_nc["WSUS Status"] == "Not Installed").sum()
    downloaded = (wsus_nc["WSUS Status"] == "Downloaded").sum()
    failed = (wsus_nc["WSUS Status"] == "Failed").sum()
    not_applicable = (wsus_nc["WSUS Status"] == "Not Applicable").sum()
    not_found = (wsus_nc["WSUS Status"] == "Not Reporting").sum()
    no_status = (wsus_nc["WSUS Status"]== "No Status").sum()

    k3, k4, k5 = st.columns(3)

#    k3.metric("Downloaded", downloaded)
#    k5.metric("Installed", installed)
#    k6.metric("No Status", no_status)


    with k3: 
        kpi_card("Downloaded", downloaded, color="#005baa", accent="#005baa")  # Tata Steel Blue
    with k4: 
        kpi_card("Installed", installed, color="#005baa", accent="#005baa")  # Tata Steel Blue
    with k5: 
        kpi_card("No Status", no_status, color="#005baa", accent="#005baa")  # Tata Steel Blue

#    k1, k2 = st.columns(2)
#    k1.metric("Not Applicable", not_applicable)
#    k2.metric("Not Installed", not_installed)
#    k4.metric("Failed", failed)

    k6, k7,k8 = st.columns(3)
    with k6: 
        kpi_card("Failed", failed, color="#005baa", accent="#005baa")  # Tata Steel Blue
    with k7: 
        kpi_card("Not Applicable", not_applicable, color="#005baa", accent="#005baa")  # Tata Steel Blue
    with k8: 
        kpi_card("Not Installed", not_installed, color="#005baa", accent="#005baa")  # Tata Steel Blue


    k7_total = st.columns([1, 2, 1])
    with k7_total[1]:
        kpi_card("Not Reporting", not_found, None, color="#005baa", accent="#005baa")
#        st.metric("Not Reporting", not_found)

with col_wsus2:
    # Donut Chart for WSUS Status
    wsus_counts = wsus_nc["WSUS Status"].value_counts().reset_index()
    wsus_counts.columns = ["WSUS Status", "Count"]

    if not wsus_counts.empty:
        fig_wsus_nc = px.pie(
            wsus_counts,
            names="WSUS Status",
            values="Count",
            hole=0.4
        )
        fig_wsus_nc.update_traces(textinfo="percent+label")
        st.plotly_chart(fig_wsus_nc, use_container_width=True)
    else:
        st.info("No WSUS status data available for non-compliant live devices.")


# =========================
# Filtered Table
# =========================
st.markdown("---")
st.subheader("📋 Filtered Asset Records")

show_cols = [
    "Asset ID", "Operating System", "WSUS Status", "LastLogonDate",
    "Owner ID", "Impact Level", "Email", "Custodian ID", "CI Type",
    "Location", "Hardware Status", "Support group", "Company", "Asset Criteria"
]

st.dataframe(filtered[show_cols])

st.download_button(
    label="📥 Download Filtered Data",
    data=filtered[show_cols].to_csv(index=False).encode("utf-8"),
    file_name="Filtered_Assets.csv",
    mime="text/csv"
)

# =========================
# Footer
# =========================
st.markdown("""
    <hr style="margin-top:40px;margin-bottom:10px;">
    <div style="text-align:center; color: grey; font-size: 14px;">
        © 2025 Kyndryl | Windows Upgrade Monitoring Dashboard
    </div>
""", unsafe_allow_html=True)



