import subprocess
import sys

def install_packages():
    required = {'streamlit-echarts', 'streamlit-shadcn-ui', 'openpyxl', 'pandas'}
    import pkg_resources
    installed = {pkg.key for pkg in pkg_resources.working_set}
    missing = required - installed
    
    if missing:
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])

try:
    install_packages()
except Exception:
    pass

"""
GCC Replenishment Engine - Elite Edition
=========================================
Advanced Streamlit application featuring Apache ECharts live animations,
Shadcn UI widgets, and automated Oracle allocation pipelines.
"""

import time
from io import BytesIO

import pandas as pd
import streamlit as st
from streamlit_echarts import st_echarts
import streamlit_shadcn_ui as ui
from openpyxl.styles import Font

# ---------------------------------------------------------------------------
# CONSTANTS & CONFIGURATION
# ---------------------------------------------------------------------------
PACK_SIZE = 12                        
WMS_RESERVE_PCT = 0.10                
FROM_LOCATION_NAME = "DIP Warehouse"  

GRADE_CUT_A_PLUS = 0.50
GRADE_CUT_A = 0.70
TOP_GROUP_PCT = 0.80
PROTECTED_GRADES = {"A+", "A", "B+"}  

REQUIRED_COLUMNS = [
    "Option", "Store ID", "Total Sold Qty", "SOH",
    "In-Transit", "Yet to Dispatch", "Target Stock", "DC Available Inventory",
]

NUMERIC_COLUMNS = [
    "Total Sold Qty", "SOH", "In-Transit", "Yet to Dispatch",
    "Target Stock", "DC Available Inventory",
]

ID_COLUMNS = ["Option", "Store ID"]

st.set_page_config(page_title="GCC Replenishment Engine", layout="wide")

# ---------------------------------------------------------------------------
# ADVANCED STYLING & KEYFRAME INJECTIONS
# ---------------------------------------------------------------------------
def inject_elite_theme():
    st.html("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');

    :root {
        --bg-main: #090D16;
        --surface: #111827;
        --surface-border: #1F2937;
        --text-primary: #F3F4F6;
        --text-secondary: #9CA3AF;
        --accent: #3B82F6;
    }

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
        background-color: var(--bg-main);
        color: var(--text-primary);
    }
    
    .main .block-container { padding-top: 2rem; max-width: 1280px; }

    @keyframes fadeInSlide {
        0% { opacity: 0; transform: translateY(20px); }
        100% { opacity: 1; transform: translateY(0); }
    }
    
    .animated-wrapper {
        animation: fadeInSlide 0.8s cubic-bezier(0.16, 1, 0.3, 1) forwards;
    }

    .hero-banner {
        background: linear-gradient(135deg, #111827 0%, #1F2937 100%);
        border: 1px solid var(--surface-border);
        border-radius: 20px;
        padding: 40px;
        margin-bottom: 30px;
        box-shadow: 0 20px 40px -15px rgba(0, 0, 0, 0.5);
        position: relative;
        overflow: hidden;
    }
    .hero-banner::after {
        content: '';
        position: absolute;
        top: 0; right: 0; width: 300px; height: 100%;
        background: radial-gradient(circle, rgba(59,130,246,0.1) 0%, transparent 70%);
        pointer-events: none;
    }
    .hero-title { font-size: 2.5rem; font-weight: 800; color: #FFFFFF; margin: 0; letter-spacing: -0.5px; }
    .hero-subtitle { font-size: 1rem; color: var(--text-secondary); margin-top: 8px; font-weight: 400; line-height: 1.5; }
    
    .stDataFrame { border-radius: 12px; border: 1px solid var(--surface-border); overflow: hidden; }
    </style>
    """)

inject_elite_theme()

# ---------------------------------------------------------------------------
# CORE BUSINESS LOGIC
# ---------------------------------------------------------------------------
def round_to_pack(qty, pack_size: int = PACK_SIZE) -> int:
    if pd.isna(qty) or qty <= 0: return 0
    remainder = qty % pack_size
    if remainder == 0: return int(qty)
    if remainder < pack_size / 2: return int(qty - remainder)
    return int(qty - remainder + pack_size)

def clean_id_column(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

def merge_uploaded_files(files_data):
    warnings, merged = [], None
    for filename, df in files_data:
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        if "Option" not in df.columns:
            warnings.append(f"'{filename}' has no 'Option' column.")
            continue
        join_keys = ["Option", "Store ID"] if "Store ID" in df.columns else ["Option"]
        df = df.drop_duplicates(subset=join_keys)
        if merged is None:
            merged = df
            continue
        merged = merged.merge(df, on=[k for k in join_keys if k in merged.columns], how="outer")
    if merged is None:
        raise ValueError("None of the uploaded files contained a valid 'Option' column.")
    return merged, warnings

def prepare_raw_data(raw_df: pd.DataFrame):
    df = raw_df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))
    df = df.dropna(subset=["Option", "Store ID"])
    for col in ID_COLUMNS: df[col] = clean_id_column(df[col])
    for col in NUMERIC_COLUMNS: df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df, []

def calculate_store_metrics(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Total Pipeline"] = df["SOH"] + df["In-Transit"] + df["Yet to Dispatch"]
    df["Base Need"] = df["Target Stock"].where(df["Total Sold Qty"] > 0, 0)
    df["Overstock Failsafe"] = df["Target Stock"] - df["Total Pipeline"]
    df["Raw Need"] = df[["Base Need", "Overstock Failsafe"]].min(axis=1)
    df.loc[df["Total Pipeline"] <= 0, "Raw Need"] = df["Target Stock"]
    df["Rounded Need"] = df["Raw Need"].apply(round_to_pack)
    return df

def assign_grades(group: pd.DataFrame) -> pd.DataFrame:
    total_sold = group["Total Sold Qty"].sum()
    if total_sold <= 0:
        group["Cumulative Sales %"], group["Grade"] = 0.0, "B"
        return group
    cum_pct = group["Total Sold Qty"].cumsum() / total_sold
    group["Cumulative Sales %"] = cum_pct
    def grade_for(pct):
        if pct <= GRADE_CUT_A_PLUS: return "A+"
        elif pct <= GRADE_CUT_A: return "A"
        elif pct <= TOP_GROUP_PCT: return "B+"
        else: return "B"
    group["Grade"] = cum_pct.apply(grade_for)
    return group

def allocate_option_group(group: pd.DataFrame) -> pd.DataFrame:
    group = group.sort_values("Total Sold Qty", ascending=False, kind="mergesort").reset_index(drop=True)
    group = assign_grades(group)
    dc_full = group["DC Available Inventory"].iloc[0]
    dc_reserved = round(dc_full * (1 - WMS_RESERVE_PCT))
    protected = group[group["Grade"].isin(PROTECTED_GRADES)].copy()
    b_grade = group[group["Grade"] == "B"].copy()
    total_need = protected["Rounded Need"].sum() + (b_grade["Rounded Need"].sum() if len(b_grade) else 0)
    pool = dc_reserved if dc_reserved >= total_need else dc_full
    
    if pool >= total_need:
        protected["Allocated Qty"] = protected["Rounded Need"]
        if len(b_grade) > 0: b_grade["Allocated Qty"] = b_grade["Rounded Need"]
    else:
        allocations, remaining, exhausted = [], pool, False
        for need in protected["Rounded Need"]:
            if not exhausted and remaining >= need:
                allocations.append(need); remaining -= need
            else:
                exhausted = True; allocations.append(0)
        protected["Allocated Qty"] = allocations
        if len(b_grade) > 0: b_grade["Allocated Qty"] = 0
    
    result = pd.concat([protected, b_grade], ignore_index=True)
    result["New WMS Inventory (90%)"] = dc_reserved
    result["Reserve Released"] = dc_reserved < total_need
    return result

def run_replenishment_engine(raw_df: pd.DataFrame):
    clean_df, warnings = prepare_raw_data(raw_df)
    metrics_df = calculate_store_metrics(clean_df)
    allocated_frames = [allocate_option_group(group) for _, group in metrics_df.groupby("Option", sort=False)]
    detail_df = pd.concat(allocated_frames, ignore_index=True).sort_values(["Option", "Store ID"]).reset_index(drop=True)
    
    triggered = detail_df[detail_df["Allocated Qty"] > 0].copy()
    triggered["FROM LOCATION"] = FROM_LOCATION_NAME
    triggered["TO LOCATION"] = triggered["Store ID"]
    triggered["ITEM"] = triggered["Option"]
    triggered["QUANTITY"] = triggered["Allocated Qty"].astype(int)
    final_df = triggered[["FROM LOCATION", "TO LOCATION", "ITEM", "QUANTITY"]].reset_index(drop=True)
    return final_df, detail_df, warnings

def convert_df_to_excel(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        for cell in ws[1]: cell.font = Font(bold=True)
    return output.getvalue()

# ---------------------------------------------------------------------------
# APPLICATION UI LAYOUT
# ---------------------------------------------------------------------------
st.html("""
<div class="hero-banner animated-wrapper">
    <div class="hero-title">GCC Replenishment Engine</div>
    <div class="hero-subtitle">High-Performance Inventory Allocation & WMS Reserve Optimization Matrix. Seamlessly routing from DIP Warehouse to regional store networks.</div>
</div>
""")

if "final_df" not in st.session_state:
    st.session_state.final_df = None
    st.session_state.detail_df = None

# Step 1: File Upload Section
st.markdown("### 📥 Step 1: Data Ingestion Pipeline")
uploaded_files = st.file_uploader("Upload pipeline and inventory reports (.xlsx)", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    files_data = [(f.name, pd.read_excel(f)) for f in uploaded_files]
    if st.button("Execute Automated Allocation Protocol", type="primary", use_container_width=True):
        progress_bar = st.progress(0, text="Initializing Secure Engine Context...")
        time.sleep(0.2)
        progress_bar.progress(30, text="Validating Pipeline Matrices & Stock Profiles...")
        time.sleep(0.3)
        progress_bar.progress(70, text="Partitioning 10% WMS Reserve & Executing Grades...")
        
        try:
            raw_df, _ = merge_uploaded_files(files_data)
            final_df, detail_df, _ = run_replenishment_engine(raw_df)
            st.session_state.final_df = final_df
            st.session_state.detail_df = detail_df
            progress_bar.progress(100, text="Computation Completed Successfully.")
            time.sleep(0.3)
            progress_bar.empty()
        except ValueError as e:
            progress_bar.empty()
            st.error(str(e))

# Step 2: Interactive Analytics Dashboard & Animated KPIs
if st.session_state.final_df is not None:
    final_df = st.session_state.final_df
    detail_df = st.session_state.detail_df

    st.markdown("---")
    st.markdown("### 📊 Step 2: Executive Summary & Live Analytics")

    # Shadcn Styled Metric Cards Grid
    col1, col2, col3 = st.columns(3)
    with col1:
        ui.metric_card(title="Total Units Allocated", content=f"{int(final_df['QUANTITY'].sum()):,}", description="Optimized transfer size", key="m1")
    with col2:
        ui.metric_card(title="Generated Transfer Lines", content=f"{len(final_df):,}", description="Oracle routing nodes", key="m2")
    with col3:
        ui.metric_card(title="Unique Items Processed", content=f"{final_df['ITEM'].nunique():,}", description="SKU style count", key="m3")

    st.markdown("<br>", unsafe_allow_html=True)

    # Apache ECharts Live Animated Visualization
    chart_summary = final_df.groupby("TO LOCATION")["QUANTITY"].sum().reset_index()
    
    echarts_options = {
        "backgroundColor": "transparent",
        "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}},
        "grid": {"top": "10%", "right": "5%", "bottom": "15%", "left": "10%"},
        "xAxis": {
            "type": "category",
            "data": chart_summary["TO LOCATION"].tolist(),
            "axisLabel": {"color": "#9CA3AF", "rotate": 30},
            "axisLine": {"lineStyle": {"color": "#1F2937"}}
        },
        "yAxis": {
            "type": "value",
            "axisLabel": {"color": "#9CA3AF"},
            "splitLine": {"lineStyle": {"color": "#1F2937", "type": "dashed"}}
        },
        "series": [{
            "data": chart_summary["QUANTITY"].tolist(),
            "type": "bar",
            "showBackground": True,
            "backgroundStyle": {"color": "rgba(31, 41, 55, 0.5)"},
            "itemStyle": {
                "color": {
                    "type": "linear",
                    "x": 0, "y": 0, "x2": 0, "y2": 1,
                    "colorStops": [
                        {"offset": 0, "color": "#60A5FA"},
                        {"offset": 1, "color": "#3B82F6"}
                    ]
                },
                "borderRadius": [6, 6, 0, 0]
            },
            "animationDelay": "function (idx) { return idx * 50; }"
        }],
        "animationEasing": "elasticOut",
        "animationDelayUpdate": "function (idx) { return idx * 20; }"
    }

    st.markdown("#### Store Allocation Volume Matrix")
    st_echarts(options=echarts_options, height="400px")

    # Step 3: Export Matrix
    st.markdown("---")
    st.markdown("### 🚀 Step 3: Oracle Transfer Export Matrix")
    
    exp_col1, exp_col2 = st.columns([4, 1])
    with exp_col2:
        plan_bytes = convert_df_to_excel(final_df, sheet_name="Replenishment")
        st.download_button("Download Plan (.xlsx)", data=plan_bytes, file_name="GCC_Replenishment_Plan.xlsx", use_container_width=True, type="primary")
    
    st.dataframe(final_df, use_container_width=True, hide_index=True)

    with st.expander("🔍 Audit Trail & WMS Reserve Breakdown Logs"):
        st.dataframe(detail_df, use_container_width=True, hide_index=True)