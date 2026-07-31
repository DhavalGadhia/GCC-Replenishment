"""
GCC Replenishment Engine
=========================================================
Advanced Streamlit application featuring heavy Apache ECharts live animations,
glassmorphic UI, pulsing glow effects, and automated Oracle allocation pipelines.
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

st.set_page_config(page_title="GCC Replenishment Engine | Command Center", layout="wide")

# ---------------------------------------------------------------------------
# NEXT-GEN CYBERPUNK / EXECUTIVE GLASSMORPHIC THEME & ANIMATIONS
# ---------------------------------------------------------------------------
def inject_next_gen_theme():
    st.html("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');

    :root {
        --bg-deep: #030712;
        --surface: rgba(17, 24, 39, 0.75);
        --surface-border: rgba(59, 130, 246, 0.2);
        --text-main: #F9FAFB;
        --text-muted: #9CA3AF;
        --accent-blue: #3B82F6;
        --accent-glow: rgba(59, 130, 246, 0.4);
    }

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
        background-color: var(--bg-deep);
        color: var(--text-main);
    }
    
    .main .block-container { padding-top: 2rem; max-width: 1350px; }

    /* Heavy Keyframe Animations */
    @keyframes entrySlideUp {
        0% { opacity: 0; transform: translateY(35px) scale(0.98); }
        100% { opacity: 1; transform: translateY(0) scale(1); }
    }
    
    @keyframes pulseNeon {
        0% { box-shadow: 0 0 15px rgba(59, 130, 246, 0.2), inset 0 0 15px rgba(59, 130, 246, 0.1); }
        50% { box-shadow: 0 0 35px rgba(59, 130, 246, 0.6), inset 0 0 25px rgba(59, 130, 246, 0.3); }
        100% { box-shadow: 0 0 15px rgba(59, 130, 246, 0.2), inset 0 0 15px rgba(59, 130, 246, 0.1); }
    }

    @keyframes floatSlow {
        0% { transform: translateY(0px); }
        50% { transform: translateY(-6px); }
        100% { transform: translateY(0px); }
    }

    .animated-entrance {
        animation: entrySlideUp 0.9s cubic-bezier(0.16, 1, 0.3, 1) forwards;
    }

    /* Glassmorphism Command Center Banner */
    .command-header {
        background: linear-gradient(135deg, rgba(17, 24, 39, 0.9) 0%, rgba(30, 41, 59, 0.8) 100%);
        backdrop-filter: blur(16px);
        border: 1px solid var(--surface-border);
        border-radius: 24px;
        padding: 45px;
        margin-bottom: 35px;
        animation: entrySlideUp 0.7s cubic-bezier(0.16, 1, 0.3, 1) forwards, pulseNeon 6s ease-in-out infinite;
        position: relative;
        overflow: hidden;
    }
    
    .command-header::before {
        content: '';
        position: absolute;
        top: -50%; left: -50%; width: 200%; height: 200%;
        background: radial-gradient(circle, rgba(59, 130, 246, 0.08) 0%, transparent 60%);
        pointer-events: none;
    }

    .header-title { 
        font-size: 2.8rem; 
        font-weight: 800; 
        background: linear-gradient(90deg, #FFFFFF, #60A5FA, #93C5FD);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 0; 
        letter-spacing: -1px; 
    }
    
    .header-subtitle { 
        font-size: 1.05rem; 
        color: var(--text-muted); 
        margin-top: 12px; 
        font-weight: 400; 
        line-height: 1.6; 
    }

    /* Status Pill */
    .status-pill {
        display: inline-flex;
        align-items: center;
        background: rgba(16, 185, 129, 0.15);
        border: 1px solid rgba(16, 185, 129, 0.4);
        color: #34D399;
        padding: 6px 14px;
        border-radius: 50px;
        font-size: 0.85rem;
        font-weight: 600;
        margin-bottom: 15px;
        animation: floatSlow 4s ease-in-out infinite;
    }
    .status-dot {
        width: 8px; height: 8px; background-color: #34D399; border-radius: 50%;
        margin-right: 8px; box-shadow: 0 0 10px #34D399;
    }

    .stDataFrame { 
        border-radius: 16px; 
        border: 1px solid var(--surface-border); 
        background: var(--surface);
        overflow: hidden; 
    }
    </style>
    """)

inject_next_gen_theme()

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

def convert_dfs_to_excel(dfs_dict: dict) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in dfs_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            for cell in ws[1]: cell.font = Font(bold=True)
    return output.getvalue()

# ---------------------------------------------------------------------------
# APPLICATION UI COMMAND CENTER
# ---------------------------------------------------------------------------
st.html("""
<div class="command-header animated-entrance">
    <div class="status-pill"><div class="status-dot"></div>Engine Active & Ready</div>
    <div class="header-title">GCC Replenishment Engine</div>
    <div class="header-subtitle">Enterprise Autonomous Inventory Allocation Matrix. Optimizing multi-store distribution vectors and enforcing the 10% WMS Reserve protocol from the DIP Warehouse.</div>
</div>
""")

if "final_df" not in st.session_state:
    st.session_state.final_df = None
    st.session_state.detail_df = None

# Step 1: Data Ingestion
st.markdown("### 📥 Data Ingestion & Pipeline Upload")
uploaded_files = st.file_uploader("Upload pipeline and inventory reports (.xlsx)", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    files_data = [(f.name, pd.read_excel(f)) for f in uploaded_files]
    if st.button("🚀 Initialize Allocation Protocol", type="primary", use_container_width=True):
        progress_bar = st.progress(0, text="Synchronizing WMS Context...")
        time.sleep(0.2)
        progress_bar.progress(35, text="Validating Stock Profiles & Pipeline Metrics...")
        time.sleep(0.3)
        progress_bar.progress(75, text="Partitioning Reserve Pool & Computing ABC Grades...")
        
        try:
            raw_df, _ = merge_uploaded_files(files_data)
            final_df, detail_df, _ = run_replenishment_engine(raw_df)
            st.session_state.final_df = final_df
            st.session_state.detail_df = detail_df
            progress_bar.progress(100, text="Execution Complete.")
            time.sleep(0.3)
            progress_bar.empty()
        except ValueError as e:
            progress_bar.empty()
            st.error(str(e))

# Step 2: Advanced Interactive Analytics & Animated Dashboard
if st.session_state.final_df is not None:
    final_df = st.session_state.final_df
    detail_df = st.session_state.detail_df

    st.markdown("---")
    
    # Phase 2 Header with Download Option Added
    p2_col1, p2_col2 = st.columns([3, 1])
    with p2_col1:
        st.markdown("### 📊 Live Executive Analytics & KPIs")
    with p2_col2:
        chart_summary = final_df.groupby("TO LOCATION")["QUANTITY"].sum().reset_index()
        sku_summary = final_df.groupby("ITEM")["QUANTITY"].sum().head(10).reset_index()
        analytics_bytes = convert_dfs_to_excel({
            "Store Distribution": chart_summary,
            "Item Share": sku_summary
        })
        st.download_button("📥 Download Analytics (.xlsx)", data=analytics_bytes, file_name="GCC_Analytics_Summary.xlsx", use_container_width=True)

    # Shadcn Metric Cards Grid with Enhanced Spacing
    col1, col2, col3 = st.columns(3)
    with col1:
        ui.metric_card(title="Total Units Allocated", content=f"{int(final_df['QUANTITY'].sum()):,}", description="Optimized transfer size", key="m1")
    with col2:
        ui.metric_card(title="Transfer Routing Lines", content=f"{len(final_df):,}", description="Oracle routing nodes", key="m2")
    with col3:
        ui.metric_card(title="Unique SKUs Processed", content=f"{final_df['ITEM'].nunique():,}", description="Option style count", key="m3")

    st.markdown("<br>", unsafe_allow_html=True)

    # Sleek, High-End Professional Apache ECharts Visualizations
    col_chart1, col_chart2 = st.columns(2)

    with col_chart1:
        st.markdown("#### 🌐 Store Distribution Volume Matrix")
        echarts_store_options = {
            "backgroundColor": "transparent",
            "tooltip": {
                "trigger": "axis",
                "axisPointer": {"type": "shadow", "shadowStyle": {"color": "rgba(59, 130, 246, 0.06)"}},
                "backgroundColor": "rgba(15, 23, 42, 0.9)",
                "borderColor": "rgba(59, 130, 246, 0.4)",
                "borderWidth": 1,
                "textStyle": {"color": "#F9FAFB", "fontFamily": "Plus Jakarta Sans"}
            },
            "grid": {"top": "15%", "right": "5%", "bottom": "18%", "left": "12%"},
            "xAxis": {
                "type": "category",
                "data": chart_summary["TO LOCATION"].tolist(),
                "axisLabel": {"color": "#9CA3AF", "rotate": 30, "fontSize": 11, "fontWeight": 500},
                "axisLine": {"lineStyle": {"color": "#374151", "width": 1.5}},
                "axisTick": {"show": False}
            },
            "yAxis": {
                "type": "value",
                "axisLabel": {"color": "#9CA3AF", "fontSize": 11},
                "splitLine": {"lineStyle": {"color": "rgba(55, 65, 81, 0.4)", "type": "dashed"}},
                "axisLine": {"show": False}
            },
            "series": [{
                "data": chart_summary["QUANTITY"].tolist(),
                "type": "bar",
                "showBackground": True,
                "backgroundStyle": {"color": "rgba(30, 41, 59, 0.4)", "borderRadius": [8, 8, 0, 0]},
                "itemStyle": {
                    "color": {
                        "type": "linear",
                        "x": 0, "y": 0, "x2": 0, "y2": 1,
                        "colorStops": [
                            {"offset": 0, "color": "#38BDF8"},  # Vibrant Cyan
                            {"offset": 0.5, "color": "#3B82F6"}, # Electric Blue
                            {"offset": 1, "color": "#1D4ED8"}    # Deep Royal Blue
                        ]
                    },
                    "borderRadius": [8, 8, 0, 0],
                    "shadowColor": "rgba(59, 130, 246, 0.4)",
                    "shadowBlur": 12
                },
                "animationDuration": 1800,
                "animationEasing": "cubicOut"
            }]
        }
        st_echarts(options=echarts_store_options, height="380px")

    with col_chart2:
        st.markdown("#### ⚡ Top Item Allocation Share")
        echarts_item_options = {
            "backgroundColor": "transparent",
            "color": [
                "#38BDF8", "#3B82F6", "#6366F1", "#8B5CF6", 
                "#EC4899", "#10B981", "#F59E0B", "#14B8A6", 
                "#A855F7", "#64748B"
            ],
            "tooltip": {
                "trigger": "item",
                "formatter": "{b}<br/><b>{c} units</b> ({d}%)",
                "backgroundColor": "rgba(15, 23, 42, 0.9)",
                "borderColor": "rgba(59, 130, 246, 0.4)",
                "borderWidth": 1,
                "textStyle": {"color": "#F9FAFB", "fontFamily": "Plus Jakarta Sans"}
            },
            "legend": {
                "type": "scroll",
                "orient": "horizontal",
                "bottom": "0%",
                "textStyle": {"color": "#9CA3AF", "fontSize": 10},
                "pageTextStyle": {"color": "#F9FAFB"}
            },
            "series": [{
                "type": "pie",
                "radius": ["38%", "68%"],
                "center": ["50%", "45%"],
                "avoidLabelOverlap": True,
                "itemStyle": {
                    "borderRadius": 8,
                    "borderColor": "#030712",
                    "borderWidth": 3
                },
                "label": {
                    "show": True,
                    "formatter": "{b}\n{d}%",
                    "color": "#9CA3AF",
                    "fontSize": 10,
                    "fontWeight": 500
                },
                "labelLine": {
                    "length": 10,
                    "length2": 10,
                    "lineStyle": {"color": "#475569"}
                },
                "emphasis": {
                    "label": {"show": True, "fontSize": "13", "fontWeight": "bold", "color": "#FFFFFF"},
                    "itemStyle": {"shadowBlur": 15, "shadowOffsetX": 0, "shadowColor": "rgba(59, 130, 246, 0.6)"}
                },
                "data": [{"value": row["QUANTITY"], "name": str(row["ITEM"])} for _, row in sku_summary.iterrows()],
                "animationDuration": 2000,
                "animationEasing": "cubicOut"
            }]
        }
        st_echarts(options=echarts_item_options, height="380px")

    # Step 3: Export Matrix
    st.markdown("---")
    st.markdown("### 🚀 Phase 3: Oracle Transfer Export Matrix")
    
    exp_col1, exp_col2 = st.columns([4, 1])
    with exp_col2:
        plan_bytes = convert_df_to_excel(final_df, sheet_name="Replenishment")
        st.download_button("📥 Download Plan (.xlsx)", data=plan_bytes, file_name="GCC_Replenishment_Plan.xlsx", use_container_width=True, type="primary")
    
    st.dataframe(final_df, use_container_width=True, hide_index=True)

    with st.expander("🔍 Deep Audit Trail & WMS Reserve Breakdown Logs"):
        audit_col1, audit_col2 = st.columns([4, 1])
        with audit_col2:
            audit_bytes = convert_df_to_excel(detail_df, sheet_name="Audit_Trail")
            st.download_button("📥 Download Audit Log (.xlsx)", data=audit_bytes, file_name="GCC_Audit_Trail.xlsx", use_container_width=True, type="primary")
        st.dataframe(detail_df, use_container_width=True, hide_index=True)
