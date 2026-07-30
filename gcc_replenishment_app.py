"""
The Hybrid S1G1 Replenishment Blueprint
========================================
Streamlit app that reads store/DC inventory data from uploaded
Excel files, runs the automated allocation logic, and exports an 
Oracle-ready transfer plan.

Architected for DIP Warehouse Operations.
"""

import time
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.styles import Font

# ---------------------------------------------------------------------------
# CONSTANTS 
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

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def round_to_pack(qty, pack_size: int = PACK_SIZE) -> int:
    if pd.isna(qty) or qty <= 0:
        return 0
    remainder = qty % pack_size
    if remainder == 0:
        return int(qty)
    if remainder < pack_size / 2:
        return int(qty - remainder)
    return int(qty - remainder + pack_size)

def clean_id_column(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

def merge_uploaded_files(files_data):
    warnings = []
    merged = None

    for filename, df in files_data:
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]

        if "Option" not in df.columns:
            warnings.append(f"'{filename}' has no 'Option' column and was skipped.")
            continue

        join_keys = ["Option", "Store ID"] if "Store ID" in df.columns else ["Option"]
        df = df.drop_duplicates(subset=join_keys)

        if merged is None:
            merged = df
            continue

        existing_cols_before = set(merged.columns)
        new_cols = [c for c in df.columns if c not in existing_cols_before]
        common_keys = [k for k in join_keys if k in existing_cols_before]

        if not common_keys:
            warnings.append(f"'{filename}' shares no join key with loaded files. Skipped.")
            continue

        to_merge = df[common_keys + new_cols]
        merged = merged.merge(to_merge, on=common_keys, how="outer")

    if merged is None:
        raise ValueError("None of the uploaded files contained an 'Option' column.")

    return merged, warnings

def prepare_raw_data(raw_df: pd.DataFrame):
    df = raw_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))

    warnings = []
    df = df.dropna(subset=["Option", "Store ID"])

    for col in ID_COLUMNS:
        df[col] = clean_id_column(df[col])

    dup_mask = df.duplicated(subset=["Option", "Store ID"], keep=False)
    if dup_mask.any():
        raise ValueError("Duplicate Option + Store ID rows found. Data must be unique.")

    for col in NUMERIC_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df, warnings

# ---------------------------------------------------------------------------
# CORE BUSINESS LOGIC
# ---------------------------------------------------------------------------
def calculate_store_metrics(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Total Pipeline"] = df["SOH"] + df["In-Transit"] + df["Yet to Dispatch"]
    df["Base Need"] = df["Target Stock"].where(df["Total Sold Qty"] > 0, 0)
    df["Overstock Failsafe"] = df["Target Stock"] - df["Total Pipeline"]
    df["Raw Need"] = df[["Base Need", "Overstock Failsafe"]].min(axis=1)

    stockout_mask = df["Total Pipeline"] <= 0
    df.loc[stockout_mask, "Raw Need"] = df.loc[stockout_mask, "Target Stock"]
    df["Rounded Need"] = df["Raw Need"].apply(round_to_pack)
    return df

def assign_grades(group: pd.DataFrame) -> pd.DataFrame:
    total_sold = group["Total Sold Qty"].sum()
    if total_sold <= 0:
        group["Cumulative Sales %"] = 0.0
        group["Grade"] = "B"
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
    reserve_released = dc_reserved < total_need

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
    result["Reserve Released"] = reserve_released
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

# ---------------------------------------------------------------------------
# EXCEL EXPORT
# ---------------------------------------------------------------------------
def convert_df_to_excel(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        for cell in ws[1]: cell.font = Font(bold=True)
        for col_cells in ws.columns:
            length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
            ws.column_dimensions[col_cells[0].column_letter].width = length + 4
    return output.getvalue()

# ---------------------------------------------------------------------------
# AVANT-GARDE UI THEME & ANIMATIONS
# ---------------------------------------------------------------------------
def inject_theme():
    st.html("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600&family=Plus+Jakarta+Sans:wght@400;500;700&display=swap');

    :root {
        --bg-color: #F4F6F8;
        --surface: #FFFFFF;
        --ink-main: #0B0F19;
        --ink-soft: #636F83;
        --accent-core: #0F172A;
        --accent-highlight: #3B82F6;
    }

    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; background-color: var(--bg-color); }
    .main .block-container { padding-top: 2rem; max-width: 1200px; }
    
    /* Animations */
    @keyframes slideFadeUp {
        0% { opacity: 0; transform: translateY(30px); }
        100% { opacity: 1; transform: translateY(0); }
    }
    .animate-1 { animation: slideFadeUp 0.6s cubic-bezier(0.16, 1, 0.3, 1) 0.1s forwards; opacity: 0; }
    .animate-2 { animation: slideFadeUp 0.6s cubic-bezier(0.16, 1, 0.3, 1) 0.3s forwards; opacity: 0; }
    .animate-3 { animation: slideFadeUp 0.6s cubic-bezier(0.16, 1, 0.3, 1) 0.5s forwards; opacity: 0; }

    /* Header */
    .hero-header {
        background: var(--surface);
        border-radius: 16px;
        padding: 40px;
        margin-bottom: 30px;
        box-shadow: 0 20px 40px rgba(0,0,0,0.04);
        border-left: 6px solid var(--accent-core);
    }
    .hero-title { font-family: 'Outfit', sans-serif; font-size: 2.6rem; font-weight: 600; color: var(--ink-main); margin: 0; letter-spacing: -1px; }
    .hero-subtitle { font-size: 1rem; color: var(--ink-soft); margin-top: 8px; font-weight: 400; }
    .architect-badge { display: inline-block; margin-top: 16px; font-size: 0.8rem; font-weight: 700; color: var(--accent-highlight); text-transform: uppercase; letter-spacing: 1px; }

    /* KPI Cards */
    .kpi-container {
        background: var(--surface);
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.03);
        text-align: center;
        border: 1px solid #E2E8F0;
        transition: transform 0.2s ease;
    }
    .kpi-container:hover { transform: translateY(-4px); box-shadow: 0 15px 35px rgba(0,0,0,0.06); }
    .kpi-value { font-family: 'Outfit', sans-serif; font-size: 2.8rem; font-weight: 600; color: var(--accent-core); margin-bottom: 4px; }
    .kpi-label { font-size: 0.85rem; font-weight: 500; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.5px; }

    /* Custom Tables */
    .stDataFrame { border-radius: 12px; border: 1px solid #E2E8F0; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }
    </style>
    """)

def render_header():
    st.html("""
    <div class="hero-header animate-1">
        <p class="hero-title">The Hybrid S1G1 Replenishment Blueprint</p>
        <p class="hero-subtitle">Automated, rule-based inventory allocation leveraging the 10% WMS Reserve.<br>Routing from the DIP Warehouse to United Arab Emirates store networks.</p>
        <div class="architect-badge">Architected by Dhaval Gadhia</div>
    </div>
    """)

# ---------------------------------------------------------------------------
# APP INITIALIZATION
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Hybrid S1G1 Engine", layout="wide")
inject_theme()
render_header()

if "final_df" not in st.session_state:
    st.session_state.final_df = None
    st.session_state.detail_df = None

# --- STEP 1: UPLOAD ---
with st.container():
    st.markdown("### Step 1: Data Ingestion", unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Upload raw pipeline and sales sheets (.xlsx)",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:
        files_data = [(f.name, pd.read_excel(f)) for f in uploaded_files]

        if st.button("Initialize S1G1 Engine", type="primary", use_container_width=True):
            
            # Simulated Progress Animation
            progress_bar = st.progress(0, text="Engine Start: Ingesting Data...")
            time.sleep(0.3)
            progress_bar.progress(25, text="Aggregating Pipeline Visibility (SOH & In-Transit)...")
            time.sleep(0.4)
            progress_bar.progress(60, text="Partitioning 10% WMS Reserve Pool...")
            
            try:
                raw_df, merge_warnings = merge_uploaded_files(files_data)
                progress_bar.progress(85, text="Executing Grade A Priority Allocations...")
                time.sleep(0.3)
                
                final_df, detail_df, warnings = run_replenishment_engine(raw_df)
                st.session_state.final_df = final_df
                st.session_state.detail_df = detail_df
                
                progress_bar.progress(100, text="Oracle Transfer Plan Generated Successfully.")
                time.sleep(0.4)
                progress_bar.empty()
                
            except ValueError as e:
                progress_bar.empty()
                st.error(str(e))

# --- STEP 2: DASHBOARD & ANALYTICS ---
if st.session_state.final_df is not None:
    final_df = st.session_state.final_df
    detail_df = st.session_state.detail_df

    st.markdown("---")
    st.html('<div class="animate-2">')
    st.markdown("### Step 2: Executive Summary")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.html(f'<div class="kpi-container"><div class="kpi-value">{int(final_df["QUANTITY"].sum()):,}</div><div class="kpi-label">Total Units Allocated</div></div>')
    with col2:
        st.html(f'<div class="kpi-container"><div class="kpi-value">{len(final_df):,}</div><div class="kpi-label">Transfer Lines Generated</div></div>')
    with col3:
        st.html(f'<div class="kpi-container"><div class="kpi-value">{final_df["ITEM"].nunique():,}</div><div class="kpi-label">Unique Options Processed</div></div>')
    st.html('</div>')

    # Interactive Analytics Viewboard
    st.markdown("<br>#### Allocation Heatmap Distribution", unsafe_allow_html=True)
    chart_data = final_df.groupby("TO LOCATION")["QUANTITY"].sum().reset_index()
    st.bar_chart(chart_data, x="TO LOCATION", y="QUANTITY", color="#3B82F6", height=320)

    # --- STEP 3: EXPORT ---
    st.markdown("---")
    st.html('<div class="animate-3">')
    st.markdown("### Step 3: Oracle Transfer Export")
    
    t_col1, t_col2 = st.columns([4, 1])
    with t_col2:
        plan_bytes = convert_df_to_excel(final_df, sheet_name="Replenishment")
        st.download_button("Download Transfer Plan", data=plan_bytes, file_name="S1G1_Transfer_Plan.xlsx", use_container_width=True, type="primary")
    
    st.dataframe(final_df, use_container_width=True, hide_index=True)

    with st.expander("Explore Deep Audit Trail & Reserve Logic Logs"):
        st.dataframe(detail_df, use_container_width=True, hide_index=True)
    st.html('</div>')