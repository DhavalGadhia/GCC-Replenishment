"""
GCC Replenishment Engine
=========================
Streamlit app that reads store/DC inventory data from one or more uploaded
Excel files, runs the replenishment + DC allocation logic below, and
exports an Oracle-ready transfer plan.

Run with:   streamlit run gcc_replenishment_app.py

For the enterprise theme colors to apply, keep the accompanying
.streamlit/config.toml file in a ".streamlit" folder next to this script.
"""

from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.styles import Font

# ---------------------------------------------------------------------------
# CONSTANTS  (adjust here if a business rule changes)
# ---------------------------------------------------------------------------
PACK_SIZE = 12                        # Hardcoded pack size (Section 1)
WMS_RESERVE_PCT = 0.10                # % of DC Available Inventory held back by default
FROM_LOCATION_NAME = "DIP Warehouse"  # Constant written to "FROM LOCATION"

# Store grading (Section 4): A+/A/B+/B by CUMULATIVE SHARE of Total Sold
# Qty within each Option, sorted highest-selling first - confirmed split:
#   A+ : cumulative 0-50% of sales
#   A  : cumulative 50-70% of sales
#   B+ : cumulative 70-80% of sales
#   B  : cumulative 80-100% of sales (bottom 20%)
# A+, A, and B+ together make up the top 80% of cumulative sales (the
# "protected" group); B Grade is the bottom 20%.
GRADE_CUT_A_PLUS = 0.50
GRADE_CUT_A = 0.70
TOP_GROUP_PCT = 0.80
PROTECTED_GRADES = {"A+", "A", "B+"}  # release-trigger group (mirrors old Tier A)

REQUIRED_COLUMNS = [
    "Option",
    "Store ID",
    "Total Sold Qty",
    "SOH",
    "In-Transit",
    "Yet to Dispatch",
    "Target Stock",
    "DC Available Inventory",
]

NUMERIC_COLUMNS = [
    "Total Sold Qty",
    "SOH",
    "In-Transit",
    "Yet to Dispatch",
    "Target Stock",
    "DC Available Inventory",
]

ID_COLUMNS = ["Option", "Store ID"]


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def round_to_pack(qty, pack_size: int = PACK_SIZE) -> int:
    """
    Section 3 - Rounding Rule.
    Rounds to the nearest multiple of `pack_size`. Remainder below half a
    pack rounds down, remainder at or above half rounds up (for pack_size=12
    that's exactly "1-5 down, 6-11 up"). Anything at/below zero -> 0
    (no order triggered).
    """
    if pd.isna(qty) or qty <= 0:
        return 0
    remainder = qty % pack_size
    if remainder == 0:
        return int(qty)
    if remainder < pack_size / 2:
        return int(qty - remainder)
    return int(qty - remainder + pack_size)


def clean_id_column(series: pd.Series) -> pd.Series:
    """Force IDs to clean strings (fixes Excel float artifacts like '1002.0')."""
    return series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


def merge_uploaded_files(files_data):
    """
    files_data: list of (filename, dataframe) tuples, in upload order.
    Stitches multiple files into one wide table by joining on whichever of
    ['Option', 'Store ID'] each file provides - a file with only 'Option'
    (e.g. a Target Stock / DC Available Inventory file) is broadcast across
    every store row for that Option. If the same column shows up in more
    than one file, the first file's values are kept and a warning is raised.
    """
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
            warnings.append(
                f"'{filename}' shares no join key ('Option'/'Store ID') with the "
                f"files already loaded and was skipped."
            )
            continue

        to_merge = df[common_keys + new_cols]
        merged = merged.merge(to_merge, on=common_keys, how="outer")

        ignored = [c for c in df.columns if c not in join_keys and c in existing_cols_before]
        if ignored:
            warnings.append(
                f"Column(s) {', '.join(ignored)} from '{filename}' were already provided "
                f"by an earlier file and were ignored (first file wins)."
            )

    if merged is None:
        raise ValueError("None of the uploaded files contained an 'Option' column.")

    return merged, warnings


def prepare_raw_data(raw_df: pd.DataFrame):
    """
    Validates and cleans the combined data before any business logic runs.
    Returns (clean_df, warnings). `warnings` are surfaced in the UI but
    don't stop the run; a ValueError is raised for problems that would make
    the run unsafe (missing columns, duplicate Option+Store rows).
    """
    df = raw_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "Missing required column(s) across all uploaded files: " + ", ".join(missing)
        )

    warnings = []

    before = len(df)
    df = df.dropna(subset=["Option", "Store ID"])
    if len(df) < before:
        warnings.append(
            f"{before - len(df)} row(s) were dropped because Option or Store ID was blank."
        )

    for col in ID_COLUMNS:
        df[col] = clean_id_column(df[col])

    # Duplicate Option+Store rows would corrupt the grading and the DC
    # deduction, so we stop rather than guess how to merge them.
    dup_mask = df.duplicated(subset=["Option", "Store ID"], keep=False)
    if dup_mask.any():
        dup_pairs = (
            df.loc[dup_mask, ["Option", "Store ID"]]
            .drop_duplicates()
            .apply(lambda r: f"{r['Option']} / {r['Store ID']}", axis=1)
            .tolist()
        )
        raise ValueError(
            "Duplicate Option + Store ID rows found (each combination must "
            "be unique - check for overlapping data across your uploaded "
            "files): " + "; ".join(dup_pairs)
        )

    # Coerce numeric columns; anything unparsable becomes 0 and is flagged.
    for col in NUMERIC_COLUMNS:
        original_na = df[col].isna()
        df[col] = pd.to_numeric(df[col], errors="coerce")
        bad = df[col].isna() & ~original_na
        if bad.any():
            warnings.append(
                f"{int(bad.sum())} value(s) in '{col}' were not numeric and were treated as 0."
            )
        df[col] = df[col].fillna(0)

    # DC Available Inventory should be one constant value per Option (it's
    # the shared DC pool for that item across all requesting stores).
    inconsistent = []
    for option, group in df.groupby("Option"):
        if group["DC Available Inventory"].nunique() > 1:
            inconsistent.append(str(option))
    if inconsistent:
        warnings.append(
            "DC Available Inventory was not consistent across all rows for "
            "Option(s): " + ", ".join(inconsistent) + ". The first value found was used."
        )

    return df, warnings


# ---------------------------------------------------------------------------
# CORE BUSINESS LOGIC
# ---------------------------------------------------------------------------
def calculate_store_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Section 2 & 3 - per-row metrics, failsafes, and rounding."""
    df = df.copy()

    # Section 2: Metric Calculations & Failsafes
    df["Total Pipeline"] = df["SOH"] + df["In-Transit"] + df["Yet to Dispatch"]

    # Base Need: any sale (>0) triggers a refill up to the Target Stock cap;
    # no sale at all -> no base demand.
    df["Base Need"] = df["Target Stock"].where(df["Total Sold Qty"] > 0, 0)

    df["Overstock Failsafe"] = df["Target Stock"] - df["Total Pipeline"]

    # Section 3: Trigger & Need Logic
    df["Raw Need"] = df[["Base Need", "Overstock Failsafe"]].min(axis=1)

    stockout_mask = df["Total Pipeline"] <= 0
    df.loc[stockout_mask, "Raw Need"] = df.loc[stockout_mask, "Target Stock"]

    df["Rounded Need"] = df["Raw Need"].apply(round_to_pack)
    return df


def assign_grades(group: pd.DataFrame) -> pd.DataFrame:
    """
    Grades stores within one Option by cumulative share of Total Sold Qty,
    sorted highest-selling first (confirmed split: A+ 0-50%, A 50-70%,
    B+ 70-80%, B 80-100%). If nobody sold anything for this Option,
    everyone defaults to Grade B - there's no sales basis to prioritize
    any store over another.
    """
    total_sold = group["Total Sold Qty"].sum()

    if total_sold <= 0:
        group["Cumulative Sales %"] = 0.0
        group["Grade"] = "B"
        return group

    cum_pct = group["Total Sold Qty"].cumsum() / total_sold
    group["Cumulative Sales %"] = cum_pct

    def grade_for(pct):
        if pct <= GRADE_CUT_A_PLUS:
            return "A+"
        elif pct <= GRADE_CUT_A:
            return "A"
        elif pct <= TOP_GROUP_PCT:
            return "B+"
        else:
            return "B"

    group["Grade"] = cum_pct.apply(grade_for)
    return group


def allocate_option_group(group: pd.DataFrame) -> pd.DataFrame:
    """
    Section 4 - Grade-based tiering (A+/A/B+/B, ranked by Total Sold Qty)
    plus the 10% WMS Reserve mechanism:

      1. Grade every store in this Option A+/A/B+/B by cumulative sales
         share. A+, A, and B+ together form the "protected" group (the top
         80% of sales, mirrors the old Tier A); B Grade is the sacrificial
         group (the bottom 20% of sales, mirrors old Tier B/C).
      2. New WMS Inventory = DC Available Inventory x (1 - WMS_RESERVE_PCT).
      3. Try to fulfil 100% of BOTH the protected group's and B Grade's
         Rounded Need from that reserved pool.
      4. If the reserved pool can't cover everyone, release the reserve and
         retry against the FULL DC Available Inventory.
      5. If even the full pool can't cover everyone, Scarcity Failsafe:
         fill the protected group sequentially in rank order (A+ first,
         then A, then B+; 100% each until the pool runs out - the moment a
         store can't be fully covered, that store AND every store ranked
         after it gets 0, no leapfrogging) and B Grade gets 0.

    Runs on a single Option's rows at a time, because DC Available
    Inventory is a shared pool that only makes sense per-item.
    """
    group = group.sort_values(
        "Total Sold Qty", ascending=False, kind="mergesort"
    ).reset_index(drop=True)
    group = assign_grades(group)

    dc_full = group["DC Available Inventory"].iloc[0]
    dc_reserved = round(dc_full * (1 - WMS_RESERVE_PCT))

    protected = group[group["Grade"].isin(PROTECTED_GRADES)].copy()
    b_grade = group[group["Grade"] == "B"].copy()

    protected_need = protected["Rounded Need"].sum()
    b_need = b_grade["Rounded Need"].sum() if len(b_grade) else 0
    total_need = protected_need + b_need

    if dc_reserved >= total_need:
        pool = dc_reserved
        reserve_released = False
    else:
        pool = dc_full
        reserve_released = True

    if pool >= total_need:
        # Enough stock (from the reserved or the released pool) to cover
        # everyone in full
        protected["Allocated Qty"] = protected["Rounded Need"]
        if len(b_grade) > 0:
            b_grade["Allocated Qty"] = b_grade["Rounded Need"]
    else:
        # Scarcity Failsafe, using the released (full) pool
        allocations = []
        remaining = pool
        exhausted = False
        for need in protected["Rounded Need"]:
            if not exhausted and remaining >= need:
                allocations.append(need)
                remaining -= need
            else:
                exhausted = True
                allocations.append(0)
        protected["Allocated Qty"] = allocations
        if len(b_grade) > 0:
            b_grade["Allocated Qty"] = 0

    result = pd.concat([protected, b_grade], ignore_index=True)
    result["New WMS Inventory (90%)"] = dc_reserved
    result["Reserve Released"] = reserve_released
    return result


def run_replenishment_engine(raw_df: pd.DataFrame):
    """
    Full pipeline, applied in the exact order of the spec:
      1. Validate & clean inputs
      2/3. Per-row metrics, failsafes, rounding
      4. Per-Option grading + the 10% WMS Reserve allocation mechanism
      5. Build the 4-column Oracle transfer output
    Returns (final_df, detail_df, warnings).
    """
    clean_df, warnings = prepare_raw_data(raw_df)
    metrics_df = calculate_store_metrics(clean_df)

    allocated_frames = [
        allocate_option_group(group)
        for _, group in metrics_df.groupby("Option", sort=False)
    ]
    detail_df = pd.concat(allocated_frames, ignore_index=True)
    detail_df = detail_df.sort_values(["Option", "Store ID"]).reset_index(drop=True)

    triggered = detail_df[detail_df["Allocated Qty"] > 0].copy()

    triggered["FROM LOCATION"] = FROM_LOCATION_NAME
    triggered["TO LOCATION"] = triggered["Store ID"]
    triggered["ITEM"] = triggered["Option"]
    triggered["QUANTITY"] = triggered["Allocated Qty"].astype(int)

    final_df = triggered[["FROM LOCATION", "TO LOCATION", "ITEM", "QUANTITY"]].reset_index(
        drop=True
    )
    return final_df, detail_df, warnings


# ---------------------------------------------------------------------------
# EXCEL EXPORT (formatted for Oracle import)
# ---------------------------------------------------------------------------
def convert_df_to_excel(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        for cell in ws[1]:
            cell.font = Font(bold=True)

        for col_cells in ws.columns:
            length = max(
                (len(str(c.value)) for c in col_cells if c.value is not None), default=8
            )
            ws.column_dimensions[col_cells[0].column_letter].width = length + 4

    return output.getvalue()


# ---------------------------------------------------------------------------
# ENTERPRISE UI THEME
# ---------------------------------------------------------------------------
def inject_theme():
    st.html("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;600&display=swap');

    :root {
        --ink: #131B33;
        --ink-muted: #5B6B82;
        --primary: #1E2A5E;
        --accent: #2F6FED;
        --teal: #0EA5A0;
        --amber: #E3A008;
        --red: #D64545;
        --border: #E2E6ED;
        --card-bg: #FFFFFF;
    }

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .main .block-container { padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1180px; }
    h1, h2, h3, h4 { font-family: 'Space Grotesk', sans-serif !important; color: var(--ink); }
    footer { visibility: hidden; }

    /* --- top header bar --- */
    .gcc-header {
        background: linear-gradient(120deg, var(--primary) 0%, #2A3D7A 100%);
        border-radius: 14px;
        padding: 26px 32px;
        margin-bottom: 28px;
        box-shadow: 0 10px 30px rgba(19,27,51,0.18);
    }
    .gcc-header-eyebrow {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.72rem;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        color: #9FB0E8;
        margin-bottom: 6px;
    }
    .gcc-header-title {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 1.9rem;
        font-weight: 700;
        color: #FFFFFF;
        margin: 0;
    }
    .gcc-header-sub {
        font-family: 'Inter', sans-serif;
        font-size: 0.92rem;
        color: #C7D2F0;
        margin-top: 6px;
    }
    .gcc-header-route {
        display: inline-block;
        margin-top: 14px;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.78rem;
        color: #E8ECFB;
        background: rgba(255,255,255,0.12);
        border: 1px solid rgba(255,255,255,0.22);
        border-radius: 999px;
        padding: 5px 14px;
    }

    /* --- section eyebrow labels --- */
    .gcc-step {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.72rem;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        color: var(--accent);
        font-weight: 600;
        margin-bottom: 2px;
    }

    /* --- KPI cards --- */
    .kpi-card {
        background: var(--card-bg);
        border-left: 4px solid var(--accent);
        border-radius: 12px;
        padding: 20px 22px;
        box-shadow: 0 1px 3px rgba(19,27,51,0.06), 0 10px 24px rgba(19,27,51,0.07);
        height: 100%;
    }
    .kpi-card.kpi-teal { border-left-color: var(--teal); }
    .kpi-card.kpi-amber { border-left-color: var(--amber); }
    .kpi-eyebrow {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.68rem;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        color: var(--ink-muted);
    }
    .kpi-value {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 2.1rem;
        font-weight: 700;
        color: var(--ink);
        margin: 6px 0 2px 0;
        line-height: 1.1;
    }
    .kpi-label { font-size: 0.86rem; color: var(--ink-muted); }
    .kpi-caption { font-size: 0.76rem; color: var(--ink-muted); margin-top: 4px; }

    /* --- attention banner --- */
    .gcc-alert {
        background: #FDF3E0;
        border: 1px solid #F0D796;
        border-radius: 10px;
        padding: 12px 18px;
        font-size: 0.88rem;
        color: #7A5A05;
        margin: 4px 0 20px 0;
    }
    .gcc-alert b { color: #5C4300; }

    /* --- badges for grade / release status (used via pandas Styler) --- */
    .stDataFrame { border-radius: 10px; overflow: hidden; }

    /* section container title row spacing */
    .gcc-section-title { font-family: 'Space Grotesk', sans-serif; font-weight: 600; font-size: 1.15rem; color: var(--ink); margin: 0; }
    .gcc-section-sub { font-size: 0.84rem; color: var(--ink-muted); margin-top: 2px; }
    </style>
    """)


def render_header():
    st.html("""
    <div class="gcc-header">
        <div class="gcc-header-eyebrow">Replenishment &middot; Allocation &middot; Oracle Export</div>
        <p class="gcc-header-title">GCC Replenishment Engine</p>
        <p class="gcc-header-sub">Grade-based store prioritization with a 10% WMS reserve, built for the UAE store network.</p>
        <span class="gcc-header-route">DIP Warehouse &rarr; UAE Stores</span>
    </div>
    """)


def render_kpi_cards(final_df: pd.DataFrame, detail_df: pd.DataFrame):
    total_units = int(final_df["QUANTITY"].sum())
    total_lines = len(final_df)
    items_covered = final_df["ITEM"].nunique()
    stores_covered = final_df["TO LOCATION"].nunique()

    col1, col2, col3 = st.columns(3)
    with col1:
        st.html(f"""
        <div class="kpi-card">
            <div class="kpi-eyebrow">Transfer Volume</div>
            <div class="kpi-value">{total_units:,}</div>
            <div class="kpi-label">Total Units to Ship</div>
        </div>
        """)
    with col2:
        st.html(f"""
        <div class="kpi-card kpi-teal">
            <div class="kpi-eyebrow">Order Lines</div>
            <div class="kpi-value">{total_lines:,}</div>
            <div class="kpi-label">Transfer Lines (Store x Item)</div>
        </div>
        """)
    with col3:
        st.html(f"""
        <div class="kpi-card kpi-amber">
            <div class="kpi-eyebrow">Coverage</div>
            <div class="kpi-value">{items_covered:,}</div>
            <div class="kpi-label">Items Covered</div>
            <div class="kpi-caption">across {stores_covered:,} stores</div>
        </div>
        """)

    # Secondary attention signal - only shown when it's actually relevant
    n_options = detail_df["Option"].nunique()
    n_released = detail_df.loc[detail_df["Reserve Released"], "Option"].nunique()
    n_bgrade_zeroed = detail_df[
        (detail_df["Grade"] == "B") & (detail_df["Reserve Released"])
    ]["Option"].nunique()
    if n_released > 0:
        st.html(f"""
        <div class="gcc-alert">
            <b>{n_released:,} of {n_options:,} items</b> needed the 10% WMS reserve released this run,
            and <b>{n_bgrade_zeroed:,} item(s)</b> had Grade B stores receive zero as a result.
            Check the audit trail below for details.
        </div>
        """)


def style_detail_table(detail_df: pd.DataFrame):
    grade_colors = {
        "A+": "background-color: #E7EAF6; color: #1E2A5E; font-weight: 600;",
        "A": "background-color: #E1F5F3; color: #0B7C77; font-weight: 600;",
        "B+": "background-color: #FCF1D8; color: #8A6608; font-weight: 600;",
        "B": "background-color: #FBE7E7; color: #B23A3A; font-weight: 600;",
    }

    def _grade(val):
        return grade_colors.get(val, "")

    def _release(val):
        if val is True:
            return "background-color: #FCF1D8; color: #8A6608; font-weight: 600;"
        return "background-color: #E1F5F3; color: #0B7C77; font-weight: 600;"

    styler = detail_df.style.map(_grade, subset=["Grade"]).map(_release, subset=["Reserve Released"])
    return styler


# ---------------------------------------------------------------------------
# STREAMLIT APP
# ---------------------------------------------------------------------------
st.set_page_config(page_title="GCC Replenishment Engine", page_icon="\U0001F4E6", layout="wide")
inject_theme()
render_header()

if "final_df" not in st.session_state:
    st.session_state.final_df = None
    st.session_state.detail_df = None

with st.container(border=True):
    st.html('<div class="gcc-step">Step 1</div>')
    st.markdown('<p class="gcc-section-title">Upload Source Data</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="gcc-section-sub">Upload one or more files - e.g. a sales file, a stock file, '
        'and a target/DC inventory file. Required fields are combined automatically by matching '
        'Option / Store ID.</p>',
        unsafe_allow_html=True,
    )
    uploaded_files = st.file_uploader(
        "Excel files (.xlsx)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if uploaded_files:
        files_data = [(f.name, pd.read_excel(f)) for f in uploaded_files]

        with st.expander(f"Preview {len(files_data)} uploaded file(s)", expanded=False):
            for name, fdf in files_data:
                st.markdown(f"**{name}**")
                st.dataframe(fdf.head(20), use_container_width=True)

        if st.button("Run Replenishment", type="primary"):
            try:
                raw_df, merge_warnings = merge_uploaded_files(files_data)
                final_df, detail_df, warnings = run_replenishment_engine(raw_df)
                st.session_state.final_df = final_df
                st.session_state.detail_df = detail_df

                for w in merge_warnings + warnings:
                    st.warning(w)

                st.success(f"Replenishment complete - {len(final_df)} transfer line(s) generated.")
            except ValueError as e:
                st.session_state.final_df = None
                st.session_state.detail_df = None
                st.error(str(e))

if st.session_state.final_df is not None:
    final_df = st.session_state.final_df
    detail_df = st.session_state.detail_df

    st.html('<div class="gcc-step" style="margin-top: 8px;">Step 2</div>')
    render_kpi_cards(final_df, detail_df)

    with st.container(border=True):
        title_col, download_col = st.columns([3, 1])
        with title_col:
            st.markdown('<p class="gcc-section-title">Oracle Transfer Plan</p>', unsafe_allow_html=True)
            st.markdown(
                '<p class="gcc-section-sub">FROM LOCATION, TO LOCATION, ITEM, QUANTITY - ready to import.</p>',
                unsafe_allow_html=True,
            )
        with download_col:
            plan_bytes = convert_df_to_excel(final_df, sheet_name="Replenishment")
            st.download_button(
                label="Download Transfer Plan",
                data=plan_bytes,
                file_name="gcc_replenishment_transfer_plan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        st.dataframe(final_df, use_container_width=True, hide_index=True)

    st.html('<div class="gcc-step" style="margin-top: 20px;">Step 3</div>')
    with st.expander("Calculation Detail - Audit Trail", expanded=False):
        st.markdown(
            '<p class="gcc-section-sub">Every intermediate column used to reach the final '
            'allocation, including each store\'s Grade and whether the 10% reserve was '
            'released - useful for spot-checking any single store against the spec.</p>',
            unsafe_allow_html=True,
        )
        detail_bytes = convert_df_to_excel(detail_df, sheet_name="CalculationDetail")
        st.download_button(
            label="Download Audit Trail",
            data=detail_bytes,
            file_name="gcc_replenishment_audit_trail.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.dataframe(style_detail_table(detail_df), use_container_width=True, hide_index=True)
