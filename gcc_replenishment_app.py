"""
GCC Replenishment Engine
=========================
Streamlit app that reads store/DC inventory data from one or more uploaded
Excel files, runs the replenishment + DC allocation logic below, and
exports an Oracle-ready transfer plan.

Run with:   streamlit run gcc_replenishment_app.py
"""

from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.styles import Font

# ---------------------------------------------------------------------------
# CONSTANTS  (adjust here if a business rule changes)
# ---------------------------------------------------------------------------
PACK_SIZE = 12                     # Hardcoded pack size (Section 1)
TIER_A_SIZE = 10                   # Top N stores by velocity get priority fill (Section 4)
FROM_LOCATION_NAME = "DIP Warehouse"  # Constant written to "FROM LOCATION"

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

    # Duplicate Option+Store rows would corrupt the Tier A ranking and the
    # DC deduction, so we stop rather than guess how to merge them.
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


def allocate_option_group(group: pd.DataFrame) -> pd.DataFrame:
    """
    Section 4 - Item-Level Sorting & Tiering, Tier A Allocation & the
    Scarcity Failsafe, and Fair Share Allocation for Tier B/C.
    Runs on a single Option's rows at a time, because DC Available
    Inventory is a shared pool that only makes sense per-item.
    """
    group = group.sort_values(
        "Total Sold Qty", ascending=False, kind="mergesort"
    ).reset_index(drop=True)

    dc_inventory = group["DC Available Inventory"].iloc[0]

    tier_a = group.iloc[:TIER_A_SIZE].copy()
    tier_bc = group.iloc[TIER_A_SIZE:].copy()
    tier_a["Tier"] = "A"

    tier_a_total_need = tier_a["Rounded Need"].sum()

    if dc_inventory >= tier_a_total_need:
        # Enough stock for all of Tier A -> fulfil 100% of each store's need
        tier_a["Allocated Qty"] = tier_a["Rounded Need"]
        scarcity = False
    else:
        # Scarcity Failsafe: sequential fill by rank. Rank 1 gets 100% of
        # its need, then rank 2, etc. The moment a store can't be fully
        # covered, that store AND every store ranked after it gets 0 - no
        # leapfrogging a higher-ranked store to fill a smaller one.
        allocations = []
        remaining = dc_inventory
        exhausted = False
        for need in tier_a["Rounded Need"]:
            if not exhausted and remaining >= need:
                allocations.append(need)
                remaining -= need
            else:
                exhausted = True
                allocations.append(0)
        tier_a["Allocated Qty"] = allocations
        scarcity = True

    remaining_dc = dc_inventory - tier_a["Allocated Qty"].sum()

    if len(tier_bc) > 0:
        tier_bc["Tier"] = "B/C"
        total_bc_need = tier_bc["Rounded Need"].sum()

        if scarcity or remaining_dc <= 0 or total_bc_need <= 0:
            # Explicit spec rule: if the Scarcity Failsafe triggered, Tier
            # B/C gets 0 regardless of any leftover crumb of DC stock.
            tier_bc["Allocated Qty"] = 0
        elif remaining_dc >= total_bc_need:
            # Enough left over to fulfil everyone in full
            tier_bc["Allocated Qty"] = tier_bc["Rounded Need"]
        else:
            # Pro-Rata Fair Share
            share_pct = tier_bc["Rounded Need"] / total_bc_need
            allocated_raw = share_pct * remaining_dc
            tier_bc["Allocated Qty"] = allocated_raw.apply(round_to_pack)

    return pd.concat([tier_a, tier_bc], ignore_index=True)


def run_replenishment_engine(raw_df: pd.DataFrame):
    """
    Full pipeline, applied in the exact order of the spec:
      1. Validate & clean inputs
      2/3. Per-row metrics, failsafes, rounding
      4. Per-Option Tier A + Scarcity Failsafe + Fair Share allocation
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
def convert_df_to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Replenishment")
        ws = writer.sheets["Replenishment"]

        for cell in ws[1]:
            cell.font = Font(bold=True)

        for col_cells in ws.columns:
            length = max(
                (len(str(c.value)) for c in col_cells if c.value is not None), default=8
            )
            ws.column_dimensions[col_cells[0].column_letter].width = length + 4

    return output.getvalue()


# ---------------------------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------------------------
st.set_page_config(page_title="GCC Replenishment Engine", layout="wide")
st.title("GCC Replenishment Engine")
st.caption(
    "Upload one or more source files, run the replenishment logic, and "
    "export an Oracle-ready transfer plan (DIP Warehouse -> UAE Stores)."
)

if "final_df" not in st.session_state:
    st.session_state.final_df = None
    st.session_state.detail_df = None

uploaded_files = st.file_uploader(
    "Upload raw data file(s) (.xlsx) - e.g. a sales file, a stock file, "
    "and a target/DC inventory file. Required fields will be combined "
    "automatically by matching Option / Store ID.",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
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

            st.success(f"Replenishment complete — {len(final_df)} transfer line(s) generated.")
        except ValueError as e:
            st.session_state.final_df = None
            st.session_state.detail_df = None
            st.error(str(e))

if st.session_state.final_df is not None:
    st.subheader("Transfer Plan")
    st.dataframe(st.session_state.final_df, use_container_width=True)

    with st.expander("Show calculation detail (audit trail)"):
        st.caption(
            "Every intermediate column used to reach the final allocation — "
            "useful for spot-checking a store's number against the spec."
        )
        st.dataframe(st.session_state.detail_df, use_container_width=True)

    excel_bytes = convert_df_to_excel(st.session_state.final_df)
    st.download_button(
        label="Download Transfer Plan (.xlsx)",
        data=excel_bytes,
        file_name="gcc_replenishment_transfer_plan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
