"""
GCC Replenishment Engine - Multi-Node
=====================================
Reads store, warehouse-inventory and (optional) inter-warehouse transit data
from Excel files, allocates stock from multiple source warehouses according
to a country-level warehouse priority grid, and exports an Oracle-ready
transfer plan.

Run with:   streamlit run gcc_replenishment_app.py

requirements.txt:
    streamlit
    pandas
    openpyxl
"""

import time
import traceback
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.styles import Font

# ---------------------------------------------------------------------------
# BUSINESS CONSTANTS
# ---------------------------------------------------------------------------
PACK_SIZE = 12                  # case pack
WMS_RESERVE_PCT = 0.10          # held back at EVERY warehouse, never released

# PACK ROUNDING THRESHOLD (% of a pack) - per country.
# A remainder at or above PACK_SIZE x threshold rounds UP to the next full
# pack; below it rounds down. 50% is the classic "1-5 down, 6-11 up" rule;
# 30% (= 3.6 units on a pack of 12) rounds up more readily.
# These are defaults only - the planner can override them in the UI.
DEFAULT_PACK_THRESHOLD_PCT = {
    "Kuwait": 30,
    "Qatar": 30,
    "Bahrain": 30,
    "Oman": 50,
    "Saudi Arabia": 50,
    "United Arab Emirates": 50,
}
FALLBACK_THRESHOLD_PCT = 50     # used if a country is missing from the grid

# Inter-warehouse movements that are already created/moving are ADDED to the
# receiving warehouse's pool, so stock on its way is not ordered a second
# time. Nothing is subtracted from the sending warehouse: the WMS feed for a
# source warehouse is already net of stock it has shipped out. Set False to
# ignore inter-warehouse transit entirely.
ADD_INBOUND_WH_TRANSIT = True

WH_NAMES = {
    "44324": "DIP WH",
    "22748": "KSA WH",
    "170001": "QAT WH",
    "242211": "JAFZA WH",
}
WH_CODE_BY_NAME = {v.upper(): k for k, v in WH_NAMES.items()}

# WAREHOUSE-REGION ELIGIBILITY & PRIORITY
# Mirrors the business grid: blank = not eligible, number = priority rank,
# 1 = tried first. Edit here if the network changes.
WH_REGION_PRIORITY = {
    "44324":  {"Oman": 1, "United Arab Emirates": 1},
    "22748":  {"Saudi Arabia": 1},
    "170001": {"Qatar": 1},
    "242211": {"Kuwait": 1, "Qatar": 2, "Bahrain": 1,
               "Oman": 2, "Saudi Arabia": 2, "United Arab Emirates": 2},
}

COUNTRY_ALIASES = {
    "uae": "United Arab Emirates", "u.a.e.": "United Arab Emirates",
    "u.a.e": "United Arab Emirates", "united arab emirates": "United Arab Emirates",
    "ksa": "Saudi Arabia", "saudi": "Saudi Arabia", "saudi arabia": "Saudi Arabia",
    "qatar": "Qatar", "qat": "Qatar",
    "kuwait": "Kuwait", "kwt": "Kuwait", "kw": "Kuwait",
    "bahrain": "Bahrain", "bhr": "Bahrain",
    "oman": "Oman", "omn": "Oman",
}

STORE_REQUIRED = [
    "Option", "Store ID", "Country", "Total Sold Qty", "SOH",
    "In-Transit", "Yet to Dispatch", "Target Stock",
]
STORE_NUMERIC = ["Total Sold Qty", "SOH", "In-Transit", "Yet to Dispatch", "Target Stock"]
WH_INV_REQUIRED = ["Option", "WH Available Inventory"]
WH_TRANSIT_REQUIRED = ["Option", "From WH", "To WH", "Qty"]


def build_country_priority():
    """country -> [wh_code, ...] ordered by priority rank."""
    out = {}
    for wh, regions in WH_REGION_PRIORITY.items():
        for country, rank in regions.items():
            out.setdefault(country, []).append((rank, wh))
    return {c: [wh for _, wh in sorted(v)] for c, v in out.items()}


COUNTRY_PRIORITY = build_country_priority()


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def round_to_pack(qty, pack_size: int = PACK_SIZE,
                  threshold_pct: float = FALLBACK_THRESHOLD_PCT) -> int:
    """
    Rounds to the nearest full pack using a country-specific threshold.

    threshold_pct is the share of a pack at which a remainder rounds UP.
    With pack_size 12:
        50%  -> cut-off 6.0 units  (remainder 1-5 down, 6-11 up)
        30%  -> cut-off 3.6 units  (remainder 1-3 down, 4-11 up)
    Anything at or below zero returns 0 (no order triggered).
    """
    if pd.isna(qty) or qty <= 0:
        return 0
    r = qty % pack_size
    if r == 0:
        return int(qty)
    cutoff = pack_size * (float(threshold_pct) / 100.0)
    return int(qty - r) if r < cutoff else int(qty - r + pack_size)


def clean_id_column(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


def normalize_country(value):
    if pd.isna(value):
        return None
    return COUNTRY_ALIASES.get(str(value).strip().lower())


def resolve_wh_code(value):
    """Accept either a WH code (44324) or a WH name (DIP WH)."""
    if pd.isna(value):
        return None
    raw = str(value).strip().replace(".0", "")
    if raw in WH_NAMES:
        return raw
    return WH_CODE_BY_NAME.get(raw.upper())


# ---------------------------------------------------------------------------
# FILE INTAKE
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _read_excel_bytes(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    return pd.read_excel(BytesIO(file_bytes))


def safe_read_uploaded_files(uploaded_files):
    files_data, read_errors = [], []
    for f in uploaded_files:
        try:
            files_data.append((f.name, _read_excel_bytes(f.getvalue(), f.name)))
        except Exception as e:
            read_errors.append(f"'{f.name}': {type(e).__name__}: {e}")
    return files_data, read_errors


def classify_files(files_data):
    """
    Sorts uploads into store data / warehouse inventory / inter-warehouse
    transit by looking at the columns each file carries.
    """
    stores, wh_inv, wh_transit, unknown = [], [], [], []
    for name, df in files_data:
        d = df.copy()
        d.columns = [str(c).strip() for c in d.columns]
        cols = set(d.columns)

        if {"From WH", "To WH"} <= cols:
            wh_transit.append((name, d))
        elif "Store ID" in cols:
            stores.append((name, d))
        elif "WH Available Inventory" in cols or {"WH Code", "WH Name"} & cols:
            wh_inv.append((name, d))
        else:
            unknown.append(name)
    return stores, wh_inv, wh_transit, unknown


def merge_store_files(store_files):
    """
    Joins store-level files on Option + Store ID. A file carrying Store ID
    but no Option (e.g. a store master with Country) is broadcast across
    every row for that store. First file wins on duplicate columns.
    """
    warnings, merged = [], None
    keyed, masters = [], []
    for name, df in store_files:
        (keyed if "Option" in df.columns else masters).append((name, df))

    for name, df in keyed:
        df = df.drop_duplicates(subset=["Option", "Store ID"])
        if merged is None:
            merged = df
            continue
        existing = set(merged.columns)
        new_cols = [c for c in df.columns if c not in existing]
        merged = merged.merge(df[["Option", "Store ID"] + new_cols],
                              on=["Option", "Store ID"], how="outer")
        ignored = [c for c in df.columns
                   if c not in ("Option", "Store ID") and c in existing]
        if ignored:
            warnings.append(
                f"Column(s) {', '.join(ignored)} from '{name}' were already provided "
                f"by an earlier file and were ignored (first file wins)."
            )

    if merged is None:
        raise ValueError(
            "No store-level file found. One uploaded file must contain both "
            "'Option' and 'Store ID' columns."
        )

    for name, df in masters:
        df = df.drop_duplicates(subset=["Store ID"])
        new_cols = [c for c in df.columns if c not in set(merged.columns)]
        if not new_cols:
            warnings.append(f"'{name}' added no new columns and was ignored.")
            continue
        merged = merged.merge(df[["Store ID"] + new_cols], on="Store ID", how="left")

    return merged, warnings


def prepare_store_data(df: pd.DataFrame):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in STORE_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(
            "Store data is missing required column(s): " + ", ".join(missing)
            + ". 'Country' can come from a separate store-master file "
              "containing Store ID and Country."
        )

    warnings = []
    before = len(df)
    df = df.dropna(subset=["Option", "Store ID"])
    if len(df) < before:
        warnings.append(f"{before - len(df)} row(s) dropped - blank Option or Store ID.")

    for c in ["Option", "Store ID"]:
        df[c] = clean_id_column(df[c])

    dup = df.duplicated(subset=["Option", "Store ID"], keep=False)
    if dup.any():
        pairs = (df.loc[dup, ["Option", "Store ID"]].drop_duplicates()
                 .apply(lambda r: f"{r['Option']} / {r['Store ID']}", axis=1).tolist())
        raise ValueError(
            "Duplicate Option + Store ID rows found (each combination must be unique): "
            + "; ".join(pairs[:25]) + (" ..." if len(pairs) > 25 else "")
        )

    for c in STORE_NUMERIC:
        na_before = df[c].isna()
        df[c] = pd.to_numeric(df[c], errors="coerce")
        bad = df[c].isna() & ~na_before
        if bad.any():
            warnings.append(f"{int(bad.sum())} value(s) in '{c}' were not numeric -> treated as 0.")
        df[c] = df[c].fillna(0)

    df["Country Clean"] = df["Country"].apply(normalize_country)
    unknown = df[df["Country Clean"].isna()]["Country"].dropna().unique().tolist()
    if unknown:
        warnings.append(
            "Country not recognised (these stores cannot be served): "
            + ", ".join(map(str, unknown[:10]))
        )
    return df, warnings


def prepare_wh_inventory(wh_files):
    if not wh_files:
        raise ValueError(
            "No warehouse inventory file found. One uploaded file must contain "
            "'Option', a warehouse column ('WH Code' or 'WH Name'), and "
            "'WH Available Inventory'."
        )
    frames, warnings = [], []
    for name, df in wh_files:
        d = df.copy()
        d.columns = [str(c).strip() for c in d.columns]
        missing = [c for c in WH_INV_REQUIRED if c not in d.columns]
        if missing:
            warnings.append(f"'{name}' skipped - missing {', '.join(missing)}.")
            continue
        wh_col = "WH Code" if "WH Code" in d.columns else (
            "WH Name" if "WH Name" in d.columns else None)
        if wh_col is None:
            warnings.append(f"'{name}' skipped - no 'WH Code' or 'WH Name' column.")
            continue
        d["Option"] = clean_id_column(d["Option"])
        d["WH Code Resolved"] = d[wh_col].apply(resolve_wh_code)
        bad = d["WH Code Resolved"].isna()
        if bad.any():
            warnings.append(
                f"{int(bad.sum())} row(s) in '{name}' had an unrecognised warehouse "
                f"and were dropped."
            )
            d = d[~bad]
        d["WH Available Inventory"] = pd.to_numeric(
            d["WH Available Inventory"], errors="coerce").fillna(0)
        frames.append(d[["Option", "WH Code Resolved", "WH Available Inventory"]])

    if not frames:
        raise ValueError("No usable warehouse inventory rows were found in the uploads.")

    out = pd.concat(frames, ignore_index=True)
    out = out.rename(columns={"WH Code Resolved": "WH Code"})
    out = out.groupby(["Option", "WH Code"], as_index=False)["WH Available Inventory"].sum()
    return out, warnings


def prepare_wh_transit(transit_files):
    if not transit_files:
        return None, []
    frames, warnings = [], []
    for name, df in transit_files:
        d = df.copy()
        d.columns = [str(c).strip() for c in d.columns]
        missing = [c for c in WH_TRANSIT_REQUIRED if c not in d.columns]
        if missing:
            warnings.append(f"'{name}' skipped - missing {', '.join(missing)}.")
            continue
        d["Option"] = clean_id_column(d["Option"])
        d["From WH Code"] = d["From WH"].apply(resolve_wh_code)
        d["To WH Code"] = d["To WH"].apply(resolve_wh_code)
        bad = d["From WH Code"].isna() | d["To WH Code"].isna()
        if bad.any():
            warnings.append(f"{int(bad.sum())} transit row(s) in '{name}' had an "
                            f"unrecognised warehouse and were dropped.")
            d = d[~bad]
        d["Qty"] = pd.to_numeric(d["Qty"], errors="coerce").fillna(0)
        frames.append(d[["Option", "From WH Code", "To WH Code", "Qty"]])

    if not frames:
        return None, warnings
    return pd.concat(frames, ignore_index=True), warnings


# ---------------------------------------------------------------------------
# CORE BUSINESS LOGIC
# ---------------------------------------------------------------------------
def calculate_store_metrics(df: pd.DataFrame, pack_thresholds: dict = None) -> pd.DataFrame:
    """
    Per-row metrics, failsafes and pack rounding. Rounding uses the
    country-specific pack threshold, so the same Raw Need can round
    differently in Kuwait than in the UAE.
    """
    thresholds = pack_thresholds or DEFAULT_PACK_THRESHOLD_PCT
    df = df.copy()
    df["Total Pipeline"] = df["SOH"] + df["In-Transit"] + df["Yet to Dispatch"]
    df["Base Need"] = df["Target Stock"].where(df["Total Sold Qty"] > 0, 0)
    df["Overstock Failsafe"] = df["Target Stock"] - df["Total Pipeline"]
    df["Raw Need"] = df[["Base Need", "Overstock Failsafe"]].min(axis=1)
    stockout = df["Total Pipeline"] <= 0
    df.loc[stockout, "Raw Need"] = df.loc[stockout, "Target Stock"]

    df["Pack Threshold %"] = (
        df["Country Clean"].map(thresholds).fillna(FALLBACK_THRESHOLD_PCT)
    )
    df["Rounded Need"] = [
        round_to_pack(q, PACK_SIZE, t)
        for q, t in zip(df["Raw Need"], df["Pack Threshold %"])
    ]
    return df


def build_wh_pools(option, wh_inv_df, wh_transit_df):
    """
    Opening stock per warehouse for one Option, then reduced by the
    permanent 10% reserve.

    Inter-warehouse stock already moving is ADDED to the receiving
    warehouse only. It is deliberately NOT subtracted from the sending
    warehouse, because a source warehouse's WMS inventory is already net
    of everything it has shipped out - subtracting again would remove the
    same units twice.

    Returns (working_pool, reserve_held, opening).
    """
    opening = {}
    sub = wh_inv_df[wh_inv_df["Option"] == option]
    for _, r in sub.iterrows():
        opening[r["WH Code"]] = opening.get(r["WH Code"], 0.0) + float(r["WH Available Inventory"])

    if ADD_INBOUND_WH_TRANSIT and wh_transit_df is not None and len(wh_transit_df):
        t = wh_transit_df[wh_transit_df["Option"] == option]
        for _, r in t.iterrows():
            dest = r["To WH Code"]
            opening[dest] = opening.get(dest, 0.0) + float(r["Qty"])

    opening = {k: max(0.0, v) for k, v in opening.items()}
    working = {k: round(v * (1 - WMS_RESERVE_PCT)) for k, v in opening.items()}
    reserve = {k: opening[k] - working[k] for k in opening}
    return working, reserve, opening


def allocate_option_group(group: pd.DataFrame, working_pool: dict) -> pd.DataFrame:
    """
    Multi-node allocation for one Option.

      1. Rank this Option's stores by Total Sold Qty, highest first, across
         all countries.
      2. Walk the ranked list. Each store looks up its country's warehouse
         priority chain from the eligibility grid and takes its FULL rounded
         need from the first eligible warehouse that can cover it in full.
      3. A store that no eligible warehouse can cover in full gets 0 and the
         run moves on to the next store.

    `working_pool` is mutated as stock is consumed, so each warehouse is
    drawn down across every country it serves.
    """
    group = group.sort_values("Total Sold Qty", ascending=False,
                              kind="mergesort").reset_index(drop=True)
    group["Sales Rank"] = range(1, len(group) + 1)

    allocated, source, status, eligible_txt = [], [], [], []

    for _, row in group.iterrows():
        need = row["Rounded Need"]
        chain = COUNTRY_PRIORITY.get(row["Country Clean"], [])
        eligible_txt.append(" > ".join(WH_NAMES.get(w, w) for w in chain) if chain else "none")

        if not chain:
            allocated.append(0); source.append(""); status.append("No eligible warehouse")
            continue
        if need <= 0:
            allocated.append(0); source.append(""); status.append("No demand")
            continue

        placed = False
        for rank, wh in enumerate(chain, start=1):
            if working_pool.get(wh, 0) >= need:
                working_pool[wh] -= need
                allocated.append(int(need))
                source.append(wh)
                status.append("Filled - priority 1" if rank == 1
                              else f"Filled - priority {rank} (fallback)")
                placed = True
                break
        if not placed:
            allocated.append(0); source.append("")
            status.append("Unfilled - no eligible warehouse had stock")

    group["Allocated Qty"] = allocated
    group["Source WH Code"] = source
    group["Source WH"] = [WH_NAMES.get(c, "") for c in source]
    group["Eligible Warehouses"] = eligible_txt
    group["Fulfilment Status"] = status
    return group


def run_replenishment_engine(store_df, wh_inv_df, wh_transit_df=None,
                             pack_thresholds=None, progress_cb=None):
    """Returns (final_df, detail_df, wh_summary_df, warnings)."""
    def report(frac, label):
        if progress_cb:
            progress_cb(frac, label)

    report(0.05, "Validating store data")
    clean, warnings = prepare_store_data(store_df)

    report(0.18, "Calculating pipeline, failsafes and pack rounding")
    metrics = calculate_store_metrics(clean, pack_thresholds)

    groups = list(metrics.groupby("Option", sort=False))
    total = len(groups)
    report(0.28, f"Building warehouse pools and holding back "
                 f"{int(WMS_RESERVE_PCT*100)}% across {total:,} option(s)")

    frames, wh_rows = [], []
    step = max(1, total // 40)
    for i, (option, group) in enumerate(groups, start=1):
        working, reserve, opening = build_wh_pools(option, wh_inv_df, wh_transit_df)
        before = dict(working)
        frames.append(allocate_option_group(group, working))
        for wh in opening:
            wh_rows.append({
                "Option": option, "WH Code": wh, "WH Name": WH_NAMES.get(wh, wh),
                "Opening (incl. inbound transit)": opening[wh],
                "Working Pool (90%)": before.get(wh, 0),
                "Reserve Held (10%)": reserve.get(wh, 0),
                "Shipped": before.get(wh, 0) - working.get(wh, 0),
                "Left in Working Pool": working.get(wh, 0),
            })
        if i % step == 0 or i == total:
            report(0.28 + 0.58 * (i / total),
                   f"Allocating across warehouses - option {i:,} of {total:,}")

    report(0.90, "Assembling audit trail")
    detail = pd.concat(frames, ignore_index=True).sort_values(
        ["Option", "Sales Rank"]).reset_index(drop=True)
    wh_summary = pd.DataFrame(wh_rows)

    report(0.95, "Building Oracle transfer plan")
    trig = detail[detail["Allocated Qty"] > 0].copy()
    trig["FROM LOCATION"] = trig["Source WH"]
    trig["TO LOCATION"] = trig["Store ID"]
    trig["ITEM"] = trig["Option"]
    trig["QUANTITY"] = trig["Allocated Qty"].astype(int)
    final = trig[["FROM LOCATION", "TO LOCATION", "ITEM", "QUANTITY"]].reset_index(drop=True)

    report(1.0, "Complete")
    return final, detail, wh_summary, warnings


# ---------------------------------------------------------------------------
# EXPORT
# ---------------------------------------------------------------------------
def convert_df_to_excel(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """Column widths from a 200-row sample, not a full-sheet scan."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        for cell in ws[1]:
            cell.font = Font(bold=True)
        n = min(len(df), 200)
        sample = df.head(n)
        for i, col in enumerate(df.columns, start=1):
            letter = ws.cell(row=1, column=i).column_letter
            mx = sample[col].astype(str).map(len).max() if n else 0
            width = max(len(str(col)), int(mx) if pd.notna(mx) else 8) + 4
            ws.column_dimensions[letter].width = min(width, 40)
    return output.getvalue()


# ---------------------------------------------------------------------------
# SUMMARY
# ---------------------------------------------------------------------------
def build_summary(final_df, detail_df, wh_summary):
    total_need = int(detail_df["Rounded Need"].sum())
    total_alloc = int(detail_df["Allocated Qty"].sum())
    fill = (total_alloc / total_need * 100) if total_need else 0.0
    fallback_lines = int(detail_df["Fulfilment Status"]
                         .str.contains("fallback", na=False).sum())
    return {
        "total_units": int(final_df["QUANTITY"].sum()) if len(final_df) else 0,
        "total_lines": len(final_df),
        "items_covered": final_df["ITEM"].nunique() if len(final_df) else 0,
        "stores_covered": final_df["TO LOCATION"].nunique() if len(final_df) else 0,
        "total_need": total_need,
        "total_alloc": total_alloc,
        "fill_rate": fill,
        "reserve_held": int(wh_summary["Reserve Held (10%)"].sum()) if len(wh_summary) else 0,
        "fallback_lines": fallback_lines,
        "wh_engaged": final_df["FROM LOCATION"].nunique() if len(final_df) else 0,
        "wh_total": int(wh_summary["WH Code"].nunique()) if len(wh_summary) else 0,
        "unfilled_lines": int(detail_df["Fulfilment Status"]
                              .str.startswith("Unfilled", na=False).sum()),
    }


def network_rollup(wh_summary: pd.DataFrame) -> pd.DataFrame:
    if not len(wh_summary):
        return pd.DataFrame()
    g = wh_summary.groupby(["WH Code", "WH Name"], as_index=False).agg({
        "Opening (incl. inbound transit)": "sum",
        "Working Pool (90%)": "sum",
        "Reserve Held (10%)": "sum",
        "Shipped": "sum",
        "Left in Working Pool": "sum",
    })
    g["Pool Utilisation %"] = (
        g["Shipped"] / g["Working Pool (90%)"].replace(0, pd.NA) * 100
    ).round(1)
    return g.sort_values("Shipped", ascending=False)


# ---------------------------------------------------------------------------
# THEME
# ---------------------------------------------------------------------------
def inject_theme():
    st.html("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;700&display=swap');

    :root{
      --ink:#0C1222; --ink-2:#141C33; --muted:#7A88A6; --line:#E4E8F0;
      --accent:#2F6FED; --teal:#0EA5A0; --amber:#E0A008; --rose:#E05252;
      --surface:#FFFFFF;
    }

    html, body, [class*="css"]{ font-family:'Inter',sans-serif; }
    .main .block-container{ padding-top:1.2rem; padding-bottom:4rem; max-width:1240px; }
    h1,h2,h3,h4{ font-family:'Space Grotesk',sans-serif !important; color:var(--ink); }
    footer{ visibility:hidden; }

    @keyframes riseIn{
      from{ opacity:0; transform:translateY(18px) scale(.985); }
      to  { opacity:1; transform:translateY(0)    scale(1);    }
    }
    @keyframes fadeIn{ from{opacity:0} to{opacity:1} }
    @keyframes laneFlow{ 0%{background-position:0% 50%} 100%{background-position:200% 50%} }
    @keyframes barGrow{ from{ width:0%; } }
    @keyframes pulseDot{
      0%,100%{ opacity:.35; transform:scale(.85); }
      50%    { opacity:1;   transform:scale(1.15); }
    }

    .rise{ animation:riseIn .55s cubic-bezier(.22,.68,.35,1) both; }
    .d1{animation-delay:.05s}.d2{animation-delay:.13s}.d3{animation-delay:.21s}
    .d4{animation-delay:.29s}.d5{animation-delay:.37s}.d6{animation-delay:.45s}

    @media (prefers-reduced-motion: reduce){
      .rise,.hero,.kpi,.lane-track,.lane-dot{ animation:none !important; }
    }

    .hero{
      position:relative; overflow:hidden;
      background:
        radial-gradient(1200px 300px at 12% -40%, rgba(47,111,237,.45), transparent 60%),
        linear-gradient(118deg,#0C1222 0%,#1B2547 58%,#243163 100%);
      border-radius:18px; padding:30px 34px 26px; margin-bottom:26px;
      box-shadow:0 18px 44px rgba(12,18,34,.28);
      animation:fadeIn .6s ease both;
    }
    .hero-eyebrow{
      font-family:'JetBrains Mono',monospace; font-size:.7rem; letter-spacing:.16em;
      text-transform:uppercase; color:#8DA2E0;
    }
    .hero-title{
      font-family:'Space Grotesk',sans-serif; font-size:2.1rem; font-weight:700;
      color:#fff; margin:8px 0 4px; letter-spacing:-.02em;
    }
    .hero-sub{ font-size:.94rem; color:#B9C6EA; margin:0; }

    .lane{ display:flex; align-items:center; gap:12px; margin-top:20px; flex-wrap:wrap; }
    .lane-node{
      font-family:'JetBrains Mono',monospace; font-size:.72rem; color:#E9EEFC;
      background:rgba(255,255,255,.10); border:1px solid rgba(255,255,255,.22);
      padding:6px 13px; border-radius:999px; white-space:nowrap;
    }
    .lane-track{
      flex:1; min-width:60px; height:3px; border-radius:3px;
      background:linear-gradient(90deg,rgba(255,255,255,.12) 0%,var(--accent) 25%,var(--teal) 50%,rgba(255,255,255,.12) 75%,rgba(255,255,255,.12) 100%);
      background-size:200% 100%; animation:laneFlow 3.4s linear infinite;
    }
    .lane-dot{ width:7px;height:7px;border-radius:50%;background:var(--teal);
      animation:pulseDot 1.8s ease-in-out infinite; }

    .step{
      font-family:'JetBrains Mono',monospace; font-size:.7rem; letter-spacing:.12em;
      text-transform:uppercase; color:var(--accent); font-weight:600; margin:2px 0 4px;
    }
    .sec-title{ font-family:'Space Grotesk',sans-serif; font-weight:600; font-size:1.14rem;
      color:var(--ink); margin:0; }
    .sec-sub{ font-size:.84rem; color:var(--muted); margin-top:3px; }

    .kpi{
      position:relative; background:var(--surface); border:1px solid var(--line);
      border-radius:14px; padding:18px 20px 16px; height:100%;
      box-shadow:0 1px 2px rgba(12,18,34,.05), 0 12px 26px rgba(12,18,34,.07);
      transition:transform .22s ease, box-shadow .22s ease; overflow:hidden;
    }
    .kpi:hover{ transform:translateY(-3px); box-shadow:0 18px 38px rgba(12,18,34,.13); }
    .kpi::before{ content:''; position:absolute; left:0; top:0; bottom:0; width:4px;
      background:var(--accent); }
    .kpi.teal::before{ background:var(--teal); }
    .kpi.amber::before{ background:var(--amber); }
    .kpi.rose::before{ background:var(--rose); }

    .kpi-eyebrow{ font-family:'JetBrains Mono',monospace; font-size:.65rem;
      letter-spacing:.12em; text-transform:uppercase; color:var(--muted); }
    .kpi-value{ font-family:'Space Grotesk',sans-serif; font-size:2rem; font-weight:700;
      color:var(--ink); line-height:1.08; margin:7px 0 3px; font-variant-numeric:tabular-nums; }
    .kpi-label{ font-size:.85rem; color:var(--ink-2); }
    .kpi-cap{ font-size:.74rem; color:var(--muted); margin-top:5px; }

    .meter{ height:5px; border-radius:4px; background:#EDF1F7; margin-top:11px; overflow:hidden; }
    .meter > span{ display:block; height:100%; border-radius:4px;
      background:linear-gradient(90deg,var(--teal),var(--accent));
      animation:barGrow 1.1s cubic-bezier(.22,.68,.35,1) both; }

    .alert{ background:#FDF6E7; border:1px solid #F0DCA4; border-radius:11px;
      padding:12px 17px; font-size:.87rem; color:#7A5B05; margin:14px 0 4px; }
    .alert b{ color:#5B4200; }

    .stDataFrame{ border-radius:11px; overflow:hidden; }
    </style>
    """)


def render_hero():
    st.html("""
    <div class="hero">
      <div class="hero-eyebrow">Multi-Node &middot; Allocation &middot; Oracle Export</div>
      <div class="hero-title">GCC Replenishment Engine</div>
      <p class="hero-sub">Sales-ranked allocation across four source warehouses, with a rolling 10% reserve held at every node.</p>
      <div class="lane">
        <span class="lane-node">DIP WH</span>
        <span class="lane-node">KSA WH</span>
        <span class="lane-node">QAT WH</span>
        <span class="lane-node">JAFZA WH</span>
        <span class="lane-dot"></span>
        <span class="lane-track"></span>
        <span class="lane-dot"></span>
        <span class="lane-node">6 GCC MARKETS</span>
      </div>
    </div>
    """)


def kpi_card(col, eyebrow, value, label, caption="", tone="", delay="d1", meter=None):
    meter_html = ""
    if meter is not None:
        meter_html = f'<div class="meter"><span style="width:{max(0,min(100,meter)):.1f}%"></span></div>'
    cap_html = f'<div class="kpi-cap">{caption}</div>' if caption else ""
    with col:
        st.html(f"""
        <div class="kpi {tone} rise {delay}">
          <div class="kpi-eyebrow">{eyebrow}</div>
          <div class="kpi-value">{value}</div>
          <div class="kpi-label">{label}</div>
          {cap_html}{meter_html}
        </div>
        """)


def render_kpis(s: dict):
    r1 = st.columns(4)
    kpi_card(r1[0], "Transfer Volume", f"{s['total_units']:,}", "Units to ship", delay="d1")
    kpi_card(r1[1], "Order Lines", f"{s['total_lines']:,}", "Store x item lines",
             tone="teal", delay="d2")
    kpi_card(r1[2], "Coverage", f"{s['items_covered']:,}", "Items covered",
             caption=f"across {s['stores_covered']:,} stores", tone="amber", delay="d3")
    kpi_card(r1[3], "Demand Fill Rate", f"{s['fill_rate']:.1f}%", "Of total rounded need",
             caption=f"{s['total_alloc']:,} of {s['total_need']:,} units",
             tone="teal", delay="d4", meter=s["fill_rate"])

    r2 = st.columns(3)
    kpi_card(r2[0], "Reserve Carried Forward", f"{s['reserve_held']:,}", "Units held across all WHs",
             caption="secured for the next run", tone="amber", delay="d4")
    kpi_card(r2[1], "Fallback Sourced", f"{s['fallback_lines']:,}", "Lines from a backup WH",
             caption="priority-1 warehouse could not cover these", tone="rose", delay="d5")
    kpi_card(r2[2], "Warehouses Engaged", f"{s['wh_engaged']}/{s['wh_total']}", "Nodes shipping",
             caption="source warehouses used this run", tone="teal", delay="d6")

    if s["unfilled_lines"]:
        st.html(f"""
        <div class="alert rise d6">
          <b>{s['unfilled_lines']:,} store-item line(s)</b> had demand but no eligible warehouse
          with enough stock to cover it. Fill rate landed at <b>{s['fill_rate']:.1f}%</b>.
          The audit trail shows which warehouses each store was allowed to draw from.
        </div>
        """)


# ---------------------------------------------------------------------------
# APP
# ---------------------------------------------------------------------------
st.set_page_config(page_title="GCC Replenishment Engine", page_icon="\U0001F4E6", layout="wide")
inject_theme()
render_hero()

for k in ("final_df", "detail_df", "wh_summary", "summary",
          "plan_bytes", "detail_bytes", "runtime"):
    if k not in st.session_state:
        st.session_state[k] = None

# --- Step 1: allocation rules ---------------------------------------------
with st.container(border=True):
    st.html('<div class="step">Step 1 &mdash; Allocation rules</div>')
    st.markdown('<p class="sec-title">Rules this run will use</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="sec-sub">Defaults are applied automatically. Switch on the toggle '
        'only if you want to override the pack rounding thresholds for this run.</p>',
        unsafe_allow_html=True,
    )

    rules_left, rules_right = st.columns([1, 1])

    with rules_left:
        st.markdown("**Warehouse eligibility & priority**")
        grid = pd.DataFrame(
            [{"Country": c,
              **{f"Priority {i+1}": WH_NAMES[w] for i, w in enumerate(chain)}}
             for c, chain in sorted(COUNTRY_PRIORITY.items())]
        ).fillna("-")
        st.dataframe(grid, use_container_width=True, hide_index=True)
        st.caption(
            f"A {int(WMS_RESERVE_PCT*100)}% reserve is held at every warehouse and is never "
            f"released. Inbound inter-warehouse transit is "
            f"{'counted' if ADD_INBOUND_WH_TRANSIT else 'ignored'}; nothing is deducted "
            f"from the sending warehouse."
        )

    with rules_right:
        st.markdown("**Pack rounding thresholds**")
        customise = st.toggle(
            "Customise thresholds for this run",
            value=False,
            help="Off: the default grid below is used. On: edit any percentage before running.",
        )

        default_grid = pd.DataFrame(
            [{"Country": c, "Threshold %": p}
             for c, p in DEFAULT_PACK_THRESHOLD_PCT.items()]
        ).sort_values("Country").reset_index(drop=True)

        if customise:
            edited = st.data_editor(
                default_grid,
                key="pack_threshold_editor",
                hide_index=True,
                use_container_width=True,
                column_config={
                    "Country": st.column_config.TextColumn("Country", disabled=True),
                    "Threshold %": st.column_config.NumberColumn(
                        "Threshold %", min_value=0, max_value=100, step=5, format="%d%%",
                        help="Share of a pack at which a remainder rounds up.",
                    ),
                },
            )
            if st.button("Reset to defaults", use_container_width=True):
                st.session_state.pop("pack_threshold_editor", None)
                st.rerun()
        else:
            st.dataframe(default_grid, use_container_width=True, hide_index=True)
            edited = default_grid

        # Clean the edited values into the dict the engine consumes
        pack_thresholds, threshold_warnings = {}, []
        for _, r in edited.iterrows():
            val = pd.to_numeric(r["Threshold %"], errors="coerce")
            if pd.isna(val):
                val = FALLBACK_THRESHOLD_PCT
                threshold_warnings.append(
                    f"{r['Country']}: blank threshold, using {FALLBACK_THRESHOLD_PCT}%."
                )
            elif not (0 <= val <= 100):
                val = min(100, max(0, float(val)))
                threshold_warnings.append(f"{r['Country']}: clamped to {val:.0f}%.")
            pack_thresholds[r["Country"]] = float(val)

        for w in threshold_warnings:
            st.warning(w)

        # Live plain-English explanation of what the current settings do
        by_pct = {}
        for c, p in sorted(pack_thresholds.items()):
            by_pct.setdefault(p, []).append(c)
        lines = []
        for p, countries in sorted(by_pct.items()):
            cutoff = PACK_SIZE * p / 100
            last_down = max(0, -(-cutoff // 1) - 1)  # highest remainder that rounds down
            lines.append(
                f"- **{', '.join(countries)}** at **{p:.0f}%** &rarr; a remainder of "
                f"**{cutoff:.1f}+ units** rounds up to the next pack of {PACK_SIZE} "
                f"(remainders 1&ndash;{int(last_down)} round down)"
            )
        st.markdown("\n".join(lines))

        if pack_thresholds != {c: float(p) for c, p in DEFAULT_PACK_THRESHOLD_PCT.items()}:
            st.info("Custom thresholds are active for this run.")

# --- Step 2: source data ---------------------------------------------------
with st.container(border=True):
    st.html('<div class="step">Step 2 &mdash; Source data</div>')
    st.markdown('<p class="sec-title">Upload source files</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="sec-sub">Files are sorted automatically by their columns. '
        'You need store data (Option, Store ID, Country, sales, stock, Target Stock) and '
        'warehouse inventory (Option, WH Code or WH Name, WH Available Inventory). '
        'An inter-warehouse transit file (Option, From WH, To WH, Qty) is optional.</p>',
        unsafe_allow_html=True,
    )

    uploaded_files = st.file_uploader(
        "Excel files (.xlsx)", type=["xlsx", "xls"],
        accept_multiple_files=True, label_visibility="collapsed",
    )

    if uploaded_files:
        with st.spinner("Reading uploaded file(s)..."):
            files_data, read_errors = safe_read_uploaded_files(uploaded_files)
        for err in read_errors:
            st.error(f"Could not read {err}")

        if files_data:
            store_files, wh_files, transit_files, unknown = classify_files(files_data)
            cols = st.columns(3)
            cols[0].metric("Store files", len(store_files))
            cols[1].metric("Warehouse inventory files", len(wh_files))
            cols[2].metric("Transit files", len(transit_files))
            for u in unknown:
                st.warning(f"'{u}' was not recognised as store, warehouse or transit data.")

            with st.expander(f"Preview {len(files_data)} uploaded file(s)", expanded=False):
                for name, fdf in files_data:
                    st.markdown(f"**{name}** &mdash; {len(fdf):,} row(s)")
                    st.dataframe(fdf.head(20), use_container_width=True)

            if st.button("Run Replenishment", type="primary"):
                bar, txt = st.empty(), st.empty()
                started = time.time()

                def progress_cb(frac, label):
                    bar.progress(min(1.0, max(0.0, frac)))
                    txt.markdown(
                        f"<span style='font-family:JetBrains Mono,monospace;font-size:.8rem;"
                        f"color:#2F6FED;'>{int(frac*100):3d}%</span> "
                        f"<span style='font-size:.86rem;color:#7A88A6;'>{label}</span>",
                        unsafe_allow_html=True,
                    )

                try:
                    progress_cb(0.02, "Merging store files")
                    store_df, w1 = merge_store_files(store_files)
                    wh_inv_df, w2 = prepare_wh_inventory(wh_files)
                    wh_transit_df, w3 = prepare_wh_transit(transit_files)

                    final_df, detail_df, wh_summary, w4 = run_replenishment_engine(
                        store_df, wh_inv_df, wh_transit_df,
                        pack_thresholds=pack_thresholds, progress_cb=progress_cb
                    )

                    progress_cb(1.0, "Preparing downloads")
                    st.session_state.final_df = final_df
                    st.session_state.detail_df = detail_df
                    st.session_state.wh_summary = wh_summary
                    st.session_state.summary = build_summary(final_df, detail_df, wh_summary)
                    st.session_state.plan_bytes = convert_df_to_excel(final_df, "Replenishment")
                    st.session_state.detail_bytes = convert_df_to_excel(detail_df, "CalculationDetail")
                    st.session_state.runtime = time.time() - started

                    bar.empty(); txt.empty()
                    for w in w1 + w2 + w3 + w4:
                        st.warning(w)
                    st.success(f"Replenishment complete - {len(final_df):,} transfer line(s) "
                               f"in {st.session_state.runtime:.1f}s.")

                except ValueError as e:
                    for k in ("final_df", "detail_df", "wh_summary", "summary"):
                        st.session_state[k] = None
                    bar.empty(); txt.empty()
                    st.error(str(e))

                except Exception as e:
                    for k in ("final_df", "detail_df", "wh_summary", "summary"):
                        st.session_state[k] = None
                    bar.empty(); txt.empty()
                    st.error(f"Something went wrong while processing this file "
                             f"({type(e).__name__}: {e}). Check the technical details below.")
                    with st.expander("Technical details"):
                        st.code(traceback.format_exc())

if st.session_state.final_df is not None:
    final_df = st.session_state.final_df
    detail_df = st.session_state.detail_df
    wh_summary = st.session_state.wh_summary
    summary = st.session_state.summary

    st.html('<div class="step" style="margin-top:16px;">Step 3 &mdash; Run summary</div>')
    render_kpis(summary)

    rollup = network_rollup(wh_summary)
    if len(rollup):
        with st.container(border=True):
            st.markdown('<p class="sec-title">Warehouse network</p>', unsafe_allow_html=True)
            st.markdown('<p class="sec-sub">What each node shipped, and what it is holding back.</p>',
                        unsafe_allow_html=True)
            st.dataframe(rollup, use_container_width=True, hide_index=True)

    st.html('<div class="step" style="margin-top:24px;">Step 4 &mdash; Export</div>')
    with st.container(border=True):
        head, dl = st.columns([3, 1])
        with head:
            st.markdown('<p class="sec-title">Oracle transfer plan</p>', unsafe_allow_html=True)
            st.markdown('<p class="sec-sub">FROM LOCATION, TO LOCATION, ITEM, QUANTITY - '
                        'FROM LOCATION is the warehouse that actually sourced each line.</p>',
                        unsafe_allow_html=True)
        with dl:
            st.download_button(
                "Download Transfer Plan", data=st.session_state.plan_bytes,
                file_name="gcc_replenishment_transfer_plan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", use_container_width=True,
            )
        st.dataframe(final_df, use_container_width=True, hide_index=True)

    with st.expander("Calculation detail - audit trail", expanded=False):
        st.markdown('<p class="sec-sub">Every intermediate column behind the allocation, '
                    'including each store\'s country, sales rank, the warehouses it was '
                    'allowed to draw from, which one sourced it, and why.</p>',
                    unsafe_allow_html=True)
        st.download_button(
            "Download Audit Trail", data=st.session_state.detail_bytes,
            file_name="gcc_replenishment_audit_trail.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        if len(detail_df) <= 20_000:
            st.dataframe(detail_df, use_container_width=True, hide_index=True)
        else:
            st.caption(f"{len(detail_df):,} rows - showing the first 20,000 for page "
                       f"performance. The download contains every row.")
            st.dataframe(detail_df.head(20_000), use_container_width=True, hide_index=True)
