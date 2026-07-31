"""
GCC Replenishment Engine
=========================
Streamlit app: reads store/DC data from one or more Excel files, runs the
replenishment + allocation logic, and exports an Oracle-ready transfer plan.

Run with:   streamlit run gcc_replenishment_app.py

requirements.txt:
    streamlit
    pandas
    openpyxl
    streamlit-echarts
    streamlit-shadcn-ui
    altair

The three UI extras (echarts / shadcn / altair) are OPTIONAL at runtime -
every one of them is imported defensively below and the app falls back to
a native Streamlit component if a package is missing or fails to import.
That way a packaging problem can never take the whole app down again.
"""

import time
import traceback
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.styles import Font

# --- optional UI libraries (never fatal) -----------------------------------
try:
    import altair as alt
    HAS_ALTAIR = True
except Exception:
    HAS_ALTAIR = False

try:
    from streamlit_echarts import st_echarts
    HAS_ECHARTS = True
except Exception:
    HAS_ECHARTS = False

try:
    import streamlit_shadcn_ui as ui
    HAS_SHADCN = True
except Exception:
    HAS_SHADCN = False


# ---------------------------------------------------------------------------
# CONSTANTS  (adjust here if a business rule changes)
# ---------------------------------------------------------------------------
PACK_SIZE = 12                        # Hardcoded pack size (Section 1)
WMS_RESERVE_PCT = 0.10                # % of DC Available Inventory held back
FROM_LOCATION_NAME = "DIP Warehouse"  # Constant written to "FROM LOCATION"

REQUIRED_COLUMNS = [
    "Option", "Store ID", "Total Sold Qty", "SOH", "In-Transit",
    "Yet to Dispatch", "Target Stock", "DC Available Inventory",
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
    """
    Rounding Rule: nearest multiple of pack_size. Remainder below half a
    pack rounds down, at/above half rounds up (for 12: "1-5 down, 6-11 up").
    Anything at/below zero -> 0 (no order triggered).
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
    files_data: list of (filename, dataframe) tuples in upload order.
    Joins on whichever of ['Option', 'Store ID'] each file provides - a file
    with only 'Option' (e.g. targets / DC inventory) is broadcast across every
    store row for that Option. First file wins on duplicate columns.
    """
    warnings, merged = [], None

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

        existing = set(merged.columns)
        new_cols = [c for c in df.columns if c not in existing]
        common_keys = [k for k in join_keys if k in existing]

        if not common_keys:
            warnings.append(
                f"'{filename}' shares no join key ('Option'/'Store ID') with the "
                f"files already loaded and was skipped."
            )
            continue

        merged = merged.merge(df[common_keys + new_cols], on=common_keys, how="outer")

        ignored = [c for c in df.columns if c not in join_keys and c in existing]
        if ignored:
            warnings.append(
                f"Column(s) {', '.join(ignored)} from '{filename}' were already "
                f"provided by an earlier file and were ignored (first file wins)."
            )

    if merged is None:
        raise ValueError("None of the uploaded files contained an 'Option' column.")
    return merged, warnings


def prepare_raw_data(raw_df: pd.DataFrame):
    """Validate + clean before any business logic runs."""
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

    # Duplicate Option+Store rows would corrupt the ranking and the DC
    # deduction, so stop rather than guess how to merge them.
    dup_mask = df.duplicated(subset=["Option", "Store ID"], keep=False)
    if dup_mask.any():
        pairs = (
            df.loc[dup_mask, ["Option", "Store ID"]].drop_duplicates()
            .apply(lambda r: f"{r['Option']} / {r['Store ID']}", axis=1).tolist()
        )
        raise ValueError(
            "Duplicate Option + Store ID rows found (each combination must be "
            "unique - check for overlapping data across your uploaded files): "
            + "; ".join(pairs[:25]) + (" ..." if len(pairs) > 25 else "")
        )

    for col in NUMERIC_COLUMNS:
        original_na = df[col].isna()
        df[col] = pd.to_numeric(df[col], errors="coerce")
        bad = df[col].isna() & ~original_na
        if bad.any():
            warnings.append(
                f"{int(bad.sum())} value(s) in '{col}' were not numeric and were treated as 0."
            )
        df[col] = df[col].fillna(0)

    inconsistent = [
        str(opt) for opt, g in df.groupby("Option")
        if g["DC Available Inventory"].nunique() > 1
    ]
    if inconsistent:
        warnings.append(
            "DC Available Inventory was not consistent across all rows for Option(s): "
            + ", ".join(inconsistent[:20]) + ". The first value found was used."
        )

    return df, warnings


# ---------------------------------------------------------------------------
# CORE BUSINESS LOGIC
# ---------------------------------------------------------------------------
def calculate_store_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Per-row metrics, failsafes, and pack rounding."""
    df = df.copy()

    df["Total Pipeline"] = df["SOH"] + df["In-Transit"] + df["Yet to Dispatch"]

    # Base Need: any sale (>0) triggers a refill up to the Target Stock cap;
    # no sale at all -> no base demand.
    df["Base Need"] = df["Target Stock"].where(df["Total Sold Qty"] > 0, 0)

    df["Overstock Failsafe"] = df["Target Stock"] - df["Total Pipeline"]

    df["Raw Need"] = df[["Base Need", "Overstock Failsafe"]].min(axis=1)

    stockout = df["Total Pipeline"] <= 0
    df.loc[stockout, "Raw Need"] = df.loc[stockout, "Target Stock"]

    df["Rounded Need"] = df["Raw Need"].apply(round_to_pack)
    return df


def allocate_option_group(group: pd.DataFrame) -> pd.DataFrame:
    """
    Allocation - pure sales-rank sequential fill from a permanently
    reserved 90% WMS pool.

      1. Sort this Option's stores by Total Sold Qty, highest first.
      2. New WMS Inventory = DC Available Inventory x (1 - WMS_RESERVE_PCT).
         This is the ONLY pool the run may draw from - the 10% reserve is
         never released, whatever the demand. It stays physically in the
         warehouse and simply reappears in tomorrow's DC Available
         Inventory snapshot, where a fresh 10% is carved off again.
      3. If the 90% pool covers 100% of total need, fill everyone.
      4. Otherwise fill sequentially by sales rank: rank 1 gets 100%, then
         rank 2, and so on. The moment a store can't be covered in full,
         that store and every store below it gets 0 (no leapfrogging).
         Stores with zero need are skipped without ending the sequence.

    Runs per Option, because DC Available Inventory is a shared pool that
    only makes sense per-item.
    """
    group = group.sort_values(
        "Total Sold Qty", ascending=False, kind="mergesort"
    ).reset_index(drop=True)
    group["Sales Rank"] = range(1, len(group) + 1)

    dc_full = group["DC Available Inventory"].iloc[0]
    pool = round(dc_full * (1 - WMS_RESERVE_PCT))   # the 90% working pool
    reserve_held = dc_full - pool                   # carried to tomorrow
    total_need = group["Rounded Need"].sum()

    if pool >= total_need:
        group["Allocated Qty"] = group["Rounded Need"]
    else:
        allocations, remaining, exhausted = [], pool, False
        for need in group["Rounded Need"]:
            if need == 0:
                allocations.append(0)
            elif not exhausted and remaining >= need:
                allocations.append(need)
                remaining -= need
            else:
                exhausted = True
                allocations.append(0)
        group["Allocated Qty"] = allocations

    group["New WMS Inventory (90%)"] = pool
    group["Reserve Held (10%)"] = reserve_held
    return group


def run_replenishment_engine(raw_df: pd.DataFrame, progress_cb=None):
    """
    Full pipeline. `progress_cb(fraction, label)` is called at each stage and
    periodically during allocation, so the UI can show real progress rather
    than a decorative animation.
    Returns (final_df, detail_df, warnings).
    """
    def report(frac, label):
        if progress_cb:
            progress_cb(frac, label)

    report(0.05, "Validating columns and cleaning input data")
    clean_df, warnings = prepare_raw_data(raw_df)

    report(0.20, "Calculating pipeline, failsafes and pack rounding")
    metrics_df = calculate_store_metrics(clean_df)

    groups = list(metrics_df.groupby("Option", sort=False))
    total_groups = len(groups)
    frames = []

    report(0.30, f"Ranking stores and holding back {int(WMS_RESERVE_PCT*100)}% WMS stock "
                 f"across {total_groups:,} option(s)")

    step = max(1, total_groups // 40)
    for i, (_, group) in enumerate(groups, start=1):
        frames.append(allocate_option_group(group))
        if i % step == 0 or i == total_groups:
            report(
                0.30 + 0.55 * (i / total_groups),
                f"Allocating stock by sales rank - option {i:,} of {total_groups:,}",
            )

    report(0.88, "Assembling audit trail")
    detail_df = pd.concat(frames, ignore_index=True)
    detail_df = detail_df.sort_values(["Option", "Sales Rank"]).reset_index(drop=True)

    report(0.94, "Building Oracle transfer plan")
    triggered = detail_df[detail_df["Allocated Qty"] > 0].copy()
    triggered["FROM LOCATION"] = FROM_LOCATION_NAME
    triggered["TO LOCATION"] = triggered["Store ID"]
    triggered["ITEM"] = triggered["Option"]
    triggered["QUANTITY"] = triggered["Allocated Qty"].astype(int)

    final_df = triggered[
        ["FROM LOCATION", "TO LOCATION", "ITEM", "QUANTITY"]
    ].reset_index(drop=True)

    report(1.0, "Complete")
    return final_df, detail_df, warnings


# ---------------------------------------------------------------------------
# EXCEL EXPORT
# ---------------------------------------------------------------------------
def convert_df_to_excel(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """
    Column widths are estimated from a 200-row sample, NOT a scan of every
    cell - scanning everything is what made large exports hang.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        for cell in ws[1]:
            cell.font = Font(bold=True)

        n = min(len(df), 200)
        sample = df.head(n)
        for col_idx, col_name in enumerate(df.columns, start=1):
            letter = ws.cell(row=1, column=col_idx).column_letter
            max_len = sample[col_name].astype(str).map(len).max() if n else 0
            width = max(len(str(col_name)), int(max_len) if pd.notna(max_len) else 8) + 4
            ws.column_dimensions[letter].width = min(width, 40)

    return output.getvalue()


@st.cache_data(show_spinner=False)
def _read_excel_bytes(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    """Cached so Streamlit's per-interaction reruns don't re-parse every file."""
    return pd.read_excel(BytesIO(file_bytes))


def safe_read_uploaded_files(uploaded_files):
    """Read each file independently so one bad upload can't crash the app."""
    files_data, read_errors = [], []
    for f in uploaded_files:
        try:
            files_data.append((f.name, _read_excel_bytes(f.getvalue(), f.name)))
        except Exception as e:
            read_errors.append(f"'{f.name}': {type(e).__name__}: {e}")
    return files_data, read_errors


# ---------------------------------------------------------------------------
# SUMMARY STATS
# ---------------------------------------------------------------------------
def build_summary(final_df: pd.DataFrame, detail_df: pd.DataFrame) -> dict:
    total_need = int(detail_df["Rounded Need"].sum())
    total_alloc = int(detail_df["Allocated Qty"].sum())
    fill_rate = (total_alloc / total_need * 100) if total_need else 0.0

    per_option = detail_df.groupby("Option", sort=False).agg(
        need=("Rounded Need", "sum"),
        alloc=("Allocated Qty", "sum"),
        pool=("New WMS Inventory (90%)", "first"),
        reserve=("Reserve Held (10%)", "first"),
    )
    short = per_option[per_option["alloc"] < per_option["need"]]

    return {
        "total_units": int(final_df["QUANTITY"].sum()) if len(final_df) else 0,
        "total_lines": len(final_df),
        "items_covered": final_df["ITEM"].nunique() if len(final_df) else 0,
        "stores_covered": final_df["TO LOCATION"].nunique() if len(final_df) else 0,
        "total_need": total_need,
        "total_alloc": total_alloc,
        "fill_rate": fill_rate,
        "options_total": int(per_option.shape[0]),
        "reserve_held": int(per_option["reserve"].sum()),
        "options_short": int(short.shape[0]),
        "stranded": int((short["pool"] - short["alloc"]).clip(lower=0).sum()),
        "unmet_lines": int(
            ((detail_df["Rounded Need"] > 0) & (detail_df["Allocated Qty"] == 0)).sum()
        ),
    }


# ---------------------------------------------------------------------------
# THEME + ANIMATION
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

    /* ---------- motion primitives ---------- */
    @keyframes riseIn{
      from{ opacity:0; transform:translateY(18px) scale(.985); }
      to  { opacity:1; transform:translateY(0)    scale(1);    }
    }
    @keyframes fadeIn{ from{opacity:0} to{opacity:1} }
    @keyframes laneFlow{
      0%  { background-position:0% 50%; }
      100%{ background-position:200% 50%; }
    }
    @keyframes barGrow{ from{ width:0%; } }
    @keyframes pulseDot{
      0%,100%{ opacity:.35; transform:scale(.85); }
      50%    { opacity:1;   transform:scale(1.15); }
    }

    .rise{ animation:riseIn .55s cubic-bezier(.22,.68,.35,1) both; }
    .d1{animation-delay:.05s}.d2{animation-delay:.13s}.d3{animation-delay:.21s}
    .d4{animation-delay:.29s}.d5{animation-delay:.37s}.d6{animation-delay:.45s}

    @media (prefers-reduced-motion: reduce){
      .rise,.hero,.kpi,.lane{ animation:none !important; }
    }

    /* ---------- hero / dispatch board ---------- */
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

    /* animated DIP -> stores lane */
    .lane{ display:flex; align-items:center; gap:14px; margin-top:20px; }
    .lane-node{
      font-family:'JetBrains Mono',monospace; font-size:.74rem; color:#E9EEFC;
      background:rgba(255,255,255,.10); border:1px solid rgba(255,255,255,.22);
      padding:6px 14px; border-radius:999px; white-space:nowrap;
    }
    .lane-track{
      flex:1; height:3px; border-radius:3px;
      background:linear-gradient(90deg,rgba(255,255,255,.12) 0%,var(--accent) 25%,var(--teal) 50%,rgba(255,255,255,.12) 75%,rgba(255,255,255,.12) 100%);
      background-size:200% 100%;
      animation:laneFlow 3.4s linear infinite;
    }
    .lane-dot{
      width:7px; height:7px; border-radius:50%; background:var(--teal);
      animation:pulseDot 1.8s ease-in-out infinite;
    }

    /* ---------- step eyebrow ---------- */
    .step{
      font-family:'JetBrains Mono',monospace; font-size:.7rem; letter-spacing:.12em;
      text-transform:uppercase; color:var(--accent); font-weight:600; margin:2px 0 4px;
    }
    .sec-title{ font-family:'Space Grotesk',sans-serif; font-weight:600; font-size:1.14rem;
      color:var(--ink); margin:0; }
    .sec-sub{ font-size:.84rem; color:var(--muted); margin-top:3px; }

    /* ---------- KPI cards ---------- */
    .kpi{
      position:relative; background:var(--surface); border:1px solid var(--line);
      border-radius:14px; padding:18px 20px 16px; height:100%;
      box-shadow:0 1px 2px rgba(12,18,34,.05), 0 12px 26px rgba(12,18,34,.07);
      transition:transform .22s ease, box-shadow .22s ease;
      overflow:hidden;
    }
    .kpi:hover{ transform:translateY(-3px); box-shadow:0 18px 38px rgba(12,18,34,.13); }
    .kpi::before{
      content:''; position:absolute; left:0; top:0; bottom:0; width:4px;
      background:var(--accent);
    }
    .kpi.teal::before{ background:var(--teal); }
    .kpi.amber::before{ background:var(--amber); }
    .kpi.rose::before{ background:var(--rose); }

    .kpi-eyebrow{
      font-family:'JetBrains Mono',monospace; font-size:.65rem; letter-spacing:.12em;
      text-transform:uppercase; color:var(--muted);
    }
    .kpi-value{
      font-family:'Space Grotesk',sans-serif; font-size:2rem; font-weight:700;
      color:var(--ink); line-height:1.08; margin:7px 0 3px;
      font-variant-numeric:tabular-nums;
    }
    .kpi-label{ font-size:.85rem; color:var(--ink-2); }
    .kpi-cap{ font-size:.74rem; color:var(--muted); margin-top:5px; }

    /* mini progress meter inside a KPI */
    .meter{ height:5px; border-radius:4px; background:#EDF1F7; margin-top:11px; overflow:hidden; }
    .meter > span{
      display:block; height:100%; border-radius:4px;
      background:linear-gradient(90deg,var(--teal),var(--accent));
      animation:barGrow 1.1s cubic-bezier(.22,.68,.35,1) both;
    }

    /* ---------- alert ---------- */
    .alert{
      background:#FDF6E7; border:1px solid #F0DCA4; border-radius:11px;
      padding:12px 17px; font-size:.87rem; color:#7A5B05; margin:14px 0 4px;
    }
    .alert b{ color:#5B4200; }

    /* ---------- rank chips in tables ---------- */
    .stDataFrame{ border-radius:11px; overflow:hidden; }
    </style>
    """)


def render_hero():
    st.html("""
    <div class="hero">
      <div class="hero-eyebrow">Replenishment &middot; Allocation &middot; Oracle Export</div>
      <div class="hero-title">GCC Replenishment Engine</div>
      <p class="hero-sub">Sales-ranked allocation from a 90% WMS pool, with a rolling 10% reserve held for top sellers across the UAE store network.</p>
      <div class="lane">
        <span class="lane-node">DIP WAREHOUSE</span>
        <span class="lane-dot"></span>
        <span class="lane-track"></span>
        <span class="lane-dot"></span>
        <span class="lane-node">UAE STORES</span>
      </div>
    </div>
    """)


def kpi_card(col, eyebrow, value, label, caption="", tone="", delay="d1", meter=None):
    meter_html = ""
    if meter is not None:
        pct = max(0, min(100, meter))
        meter_html = f'<div class="meter"><span style="width:{pct:.1f}%"></span></div>'
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
    kpi_card(r1[0], "Transfer Volume", f"{s['total_units']:,}", "Units to ship", tone="", delay="d1")
    kpi_card(r1[1], "Order Lines", f"{s['total_lines']:,}", "Store x item lines", tone="teal", delay="d2")
    kpi_card(r1[2], "Coverage", f"{s['items_covered']:,}", "Items covered",
             caption=f"across {s['stores_covered']:,} stores", tone="amber", delay="d3")
    kpi_card(r1[3], "Demand Fill Rate", f"{s['fill_rate']:.1f}%", "Of total rounded need",
             caption=f"{s['total_alloc']:,} of {s['total_need']:,} units",
             tone="teal", delay="d4", meter=s["fill_rate"])

    r2 = st.columns(3)
    kpi_card(r2[0], "Reserve Carried Forward", f"{s['reserve_held']:,}", "Units held in WMS",
             caption=f"secured across {s['options_total']:,} options for the next run",
             tone="amber", delay="d4")
    kpi_card(r2[1], "Constrained Items", f"{s['options_short']:,}", "Options short of full demand",
             caption=f"{s['unmet_lines']:,} store-item lines unfilled", tone="rose", delay="d5")
    kpi_card(r2[2], "Stranded Stock", f"{s['stranded']:,}", "Units in pool but unshipped",
             caption="left behind by the no-leapfrog rule", tone="rose", delay="d6")

    if s["options_short"]:
        st.html(f"""
        <div class="alert rise d6">
          <b>{s['options_short']:,} of {s['options_total']:,} options</b> could not cover full
          demand from the 90% pool, leaving <b>{s['unmet_lines']:,}</b> store-item lines unfilled.
          The 10% reserve (<b>{s['reserve_held']:,} units</b>) stays in the warehouse and rolls
          into the next run. Fill rate landed at <b>{s['fill_rate']:.1f}%</b>.
        </div>
        """)


# ---------------------------------------------------------------------------
# CHARTS
# ---------------------------------------------------------------------------
def render_charts(final_df: pd.DataFrame, detail_df: pd.DataFrame, s: dict):
    left, right = st.columns([3, 2])

    # --- top stores by allocated units -----------------------------------
    with left:
        st.markdown('<p class="sec-title">Top stores by allocated units</p>', unsafe_allow_html=True)
        st.markdown('<p class="sec-sub">Where the volume is actually going this run.</p>',
                    unsafe_allow_html=True)
        top = (
            final_df.groupby("TO LOCATION", as_index=False)["QUANTITY"].sum()
            .sort_values("QUANTITY", ascending=False).head(12)
        )
        if HAS_ALTAIR and len(top):
            chart = (
                alt.Chart(top)
                .mark_bar(cornerRadiusEnd=5, height=17)
                .encode(
                    x=alt.X("QUANTITY:Q", title="Units allocated"),
                    y=alt.Y("TO LOCATION:N", sort="-x", title=None),
                    color=alt.Color("QUANTITY:Q", legend=None,
                                    scale=alt.Scale(range=["#8FB4F5", "#2F6FED", "#0EA5A0"])),
                    tooltip=[alt.Tooltip("TO LOCATION:N", title="Store"),
                             alt.Tooltip("QUANTITY:Q", title="Units", format=",")],
                )
                .properties(height=320)
            )
            st.altair_chart(chart, use_container_width=True)
        else:
            st.bar_chart(top.set_index("TO LOCATION"), height=320)

    # --- fill-rate gauge --------------------------------------------------
    with right:
        st.markdown('<p class="sec-title">Demand fill rate</p>', unsafe_allow_html=True)
        st.markdown('<p class="sec-sub">Allocated units as a share of total rounded need.</p>',
                    unsafe_allow_html=True)
        if HAS_ECHARTS:
            st_echarts(
                options={
                    "series": [{
                        "type": "gauge",
                        "startAngle": 200, "endAngle": -20,
                        "min": 0, "max": 100,
                        "progress": {"show": True, "width": 16,
                                     "itemStyle": {"color": "#2F6FED"}},
                        "axisLine": {"lineStyle": {"width": 16, "color": [[1, "#EDF1F7"]]}},
                        "axisTick": {"show": False},
                        "splitLine": {"show": False},
                        "axisLabel": {"show": False},
                        "pointer": {"show": False},
                        "anchor": {"show": False},
                        "title": {"show": False},
                        "detail": {
                            "valueAnimation": True,
                            "fontSize": 34, "fontWeight": "bold",
                            "offsetCenter": [0, "5%"],
                            "formatter": "{value}%", "color": "#0C1222",
                        },
                        "data": [{"value": round(s["fill_rate"], 1)}],
                    }]
                },
                height="320px",
            )
        else:
            st.metric("Fill rate", f"{s['fill_rate']:.1f}%")
            st.progress(min(1.0, s["fill_rate"] / 100))
            st.caption(f"{s['total_alloc']:,} of {s['total_need']:,} units allocated")

    # --- allocation outcome mix ------------------------------------------
    st.markdown('<p class="sec-title" style="margin-top:18px;">Allocation outcome by option</p>',
                unsafe_allow_html=True)
    st.markdown('<p class="sec-sub">How each option resolved against its WMS pool.</p>',
                unsafe_allow_html=True)

    per_opt = detail_df.groupby("Option", sort=False).agg(
        need=("Rounded Need", "sum"), alloc=("Allocated Qty", "sum"),
    )

    def outcome(r):
        if r["need"] == 0:
            return "No demand this run"
        if r["alloc"] < r["need"]:
            return "Short - sequential cutoff"
        return "Filled from 90% pool"

    mix = (
        per_opt.apply(outcome, axis=1).value_counts()
        .rename_axis("Outcome").reset_index(name="Options")
    )

    if HAS_ALTAIR and len(mix):
        donut = (
            alt.Chart(mix)
            .mark_arc(innerRadius=68, cornerRadius=4, stroke="#fff", strokeWidth=2)
            .encode(
                theta=alt.Theta("Options:Q"),
                color=alt.Color(
                    "Outcome:N", title=None,
                    scale=alt.Scale(
                        domain=["Filled from 90% pool", "Short - sequential cutoff",
                                "No demand this run"],
                        range=["#0EA5A0", "#E05252", "#C3CCDC"]),
                    legend=alt.Legend(orient="right"),
                ),
                tooltip=[alt.Tooltip("Outcome:N"), alt.Tooltip("Options:Q", format=",")],
            )
            .properties(height=260)
        )
        st.altair_chart(donut, use_container_width=True)
    else:
        st.dataframe(mix, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# APP
# ---------------------------------------------------------------------------
st.set_page_config(page_title="GCC Replenishment Engine", page_icon="\U0001F4E6", layout="wide")
inject_theme()
render_hero()

for k in ("final_df", "detail_df", "plan_bytes", "detail_bytes", "summary", "runtime"):
    if k not in st.session_state:
        st.session_state[k] = None

with st.container(border=True):
    st.html('<div class="step">Step 1 &mdash; Source data</div>')
    st.markdown('<p class="sec-title">Upload source files</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="sec-sub">One or more files - e.g. sales, stock, and targets / DC inventory. '
        'Fields are combined automatically by matching Option and Store ID.</p>',
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
            with st.expander(f"Preview {len(files_data)} uploaded file(s)", expanded=False):
                for name, fdf in files_data:
                    st.markdown(f"**{name}** &mdash; {len(fdf):,} row(s)")
                    st.dataframe(fdf.head(20), use_container_width=True)

            if st.button("Run Replenishment", type="primary"):
                bar_slot = st.empty()
                text_slot = st.empty()
                started = time.time()

                def progress_cb(frac, label):
                    bar_slot.progress(min(1.0, max(0.0, frac)))
                    text_slot.markdown(
                        f"<span style='font-family:JetBrains Mono,monospace;font-size:.8rem;"
                        f"color:#2F6FED;'>{int(frac*100):3d}%</span> "
                        f"<span style='font-size:.86rem;color:#7A88A6;'>{label}</span>",
                        unsafe_allow_html=True,
                    )

                try:
                    progress_cb(0.02, "Merging uploaded files")
                    raw_df, merge_warnings = merge_uploaded_files(files_data)

                    final_df, detail_df, warnings = run_replenishment_engine(
                        raw_df, progress_cb=progress_cb
                    )

                    progress_cb(1.0, "Preparing downloads")
                    plan_bytes = convert_df_to_excel(final_df, "Replenishment")
                    detail_bytes = convert_df_to_excel(detail_df, "CalculationDetail")

                    st.session_state.final_df = final_df
                    st.session_state.detail_df = detail_df
                    st.session_state.plan_bytes = plan_bytes
                    st.session_state.detail_bytes = detail_bytes
                    st.session_state.summary = build_summary(final_df, detail_df)
                    st.session_state.runtime = time.time() - started

                    bar_slot.empty()
                    text_slot.empty()

                    for w in merge_warnings + warnings:
                        st.warning(w)
                    st.success(
                        f"Replenishment complete - {len(final_df):,} transfer line(s) "
                        f"in {st.session_state.runtime:.1f}s."
                    )

                except ValueError as e:
                    for k in ("final_df", "detail_df", "plan_bytes", "detail_bytes", "summary"):
                        st.session_state[k] = None
                    bar_slot.empty(); text_slot.empty()
                    st.error(str(e))

                except Exception as e:
                    for k in ("final_df", "detail_df", "plan_bytes", "detail_bytes", "summary"):
                        st.session_state[k] = None
                    bar_slot.empty(); text_slot.empty()
                    st.error(
                        f"Something went wrong while processing this file "
                        f"({type(e).__name__}: {e}). Check the technical details below."
                    )
                    with st.expander("Technical details"):
                        st.code(traceback.format_exc())

if st.session_state.final_df is not None:
    final_df = st.session_state.final_df
    detail_df = st.session_state.detail_df
    summary = st.session_state.summary

    st.html('<div class="step" style="margin-top:16px;">Step 2 &mdash; Run summary</div>')
    render_kpis(summary)

    st.html('<div class="step" style="margin-top:24px;">Step 3 &mdash; Analysis</div>')
    with st.container(border=True):
        render_charts(final_df, detail_df, summary)

    st.html('<div class="step" style="margin-top:24px;">Step 4 &mdash; Export</div>')
    with st.container(border=True):
        head, dl = st.columns([3, 1])
        with head:
            st.markdown('<p class="sec-title">Oracle transfer plan</p>', unsafe_allow_html=True)
            st.markdown(
                '<p class="sec-sub">FROM LOCATION, TO LOCATION, ITEM, QUANTITY - ready to import.</p>',
                unsafe_allow_html=True,
            )
        with dl:
            st.download_button(
                "Download Transfer Plan",
                data=st.session_state.plan_bytes,
                file_name="gcc_replenishment_transfer_plan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", use_container_width=True,
            )
        st.dataframe(final_df, use_container_width=True, hide_index=True)

    with st.expander("Calculation detail - audit trail", expanded=False):
        st.markdown(
            '<p class="sec-sub">Every intermediate column behind the final allocation, '
            'including each store\'s sales rank, the 90% WMS working pool, and the 10% '
            'held back for the next run.</p>',
            unsafe_allow_html=True,
        )
        st.download_button(
            "Download Audit Trail",
            data=st.session_state.detail_bytes,
            file_name="gcc_replenishment_audit_trail.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        if len(detail_df) <= 20_000:
            st.dataframe(detail_df, use_container_width=True, hide_index=True)
        else:
            st.caption(
                f"{len(detail_df):,} rows - showing the first 20,000 for page performance. "
                f"The download above contains every row."
            )
            st.dataframe(detail_df.head(20_000), use_container_width=True, hide_index=True)
