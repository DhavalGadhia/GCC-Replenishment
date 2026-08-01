"""
GCC Replenishment Engine - Multi-Node (production)
==================================================
Reads the five source extracts (DEPTH, RMS, SALE, SOH, WMS) straight out of
the reporting system, allocates stock from multiple source warehouses using a
country-level warehouse priority grid, and exports an Oracle-ready transfer
plan.

Run with:   streamlit run gcc_replenishment_app.py

requirements.txt:
    streamlit
    pandas
    numpy
    openpyxl
"""

import time
import traceback
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font

# ---------------------------------------------------------------------------
# BUSINESS CONSTANTS
# ---------------------------------------------------------------------------
PACK_SIZE = 12
WMS_RESERVE_PCT = 0.10          # held at every warehouse, never released

# Warehouses in the eligibility grid, keyed by WMS "Loc Key".
GRID_WH = {
    "44324":  "DIP WH",
    "22748":  "KSA WH",
    "170001": "QAT WH",
    "242211": "JAFZA WH",
}

# WAREHOUSE-REGION ELIGIBILITY & PRIORITY (1 = tried first, absent = ineligible)
WH_REGION_PRIORITY = {
    "44324":  {"Oman": 1, "United Arab Emirates": 1},
    "22748":  {"Saudi Arabia": 1},
    "170001": {"Qatar": 1},
    "242211": {"Kuwait": 1, "Qatar": 2, "Bahrain": 1,
               "Oman": 2, "Saudi Arabia": 2, "United Arab Emirates": 2},
}

# Pack rounding threshold (% of a pack at which a remainder rounds up)
DEFAULT_PACK_THRESHOLD_PCT = {
    "Kuwait": 30, "Qatar": 30, "Bahrain": 30,
    "Oman": 50, "Saudi Arabia": 50, "United Arab Emirates": 50,
}
FALLBACK_THRESHOLD_PCT = 50

# Fallback Target Stock used when an Option + Store combination has no row in
# the DEPTH file. Without this the combination cannot be sized and is skipped
# entirely, so this is what brings those stores into the run.
DEFAULT_TARGET_STOCK = {
    "Kuwait": 24, "Qatar": 24, "Bahrain": 24,
    "Oman": 24, "Saudi Arabia": 24, "United Arab Emirates": 24,
}
FALLBACK_TARGET_STOCK = 24

# Only these columns are pulled from each workbook - reading all 44/56 columns
# of a 55 MB extract is the single biggest avoidable cost.
FILE_SPECS = {
    "DEPTH": {
        "signature": {"OPTION", "New Depth"},
        "header_row": 3,                       # two blank rows precede the header
        "cols": ["OPTION", "Store code", "New Depth"],
    },
    "RMS": {
        "signature": {"Open Order Qty", "From Loc Key"},
        "header_row": 1,
        "cols": ["From Loc Key", "To Loc Key", "To Loc Type", "Option", "Open Order Qty"],
    },
    "SALE": {
        "signature": {"Net Sales Qty", "Location Code"},
        "header_row": 1,
        "cols": ["Country", "Location Code", "Location", "Item Style Code",
                 "Item Color", "Net Sales Qty"],
    },
    "WMS": {
        "signature": {"Pack Available Qty", "Loc Key"},
        "header_row": 1,
        "cols": ["Country", "Loc Key", "Location", "Option", "Pack Available Qty"],
    },
    # SOH is a wide pivot and is handled by its own reader.
    "SOH": {"signature": {"Pack PhysicalQty", "option"}, "header_row": 4, "cols": None},
}


def build_country_priority():
    out = {}
    for wh, regions in WH_REGION_PRIORITY.items():
        for country, rank in regions.items():
            out.setdefault(country, []).append((rank, wh))
    return {c: [wh for _, wh in sorted(v)] for c, v in out.items()}


COUNTRY_PRIORITY = build_country_priority()

_S = lambda s: s.astype(str).str.strip()


# ---------------------------------------------------------------------------
# ROUNDING
# ---------------------------------------------------------------------------
def round_to_pack(qty, pack_size=PACK_SIZE, threshold_pct=FALLBACK_THRESHOLD_PCT) -> int:
    """Scalar form. Remainder at/above pack_size x threshold rounds up."""
    if pd.isna(qty) or qty <= 0:
        return 0
    r = qty % pack_size
    if r == 0:
        return int(qty)
    return int(qty - r) if r < pack_size * threshold_pct / 100.0 else int(qty - r + pack_size)


def round_to_pack_vec(qty, threshold_pct, pack_size=PACK_SIZE):
    """Vectorised form used by the engine - same rule, applied row-wise."""
    q = np.asarray(qty, dtype=float)
    t = np.asarray(threshold_pct, dtype=float)
    r = np.mod(q, pack_size)
    cut = pack_size * t / 100.0
    rounded = np.where(r == 0, q, np.where(r < cut, q - r, q - r + pack_size))
    return np.where(q > 0, rounded, 0).astype(np.int64)


# ---------------------------------------------------------------------------
# FILE INTAKE
# ---------------------------------------------------------------------------
def _sheet_header(path_or_buf, header_row):
    wb = load_workbook(path_or_buf, read_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    for _ in range(header_row - 1):
        next(it)
    hdr = list(next(it))
    wb.close()
    return hdr


def identify_file(buf):
    """Return the file kind by matching signature columns in the first rows."""
    try:
        wb = load_workbook(buf, read_only=True)
        ws = wb[wb.sheetnames[0]]
        seen = set()
        for i, row in enumerate(ws.iter_rows(max_row=5, values_only=True)):
            seen |= {str(c).strip() for c in row if c is not None}
            if i >= 4:
                break
        wb.close()
    except Exception:
        return None
    for kind, spec in FILE_SPECS.items():
        if spec["signature"] <= seen:
            return kind
    return None


def read_flat(buf, kind):
    """Read only the needed columns from a flat sheet."""
    spec = FILE_SPECS[kind]
    wb = load_workbook(buf, read_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    for _ in range(spec["header_row"] - 1):
        next(it)
    hdr = list(next(it))
    idx = {c: i for i, c in enumerate(hdr) if c is not None}
    missing = [c for c in spec["cols"] if c not in idx]
    if missing:
        wb.close()
        raise ValueError(f"{kind} file is missing column(s): {', '.join(missing)}")
    ii = [idx[c] for c in spec["cols"]]
    rows = [tuple(r[i] for i in ii) for r in it]
    wb.close()
    return pd.DataFrame(rows, columns=spec["cols"])


def read_soh_pivot(buf):
    """
    SOH is a wide crosstab: rows 1-3 carry Country / Store name / Store code
    across repeating per-store metric blocks, row 4 is the metric header, and
    the leading columns are item attributes. Unpivot to Option x Store x SOH,
    skipping zero cells (that alone removes ~99% of the cells).
    """
    wb = load_workbook(buf, read_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    r1, r2, r3, r4 = (list(next(it)) for _ in range(4))

    n = len(r4)
    attr_end = 0
    while attr_end < n and r1[attr_end] is None:
        attr_end += 1

    if "option" in r4:
        opt_i = r4.index("option")
    elif "Option" in r4:
        opt_i = r4.index("Option")
    else:
        wb.close()
        raise ValueError("SOH file has no 'option' column in its header row.")

    blocks, j = [], attr_end
    while j < n:
        code, start = r3[j], j
        while j < n and r3[j] == code:
            j += 1
        metrics = {r4[k]: k for k in range(start, j)}
        if "Pack PhysicalQty" in metrics:
            blocks.append((r1[start], r2[start], code, metrics["Pack PhysicalQty"]))

    recs = []
    for r in it:
        opt = r[opt_i]
        if opt is None:
            continue
        for country, store_name, code, ci in blocks:
            q = r[ci]
            if q:
                recs.append((opt, code, store_name, country, q))
    wb.close()
    return pd.DataFrame(recs, columns=["Option", "Store code", "Location", "Country", "SOH"])


@st.cache_data(show_spinner=False, max_entries=3)
def load_sources(file_payloads):
    """
    file_payloads: tuple of (name, bytes). Cached on the raw bytes so a rerun
    (any widget click) never re-parses the workbooks.
    Returns (frames dict, per-file read seconds, warnings).
    """
    frames, timings, warnings = {}, {}, []
    for name, data in file_payloads:
        buf = BytesIO(data)
        kind = identify_file(buf)
        if kind is None:
            warnings.append(f"'{name}' was not recognised as DEPTH, RMS, SALE, SOH or WMS.")
            continue
        t0 = time.time()
        buf.seek(0)
        try:
            df = read_soh_pivot(buf) if kind == "SOH" else read_flat(buf, kind)
        except Exception as e:
            warnings.append(f"'{name}' ({kind}) could not be read: {type(e).__name__}: {e}")
            continue
        if kind in frames:
            frames[kind] = pd.concat([frames[kind], df], ignore_index=True)
        else:
            frames[kind] = df
        timings[kind] = timings.get(kind, 0) + (time.time() - t0)
    return frames, timings, warnings


# ---------------------------------------------------------------------------
# ENGINE
# ---------------------------------------------------------------------------
def prepare(frames, active_wh):
    """Normalise the five extracts into the tables the allocator consumes."""
    missing = [k for k in ("DEPTH", "SALE", "SOH", "WMS") if k not in frames]
    if missing:
        raise ValueError(
            "Missing required file(s): " + ", ".join(missing)
            + ". Upload the DEPTH, SALE, SOH and WMS extracts (RMS is optional "
              "but strongly recommended - without it, in-transit stock is unknown "
              "and the engine will over-order)."
        )

    depth, sale = frames["DEPTH"], frames["SALE"]
    soh, wms = frames["SOH"], frames["WMS"]
    rms = frames.get("RMS")
    warnings = []

    depth = depth.assign(
        Option=_S(depth["OPTION"]), Store=_S(depth["Store code"]),
        Target=pd.to_numeric(depth["New Depth"], errors="coerce").fillna(0),
    ).groupby(["Option", "Store"], as_index=False)["Target"].max()

    sale = sale.assign(
        Option=_S(sale["Item Style Code"]) + _S(sale["Item Color"]),
        Store=_S(sale["Location Code"]),
        Qty=pd.to_numeric(sale["Net Sales Qty"], errors="coerce").fillna(0),
    )
    country_map = pd.concat([
        sale[["Store", "Country"]],
        soh.assign(Store=_S(soh["Store code"]))[["Store", "Country"]],
    ]).dropna().drop_duplicates("Store")

    sales = sale.groupby(["Option", "Store"], as_index=False)["Qty"].sum()
    sales = sales[sales["Qty"] > 0]

    soh_ag = soh.assign(
        Option=_S(soh["Option"]), Store=_S(soh["Store code"]),
        SOH=pd.to_numeric(soh["SOH"], errors="coerce").fillna(0),
    ).groupby(["Option", "Store"], as_index=False)["SOH"].sum()

    if rms is not None and len(rms):
        rms = rms.assign(
            Option=_S(rms["Option"]), To=_S(rms["To Loc Key"]),
            Open=pd.to_numeric(rms["Open Order Qty"], errors="coerce").fillna(0),
        )
        in_transit = (rms[rms["To Loc Type"] == "S"]
                      .groupby(["Option", "To"], as_index=False)["Open"].sum()
                      .rename(columns={"To": "Store", "Open": "InTransit"}))
        inbound = (rms[(rms["To Loc Type"] == "W") & (rms["To"].isin(active_wh))]
                   .groupby(["Option", "To"], as_index=False)["Open"].sum()
                   .rename(columns={"To": "WH", "Open": "Inbound"}))
    else:
        warnings.append(
            "No RMS file uploaded - in-transit and yet-to-dispatch stock is being "
            "treated as zero, which will overstate demand."
        )
        in_transit = pd.DataFrame(columns=["Option", "Store", "InTransit"])
        inbound = pd.DataFrame(columns=["Option", "WH", "Inbound"])

    wms_ag = wms.assign(
        Option=_S(wms["Option"]), WH=_S(wms["Loc Key"]),
        Avail=pd.to_numeric(wms["Pack Available Qty"], errors="coerce").fillna(0),
    )
    wms_ag = wms_ag[wms_ag["WH"].isin(active_wh)]
    wms_ag = wms_ag.groupby(["Option", "WH"], as_index=False)["Avail"].sum()
    if len(inbound):
        wms_ag = wms_ag.merge(inbound, on=["Option", "WH"], how="left")
    else:
        wms_ag["Inbound"] = 0.0
    wms_ag["Inbound"] = wms_ag["Inbound"].fillna(0)
    wms_ag["Opening"] = wms_ag["Avail"] + wms_ag["Inbound"]
    wms_ag["Working Pool (90%)"] = (wms_ag["Opening"] * (1 - WMS_RESERVE_PCT)).round().astype(np.int64)
    wms_ag["Reserve Held (10%)"] = wms_ag["Opening"] - wms_ag["Working Pool (90%)"]

    return sales, depth, soh_ag, in_transit, wms_ag, country_map, warnings


def build_universe(sales, depth, soh_ag, in_transit, country_map,
                   pack_thresholds, default_targets=None):
    """
    Scope = option x store with sales > 0. Target Stock comes from DEPTH where
    a row exists; otherwise the country's default Target Stock is applied, so
    a selling store is never skipped just because DEPTH has no row for it.
    """
    # Country must be attached BEFORE the target is resolved, because the
    # fallback target is country-specific.
    u = sales.merge(country_map, on="Store", how="left")
    u["Country"] = u["Country"].fillna("UNKNOWN")

    u = u.merge(depth, on=["Option", "Store"], how="left")

    dflt = default_targets or DEFAULT_TARGET_STOCK
    missing = u["Target"].isna()
    n_defaulted = int(missing.sum())
    u["Target Source"] = np.where(missing, "Default grid", "DEPTH file")
    u.loc[missing, "Target"] = (
        u.loc[missing, "Country"].map(dflt).fillna(FALLBACK_TARGET_STOCK)
    )

    u = (u.merge(soh_ag, on=["Option", "Store"], how="left")
           .merge(in_transit, on=["Option", "Store"], how="left"))
    u[["SOH", "InTransit"]] = u[["SOH", "InTransit"]].fillna(0)

    # Open Order Qty already covers in-transit AND yet-to-dispatch, so there is
    # no separate yet-to-dispatch term to add here.
    u["Total Pipeline"] = u["SOH"] + u["InTransit"]
    u["Base Need"] = np.where(u["Qty"] > 0, u["Target"], 0)
    u["Overstock Failsafe"] = u["Target"] - u["Total Pipeline"]
    u["Raw Need"] = np.minimum(u["Base Need"], u["Overstock Failsafe"])
    stockout = u["Total Pipeline"] <= 0
    u.loc[stockout, "Raw Need"] = u.loc[stockout, "Target"]

    thr = pack_thresholds or DEFAULT_PACK_THRESHOLD_PCT
    u["Pack Threshold %"] = u["Country"].map(thr).fillna(FALLBACK_THRESHOLD_PCT)
    u["Rounded Need"] = round_to_pack_vec(u["Raw Need"], u["Pack Threshold %"])
    return u, n_defaulted


def allocate(u, wms_ag):
    """
    Rank every store by Total Sold Qty (highest first) within its Option, then
    walk that Option's country priority chain and take the full rounded need
    from the first eligible warehouse that can cover it. Pools are per
    Option x Warehouse and are drawn down as the pass proceeds.
    """
    pools = dict(zip(
        zip(wms_ag["Option"].to_numpy(), wms_ag["WH"].to_numpy()),
        wms_ag["Working Pool (90%)"].to_numpy().astype(int),
    ))

    u = u.sort_values(["Option", "Qty"], ascending=[True, False], kind="mergesort")
    u["Sales Rank"] = u.groupby("Option").cumcount() + 1

    options = u["Option"].to_numpy()
    countries = u["Country"].to_numpy()
    needs = u["Rounded Need"].to_numpy()

    alloc = np.zeros(len(u), dtype=np.int64)
    source = np.empty(len(u), dtype=object)
    status = np.empty(len(u), dtype=object)

    for i in range(len(u)):
        need = needs[i]
        if need <= 0:
            source[i] = ""; status[i] = "No demand"; continue
        chain = COUNTRY_PRIORITY.get(countries[i])
        if not chain:
            source[i] = ""; status[i] = "No eligible warehouse"; continue
        placed = False
        for rank, wh in enumerate(chain, start=1):
            key = (options[i], wh)
            if pools.get(key, 0) >= need:
                pools[key] -= need
                alloc[i] = need
                source[i] = wh
                status[i] = ("Filled - priority 1" if rank == 1
                             else f"Filled - priority {rank} (fallback)")
                placed = True
                break
        if not placed:
            source[i] = ""; status[i] = "Unfilled - no eligible warehouse had stock"

    u["Allocated Qty"] = alloc
    u["Source WH Code"] = source
    u["Source WH"] = pd.Series(source, index=u.index).map(GRID_WH).fillna("")
    u["Fulfilment Status"] = status
    return u, pools


def run_engine(frames, active_wh, pack_thresholds=None, default_targets=None,
               progress_cb=None):
    def report(f, label):
        if progress_cb:
            progress_cb(f, label)

    t0 = time.time()
    report(0.10, "Normalising extracts and mapping stores to countries")
    sales, depth, soh_ag, in_transit, wms_ag, cmap, warns = prepare(frames, active_wh)

    report(0.35, "Calculating pipeline, failsafes and pack rounding")
    u, n_defaulted = build_universe(sales, depth, soh_ag, in_transit, cmap,
                                    pack_thresholds, default_targets)

    report(0.60, f"Allocating across warehouses - {u['Option'].nunique():,} options")
    detail, pools = allocate(u, wms_ag)

    report(0.90, "Building Oracle transfer plan")
    hit = detail[detail["Allocated Qty"] > 0]
    final = pd.DataFrame({
        "FROM LOCATION": hit["Source WH Code"].to_numpy(),
        "TO LOCATION": hit["Store"].to_numpy(),
        "ITEM": hit["Option"].to_numpy(),
        "QUANTITY": hit["Allocated Qty"].to_numpy().astype(int),
    })

    shipped = (final.groupby("FROM LOCATION")["QUANTITY"].sum()
               if len(final) else pd.Series(dtype=float))
    wh_summary = wms_ag.groupby("WH", as_index=False).agg(
        Opening=("Opening", "sum"),
        **{"Working Pool (90%)": ("Working Pool (90%)", "sum"),
           "Reserve Held (10%)": ("Reserve Held (10%)", "sum")})
    wh_summary["WH Name"] = wh_summary["WH"].map(GRID_WH)
    wh_summary["Shipped"] = wh_summary["WH"].map(shipped).fillna(0).astype(int)
    wh_summary["Unused Pool"] = wh_summary["Working Pool (90%)"] - wh_summary["Shipped"]
    wh_summary["Pool Utilisation %"] = (
        wh_summary["Shipped"] / wh_summary["Working Pool (90%)"].replace(0, np.nan) * 100
    ).round(1)
    wh_summary = wh_summary[["WH", "WH Name", "Opening", "Working Pool (90%)",
                             "Reserve Held (10%)", "Shipped", "Unused Pool",
                             "Pool Utilisation %"]]

    report(1.0, "Complete")
    return final, detail, wh_summary, warns, n_defaulted, time.time() - t0


# ---------------------------------------------------------------------------
# EXPORT
# ---------------------------------------------------------------------------
def to_excel(df, sheet_name="Sheet1"):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        for cell in ws[1]:
            cell.font = Font(bold=True)
        n = min(len(df), 200)
        sample = df.head(n)
        for i, col in enumerate(df.columns, start=1):
            letter = ws.cell(row=1, column=i).column_letter
            mx = sample[col].astype(str).map(len).max() if n else 0
            ws.column_dimensions[letter].width = min(
                max(len(str(col)), int(mx) if pd.notna(mx) else 8) + 4, 40)
    return out.getvalue()


def summarise(final, detail, wh_summary, n_defaulted):
    need = int(detail["Rounded Need"].sum())
    got = int(detail["Allocated Qty"].sum())
    return {
        "units": int(final["QUANTITY"].sum()) if len(final) else 0,
        "lines": len(final),
        "items": final["ITEM"].nunique() if len(final) else 0,
        "stores": final["TO LOCATION"].nunique() if len(final) else 0,
        "need": need, "got": got,
        "fill": (got / need * 100) if need else 0.0,
        "reserve": int(wh_summary["Reserve Held (10%)"].sum()),
        "fallback": int(detail["Fulfilment Status"].str.contains("fallback", na=False).sum()),
        "unfilled": int(detail["Fulfilment Status"].str.startswith("Unfilled", na=False).sum()),
        "unmet_units": int(detail.loc[
            detail["Fulfilment Status"].str.startswith("Unfilled", na=False), "Rounded Need"].sum()),
        "defaulted": n_defaulted,
        "defaulted_units": int(detail.loc[detail["Target Source"] == "Default grid", "Allocated Qty"].sum()),
        "wh_used": final["FROM LOCATION"].nunique() if len(final) else 0,
    }


# ---------------------------------------------------------------------------
# THEME
# ---------------------------------------------------------------------------
def inject_theme():
    st.html("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;700&display=swap');
    :root{ --ink:#0C1222; --ink-2:#141C33; --muted:#7A88A6; --line:#E4E8F0;
           --accent:#2F6FED; --teal:#0EA5A0; --amber:#E0A008; --rose:#E05252; --surface:#fff; }
    html, body, [class*="css"]{ font-family:'Inter',sans-serif; }
    .main .block-container{ padding-top:1.2rem; padding-bottom:4rem; max-width:1240px; }
    h1,h2,h3,h4{ font-family:'Space Grotesk',sans-serif !important; color:var(--ink); }
    footer{ visibility:hidden; }

    @keyframes riseIn{ from{opacity:0; transform:translateY(18px) scale(.985);} to{opacity:1; transform:none;} }
    @keyframes fadeIn{ from{opacity:0} to{opacity:1} }
    @keyframes laneFlow{ 0%{background-position:0% 50%} 100%{background-position:200% 50%} }
    @keyframes barGrow{ from{width:0%} }
    @keyframes pulseDot{ 0%,100%{opacity:.35;transform:scale(.85)} 50%{opacity:1;transform:scale(1.15)} }
    .rise{ animation:riseIn .55s cubic-bezier(.22,.68,.35,1) both; }
    .d1{animation-delay:.05s}.d2{animation-delay:.13s}.d3{animation-delay:.21s}
    .d4{animation-delay:.29s}.d5{animation-delay:.37s}.d6{animation-delay:.45s}
    @media (prefers-reduced-motion: reduce){ .rise,.hero,.kpi,.lane-track,.lane-dot{animation:none !important} }

    .hero{ position:relative; overflow:hidden;
      background: radial-gradient(1200px 300px at 12% -40%, rgba(47,111,237,.45), transparent 60%),
                  linear-gradient(118deg,#0C1222 0%,#1B2547 58%,#243163 100%);
      border-radius:18px; padding:30px 34px 26px; margin-bottom:26px;
      box-shadow:0 18px 44px rgba(12,18,34,.28); animation:fadeIn .6s ease both; }
    .hero-eyebrow{ font-family:'JetBrains Mono',monospace; font-size:.7rem; letter-spacing:.16em;
      text-transform:uppercase; color:#8DA2E0; }
    .hero-title{ font-family:'Space Grotesk',sans-serif; font-size:2.1rem; font-weight:700;
      color:#fff; margin:8px 0 4px; letter-spacing:-.02em; }
    .hero-sub{ font-size:.94rem; color:#B9C6EA; margin:0; }
    .lane{ display:flex; align-items:center; gap:12px; margin-top:20px; flex-wrap:wrap; }
    .lane-node{ font-family:'JetBrains Mono',monospace; font-size:.72rem; color:#E9EEFC;
      background:rgba(255,255,255,.10); border:1px solid rgba(255,255,255,.22);
      padding:6px 13px; border-radius:999px; white-space:nowrap; }
    .lane-track{ flex:1; min-width:60px; height:3px; border-radius:3px;
      background:linear-gradient(90deg,rgba(255,255,255,.12) 0%,var(--accent) 25%,var(--teal) 50%,rgba(255,255,255,.12) 75%);
      background-size:200% 100%; animation:laneFlow 3.4s linear infinite; }
    .lane-dot{ width:7px;height:7px;border-radius:50%;background:var(--teal);
      animation:pulseDot 1.8s ease-in-out infinite; }

    .step{ font-family:'JetBrains Mono',monospace; font-size:.7rem; letter-spacing:.12em;
      text-transform:uppercase; color:var(--accent); font-weight:600; margin:2px 0 4px; }
    .sec-title{ font-family:'Space Grotesk',sans-serif; font-weight:600; font-size:1.14rem;
      color:var(--ink); margin:0; }
    .sec-sub{ font-size:.84rem; color:var(--muted); margin-top:3px; }

    .kpi{ position:relative; background:var(--surface); border:1px solid var(--line);
      border-radius:14px; padding:18px 20px 16px; height:100%;
      box-shadow:0 1px 2px rgba(12,18,34,.05), 0 12px 26px rgba(12,18,34,.07);
      transition:transform .22s ease, box-shadow .22s ease; overflow:hidden; }
    .kpi:hover{ transform:translateY(-3px); box-shadow:0 18px 38px rgba(12,18,34,.13); }
    .kpi::before{ content:''; position:absolute; left:0; top:0; bottom:0; width:4px; background:var(--accent); }
    .kpi.teal::before{background:var(--teal)} .kpi.amber::before{background:var(--amber)}
    .kpi.rose::before{background:var(--rose)}
    .kpi-eyebrow{ font-family:'JetBrains Mono',monospace; font-size:.65rem; letter-spacing:.12em;
      text-transform:uppercase; color:var(--muted); }
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

    /* ============ Step 1 control deck ============ */
    .grp{ font-family:'JetBrains Mono',monospace; font-size:.62rem; letter-spacing:.18em;
      text-transform:uppercase; color:var(--muted); font-weight:600;
      margin:18px 0 8px; padding-bottom:5px; border-bottom:1px solid var(--line); }

    .ribbon{ display:flex; flex-wrap:wrap; gap:7px; margin:2px 0 4px; }
    .ribbon .rb{ font-family:'JetBrains Mono',monospace; font-size:.68rem; color:var(--muted);
      background:#F4F7FC; border:1px solid var(--line); border-radius:999px; padding:4px 11px;
      white-space:nowrap; }
    .ribbon .rb b{ color:var(--ink); font-weight:700; }

    /* warehouse node cards */
    .node{ border:1px solid var(--line); border-radius:12px; padding:11px 13px 10px;
      background:var(--surface); position:relative; overflow:hidden; margin-top:-6px;
      transition:opacity .2s ease, border-color .2s ease, transform .2s ease; }
    .node::before{ content:''; position:absolute; left:0; top:0; bottom:0; width:3px;
      background:var(--nc); }
    .node:hover{ transform:translateY(-2px); border-color:var(--nc); }
    .node.off{ opacity:.34; filter:grayscale(1); }
    .node-code{ font-family:'JetBrains Mono',monospace; font-size:.74rem; color:var(--ink);
      font-weight:700; letter-spacing:.04em; }
    .node-serves{ font-family:'JetBrains Mono',monospace; font-size:.62rem; color:var(--nc);
      letter-spacing:.09em; margin-top:4px; font-weight:600; }
    .node-meta{ font-size:.66rem; color:var(--muted); margin-top:3px; }

    /* routing lanes */
    .lanes{ display:flex; flex-direction:column; gap:5px; }
    .lane-row{ display:flex; align-items:center; gap:8px; padding:6px 10px;
      border:1px solid var(--line); border-radius:10px; background:var(--surface);
      transition:border-color .2s ease, background .2s ease; }
    .lane-row:hover{ border-color:#CBD6E8; background:#FBFCFE; }
    .lane-row.broken{ background:#FDF3F3; border-color:#F2C9C9; }
    .ctry{ font-family:'JetBrains Mono',monospace; font-size:.7rem; font-weight:700;
      color:var(--ink); width:38px; flex:none; }
    .ctry-full{ font-size:.75rem; color:var(--muted); width:132px; flex:none;
      overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
    .rail{ flex:1; height:1px; background:repeating-linear-gradient(90deg,
      var(--line) 0 5px, transparent 5px 10px); min-width:14px; }
    .hop{ font-family:'JetBrains Mono',monospace; font-size:.65rem; font-weight:600;
      color:var(--nc); border:1px solid var(--nc); background:color-mix(in srgb, var(--nc) 8%, white);
      border-radius:999px; padding:3px 10px; white-space:nowrap; }
    .hop.bk{ border-style:dashed; opacity:.8; }
    .hop.dead{ color:#B6BFCE; border-color:#DFE4EC; background:#F6F8FB;
      text-decoration:line-through; }
    .arw{ color:#C3CCDC; font-size:.85rem; }
    .warn{ font-family:'JetBrains Mono',monospace; font-size:.6rem; color:var(--rose);
      background:#FBE7E7; border-radius:999px; padding:3px 9px; letter-spacing:.06em; }

    /* rule strips (pack rounding + default target) */
    .rules{ display:flex; flex-direction:column; gap:9px; margin-top:4px; }
    .rule{ border:1px solid var(--line); border-radius:11px; padding:10px 13px;
      background:var(--surface); }
    .rule-head{ display:flex; align-items:baseline; gap:9px; margin-bottom:8px; }
    .rule-pct{ font-family:'Space Grotesk',sans-serif; font-size:1.15rem; font-weight:700;
      color:var(--ink); font-variant-numeric:tabular-nums; }
    .rule-cs{ font-family:'JetBrains Mono',monospace; font-size:.63rem; color:var(--muted);
      letter-spacing:.1em; }
    .strip{ display:flex; align-items:center; gap:4px; flex-wrap:wrap; }
    .pd{ width:9px; height:9px; border-radius:50%; flex:none; }
    .pd.dn{ background:#EDF1F7; border:1px solid #DCE3EE; }
    .pd.up{ background:linear-gradient(180deg,var(--teal),var(--accent)); }
    .pk{ width:15px; height:11px; border-radius:3px; flex:none;
      background:linear-gradient(180deg,var(--teal),var(--accent)); }
    .pk-more{ font-family:'JetBrains Mono',monospace; font-size:.62rem; color:var(--muted); }
    .strip-cap{ font-family:'JetBrains Mono',monospace; font-size:.61rem; color:var(--muted);
      margin-left:7px; }
    </style>
    """)


def render_hero():
    st.html("""
    <div class="hero">
      <div class="hero-eyebrow">Multi-Node &middot; Allocation &middot; Oracle Export</div>
      <div class="hero-title">GCC Replenishment Engine</div>
      <p class="hero-sub">Sales-ranked allocation across the source warehouse network, with a rolling 10% reserve held at every node.</p>
      <div class="lane">
        <span class="lane-node">DIP</span><span class="lane-node">KSA</span>
        <span class="lane-node">QAT</span><span class="lane-node">JAFZA</span>
        <span class="lane-dot"></span><span class="lane-track"></span><span class="lane-dot"></span>
        <span class="lane-node">6 GCC MARKETS</span>
      </div>
    </div>
    """)


def kpi(col, eyebrow, value, label, caption="", tone="", delay="d1", meter=None):
    m = f'<div class="meter"><span style="width:{max(0,min(100,meter)):.1f}%"></span></div>' if meter is not None else ""
    c = f'<div class="kpi-cap">{caption}</div>' if caption else ""
    with col:
        st.html(f"""<div class="kpi {tone} rise {delay}">
          <div class="kpi-eyebrow">{eyebrow}</div><div class="kpi-value">{value}</div>
          <div class="kpi-label">{label}</div>{c}{m}</div>""")


def render_kpis(s, runtime):
    r1 = st.columns(4)
    kpi(r1[0], "Transfer Volume", f"{s['units']:,}", "Units to ship", delay="d1")
    kpi(r1[1], "Order Lines", f"{s['lines']:,}", "Store x item lines", tone="teal", delay="d2")
    kpi(r1[2], "Coverage", f"{s['items']:,}", "Items covered",
        caption=f"across {s['stores']:,} stores", tone="amber", delay="d3")
    kpi(r1[3], "Demand Fill Rate", f"{s['fill']:.1f}%", "Of total rounded need",
        caption=f"{s['got']:,} of {s['need']:,} units", tone="teal", delay="d4", meter=s["fill"])
    r2 = st.columns(4)
    kpi(r2[0], "Reserve Carried Forward", f"{s['reserve']:,}", "Units held across all WHs",
        caption="secured for the next run", tone="amber", delay="d4")
    kpi(r2[1], "Fallback Sourced", f"{s['fallback']:,}", "Lines from a backup WH",
        caption="priority-1 warehouse could not cover", tone="rose", delay="d5")
    kpi(r2[2], "Unmet Demand", f"{s['unmet_units']:,}", "Units with no stock",
        caption=f"{s['unfilled']:,} store-item lines", tone="rose", delay="d6")
    kpi(r2[3], "Run Time", f"{runtime:.1f}s", "Engine execution",
        caption="excludes file reading", tone="teal", delay="d6")

    if s["defaulted"]:
        st.html(f"""<div class="alert rise d6">
          <b>{s['defaulted']:,} option x store combinations sold but have no row in the DEPTH
          file</b>, so the default target stock grid was applied to size them. They contributed
          <b>{s['defaulted_units']:,} units</b> to this plan. Check the Target Source column in
          the audit trail to see which lines these are.</div>""")


# ---------------------------------------------------------------------------
# APP
# ---------------------------------------------------------------------------
st.set_page_config(page_title="GCC Replenishment Engine", page_icon="\U0001F4E6", layout="wide")
inject_theme()
render_hero()

for k in ("final", "detail", "wh_summary", "summary", "plan_bytes",
          "detail_bytes", "runtime", "read_time"):
    st.session_state.setdefault(k, None)

# --- Step 1: rules --------------------------------------------------------
WH_ACCENT = {"44324": "#2F6FED", "22748": "#0EA5A0", "170001": "#E0A008", "242211": "#8B5CF6"}
COUNTRY_SHORT = {"Kuwait": "KUW", "Qatar": "QAT", "Bahrain": "BAH",
                 "Oman": "OMN", "Saudi Arabia": "KSA", "United Arab Emirates": "UAE"}


def dot_strip(threshold_pct):
    """A pack of 12 drawn as 11 remainder dots: hollow = rounds down, solid = rounds up."""
    cut = PACK_SIZE * threshold_pct / 100.0
    dots = []
    for r in range(1, PACK_SIZE):
        up = r >= cut
        dots.append(f'<span class="pd {"up" if up else "dn"}"></span>')
    return (f'<div class="strip">{"".join(dots)}'
            f'<span class="strip-cap">&ge;{cut:.1f} rounds up</span></div>')


with st.container(border=True):
    st.html('<div class="step">Step 1 &mdash; Allocation rules</div>')
    st.markdown('<p class="sec-title">Control deck</p>', unsafe_allow_html=True)
    st.markdown('<p class="sec-sub">Everything below has a working default. Touch it only '
                'to override this run.</p>', unsafe_allow_html=True)

    ribbon = st.empty()

    # ---- source warehouses -------------------------------------------------
    st.html('<div class="grp">Source warehouses</div>')
    wh_cols = st.columns(len(GRID_WH))
    active_wh = []
    for i, (code, name) in enumerate(GRID_WH.items()):
        serves = [COUNTRY_SHORT[c] for c, ch in sorted(COUNTRY_PRIORITY.items()) if code in ch]
        primary = [COUNTRY_SHORT[c] for c, ch in sorted(COUNTRY_PRIORITY.items())
                   if ch and ch[0] == code]
        with wh_cols[i]:
            on = st.checkbox(name, value=True, key=f"wh_{code}")
            st.html(f"""
            <div class="node {'' if on else 'off'}" style="--nc:{WH_ACCENT[code]}">
              <div class="node-code">{code}</div>
              <div class="node-serves">{' '.join(serves) if serves else '&mdash;'}</div>
              <div class="node-meta">{len(primary)} primary &middot; {len(serves)-len(primary)} backup</div>
            </div>""")
        if on:
            active_wh.append(code)

    # ---- routing lanes -----------------------------------------------------
    st.html('<div class="grp">Routing</div>')
    lane_html = ['<div class="lanes">']
    for country, chain in sorted(COUNTRY_PRIORITY.items()):
        usable = [w for w in chain if w in active_wh]
        broken = not usable
        hops = []
        for rank, w in enumerate(chain, start=1):
            live = w in active_wh
            cls = "hop" + ("" if live else " dead") + (" bk" if rank > 1 else "")
            hops.append(f'<span class="{cls}" style="--nc:{WH_ACCENT[w]}">{GRID_WH[w]}</span>')
            if rank < len(chain):
                hops.append('<span class="arw">&rsaquo;</span>')
        lane_html.append(
            f'<div class="lane-row{" broken" if broken else ""}">'
            f'<span class="ctry">{COUNTRY_SHORT[country]}</span>'
            f'<span class="ctry-full">{country}</span>'
            f'<span class="rail"></span>{"".join(hops)}'
            f'{"<span class=warn>no source</span>" if broken else ""}</div>')
    lane_html.append('</div>')
    st.html("".join(lane_html))

    # ---- rounding + targets ------------------------------------------------
    c_left, c_right = st.columns([1, 1])

    with c_left:
        st.html('<div class="grp">Pack rounding</div>')
        customise = st.toggle("Override thresholds", value=False,
                              help="A remainder at or above the cut rounds up to a full pack.")
        grid_df = pd.DataFrame([{"Country": c, "Threshold %": p}
                                for c, p in DEFAULT_PACK_THRESHOLD_PCT.items()]
                               ).sort_values("Country").reset_index(drop=True)
        if customise:
            edited = st.data_editor(
                grid_df, key="thr_editor", hide_index=True, use_container_width=True,
                column_config={
                    "Country": st.column_config.TextColumn("Country", disabled=True),
                    "Threshold %": st.column_config.NumberColumn(
                        "Threshold %", min_value=0, max_value=100, step=5, format="%d%%"),
                })
            if st.button("Reset thresholds", use_container_width=True):
                st.session_state.pop("thr_editor", None)
                st.rerun()
        else:
            edited = grid_df

        pack_thresholds = {}
        for _, r in edited.iterrows():
            v = pd.to_numeric(r["Threshold %"], errors="coerce")
            v = FALLBACK_THRESHOLD_PCT if pd.isna(v) else min(100, max(0, float(v)))
            pack_thresholds[r["Country"]] = v

        groups = {}
        for c, p in sorted(pack_thresholds.items()):
            groups.setdefault(p, []).append(COUNTRY_SHORT.get(c, c))
        blocks = []
        for p, cs in sorted(groups.items()):
            blocks.append(f'<div class="rule"><div class="rule-head">'
                          f'<span class="rule-pct">{p:.0f}%</span>'
                          f'<span class="rule-cs">{" ".join(cs)}</span></div>'
                          f'{dot_strip(p)}</div>')
        st.html(f'<div class="rules">{"".join(blocks)}</div>')

    with c_right:
        st.html('<div class="grp">Default target stock</div>')
        tgt_customise = st.toggle("Override targets", value=False,
                                  help="Used when an Option + Store has no row in DEPTH.")
        tgt_df = pd.DataFrame([{"Country": c, "Default Target Stock": v}
                               for c, v in DEFAULT_TARGET_STOCK.items()]
                              ).sort_values("Country").reset_index(drop=True)
        if tgt_customise:
            tgt_edited = st.data_editor(
                tgt_df, key="tgt_editor", hide_index=True, use_container_width=True,
                column_config={
                    "Country": st.column_config.TextColumn("Country", disabled=True),
                    "Default Target Stock": st.column_config.NumberColumn(
                        "Default Target Stock", min_value=0, max_value=999, step=12),
                })
            if st.button("Reset targets", use_container_width=True):
                st.session_state.pop("tgt_editor", None)
                st.rerun()
        else:
            tgt_edited = tgt_df

        default_targets = {}
        for _, r in tgt_edited.iterrows():
            v = pd.to_numeric(r["Default Target Stock"], errors="coerce")
            v = FALLBACK_TARGET_STOCK if pd.isna(v) else max(0, float(v))
            default_targets[r["Country"]] = v

        tgroups = {}
        for c, v in sorted(default_targets.items()):
            tgroups.setdefault(v, []).append(COUNTRY_SHORT.get(c, c))
        tblocks = []
        for v, cs in sorted(tgroups.items()):
            packs = int(v // PACK_SIZE)
            rem = v - packs * PACK_SIZE
            cells = "".join('<span class="pk"></span>' for _ in range(min(packs, 12)))
            if packs > 12:
                cells += f'<span class="pk-more">+{packs-12}</span>'
            tblocks.append(
                f'<div class="rule"><div class="rule-head">'
                f'<span class="rule-pct">{v:.0f}</span>'
                f'<span class="rule-cs">{" ".join(cs)}</span></div>'
                f'<div class="strip">{cells}'
                f'<span class="strip-cap">{packs} pack{"s" if packs!=1 else ""} of {PACK_SIZE}'
                f'{f" + {rem:.0f}" if rem else ""}</span></div></div>')
        st.html(f'<div class="rules">{"".join(tblocks)}</div>')

    # ---- live config ribbon -------------------------------------------------
    thr_txt = " / ".join(f"{p:.0f}%" for p in sorted(groups))
    tgt_txt = " / ".join(f"{v:.0f}" for v in sorted(tgroups))
    reachable = sum(1 for ch in COUNTRY_PRIORITY.values() if any(w in active_wh for w in ch))
    ribbon.html(f"""
    <div class="ribbon">
      <span class="rb"><b>{len(active_wh)}</b>/{len(GRID_WH)} warehouses</span>
      <span class="rb"><b>{reachable}</b>/{len(COUNTRY_PRIORITY)} markets reachable</span>
      <span class="rb">pack <b>{PACK_SIZE}</b></span>
      <span class="rb">thresholds <b>{thr_txt}</b></span>
      <span class="rb">default target <b>{tgt_txt}</b></span>
      <span class="rb">reserve <b>{int(WMS_RESERVE_PCT*100)}%</b> held</span>
    </div>""")
# --- Step 2: files --------------------------------------------------------
with st.container(border=True):
    st.html('<div class="step">Step 2 &mdash; Source data</div>')
    st.markdown('<p class="sec-title">Upload the five extracts</p>', unsafe_allow_html=True)
    st.markdown('<p class="sec-sub">DEPTH, RMS, SALE, SOH and WMS. Each file is identified '
                'automatically from its columns, so upload order does not matter.</p>',
                unsafe_allow_html=True)

    uploads = st.file_uploader("Excel files (.xlsx)", type=["xlsx", "xls"],
                               accept_multiple_files=True, label_visibility="collapsed")

    if uploads:
        payload = tuple((f.name, f.getvalue()) for f in uploads)
        t_read = time.time()
        with st.spinner("Reading and identifying files..."):
            frames, timings, read_warnings = load_sources(payload)
        read_secs = time.time() - t_read
        st.session_state.read_time = read_secs

        for w in read_warnings:
            st.warning(w)

        cols = st.columns(5)
        for i, kind in enumerate(["DEPTH", "RMS", "SALE", "SOH", "WMS"]):
            with cols[i]:
                if kind in frames:
                    st.metric(kind, f"{len(frames[kind]):,}",
                              help=f"rows parsed in {timings.get(kind,0):.1f}s")
                else:
                    st.metric(kind, "—")
        st.caption(f"Parsed in {read_secs:.1f}s. Files are cached, so re-running does not "
                   f"re-read them.")

        if st.button("Run Replenishment", type="primary", disabled=not frames):
            bar, txt = st.empty(), st.empty()

            def progress_cb(frac, label):
                bar.progress(min(1.0, max(0.0, frac)))
                txt.markdown(
                    f"<span style='font-family:JetBrains Mono,monospace;font-size:.8rem;"
                    f"color:#2F6FED;'>{int(frac*100):3d}%</span> "
                    f"<span style='font-size:.86rem;color:#7A88A6;'>{label}</span>",
                    unsafe_allow_html=True)

            try:
                final, detail, wh_summary, warns, n_defaulted, secs = run_engine(
                    frames, set(active_wh), pack_thresholds, default_targets, progress_cb)

                st.session_state.final = final
                st.session_state.detail = detail
                st.session_state.wh_summary = wh_summary
                st.session_state.summary = summarise(final, detail, wh_summary, n_defaulted)
                st.session_state.runtime = secs
                st.session_state.plan_bytes = to_excel(final, "Replenishment")
                st.session_state.detail_bytes = to_excel(detail, "CalculationDetail")

                bar.empty(); txt.empty()
                for w in warns:
                    st.warning(w)
                st.success(f"Replenishment complete - {len(final):,} transfer lines in {secs:.1f}s "
                           f"(files parsed in {read_secs:.1f}s).")
            except ValueError as e:
                for k in ("final", "detail", "wh_summary", "summary"):
                    st.session_state[k] = None
                bar.empty(); txt.empty(); st.error(str(e))
            except Exception as e:
                for k in ("final", "detail", "wh_summary", "summary"):
                    st.session_state[k] = None
                bar.empty(); txt.empty()
                st.error(f"Something went wrong ({type(e).__name__}: {e}).")
                with st.expander("Technical details"):
                    st.code(traceback.format_exc())

# --- Step 3 / 4 -----------------------------------------------------------
if st.session_state.final is not None:
    final = st.session_state.final
    detail = st.session_state.detail
    s = st.session_state.summary

    st.html('<div class="step" style="margin-top:16px;">Step 3 &mdash; Run summary</div>')
    render_kpis(s, st.session_state.runtime)

    with st.container(border=True):
        st.markdown('<p class="sec-title">Warehouse network</p>', unsafe_allow_html=True)
        st.markdown('<p class="sec-sub">What each node shipped, and what it is holding back.</p>',
                    unsafe_allow_html=True)
        st.dataframe(st.session_state.wh_summary, use_container_width=True, hide_index=True)

    with st.container(border=True):
        st.markdown('<p class="sec-title">By country</p>', unsafe_allow_html=True)
        by_c = detail.groupby("Country", as_index=False).agg(
            Lines=("Option", "size"), Need=("Rounded Need", "sum"),
            Allocated=("Allocated Qty", "sum"))
        by_c["Fill %"] = (by_c["Allocated"] / by_c["Need"].replace(0, np.nan) * 100).round(1)
        st.dataframe(by_c, use_container_width=True, hide_index=True)

    st.html('<div class="step" style="margin-top:24px;">Step 4 &mdash; Export</div>')
    with st.container(border=True):
        head, dl = st.columns([3, 1])
        with head:
            st.markdown('<p class="sec-title">Oracle transfer plan</p>', unsafe_allow_html=True)
            st.markdown('<p class="sec-sub">FROM LOCATION and TO LOCATION are location codes '
                        '(Loc Key), ready to import.</p>', unsafe_allow_html=True)
        with dl:
            st.download_button("Download Transfer Plan", data=st.session_state.plan_bytes,
                               file_name="gcc_replenishment_transfer_plan.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               type="primary", use_container_width=True)
        st.dataframe(final.head(5000), use_container_width=True, hide_index=True)
        if len(final) > 5000:
            st.caption(f"Showing the first 5,000 of {len(final):,} lines. The download has all.")

    with st.expander("Calculation detail - audit trail", expanded=False):
        st.markdown('<p class="sec-sub">Every intermediate column behind the allocation: '
                    'country, sales rank, pipeline, failsafes, pack threshold, source '
                    'warehouse and why.</p>', unsafe_allow_html=True)
        st.download_button("Download Audit Trail", data=st.session_state.detail_bytes,
                           file_name="gcc_replenishment_audit_trail.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(detail.head(20000), use_container_width=True, hide_index=True)
        if len(detail) > 20000:
            st.caption(f"Showing the first 20,000 of {len(detail):,} rows for page performance.")
