"""
Focus Factory Failure Drivers Dashboard — SANITIZED PUBLIC VERSION
==================================================================
Streamlit dashboard that ingests a manufacturing test-data export (Excel) and
computes TRUE First Pass Yield and failure-driver analytics across ALL
departments ("Focus Factories") present in the file:

  * True FPY         — a serial counts as passed only if it had ZERO failures
                       across ALL process stages
  * FPY by department, yield gauges, DPMO
  * Top failure drivers per department (stage, station, codes, parts, ATE step...)
  * Combo drivers, Pareto 80/20, SPC control charts, trend analysis
  * Repeat-offender detection (chronic Part + Station + Failure Code combos)
  * Interactive root-cause drill-down tree per department
  * Failure-analysis summary tables with editable Actions/Status
  * Multi-sheet Excel exports of every view

Data handling is schema-tolerant: column-name aliases are matched to a
canonical schema, the best sheet/header row is auto-detected, and optional
serial-level deduplication keeps only the FIRST test per serial + stage.

All examples, comments, and the bundled sample data use fictional
placeholders — no real product names, part numbers, serials, or personnel.

Run with:  streamlit run fpy_dashboard.py
Sample data: python generate_sample_data.py  (creates SAMPLE_KPI_DATA.xlsx)
"""

import io
import re
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# -----------------------------
# Column aliases (handles slightly different exports)
# -----------------------------
COLUMN_ALIASES = {
    "FF": ["ff", "focus factory", "department", "dept", "focus_factory"],
    "Pass or Fail": ["pass or fail", "result", "status", "pass/fail", "passfail"],
    "Process Stage": ["process stage", "stage", "test stage", "process", "test phase", "phase"],
    "Part No": ["part no", "part number", "pn", "p/n", "assy pn", "assembly part number"],
    "Failure Code": ["failure code", "fail code", "failure", "failure mode"],
    "Defect Code": ["defect code", "defect", "defect type"],
    "Station": ["station", "test station", "tester", "bench"],
    "Failed ATE Step": ["failed ate step", "ate step", "failed step", "test step", "ate fail step"],
    "Failed Assy": ["failed assy", "failed assembly", "assembly", "assy"],
    "Reference Designator": ["reference designator", "ref des", "refdes", "reference designator(s)", "ref designator"],
    "Tested By": ["tested by", "operator", "tester name", "tech", "tested_by"],
    "Troubleshoot Completed By": ["troubleshoot completed by", "troubleshoot by", "repaired by", "tech troubleshooting"],
    "Date Tested": ["date tested", "test date", "date", "timestamp", "date/time", "datetime"],
    "Tech Comments": ["tech comments", "comments", "notes", "failure notes", "observations"],
    "Rework Instructions": ["rework instructions", "rework", "action", "corrective action", "repair action"],
    "Serial Number": ["serial number", "serial no", "serial", "sn", "s/n", "serial_number", "serial #", "unit serial", "unit sn"],
}

DRIVER_FIELDS = [
    "Process Stage",
    "Failure Code",
    "Defect Code",
    "Part No",
    "Station",
    "Failed ATE Step",
    "Failed Assy",
    "Reference Designator",
    "Tested By",
    "Troubleshoot Completed By",
]

COMBO_DEFS = {
    "Stage + Failure + Defect": ["Process Stage", "Failure Code", "Defect Code"],
    "Stage + Station": ["Process Stage", "Station"],
    "Part No + Failure Code": ["Part No", "Failure Code"],
}

# Alert thresholds for KPIs
DEFAULT_THRESHOLDS = {
    "yield_warning": 95.0,  # Yellow below this
    "yield_critical": 90.0,  # Red below this
    "failure_rate_warning": 5.0,  # Yellow above this %
    "failure_rate_critical": 10.0,  # Red above this %
}


# -----------------------------
# Helpers
# -----------------------------
def normalize_colname(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def find_header_row(df_head: pd.DataFrame, required_norm_names):
    best_row, best_score = None, -1
    for r in range(min(len(df_head), 50)):
        row_vals = [normalize_colname(x) for x in df_head.iloc[r].tolist()]
        row_set = set([v for v in row_vals if v])
        score = sum(1 for req in required_norm_names if req in row_set)
        if score > best_score:
            best_score, best_row = score, r
        if best_score >= len(required_norm_names):
            break
    return best_row, best_score


def build_canonical_column_map(columns):
    norm_to_actual = {normalize_colname(c): c for c in columns}
    canon_map = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        for a in aliases:
            a_norm = normalize_colname(a)
            if a_norm in norm_to_actual:
                canon_map[canonical] = norm_to_actual[a_norm]
                break
    return canon_map


def load_best_sheet(file_like) -> pd.DataFrame:
    xls = pd.ExcelFile(file_like, engine="openpyxl")
    required_candidates = ["part no", "process stage", "pass or fail"]

    best = {"sheet": None, "header_row": None, "score": -1}
    for sheet in xls.sheet_names:
        tmp = pd.read_excel(file_like, sheet_name=sheet, header=None, nrows=50, engine="openpyxl")
        header_row, score = find_header_row(tmp, required_candidates)
        if score > best["score"]:
            best = {"sheet": sheet, "header_row": header_row, "score": score}

    if best["sheet"] is None:
        raise ValueError("No sheets found.")

    df = pd.read_excel(file_like, sheet_name=best["sheet"], header=best["header_row"], engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def clean_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def standardize_df(df: pd.DataFrame):
    canon_map = build_canonical_column_map(df.columns)

    for canonical, actual in canon_map.items():
        df[canonical] = df[actual]

    # Ensure minimum columns exist
    for needed in ["FF", "Pass or Fail", "Process Stage", "Part No"]:
        if needed not in df.columns:
            df[needed] = ""

    # Clean key cols
    for c in ["FF", "Pass or Fail", "Process Stage", "Part No"]:
        df[c] = df[c].apply(clean_str)

    df["Pass or Fail"] = df["Pass or Fail"].str.upper()
    df["FF"] = df["FF"].replace("", "UNKNOWN").fillna("UNKNOWN")

    # Date parse if present
    if "Date Tested" in df.columns:
        df["Date Tested"] = pd.to_datetime(df["Date Tested"], errors="coerce")

    return df, canon_map


def deduplicate_for_fpy(df: pd.DataFrame, id_col: str = "Serial Number"):
    """Keep only the FIRST test result per serial number (+ process stage) for true FPY.

    For each unique serial at each process stage, sort by Date Tested and keep
    the first record.  If the first attempt is a FAIL it stays FAIL;
    if it's a PASS it stays PASS.  Subsequent retests / duplicate passes are dropped.
    """
    if id_col not in df.columns:
        return df.copy()

    tmp = df.copy()

    # Sort by date so the earliest test comes first
    if "Date Tested" in tmp.columns and not tmp["Date Tested"].isna().all():
        tmp = tmp.sort_values("Date Tested", ascending=True)

    # Build dedup key: serial + process stage (if available) so a unit tested
    # at multiple stages keeps one record per stage.
    dedup_cols = [id_col]
    if "Process Stage" in tmp.columns:
        dedup_cols.append("Process Stage")

    # Drop rows where serial is blank/NaN — can't dedup without an identifier
    tmp["_serial_clean"] = tmp[id_col].fillna("").astype(str).str.strip()
    has_serial = tmp["_serial_clean"] != ""
    df_with_serial = tmp[has_serial].drop_duplicates(subset=dedup_cols, keep="first")
    df_without_serial = tmp[~has_serial]  # keep rows with no serial as-is

    result = pd.concat([df_with_serial, df_without_serial], ignore_index=True)
    result = result.drop(columns=["_serial_clean"])
    return result


def filter_failures(df: pd.DataFrame, include_all=False):
    if include_all:
        return df.copy()

    # If Pass/Fail mostly blank, assume it's failure-only export
    non_blank_ratio = (df["Pass or Fail"].astype(str).str.strip() != "").mean()
    if non_blank_ratio < 0.2:
        return df.copy()

    return df[df["Pass or Fail"].isin(["FAIL", "F", "FAILED"])].copy()


def top_n_counts(df, group_cols, top_n=10):
    if isinstance(group_cols, str):
        group_cols = [group_cols]

    tmp = df.copy()
    for c in group_cols:
        if c not in tmp.columns:
            tmp[c] = ""
        tmp[c] = tmp[c].fillna("").astype(str).str.strip().replace("", "(blank)")

    counts = (
        tmp.groupby(["FF"] + group_cols)
           .size()
           .reset_index(name="Count")
           .sort_values(["FF", "Count"], ascending=[True, False])
    )
    counts["Rank"] = counts.groupby("FF")["Count"].rank(method="first", ascending=False).astype(int)
    counts = counts[counts["Rank"] <= top_n].copy()

    totals = tmp.groupby("FF").size().rename("FF_Total").reset_index()
    counts = counts.merge(totals, on="FF", how="left")
    counts["% of FF"] = (counts["Count"] / counts["FF_Total"]).replace([np.inf, -np.inf], np.nan)

    if len(group_cols) > 1:
        counts["Item"] = counts[group_cols].agg(" | ".join, axis=1)
    else:
        counts["Item"] = counts[group_cols[0]]

    return counts[["FF", "Rank", "Item", "Count", "% of FF"]]


def to_excel_download(dfs: dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in dfs.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    output.seek(0)
    return output


# -----------------------------
# Advanced KPI Calculations
# -----------------------------
def calculate_yield_metrics(df_all: pd.DataFrame, df_fail: pd.DataFrame):
    """Calculate true First Pass Yield — a serial passes only if it had
    ZERO failures across ALL process stages."""
    metrics = {}

    id_col = "Serial Number"

    # --- Fallback to row-level if no Serial Number column ---
    if id_col not in df_all.columns:
        total_tested = len(df_all)
        total_failed = len(df_fail)
        total_passed = total_tested - total_failed
        metrics["fpy"] = (total_passed / total_tested * 100) if total_tested else 0
        metrics["failure_rate"] = (total_failed / total_tested * 100) if total_tested else 0
        metrics["dpmo"] = (total_failed / total_tested * 1_000_000) if total_tested else 0
        metrics["total_tested"] = total_tested
        metrics["total_failed"] = total_failed
        metrics["total_passed"] = total_passed
        metrics["note"] = "Row-level (no Serial Number column)"
        return metrics

    # --- True FPY: serial-level ---
    # All unique serials that were tested (exclude blank serials)
    all_serials = (
        df_all[id_col]
        .fillna("").astype(str).str.strip()
    )
    all_serials = set(all_serials[all_serials != ""])

    # Serials that had at least one FAIL
    failed_serials = (
        df_fail[id_col]
        .fillna("").astype(str).str.strip()
    )
    failed_serials = set(failed_serials[failed_serials != ""])

    # Serials with ZERO failures across all stages
    passed_serials = all_serials - failed_serials

    total_tested = len(all_serials)
    total_passed = len(passed_serials)
    total_failed = len(failed_serials)

    metrics["fpy"] = (total_passed / total_tested * 100) if total_tested else 0
    metrics["failure_rate"] = (total_failed / total_tested * 100) if total_tested else 0
    metrics["dpmo"] = (total_failed / total_tested * 1_000_000) if total_tested else 0
    metrics["total_tested"] = total_tested
    metrics["total_failed"] = total_failed
    metrics["total_passed"] = total_passed
    metrics["note"] = "Serial-level FPY"

    return metrics

def calculate_yield_by_group(df_all: pd.DataFrame, group_col: str):
    """Calculate true FPY grouped by a column (e.g., FF).
    A serial is 'passed' only if it has ZERO failures within that group."""
    if group_col not in df_all.columns:
        return pd.DataFrame()

    id_col = "Serial Number"
    has_serial = (
        id_col in df_all.columns
        and df_all[id_col].fillna("").astype(str).str.strip().ne("").any()
    )

    fail_mask = df_all["Pass or Fail"].isin(["FAIL", "F", "FAILED"])

    rows = []
    for group_val, grp in df_all.groupby(group_col):
        if has_serial:
            serials = set(
                grp[id_col].fillna("").astype(str).str.strip()
            ) - {""}
            grp_fails = grp[grp["Pass or Fail"].isin(["FAIL", "F", "FAILED"])]
            failed_serials = set(
                grp_fails[id_col].fillna("").astype(str).str.strip()
            ) - {""}

            total = len(serials)
            failed = len(failed_serials)
            passed = total - failed
        else:
            total = len(grp)
            failed = grp["Pass or Fail"].isin(["FAIL", "F", "FAILED"]).sum()
            passed = total - failed

        rows.append({
            group_col: group_val,
            "Total_Tested": total,
            "Total_Failed": failed,
            "Total_Passed": passed,
            "FPY_%": round(passed / total * 100, 2) if total else 0,
            "Failure_Rate_%": round(failed / total * 100, 2) if total else 0,
            "DPMO": round(failed / total * 1_000_000) if total else 0,
        })

    return pd.DataFrame(rows).sort_values("FPY_%", ascending=True)


def calculate_trend_metrics(df: pd.DataFrame, date_col: str, period: str = "W"):
    """Calculate weekly or monthly trend with WoW/MoM changes."""
    if date_col not in df.columns or df[date_col].isna().all():
        return pd.DataFrame()

    tmp = df.copy()
    tmp[date_col] = pd.to_datetime(tmp[date_col], errors="coerce")
    tmp = tmp.dropna(subset=[date_col])

    if period == "W":
        tmp["Period"] = tmp[date_col].dt.to_period("W").dt.start_time
        label = "Week"
    else:
        tmp["Period"] = tmp[date_col].dt.to_period("M").dt.start_time
        label = "Month"

    trend = tmp.groupby("Period").size().reset_index(name="Count")
    trend = trend.sort_values("Period")

    # Calculate period-over-period change
    trend["Prev_Count"] = trend["Count"].shift(1)
    trend["Change"] = trend["Count"] - trend["Prev_Count"]
    trend["Change_%"] = ((trend["Change"] / trend["Prev_Count"]) * 100).round(1)
    trend["Change_%"] = trend["Change_%"].fillna(0)

    # Rolling average (4-period)
    trend["Rolling_Avg"] = trend["Count"].rolling(window=4, min_periods=1).mean().round(1)

    return trend


def detect_repeat_failures(df: pd.DataFrame, identifier_cols: list):
    """Detect parts/serials that have failed multiple times."""
    available_cols = [c for c in identifier_cols if c in df.columns]
    if not available_cols:
        return pd.DataFrame()

    # Use first available identifier
    id_col = available_cols[0]

    repeat_counts = df.groupby(id_col).size().reset_index(name="Failure_Count")
    repeat_counts = repeat_counts[repeat_counts["Failure_Count"] > 1]
    repeat_counts = repeat_counts.sort_values("Failure_Count", ascending=False)

    return repeat_counts


def calculate_pareto_80_20(df: pd.DataFrame, column: str):
    """Calculate Pareto with 80/20 rule identification."""
    if column not in df.columns:
        return pd.DataFrame(), 0

    tmp = df[column].fillna("").astype(str).str.strip().replace("", "(blank)")
    vc = tmp.value_counts().reset_index()
    vc.columns = [column, "Count"]

    total = vc["Count"].sum()
    vc["Cumulative"] = vc["Count"].cumsum()
    vc["Cum_%"] = (vc["Cumulative"] / total * 100).round(1)
    vc["Contribution_%"] = (vc["Count"] / total * 100).round(1)

    # Find how many items make up 80% of failures
    items_for_80 = len(vc[vc["Cum_%"] <= 80]) + 1

    return vc, items_for_80


def calculate_spc_limits(series: pd.Series):
    """Calculate Statistical Process Control limits."""
    mean = series.mean()
    std = series.std()

    ucl = mean + 3 * std  # Upper Control Limit
    lcl = max(0, mean - 3 * std)  # Lower Control Limit (can't be negative for counts)
    uwl = mean + 2 * std  # Upper Warning Limit
    lwl = max(0, mean - 2 * std)  # Lower Warning Limit

    return {
        "mean": mean,
        "ucl": ucl,
        "lcl": lcl,
        "uwl": uwl,
        "lwl": lwl
    }


def create_gauge_chart(value: float, title: str, thresholds: dict = None):
    """Create a gauge chart for yield or other percentage metrics."""
    if thresholds is None:
        thresholds = {"warning": 95, "critical": 90}

    # Determine color based on thresholds
    if value >= thresholds["warning"]:
        color = "green"
    elif value >= thresholds["critical"]:
        color = "orange"
    else:
        color = "red"

    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=value,
        domain={'x': [0, 1], 'y': [0, 1]},
        title={'text': title, 'font': {'size': 16}},
        number={'suffix': '%', 'font': {'size': 28}},
        gauge={
            'axis': {'range': [0, 100], 'tickwidth': 1},
            'bar': {'color': color},
            'bgcolor': "white",
            'borderwidth': 2,
            'steps': [
                {'range': [0, thresholds["critical"]], 'color': 'rgba(255,0,0,0.2)'},
                {'range': [thresholds["critical"], thresholds["warning"]], 'color': 'rgba(255,165,0,0.2)'},
                {'range': [thresholds["warning"], 100], 'color': 'rgba(0,255,0,0.2)'}
            ],
            'threshold': {
                'line': {'color': "black", 'width': 2},
                'thickness': 0.75,
                'value': value
            }
        }
    ))
    fig.update_layout(height=250, margin=dict(l=20, r=20, t=40, b=20))
    return fig


def create_control_chart(trend_df: pd.DataFrame, value_col: str, date_col: str):
    """Create an SPC control chart."""
    if trend_df.empty:
        return None

    limits = calculate_spc_limits(trend_df[value_col])

    fig = go.Figure()

    # Main data line
    fig.add_trace(go.Scatter(
        x=trend_df[date_col],
        y=trend_df[value_col],
        mode='lines+markers',
        name='Failures',
        line=dict(color='blue', width=2),
        marker=dict(size=8)
    ))

    # Mean line
    fig.add_hline(y=limits["mean"], line_dash="solid", line_color="green",
                  annotation_text=f"Mean: {limits['mean']:.1f}")

    # Control limits
    fig.add_hline(y=limits["ucl"], line_dash="dash", line_color="red",
                  annotation_text=f"UCL: {limits['ucl']:.1f}")
    fig.add_hline(y=limits["lcl"], line_dash="dash", line_color="red",
                  annotation_text=f"LCL: {limits['lcl']:.1f}")

    # Warning limits
    fig.add_hline(y=limits["uwl"], line_dash="dot", line_color="orange",
                  annotation_text=f"UWL: {limits['uwl']:.1f}")
    fig.add_hline(y=limits["lwl"], line_dash="dot", line_color="orange")

    # Highlight out-of-control points
    ooc_mask = (trend_df[value_col] > limits["ucl"]) | (trend_df[value_col] < limits["lcl"])
    if ooc_mask.any():
        fig.add_trace(go.Scatter(
            x=trend_df.loc[ooc_mask, date_col],
            y=trend_df.loc[ooc_mask, value_col],
            mode='markers',
            name='Out of Control',
            marker=dict(color='red', size=12, symbol='x')
        ))

    fig.update_layout(
        title="Statistical Process Control Chart",
        xaxis_title="Period",
        yaxis_title="Failure Count",
        height=400
    )

    return fig


def create_sunburst_chart(df: pd.DataFrame, path_cols: list, values_col: str = None):
    """Create a sunburst chart for hierarchical failure breakdown."""
    available_cols = [c for c in path_cols if c in df.columns]
    if len(available_cols) < 2:
        return None

    tmp = df.copy()
    for c in available_cols:
        tmp[c] = tmp[c].fillna("").astype(str).str.strip().replace("", "(blank)")

    # Aggregate counts
    agg = tmp.groupby(available_cols).size().reset_index(name="Count")

    fig = px.sunburst(
        agg,
        path=available_cols,
        values="Count",
        title="Hierarchical Failure Breakdown"
    )
    fig.update_layout(height=500)
    return fig


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Focus Factory Failure Drivers Dashboard", layout="wide")

st.title("📊 Focus Factory Failure Drivers — Interactive Dashboard")
st.caption("Upload your failure export (same structure/headers). Filter by FF, date, stage, station, codes, parts. View Top drivers + combos + trends.")

with st.sidebar:
    st.header("1) Load Data")
    uploaded = st.file_uploader("Upload Excel export", type=["xlsx", "xlsm", "xls"])
    include_all = st.checkbox("Include PASS rows (usually leave OFF)", value=False)
    dedup_fpy = st.checkbox("Deduplicate for true FPY (first test per serial/part)", value=True,
                            help="Keeps only the first test result per serial/part number. "
                                 "Removes duplicate PASS and re-test records so yield reflects first-pass only.")
    top_n = st.slider("Top N drivers per metric", min_value=5, max_value=30, value=10, step=1)

    st.divider()
    st.header("⚙️ Alert Thresholds")
    yield_warning = st.number_input("Yield Warning (%)", min_value=0.0, max_value=100.0, value=95.0, step=0.5)
    yield_critical = st.number_input("Yield Critical (%)", min_value=0.0, max_value=100.0, value=90.0, step=0.5)

if not uploaded:
    st.info("Upload an Excel file to begin.\n\n"
            "Don't have data? Run `python generate_sample_data.py` to create "
            "a synthetic SAMPLE_KPI_DATA.xlsx covering multiple departments.")
    st.stop()

# Load and standardize
try:
    df_raw = load_best_sheet(uploaded)
    df, colmap = standardize_df(df_raw)

    # Determine the best identifier column for deduplication
    _fpy_id_col = None
    if dedup_fpy:
        # Must use Serial Number — Part No is a part type, not a unique unit ID
        if "Serial Number" in df.columns:
            _non_blank = df["Serial Number"].fillna("").astype(str).str.strip().ne("").sum()
            if _non_blank > 0:
                _fpy_id_col = "Serial Number"

    if dedup_fpy and _fpy_id_col:
        _before = len(df)
        df = deduplicate_for_fpy(df, id_col=_fpy_id_col)
        _after = len(df)
        _dropped = _before - _after
        if _dropped > 0:
            st.sidebar.success(
                f"FPY dedup on **{_fpy_id_col}** + Process Stage: "
                f"kept {_after:,} of {_before:,} rows "
                f"({_dropped:,} duplicate retests removed)"
            )
        else:
            st.sidebar.info("FPY dedup enabled but no duplicates found.")
    elif dedup_fpy:
        st.sidebar.warning(
            "No **Serial Number** (or Serial No / SN) column found in your data — "
            "dedup requires a unique unit identifier. Check your column headers."
        )

    df_fail = filter_failures(df, include_all=include_all)

    # Calculate yield metrics (need full dataset)
    yield_metrics = calculate_yield_metrics(df, df_fail)
except Exception as e:
    st.error(f"Failed to load/analyze file: {e}")
    st.stop()

# Filters
with st.sidebar:
    st.header("2) Filters")
    ff_list = sorted(df_fail["FF"].dropna().unique().tolist())
    selected_ff = st.multiselect("Focus Factory (FF)", options=ff_list, default=ff_list)

    # Date filter if available
    if "Date Tested" in df_fail.columns and not df_fail["Date Tested"].isna().all():
        min_date = df_fail["Date Tested"].min()
        max_date = df_fail["Date Tested"].max()
        date_range = st.date_input("Date Tested range", value=(min_date.date(), max_date.date()))
    else:
        date_range = None

    # Additional slicers (only if columns exist)
    def multisel(col):
        if col in df_fail.columns:
            opts = sorted([x for x in df_fail[col].fillna("").astype(str).str.strip().replace("", "(blank)").unique()])
            return st.multiselect(col, options=opts, default=opts)
        return None

    sel_stage = multisel("Process Stage")
    sel_station = multisel("Station")
    sel_failcode = multisel("Failure Code")
    sel_defcode = multisel("Defect Code")
    sel_part = multisel("Part No")

# Apply filters
f = df_fail.copy()
f = f[f["FF"].isin(selected_ff)]

if date_range and len(date_range) == 2 and "Date Tested" in f.columns:
    start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    f = f[(f["Date Tested"].isna()) | ((f["Date Tested"] >= start) & (f["Date Tested"] <= end))]

def apply_multi(col, sel):
    global f
    if sel is None or col not in f.columns:
        return
    # normalize blanks
    series = f[col].fillna("").astype(str).str.strip().replace("", "(blank)")
    f[col] = series
    f = f[f[col].isin(sel)]

apply_multi("Process Stage", sel_stage)
apply_multi("Station", sel_station)
apply_multi("Failure Code", sel_failcode)
apply_multi("Defect Code", sel_defcode)
apply_multi("Part No", sel_part)

# KPIs
total_rows = len(f)
unique_parts = f["Part No"].nunique() if "Part No" in f.columns else 0
unique_ff = f["FF"].nunique()

def top_value(col):
    if col not in f.columns or f.empty:
        return ""
    s = f[col].fillna("").astype(str).str.strip().replace("", "(blank)")
    vc = s.value_counts()
    return f"{vc.index[0]}  ({vc.iloc[0]})" if len(vc) else ""

# Enhanced KPI display with yield metrics
st.subheader("📈 Key Performance Indicators")
st.caption(f"Unit-level KPIs — {yield_metrics.get('note', 'Serial-level FPY')}: each physical unit (serial number) is counted once. "
           "Unique Serial Numbers Tested = Units Passed + Units Failed. DPMO = failed units ÷ tested units × 1,000,000.")

# Row 1: Core metrics (all at serial/unit level so the numbers reconcile)
kpi_row1 = st.columns(6)
kpi_row1[0].metric("Unique Serial Numbers Tested", f"{yield_metrics['total_tested']:,}")
kpi_row1[1].metric("Units Failed", f"{yield_metrics['total_failed']:,}", delta=f"{yield_metrics['failure_rate']:.1f}% rate" if yield_metrics['total_tested'] > 0 else None, delta_color="inverse")
kpi_row1[2].metric("Units Passed", f"{yield_metrics['total_passed']:,}")
kpi_row1[3].metric("Focus Factories", f"{unique_ff:,}")
kpi_row1[4].metric("Unique Parts", f"{unique_parts:,}")
kpi_row1[5].metric("DPMO", f"{yield_metrics['dpmo']:,.0f}")

# Row 2: Yield gauges
st.markdown("---")
gauge_cols = st.columns(3)

with gauge_cols[0]:
    fig_fpy = create_gauge_chart(
        yield_metrics["fpy"],
        "First Pass Yield (FPY)",
        {"warning": yield_warning, "critical": yield_critical}
    )
    st.plotly_chart(fig_fpy, use_container_width=True)

with gauge_cols[1]:
    # Failure rate gauge (inverted - lower is better)
    fig_fr = go.Figure(go.Indicator(
        mode="gauge+number",
        value=yield_metrics["failure_rate"],
        domain={'x': [0, 1], 'y': [0, 1]},
        title={'text': "Failure Rate", 'font': {'size': 16}},
        number={'suffix': '%', 'font': {'size': 28}},
        gauge={
            'axis': {'range': [0, max(20, yield_metrics["failure_rate"] * 1.2)]},
            'bar': {'color': "red" if yield_metrics["failure_rate"] > 10 else "orange" if yield_metrics["failure_rate"] > 5 else "green"},
            'steps': [
                {'range': [0, 5], 'color': 'rgba(0,255,0,0.2)'},
                {'range': [5, 10], 'color': 'rgba(255,165,0,0.2)'},
                {'range': [10, 20], 'color': 'rgba(255,0,0,0.2)'}
            ]
        }
    ))
    fig_fr.update_layout(height=250, margin=dict(l=20, r=20, t=40, b=20))
    st.plotly_chart(fig_fr, use_container_width=True)

with gauge_cols[2]:
    st.markdown("**Top Drivers Summary**")
    st.write(f"🔧 **Top Stage:** {top_value('Process Stage')}")
    st.write(f"❌ **Top Failure Code:** {top_value('Failure Code')}")
    st.write(f"🔴 **Top Defect Code:** {top_value('Defect Code')}")
    st.write(f"🖥️ **Top Station:** {top_value('Station')}")
    st.write(f"📦 **Top Part No:** {top_value('Part No')}")

st.divider()

# FPY Bar Chart
st.subheader("First Pass Yield (FPY) by Focus Factory")
yield_by_ff_main = calculate_yield_by_group(df, "FF")
if not yield_by_ff_main.empty:
    fig_fpy_bar = px.bar(
        yield_by_ff_main.sort_values("FPY_%", ascending=False),
        x="FF",
        y="FPY_%",
        text=yield_by_ff_main.sort_values("FPY_%", ascending=False)["FPY_%"].apply(lambda v: f"{v:.1f}%"),
        color="FPY_%",
        color_continuous_scale=["red", "orange", "green"],
        range_color=[80, 100],
    )
    fig_fpy_bar.add_hline(y=yield_warning, line_dash="dash", line_color="orange",
                          annotation_text=f"Warning: {yield_warning}%")
    fig_fpy_bar.add_hline(y=yield_critical, line_dash="dash", line_color="red",
                          annotation_text=f"Critical: {yield_critical}%")
    fig_fpy_bar.update_layout(xaxis_title="Focus Factory", yaxis_title="FPY %",
                               yaxis_range=[0, 105], coloraxis_showscale=False)
    st.plotly_chart(fig_fpy_bar, use_container_width=True)
else:
    st.info("Need Pass/Fail data to calculate FPY.")

st.divider()

# Charts Row 1
left, right = st.columns(2)

with left:
    st.subheader("FAILs by Focus Factory (FF)")
    ff_counts = f.groupby("FF").size().reset_index(name="Count").sort_values("Count", ascending=False)
    fig = px.bar(ff_counts, x="FF", y="Count", text="Count")
    fig.update_layout(xaxis_title="FF", yaxis_title="FAIL Count")
    st.plotly_chart(fig, use_container_width=True)

with right:
    st.subheader("Process Stage by FF (stacked)")
    if "Process Stage" in f.columns:
        stage_ff = f.groupby(["FF", "Process Stage"]).size().reset_index(name="Count")
        fig2 = px.bar(stage_ff, x="FF", y="Count", color="Process Stage", barmode="stack")
        fig2.update_layout(xaxis_title="FF", yaxis_title="FAIL Count")
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("No Process Stage column detected.")

# Charts Row 2
left2, right2 = st.columns(2)

with left2:
    st.subheader("Pareto: Failure Code (selected filters)")
    if "Failure Code" in f.columns and not f.empty:
        vc = f["Failure Code"].fillna("").astype(str).str.strip().replace("", "(blank)").value_counts().reset_index()
        vc.columns = ["Failure Code", "Count"]
        vc["Cum%"] = vc["Count"].cumsum() / vc["Count"].sum()
        fig3 = px.bar(vc.head(20), x="Failure Code", y="Count", text="Count")
        st.plotly_chart(fig3, use_container_width=True)

        fig3b = px.line(vc.head(20), x="Failure Code", y="Cum%", markers=True)
        fig3b.update_layout(yaxis_tickformat=".0%")
        st.plotly_chart(fig3b, use_container_width=True)
    else:
        st.info("No Failure Code data available after filtering.")

with right2:
    st.subheader("Top Parts (selected filters)")
    if "Part No" in f.columns and not f.empty:
        part_counts = f["Part No"].value_counts().head(20).reset_index()
        part_counts.columns = ["Part No", "Count"]
        fig4 = px.bar(part_counts, x="Part No", y="Count", text="Count")
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("No Part No data available.")

st.divider()

# Heatmap: Station x Stage
st.subheader("Heatmap: Station × Process Stage (FAIL Count)")
if "Station" in f.columns and "Process Stage" in f.columns and not f.empty:
    pivot = pd.pivot_table(
        f, index="Station", columns="Process Stage", values="FF",
        aggfunc="count", fill_value=0
    )
    fig_hm = px.imshow(pivot, aspect="auto", color_continuous_scale="Blues")
    fig_hm.update_layout(xaxis_title="Process Stage", yaxis_title="Station")
    st.plotly_chart(fig_hm, use_container_width=True)
else:
    st.info("Need Station + Process Stage columns (and data) to show this heatmap.")

st.divider()

# Drivers tables
st.subheader("Top Drivers per FF (multi-metric)")
tabs = st.tabs(["Overview", "Yield by FF", "Drivers (Top N)", "Combo Drivers", "Pareto Analysis", "SPC Control Chart", "Trend Analysis", "Repeat Failures", "Sunburst View", "Column Mapping"])

with tabs[0]:
    # FF overview table
    def top_per_ff(df_in, col):
        if col not in df_in.columns:
            return pd.DataFrame({"FF": df_in["FF"].unique(), f"Top_{col}": ""})
        tmp = (df_in.groupby(["FF", col]).size().reset_index(name="Count")
               .sort_values(["FF", "Count"], ascending=[True, False]))
        top = tmp.groupby("FF").head(1)[["FF", col, "Count"]]
        top = top.rename(columns={col: f"Top_{col}", "Count": f"Top_{col}_Count"})
        return top

    overview = f.groupby("FF").size().rename("Total_FAILs").reset_index()
    for col in ["Process Stage", "Failure Code", "Defect Code", "Part No", "Station", "Failed ATE Step", "Failed Assy"]:
        if col in f.columns:
            overview = overview.merge(top_per_ff(f, col), on="FF", how="left")
    overview = overview.sort_values("Total_FAILs", ascending=False)

    st.dataframe(overview, use_container_width=True, height=380)

with tabs[1]:
    st.markdown("### 📊 Yield Metrics by Focus Factory")
    st.caption("First Pass Yield (FPY) and failure rates broken down by FF")

    yield_by_ff = calculate_yield_by_group(df, "FF")
    if not yield_by_ff.empty:
        # Color code based on thresholds
        def highlight_yield(row):
            fpy = row["FPY_%"]
            if fpy >= yield_warning:
                return ["background-color: rgba(0,255,0,0.2)"] * len(row)
            elif fpy >= yield_critical:
                return ["background-color: rgba(255,165,0,0.2)"] * len(row)
            else:
                return ["background-color: rgba(255,0,0,0.2)"] * len(row)

        styled_yield = yield_by_ff.style.apply(highlight_yield, axis=1)
        st.dataframe(styled_yield, use_container_width=True, height=350)

        # Yield bar chart
        fig_yield = px.bar(
            yield_by_ff.sort_values("FPY_%"),
            x="FF",
            y="FPY_%",
            color="FPY_%",
            color_continuous_scale=["red", "orange", "green"],
            range_color=[80, 100],
            title="First Pass Yield by Focus Factory"
        )
        fig_yield.add_hline(y=yield_warning, line_dash="dash", line_color="orange", annotation_text=f"Warning: {yield_warning}%")
        fig_yield.add_hline(y=yield_critical, line_dash="dash", line_color="red", annotation_text=f"Critical: {yield_critical}%")
        fig_yield.update_layout(xaxis_title="Focus Factory", yaxis_title="FPY %")
        st.plotly_chart(fig_yield, use_container_width=True)
    else:
        st.info("Need Pass/Fail data to calculate yield metrics.")

with tabs[2]:
    all_drivers = []
    for field in DRIVER_FIELDS:
        if field in f.columns:
            d = top_n_counts(f, field, top_n=top_n)
            d.insert(1, "Metric", field)
            all_drivers.append(d)
    drivers_long = pd.concat(all_drivers, ignore_index=True) if all_drivers else pd.DataFrame()
    st.dataframe(drivers_long, use_container_width=True, height=450)

with tabs[3]:
    combos = []
    for combo_name, cols in COMBO_DEFS.items():
        if all(c in f.columns for c in cols):
            cdf = top_n_counts(f, cols, top_n=top_n)
            cdf.insert(1, "Combo", combo_name)
            combos.append(cdf)
    combo_df = pd.concat(combos, ignore_index=True) if combos else pd.DataFrame()
    st.dataframe(combo_df, use_container_width=True, height=450)

with tabs[4]:
    st.markdown("### 📊 Pareto Analysis (80/20 Rule)")
    st.caption("Identify the vital few causes driving the majority of failures")

    pareto_col = st.selectbox(
        "Select column for Pareto analysis:",
        ["Failure Code", "Defect Code", "Part No", "Process Stage", "Station", "Failed ATE Step"],
        index=0
    )

    if pareto_col in f.columns:
        pareto_df, items_80 = calculate_pareto_80_20(f, pareto_col)

        if not pareto_df.empty:
            st.info(f"🎯 **80/20 Insight:** The top **{items_80}** items account for 80% of all failures")

            # Combined Pareto chart
            fig_pareto = make_subplots(specs=[[{"secondary_y": True}]])

            fig_pareto.add_trace(
                go.Bar(x=pareto_df[pareto_col].head(20), y=pareto_df["Count"].head(20), name="Count", marker_color="steelblue"),
                secondary_y=False
            )
            fig_pareto.add_trace(
                go.Scatter(x=pareto_df[pareto_col].head(20), y=pareto_df["Cum_%"].head(20), name="Cumulative %",
                          mode="lines+markers", marker_color="red", line=dict(width=2)),
                secondary_y=True
            )

            # Add 80% line
            fig_pareto.add_hline(y=80, line_dash="dash", line_color="green", secondary_y=True,
                                annotation_text="80% threshold")

            fig_pareto.update_layout(title=f"Pareto Chart: {pareto_col}", height=450)
            fig_pareto.update_yaxes(title_text="Count", secondary_y=False)
            fig_pareto.update_yaxes(title_text="Cumulative %", secondary_y=True, range=[0, 105])

            st.plotly_chart(fig_pareto, use_container_width=True)

            # Data table
            st.dataframe(pareto_df.head(20), use_container_width=True)
    else:
        st.info(f"No {pareto_col} column found.")

with tabs[5]:
    st.markdown("### 📉 Statistical Process Control (SPC) Chart")
    st.caption("Monitor process stability and detect out-of-control conditions")

    spc_period = st.radio("Select period:", ["Weekly", "Monthly"], horizontal=True)
    period_code = "W" if spc_period == "Weekly" else "M"

    if "Date Tested" in f.columns and not f["Date Tested"].isna().all():
        trend_data = calculate_trend_metrics(f, "Date Tested", period=period_code)

        if not trend_data.empty and len(trend_data) >= 3:
            spc_chart = create_control_chart(trend_data, "Count", "Period")
            if spc_chart:
                st.plotly_chart(spc_chart, use_container_width=True)

                # SPC Statistics
                limits = calculate_spc_limits(trend_data["Count"])
                spc_stats = st.columns(5)
                spc_stats[0].metric("Mean", f"{limits['mean']:.1f}")
                spc_stats[1].metric("UCL", f"{limits['ucl']:.1f}")
                spc_stats[2].metric("LCL", f"{limits['lcl']:.1f}")
                spc_stats[3].metric("Std Dev", f"{trend_data['Count'].std():.1f}")

                ooc_count = ((trend_data["Count"] > limits["ucl"]) | (trend_data["Count"] < limits["lcl"])).sum()
                spc_stats[4].metric("Out of Control Points", ooc_count, delta="⚠️" if ooc_count > 0 else "✅")

                st.dataframe(trend_data, use_container_width=True)
        else:
            st.info("Need at least 3 periods of data for SPC analysis.")
    else:
        st.info("No Date Tested column available for SPC analysis.")

with tabs[6]:
    st.markdown("### 📈 Trend Analysis with Period-over-Period Changes")

    if "Date Tested" in f.columns and not f["Date Tested"].isna().all():
        trend_period = st.radio("Trend period:", ["Weekly", "Monthly"], horizontal=True, key="trend_period")
        trend_code = "W" if trend_period == "Weekly" else "M"

        trend_df = calculate_trend_metrics(f, "Date Tested", period=trend_code)

        if not trend_df.empty:
            # Trend with rolling average
            fig_trend = go.Figure()

            fig_trend.add_trace(go.Bar(
                x=trend_df["Period"],
                y=trend_df["Count"],
                name="Actual",
                marker_color="steelblue"
            ))

            fig_trend.add_trace(go.Scatter(
                x=trend_df["Period"],
                y=trend_df["Rolling_Avg"],
                name="4-Period Rolling Avg",
                mode="lines",
                line=dict(color="red", width=2)
            ))

            fig_trend.update_layout(
                title=f"{trend_period} Failure Trend with Rolling Average",
                xaxis_title="Period",
                yaxis_title="Failure Count",
                height=400
            )
            st.plotly_chart(fig_trend, use_container_width=True)

            # Period-over-period table
            st.markdown("**Period-over-Period Changes:**")
            trend_display = trend_df[["Period", "Count", "Change", "Change_%", "Rolling_Avg"]].copy()
            trend_display["Period"] = trend_display["Period"].dt.strftime("%Y-%m-%d")
            trend_display = trend_display.sort_values("Period", ascending=False)
            st.dataframe(trend_display, use_container_width=True)

            # Trend by FF
            st.markdown("**Trend by Focus Factory:**")
            tmp = f.copy()
            tmp["Period"] = tmp["Date Tested"].dt.to_period(trend_code[0]).dt.to_timestamp()
            ff_trend = tmp.groupby(["FF", "Period"]).size().reset_index(name="Count")

            fig_ff_trend = px.line(ff_trend, x="Period", y="Count", color="FF", markers=True)
            fig_ff_trend.update_layout(title="Failure Trend by Focus Factory", height=400)
            st.plotly_chart(fig_ff_trend, use_container_width=True)
    else:
        st.info("No usable Date Tested column detected.")

with tabs[7]:
    st.markdown("### 🔄 Repeat Failure Detection")
    st.caption("Identify parts or serial numbers that have failed multiple times")

    identifier_options = ["Part No", "Failed Assy", "Reference Designator"]
    available_ids = [c for c in identifier_options if c in f.columns]

    if available_ids:
        selected_id = st.selectbox("Select identifier column:", available_ids)

        repeat_df = detect_repeat_failures(f, [selected_id])

        if not repeat_df.empty:
            total_repeats = len(repeat_df)
            total_repeat_failures = repeat_df["Failure_Count"].sum()

            repeat_cols = st.columns(3)
            repeat_cols[0].metric("Items with Repeat Failures", f"{total_repeats:,}")
            repeat_cols[1].metric("Total Repeat Failure Count", f"{total_repeat_failures:,}")
            repeat_cols[2].metric("% of All Failures", f"{total_repeat_failures/len(f)*100:.1f}%" if len(f) > 0 else "0%")

            st.dataframe(repeat_df.head(50), use_container_width=True, height=350)

            # Histogram of repeat counts
            fig_repeat = px.histogram(repeat_df, x="Failure_Count", nbins=20,
                                      title="Distribution of Repeat Failure Counts")
            fig_repeat.update_layout(xaxis_title="Number of Failures", yaxis_title="Count of Items")
            st.plotly_chart(fig_repeat, use_container_width=True)
        else:
            st.success("✅ No repeat failures detected!")
    else:
        st.info("No suitable identifier columns found for repeat failure analysis.")

with tabs[8]:
    st.markdown("### 🌐 Sunburst / Hierarchical View")
    st.caption("Visualize failure breakdown hierarchically")

    hierarchy_options = [
        ["FF", "Process Stage", "Failure Code"],
        ["FF", "Process Stage", "Station"],
        ["FF", "Part No", "Failure Code"],
        ["Process Stage", "Failure Code", "Defect Code"],
    ]

    hierarchy_labels = [" → ".join(h) for h in hierarchy_options]
    selected_hierarchy_idx = st.selectbox("Select hierarchy:", range(len(hierarchy_labels)),
                                          format_func=lambda x: hierarchy_labels[x])

    selected_hierarchy = hierarchy_options[selected_hierarchy_idx]
    sunburst = create_sunburst_chart(f, selected_hierarchy)

    if sunburst:
        st.plotly_chart(sunburst, use_container_width=True)
    else:
        st.info("Not enough data/columns for this hierarchy view.")

    # Also add a treemap
    available_cols = [c for c in selected_hierarchy if c in f.columns]
    if len(available_cols) >= 2:
        tmp = f.copy()
        for c in available_cols:
            tmp[c] = tmp[c].fillna("").astype(str).str.strip().replace("", "(blank)")
        agg = tmp.groupby(available_cols).size().reset_index(name="Count")

        fig_tree = px.treemap(agg, path=available_cols, values="Count",
                             title="Treemap: Failure Distribution")
        fig_tree.update_layout(height=500)
        st.plotly_chart(fig_tree, use_container_width=True)

with tabs[9]:
    colmap_df = pd.DataFrame({
        "Canonical Column": list(colmap.keys()),
        "Matched Actual Column": list(colmap.values()),
    })
    st.dataframe(colmap_df, use_container_width=True)

# -----------------------------
# Focus Factory Deep Dive Analysis
# -----------------------------
st.divider()
st.header("🔍 Focus Factory Deep Dive — Worst Performers per FF")
st.caption("Comprehensive breakdown of top failure drivers for each Focus Factory")

# Get list of FFs for individual analysis
ff_in_data = sorted(f["FF"].unique().tolist())

# Create tabs for each analysis type
analysis_tabs = st.tabs(["📊 Summary Table", "🏭 By Focus Factory", "📈 Comparison Charts"])

with analysis_tabs[0]:
    st.subheader("Worst Performers Summary — All Focus Factories")

    # Build comprehensive summary table
    summary_rows = []

    for ff in ff_in_data:
        ff_data = f[f["FF"] == ff]
        ff_total = len(ff_data)

        row = {"Focus Factory": ff, "Total Failures": ff_total}

        # Get #1 worst for each metric
        for col, label in [
            ("Part No", "Worst Part No"),
            ("Process Stage", "Worst Process Stage"),
            ("Station", "Worst Station"),
            ("Failure Code", "Worst Failure Code"),
            ("Defect Code", "Worst Defect Code"),
            ("Failed ATE Step", "Worst ATE Step"),
            ("Failed Assy", "Worst Failed Assy"),
            ("Tested By", "Worst Tested By"),
        ]:
            if col in ff_data.columns:
                vc = ff_data[col].fillna("").astype(str).str.strip().replace("", "(blank)").value_counts()
                if len(vc) > 0:
                    worst_item = vc.index[0]
                    worst_count = vc.iloc[0]
                    worst_pct = (worst_count / ff_total * 100) if ff_total > 0 else 0
                    row[label] = f"{worst_item}"
                    row[f"{label} Count"] = worst_count
                    row[f"{label} %"] = f"{worst_pct:.1f}%"
                else:
                    row[label] = ""
                    row[f"{label} Count"] = 0
                    row[f"{label} %"] = ""

        summary_rows.append(row)

    summary_df = pd.DataFrame(summary_rows)

    # Display with highlighting
    st.dataframe(summary_df, use_container_width=True, height=400)

    # Key insights
    st.subheader("🎯 Key Insights")
    col_a, col_b, col_c = st.columns(3)

    with col_a:
        st.markdown("**Highest Failure FF:**")
        if len(summary_df) > 0:
            worst_ff = summary_df.loc[summary_df["Total Failures"].idxmax()]
            st.metric(worst_ff["Focus Factory"], f"{worst_ff['Total Failures']:,} failures")

    with col_b:
        if "Worst Part No" in summary_df.columns:
            st.markdown("**Most Common Worst Part:**")
            part_counts = summary_df["Worst Part No"].value_counts()
            if len(part_counts) > 0:
                st.write(f"**{part_counts.index[0]}** appears as worst in {part_counts.iloc[0]} FF(s)")

    with col_c:
        if "Worst Process Stage" in summary_df.columns:
            st.markdown("**Most Common Worst Stage:**")
            stage_counts = summary_df["Worst Process Stage"].value_counts()
            if len(stage_counts) > 0:
                st.write(f"**{stage_counts.index[0]}** appears as worst in {stage_counts.iloc[0]} FF(s)")

with analysis_tabs[1]:
    st.subheader("Detailed Analysis by Focus Factory")

    # Dropdown to select FF
    selected_ff_detail = st.selectbox("Select Focus Factory for detailed view:", ff_in_data)

    if selected_ff_detail:
        ff_subset = f[f["FF"] == selected_ff_detail]
        ff_total_fails = len(ff_subset)

        st.metric(f"Total Failures in {selected_ff_detail}", f"{ff_total_fails:,}")
        st.divider()

        # Create columns for each metric
        metric_cols = st.columns(2)

        metrics_to_show = [
            ("Part No", "🔧 Top 10 Failing Parts"),
            ("Process Stage", "⚙️ Top Process Stages"),
            ("Station", "🖥️ Top Failing Stations"),
            ("Failure Code", "❌ Top Failure Codes"),
            ("Defect Code", "🔴 Top Defect Codes"),
            ("Failed ATE Step", "📋 Top Failed ATE Steps"),
        ]

        for idx, (col, title) in enumerate(metrics_to_show):
            if col in ff_subset.columns:
                with metric_cols[idx % 2]:
                    st.markdown(f"**{title}**")
                    vc = ff_subset[col].fillna("").astype(str).str.strip().replace("", "(blank)").value_counts().head(10)
                    if len(vc) > 0:
                        vc_df = vc.reset_index()
                        vc_df.columns = [col, "Count"]
                        vc_df["% of FF"] = (vc_df["Count"] / ff_total_fails * 100).round(1).astype(str) + "%"
                        vc_df.index = range(1, len(vc_df) + 1)
                        vc_df.index.name = "Rank"
                        st.dataframe(vc_df, use_container_width=True)

                        # Mini bar chart
                        fig = px.bar(vc_df.head(5), x=col, y="Count", text="Count")
                        fig.update_layout(height=250, showlegend=False, xaxis_title="", yaxis_title="Failures")
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info(f"No {col} data available.")

        # --- Drill-down: Top 3 Stations with Parts & Failure Codes ---
        if "Station" in ff_subset.columns:
            st.divider()
            st.markdown("### 🔎 Top 3 Stations — Parts & Failure Codes Breakdown")
            station_vc = ff_subset["Station"].fillna("").astype(str).str.strip().replace("", "(blank)").value_counts().head(3)
            if len(station_vc) > 0:
                station_tabs = st.tabs([f"#{r+1}  {stn} ({cnt})" for r, (stn, cnt) in enumerate(station_vc.items())])
                for tab, (stn_name, stn_count) in zip(station_tabs, station_vc.items()):
                    with tab:
                        stn_data = ff_subset[ff_subset["Station"].fillna("").astype(str).str.strip().replace("", "(blank)") == stn_name]
                        st.markdown(f"**Station: {stn_name}** — {stn_count:,} failures ({stn_count / ff_total_fails * 100:.1f}% of FF)")
                        detail_c1, detail_c2 = st.columns(2)

                        # Failing Parts for this station
                        with detail_c1:
                            st.markdown("**🔧 Failing Parts**")
                            if "Part No" in stn_data.columns:
                                parts_vc = stn_data["Part No"].fillna("").astype(str).str.strip().replace("", "(blank)").value_counts().head(10)
                                if len(parts_vc) > 0:
                                    parts_df = parts_vc.reset_index()
                                    parts_df.columns = ["Part No", "Count"]
                                    parts_df["% of Station"] = (parts_df["Count"] / stn_count * 100).round(1).astype(str) + "%"
                                    parts_df.index = range(1, len(parts_df) + 1)
                                    parts_df.index.name = "Rank"
                                    st.dataframe(parts_df, use_container_width=True)
                                else:
                                    st.info("No Part No data.")
                            else:
                                st.info("Part No column not available.")

                        # Failure Codes for this station
                        with detail_c2:
                            st.markdown("**❌ Failure Codes**")
                            if "Failure Code" in stn_data.columns:
                                fc_vc = stn_data["Failure Code"].fillna("").astype(str).str.strip().replace("", "(blank)").value_counts().head(10)
                                if len(fc_vc) > 0:
                                    fc_df = fc_vc.reset_index()
                                    fc_df.columns = ["Failure Code", "Count"]
                                    fc_df["% of Station"] = (fc_df["Count"] / stn_count * 100).round(1).astype(str) + "%"
                                    fc_df.index = range(1, len(fc_df) + 1)
                                    fc_df.index.name = "Rank"
                                    st.dataframe(fc_df, use_container_width=True)
                                else:
                                    st.info("No Failure Code data.")
                            else:
                                st.info("Failure Code column not available.")

with analysis_tabs[2]:
    st.subheader("Cross-FF Comparison Charts")

    compare_metric = st.selectbox(
        "Select metric to compare across Focus Factories:",
        ["Part No", "Process Stage", "Station", "Failure Code", "Defect Code"],
        index=1
    )

    if compare_metric in f.columns:
        # Get top items for this metric across all FFs
        tmp = f.copy()
        tmp[compare_metric] = tmp[compare_metric].fillna("").astype(str).str.strip().replace("", "(blank)")

        # Pivot: FF vs metric item
        pivot_data = tmp.groupby(["FF", compare_metric]).size().reset_index(name="Count")

        # Get top 10 items overall
        top_items = tmp[compare_metric].value_counts().head(10).index.tolist()
        pivot_filtered = pivot_data[pivot_data[compare_metric].isin(top_items)]

        # Grouped bar chart
        fig_compare = px.bar(
            pivot_filtered,
            x="FF",
            y="Count",
            color=compare_metric,
            barmode="group",
            title=f"Top 10 {compare_metric} by Focus Factory"
        )
        fig_compare.update_layout(xaxis_title="Focus Factory", yaxis_title="Failure Count")
        st.plotly_chart(fig_compare, use_container_width=True)

        # Heatmap view
        st.markdown("**Heatmap View:**")
        pivot_table = pivot_filtered.pivot_table(
            index=compare_metric,
            columns="FF",
            values="Count",
            fill_value=0
        )
        fig_hm2 = px.imshow(
            pivot_table,
            aspect="auto",
            color_continuous_scale="Reds",
            title=f"{compare_metric} Distribution Across Focus Factories"
        )
        st.plotly_chart(fig_hm2, use_container_width=True)

        # Data table
        st.markdown("**Data Table:**")
        st.dataframe(pivot_filtered.sort_values(["FF", "Count"], ascending=[True, False]), use_container_width=True)

# =====================================================================
# Repeat Offender Dashboard per FF
# =====================================================================
st.divider()
st.header("🔁 Repeat Offender Dashboard — Chronic Failures per Department")
st.caption(
    "Surfaces recurring Part + Station + Failure Code combinations within each "
    "department so engineers can target systemic, repeating problems."
)

# Determine which key columns are available for the combo key
_repeat_key_cols = [c for c in ["Part No", "Station", "Failure Code"] if c in f.columns]

if len(_repeat_key_cols) >= 2:
    # Let user control the lookback window if dates are available
    _has_date = "Date Tested" in f.columns and not f["Date Tested"].isna().all()
    if _has_date:
        _lookback = st.slider(
            "Lookback window (days) — only count failures within this many recent days:",
            min_value=7, max_value=365, value=90, step=7, key="repeat_lookback"
        )
        _cutoff = pd.Timestamp.now() - pd.Timedelta(days=_lookback)
        _repeat_src = f[f["Date Tested"] >= _cutoff].copy()
        st.info(f"Analyzing failures from the last **{_lookback} days** ({len(_repeat_src):,} records)")
    else:
        _repeat_src = f.copy()

    _min_occurrences = st.slider(
        "Minimum occurrences to flag as repeat offender:",
        min_value=2, max_value=20, value=3, step=1, key="repeat_min"
    )

    # Build combo key per FF
    _repeat_src["_combo_key"] = _repeat_src[_repeat_key_cols].fillna("").astype(str).apply(
        lambda row: " | ".join(row), axis=1
    )

    _repeat_agg = (
        _repeat_src.groupby(["FF", "_combo_key"] + _repeat_key_cols)
        .size()
        .reset_index(name="Occurrences")
    )
    _repeat_agg = _repeat_agg[_repeat_agg["Occurrences"] >= _min_occurrences]
    _repeat_agg = _repeat_agg.sort_values(["FF", "Occurrences"], ascending=[True, False])

    if _repeat_agg.empty:
        st.success(f"✅ No combinations with {_min_occurrences}+ occurrences found — no chronic repeat offenders!")
    else:
        # Summary metrics
        _ro_cols = st.columns(4)
        _ro_cols[0].metric("Repeat Combos Found", f"{len(_repeat_agg):,}")
        _ro_cols[1].metric("FFs Affected", f"{_repeat_agg['FF'].nunique()}")
        _ro_cols[2].metric("Total Repeat Failures", f"{_repeat_agg['Occurrences'].sum():,}")
        _worst_combo = _repeat_agg.iloc[0]
        _ro_cols[3].metric("Worst Combo", f"{int(_worst_combo['Occurrences'])}× in {_worst_combo['FF']}")

        # Per-FF tabs
        _ro_ffs = sorted(_repeat_agg["FF"].unique().tolist())
        _ro_ff_tabs = st.tabs([f"{ff_name} ({len(_repeat_agg[_repeat_agg['FF'] == ff_name])})" for ff_name in _ro_ffs])

        for _ro_tab, _ro_ff in zip(_ro_ff_tabs, _ro_ffs):
            with _ro_tab:
                _ff_repeats = _repeat_agg[_repeat_agg["FF"] == _ro_ff].drop(columns=["_combo_key"]).reset_index(drop=True)
                _ff_repeats.index = range(1, len(_ff_repeats) + 1)
                _ff_repeats.index.name = "Rank"

                st.markdown(f"**{_ro_ff}** — {len(_ff_repeats)} chronic combos, {_ff_repeats['Occurrences'].sum():,} total repeat failures")

                # Table
                st.dataframe(_ff_repeats, use_container_width=True, height=350)

                # Top 10 horizontal bar
                _top10 = _ff_repeats.head(10).copy()
                _top10["Label"] = _top10[_repeat_key_cols].astype(str).apply(lambda r: " | ".join(r), axis=1)
                fig_ro = px.bar(
                    _top10, y="Label", x="Occurrences", orientation="h",
                    text="Occurrences", color="Occurrences",
                    color_continuous_scale="Reds",
                    title=f"Top 10 Repeat Offenders — {_ro_ff}"
                )
                fig_ro.update_layout(
                    height=400, yaxis=dict(autorange="reversed"),
                    xaxis_title="Occurrences", yaxis_title=""
                )
                st.plotly_chart(fig_ro, use_container_width=True)

                # If dates available, show timeline for the worst offender
                if _has_date and len(_top10) > 0:
                    _worst_key_vals = {c: str(_top10.iloc[0][c]) for c in _repeat_key_cols}
                    _mask = _repeat_src["FF"] == _ro_ff
                    for _kc, _kv in _worst_key_vals.items():
                        _mask = _mask & (_repeat_src[_kc].fillna("").astype(str) == _kv)
                    _timeline_data = _repeat_src[_mask].copy()
                    if not _timeline_data.empty:
                        _timeline_data = _timeline_data.sort_values("Date Tested")
                        _timeline_grouped = _timeline_data.groupby(
                            _timeline_data["Date Tested"].dt.date
                        ).size().reset_index(name="Failures")
                        _timeline_grouped.columns = ["Date", "Failures"]
                        st.markdown(f"**Timeline for worst combo:** {' | '.join(_worst_key_vals.values())}")
                        fig_tl = px.bar(_timeline_grouped, x="Date", y="Failures", text="Failures")
                        fig_tl.update_layout(height=250, xaxis_title="Date", yaxis_title="Failures")
                        st.plotly_chart(fig_tl, use_container_width=True)
else:
    st.info("Need at least 2 of Part No, Station, Failure Code columns for repeat offender analysis.")


# =====================================================================
# Root Cause Drill-Down Tree per FF
# =====================================================================
st.divider()
st.header("🌳 Root Cause Drill-Down — Interactive Failure Tree per Department")
st.caption(
    "Step through the failure hierarchy level by level: "
    "FF → Process Stage → Station → Failure Code → Part No → Reference Designator. "
    "Each level shows count and % contribution — follow the biggest bar to the root cause. "
    "Select multiple Focus Factories to compare side by side, and export one Excel sheet per FF."
)

# Define hierarchy levels
_DRILL_LEVELS = ["Process Stage", "Station", "Failure Code", "Defect Code", "Part No",
                 "Failed ATE Step", "Failed Assy", "Reference Designator"]
_available_levels = [c for c in _DRILL_LEVELS if c in f.columns]

if len(_available_levels) >= 2:
    # Pick one or more FFs
    _drill_ff_options = sorted(f["FF"].unique().tolist())
    _drill_ffs = st.multiselect(
        "Select Focus Factory/Factories to drill into:", _drill_ff_options,
        default=_drill_ff_options[:1],
        key="drill_ffs"
    )

    # Let user choose hierarchy depth & order (shared across all selected FFs)
    _selected_levels = st.multiselect(
        "Choose drill-down levels (in order, top → bottom):",
        _available_levels,
        default=_available_levels[:4],
        key="drill_levels"
    )

    if not _drill_ffs:
        st.info("Select at least one Focus Factory to drill into.")
    elif len(_selected_levels) == 1:
        st.warning("Select at least 2 levels for a meaningful drill-down.")
    elif len(_selected_levels) >= 2:
        # One tab per selected Focus Factory — each keeps its own drill-down path
        _drill_tabs = st.tabs([str(_ff) for _ff in _drill_ffs])
        _drill_sheets = {}

        for _drill_tab, _drill_ff in zip(_drill_tabs, _drill_ffs):
            with _drill_tab:
                _drill_data = f[f["FF"] == _drill_ff].copy()
                _drill_total = len(_drill_data)

                st.metric(f"Total Failures in {_drill_ff}", f"{_drill_total:,}")

                # --- Sunburst / Treemap ---
                _drill_tmp = _drill_data.copy()
                for _lv in _selected_levels:
                    _drill_tmp[_lv] = _drill_tmp[_lv].fillna("(blank)").astype(str).str.strip().replace("", "(blank)")

                _drill_agg = _drill_tmp.groupby(_selected_levels).size().reset_index(name="Failures")

                _sun_tree_view = st.radio(
                    "Visualization style:", ["Sunburst", "Treemap", "Both"],
                    horizontal=True, key=f"drill_viz_{_drill_ff}"
                )

                if _sun_tree_view in ("Sunburst", "Both"):
                    fig_sun = px.sunburst(
                        _drill_agg, path=_selected_levels, values="Failures",
                        title=f"Sunburst — {_drill_ff}: {' → '.join(_selected_levels)}",
                        color="Failures", color_continuous_scale="Reds"
                    )
                    fig_sun.update_layout(height=600)
                    st.plotly_chart(fig_sun, use_container_width=True)

                if _sun_tree_view in ("Treemap", "Both"):
                    fig_tree = px.treemap(
                        _drill_agg, path=_selected_levels, values="Failures",
                        title=f"Treemap — {_drill_ff}: {' → '.join(_selected_levels)}",
                        color="Failures", color_continuous_scale="OrRd"
                    )
                    fig_tree.update_layout(height=600)
                    st.plotly_chart(fig_tree, use_container_width=True)

                # --- Step-by-step drill-down with linked filters ---
                st.markdown("---")
                st.markdown("### 🔽 Step-by-Step Drill-Down")
                st.caption("Select a value at each level to narrow down to the root cause.")

                _filtered = _drill_data.copy()
                _breadcrumb_parts = [f"**{_drill_ff}** ({_drill_total:,})"]

                for _lvl_idx, _lvl_col in enumerate(_selected_levels):
                    if _lvl_col not in _filtered.columns or _filtered.empty:
                        break

                    _lvl_vc = (
                        _filtered[_lvl_col]
                        .fillna("(blank)").astype(str).str.strip().replace("", "(blank)")
                        .value_counts()
                    )
                    if _lvl_vc.empty:
                        break

                    _lvl_df = _lvl_vc.reset_index()
                    _lvl_df.columns = [_lvl_col, "Failures"]
                    _lvl_df["% of Selection"] = (_lvl_df["Failures"] / _lvl_df["Failures"].sum() * 100).round(1)
                    _lvl_df["Cum %"] = _lvl_df["% of Selection"].cumsum().round(1)
                    _lvl_df.index = range(1, len(_lvl_df) + 1)

                    st.markdown(f"**Level {_lvl_idx + 1}: {_lvl_col}** — {len(_lvl_df)} unique values, {_lvl_df['Failures'].sum():,} failures")
                    st.markdown(f"Breadcrumb: {' → '.join(_breadcrumb_parts)}")

                    _col_chart, _col_table = st.columns([3, 2])
                    with _col_chart:
                        _bar_data = _lvl_df.head(15).copy()
                        fig_lvl = px.bar(
                            _bar_data, x=_lvl_col, y="Failures", text="Failures",
                            color="% of Selection", color_continuous_scale="Blues",
                            title=f"{_lvl_col} breakdown (top 15)"
                        )
                        fig_lvl.update_layout(height=350, xaxis_title="", yaxis_title="Failures")
                        st.plotly_chart(fig_lvl, use_container_width=True)

                    with _col_table:
                        st.dataframe(_lvl_df.head(15), use_container_width=True, height=350)

                    # Let user pick a value to drill deeper (only if there's a next level)
                    if _lvl_idx < len(_selected_levels) - 1:
                        _pick_options = ["(show all — don't filter)"] + _lvl_vc.index.tolist()
                        _picked = st.selectbox(
                            f"Select a {_lvl_col} to drill into:",
                            _pick_options,
                            key=f"drill_pick_{_drill_ff}_{_lvl_idx}"
                        )
                        if _picked != "(show all — don't filter)":
                            _filtered = _filtered[
                                _filtered[_lvl_col].fillna("(blank)").astype(str).str.strip().replace("", "(blank)") == _picked
                            ]
                            _breadcrumb_parts.append(f"**{_picked}** ({len(_filtered):,})")
                        else:
                            _breadcrumb_parts.append(f"*all {_lvl_col}*")

                # Final summary of drilled-down selection
                if len(_filtered) > 0 and len(_filtered) < _drill_total:
                    st.markdown("---")
                    st.markdown(f"### 🎯 Drill-Down Result: {len(_filtered):,} failures")
                    st.markdown(f"Path: {' → '.join(_breadcrumb_parts)}")
                    st.markdown(f"**These {len(_filtered):,} failures represent {len(_filtered) / _drill_total * 100:.1f}% of all failures in {_drill_ff}.**")

                    # Serial Number summary for the drilled-down selection
                    _sn_col = "Serial Number"
                    if _sn_col in _filtered.columns:
                        _serials = (
                            _filtered[_sn_col]
                            .fillna("").astype(str).str.strip()
                            .replace("", pd.NA)
                            .dropna()
                            .unique()
                        )
                        st.markdown(f"**🔢 Unique Serial Numbers ({len(_serials):,}):**")
                        _sn_detail_cols = [c for c in _selected_levels + [_sn_col, "Date Tested", "Pass or Fail"] if c in _filtered.columns]
                        seen_sn = set()
                        _sn_detail_cols = [c for c in _sn_detail_cols if not (c in seen_sn or seen_sn.add(c))]
                        _sn_df = (
                            _filtered[_sn_detail_cols]
                            .sort_values(_sn_col)
                            .reset_index(drop=True)
                        )
                        _sn_df.index = range(1, len(_sn_df) + 1)
                        st.dataframe(_sn_df, use_container_width=True, height=300)

                    # Show the remaining detail columns if useful
                    _detail_cols = ["Serial Number", "Date Tested", "Tested By", "Tech Comments", "Rework Instructions",
                                    "Troubleshoot Completed By", "Reference Designator"]
                    _avail_detail = [c for c in _detail_cols if c in _filtered.columns]
                    if _avail_detail:
                        st.markdown("**Detail records:**")
                        _show_cols = _selected_levels + _avail_detail
                        _show_cols = [c for c in _show_cols if c in _filtered.columns]
                        # Deduplicate while preserving order
                        seen = set()
                        _show_cols = [c for c in _show_cols if not (c in seen or seen.add(c))]
                        st.dataframe(
                            _filtered[_show_cols].head(100).reset_index(drop=True),
                            use_container_width=True, height=350
                        )
                elif len(_filtered) > 0:
                    st.info("No drill-down filters applied — the export will include all records for this Focus Factory.")

                # Capture this FF's drilled-down records for the multi-sheet export
                if len(_filtered) > 0:
                    _drill_sheets[str(_drill_ff)] = _filtered.reset_index(drop=True)

        # Export — one sheet per selected Focus Factory (drilled-down records)
        if _drill_sheets:
            st.markdown("---")
            _drill_out = io.BytesIO()
            with pd.ExcelWriter(_drill_out, engine="openpyxl") as _wr:
                _used_sheet_names = set()
                for _dept, _df_sheet in _drill_sheets.items():
                    _safe = re.sub(r"[\\/*?:\[\]]", "_", _dept)[:31] or "Sheet"
                    if _safe in _used_sheet_names:
                        _safe = f"{_safe[:28]}_{len(_used_sheet_names)}"
                    _used_sheet_names.add(_safe)
                    _df_sheet.to_excel(_wr, sheet_name=_safe, index=False)
            _drill_out.seek(0)
            st.download_button(
                "⬇️ Download drilled-down results (Excel — one sheet per Focus Factory)",
                data=_drill_out,
                file_name=f"DrillDown_By_FF_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="drill_download_btn",
            )
else:
    st.info("Need at least 2 driver columns (Process Stage, Station, Failure Code, etc.) for drill-down analysis.")


# =====================================================================
# FAILURE ANALYSIS SUMMARY TABLE BY DEPARTMENT
# =====================================================================
st.divider()
st.header("📋 Failure Analysis Summary by Department")
st.caption(
    "Select one or more departments to see the top failure combinations — "
    "Top Station, Part #, Failure Code, Defect Code, and most recurring causes — "
    "ready to present to the team. Export produces one Excel sheet per department."
)

_fa_ff_list = sorted(f["FF"].unique().tolist())
if _fa_ff_list:
    from collections import Counter

    # Noise values that carry no useful information
    _NOISE = {
        "", "nan", "none", "n/a", "na", "(blank)", "tbd",
        "tbd - to be determined", "to be determined",
        "see tech comments", "others - see tech comments",
        "other - see tech comments", "other", "others",
        "unknown", "unk", "no comment", "no comments", "pending",
    }

    _FA_STOP = {
        "the", "a", "an", "and", "or", "but", "in", "on", "at", "to", "for",
        "of", "with", "is", "was", "are", "were", "be", "been", "has", "have",
        "had", "do", "did", "will", "would", "could", "should", "may", "might",
        "this", "that", "these", "those", "it", "its", "we", "i", "you", "he",
        "she", "they", "unit", "tested", "found", "replaced", "part", "per",
        "check", "checked", "no", "not", "ok", "see", "ref", "via",
        "due", "after", "before", "during", "need", "needs", "item", "new",
        "all", "also", "back", "both", "each", "few", "from", "get",
        "got", "into", "just", "like", "made", "make", "more", "most",
        "over", "same", "some", "such", "than", "then", "them", "their",
        "there", "time", "too", "use", "used", "very",
        "what", "when", "which", "who", "with", "your",
    }

    def _summarize_causes(comments, fallback_cause=""):
        """Return a concise theme summary from a list of free-text comments."""
        # Filter noise
        clean_raw = [
            str(c).strip()
            for c in comments
            if str(c).strip().lower() not in _NOISE
        ]
        if not clean_raw:
            return fallback_cause  # use structured cause data as fallback

        # If one comment dominates (≥40% of records) just show it truncated
        freq = Counter(clean_raw)
        top_comment, top_count = freq.most_common(1)[0]
        if top_count / len(clean_raw) >= 0.4 or len(freq) <= 2:
            return top_comment[:100] if len(top_comment) > 100 else top_comment

        # Otherwise tokenize and find the most frequent meaningful phrases
        normalized = [
            re.sub(r"[^\w\s]", " ", c.lower()).strip()
            for c in clean_raw
        ]

        def _tokens(text):
            return [w for w in text.split() if len(w) > 2 and w not in _FA_STOP]

        token_lists = [_tokens(t) for t in normalized]
        flat_tokens = [w for toks in token_lists for w in toks]

        if not flat_tokens:
            return top_comment[:100]

        word_freq = Counter(flat_tokens)

        # Try bigrams first (threshold: appears in at least 2 records)
        bigrams = []
        for toks in token_lists:
            for i in range(len(toks) - 1):
                bigrams.append(f"{toks[i]} {toks[i+1]}")
        bigram_freq = Counter(bigrams)
        top_bigrams = [bg for bg, cnt in bigram_freq.most_common(10) if cnt >= 2]

        if top_bigrams:
            chosen, used_words = [], set()
            for bg in top_bigrams:
                parts = set(bg.split())
                if not parts & used_words:
                    chosen.append(bg.title())
                    used_words |= parts
                if len(chosen) == 2:
                    break
            return "; ".join(chosen)

        # Fall back to top 3 individual keywords
        top_words = [w.title() for w, _ in word_freq.most_common(10) if len(w) > 3][:3]
        return ", ".join(top_words) if top_words else top_comment[:100]

    def _build_fa_display(_fa_ff, _fa_top_n):
        """Build the aggregated failure summary table for one department.
        Returns (display_df, dept_total) or (None, dept_total) if columns are missing."""
        _fa_src = f[f["FF"] == _fa_ff].copy()
        _fa_total = len(_fa_src)

        # Clean key columns
        for _c in ["Station", "Part No", "Failure Code", "Defect Code", "Tech Comments", "Rework Instructions"]:
            if _c in _fa_src.columns:
                _fa_src[_c] = _fa_src[_c].fillna("").astype(str).str.strip().replace("", "(blank)")

        # Determine available grouping cols
        _fa_group_cols = [c for c in ["Station", "Part No", "Failure Code", "Defect Code"] if c in _fa_src.columns]
        if len(_fa_group_cols) < 2:
            return None, _fa_total

        # Aggregate by the available group columns
        _fa_agg = (
            _fa_src.groupby(_fa_group_cols)
            .size()
            .reset_index(name="Occurrences")
            .sort_values("Occurrences", ascending=False)
            .head(_fa_top_n)
            .reset_index(drop=True)
        )
        _fa_agg["% of Dept"] = (_fa_agg["Occurrences"] / _fa_total * 100).round(1).astype(str) + "%"

        def _top_causes(grp_row):
            mask = pd.Series([True] * len(_fa_src))
            for _col in _fa_group_cols:
                mask = mask & (_fa_src[_col] == grp_row[_col])
            subset = _fa_src[mask]
            all_comments = []
            for _src_col in ["Tech Comments", "Rework Instructions"]:
                if _src_col in subset.columns:
                    all_comments += subset[_src_col].fillna("").astype(str).tolist()

            # Build a fallback from structured fields when no useful comment text exists
            fallback_parts = []
            for _fb_col in ["Defect Code", "Failure Code"]:
                if _fb_col in grp_row:
                    val = str(grp_row[_fb_col]).strip()
                    if val.lower() not in _NOISE:
                        fallback_parts.append(val)
            fallback = " / ".join(fallback_parts) if fallback_parts else ""

            return _summarize_causes(all_comments, fallback_cause=fallback)

        _fa_agg["Most Recurring Causes"] = _fa_agg.apply(_top_causes, axis=1)
        _fa_agg["Actions"] = ""
        _fa_agg["Status"] = ""

        # Rename columns to match the requested format
        _rename_map = {
            "Station": "Top Station",
            "Part No": "Part #",
            "Failure Code": "Failure Code",
            "Defect Code": "Defect Code",
            "Occurrences": "Occurrences",
            "% of Dept": "% of Dept",
        }
        _fa_display = _fa_agg.rename(columns={k: v for k, v in _rename_map.items() if k in _fa_agg.columns})
        return _fa_display, _fa_total

    # Column config: data cols read-only, Actions/Status editable
    _fa_col_cfg = {
        col: st.column_config.TextColumn(col, disabled=True)
        for col in ["Top Station", "Part #", "Failure Code", "Defect Code"]
    }
    _fa_col_cfg["Occurrences"] = st.column_config.NumberColumn("Occurrences", disabled=True)
    _fa_col_cfg["% of Dept"] = st.column_config.TextColumn("% of Dept", disabled=True)
    _fa_col_cfg["Most Recurring Causes"] = st.column_config.TextColumn("Most Recurring Causes", disabled=True, width="large")
    _fa_col_cfg["Actions"] = st.column_config.TextColumn("Actions", width="large")
    _fa_col_cfg["Status"] = st.column_config.TextColumn("Status", width="medium")

    _fa_ffs = st.multiselect(
        "Select Department(s) (FF):",
        _fa_ff_list,
        default=_fa_ff_list,
        key="fa_ff_select",
    )

    if not _fa_ffs:
        st.info("Select at least one department to see the summary.")
    else:
        _fa_top_n = st.slider("Number of rows to show:", min_value=5, max_value=50, value=15, step=5, key="fa_top_n")

        # One tab per selected department — all respect the sidebar drill-down filters
        _fa_tabs = st.tabs([str(_d) for _d in _fa_ffs])
        _fa_sheets = {}
        for _fa_tab, _fa_ff in zip(_fa_tabs, _fa_ffs):
            with _fa_tab:
                _fa_display, _fa_total = _build_fa_display(_fa_ff, _fa_top_n)
                if _fa_display is None:
                    st.info("Need at least Station and Part No columns for this analysis.")
                    continue

                st.markdown(f"**{_fa_ff}** — {_fa_total:,} total failures | showing top {min(_fa_top_n, len(_fa_display))} combinations")

                _fa_edited = st.data_editor(
                    _fa_display,
                    column_config=_fa_col_cfg,
                    use_container_width=True,
                    height=min(600, 45 + len(_fa_display) * 38),
                    hide_index=True,
                    key=f"fa_data_editor_{_fa_ff}",
                )
                _fa_sheets[str(_fa_ff)] = _fa_edited

        # Download — one sheet per selected department (includes any Actions/Status edits)
        if _fa_sheets:
            _fa_out = io.BytesIO()
            with pd.ExcelWriter(_fa_out, engine="openpyxl") as _wr:
                for _fa_dept, _fa_sheet_df in _fa_sheets.items():
                    _fa_safe = re.sub(r"[\\/*?:\[\]]", "_", _fa_dept)[:31] or "Sheet"
                    _fa_sheet_df.to_excel(_wr, sheet_name=_fa_safe, index=False)
            _fa_out.seek(0)
            st.download_button(
                "⬇️ Download all departments (Excel — one sheet per department)",
                data=_fa_out,
                file_name=f"FA_Summary_By_Department_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="fa_download_btn",
            )
else:
    st.info("No failure data available for this analysis.")


# =====================================================================
# TOP STATION + TOP PART — FAILURE DETAILS PER DEPARTMENT
# =====================================================================
st.divider()
st.header("🔧 Top Station + Top Part — Failure Details by Department")
st.caption(
    "For each selected department, the station with the most failures and the part with the most failures "
    "are identified, and the matching failure records are shown in detail. "
    "Export produces one Excel sheet per department containing these detail records."
)

_ts_ff_list = sorted(f["FF"].unique().tolist())
if not _ts_ff_list:
    st.info("No failure data available for this analysis.")
elif "Station" not in f.columns or "Part No" not in f.columns:
    st.info("Need both Station and Part No columns for this analysis.")
else:
    _ts_ffs = st.multiselect(
        "Select Department(s) (FF):",
        _ts_ff_list,
        default=_ts_ff_list,
        key="ts_ff_select",
    )

    if not _ts_ffs:
        st.info("Select at least one department to see the details.")
    else:
        _TS_DETAIL_COLS = [
            "Station", "Part No", "Serial Number", "Date Tested", "Tested By",
            "Tech Comments", "Rework Instructions", "Troubleshoot Completed By",
            "Reference Designator",
        ]

        def _build_ts_details(_ts_ff):
            """Return (station_detail_df, part_detail_df, top_station, top_part, dept_total)."""
            _src = f[f["FF"] == _ts_ff].copy()
            _dept_total = len(_src)

            _src["_st_clean"] = _src["Station"].fillna("(blank)").astype(str).str.strip().replace("", "(blank)")
            _src["_pn_clean"] = _src["Part No"].fillna("(blank)").astype(str).str.strip().replace("", "(blank)")

            _top_station = _src["_st_clean"].value_counts().idxmax()
            _top_part = _src["_pn_clean"].value_counts().idxmax()

            _cols = [c for c in _TS_DETAIL_COLS if c in _src.columns]
            _st_detail = _src[_src["_st_clean"] == _top_station][_cols].sort_values(
                "Date Tested" if "Date Tested" in _cols else _cols[0]
            ).reset_index(drop=True)
            _pn_detail = _src[_src["_pn_clean"] == _top_part][_cols].sort_values(
                "Date Tested" if "Date Tested" in _cols else _cols[0]
            ).reset_index(drop=True)

            return _st_detail, _pn_detail, _top_station, _top_part, _dept_total

        # Summary overview across all selected departments
        _ts_summary_rows = []
        _ts_per_dept = {}
        for _ts_ff in _ts_ffs:
            _st_det, _pn_det, _top_st, _top_pn, _tot = _build_ts_details(_ts_ff)
            _ts_per_dept[_ts_ff] = (_st_det, _pn_det, _top_st, _top_pn, _tot)
            _ts_summary_rows.append({
                "Department": _ts_ff,
                "Total Failures": _tot,
                "Top Station": _top_st,
                "Station Failures": len(_st_det),
                "Top Part No": _top_pn,
                "Part Failures": len(_pn_det),
            })
        st.dataframe(pd.DataFrame(_ts_summary_rows), use_container_width=True, hide_index=True)

        # Per-department detail tabs
        _ts_tabs = st.tabs([str(_d) for _d in _ts_ffs])
        for _ts_tab, _ts_ff in zip(_ts_tabs, _ts_ffs):
            _st_det, _pn_det, _top_st, _top_pn, _tot = _ts_per_dept[_ts_ff]
            with _ts_tab:
                _c1, _c2, _c3 = st.columns(3)
                _c1.metric("Total Failures", f"{_tot:,}")
                _c2.metric("Top Station", str(_top_st), delta=f"{len(_st_det):,} failures", delta_color="off")
                _c3.metric("Top Part No", str(_top_pn), delta=f"{len(_pn_det):,} failures", delta_color="off")

                st.markdown(f"**🔴 Station {_top_st} — {len(_st_det):,} failures** "
                            f"({len(_st_det) / _tot * 100:.1f}% of dept)")
                st.dataframe(_st_det, use_container_width=True, height=300)

                st.markdown(f"**🔴 Part {_top_pn} — {len(_pn_det):,} failures** "
                            f"({len(_pn_det) / _tot * 100:.1f}% of dept)")
                st.dataframe(_pn_det, use_container_width=True, height=300)

        # Export — one sheet per department
        _ts_out = io.BytesIO()
        with pd.ExcelWriter(_ts_out, engine="openpyxl") as _wr:
            _used_ts_names = set()
            pd.DataFrame(_ts_summary_rows).to_excel(_wr, sheet_name="Summary", index=False)
            _used_ts_names.add("Summary")
            for _ts_ff in _ts_ffs:
                _st_det, _pn_det, _top_st, _top_pn, _tot = _ts_per_dept[_ts_ff]
                _ts_safe = re.sub(r"[\\/*?:\[\]]", "_", str(_ts_ff))[:31] or "Sheet"
                if _ts_safe in _used_ts_names:
                    _ts_safe = f"{_ts_safe[:28]}_{len(_used_ts_names)}"
                _used_ts_names.add(_ts_safe)

                _sheet = pd.concat(
                    [
                        _st_det.assign(**{"__Section__": f"Top Station: {_top_st}"}),
                        _pn_det.assign(**{"__Section__": f"Top Part No: {_top_pn}"}),
                    ],
                    ignore_index=True,
                ).rename(columns={"__Section__": "Section"})
                _sheet.to_excel(_wr, sheet_name=_ts_safe, index=False)
        _ts_out.seek(0)
        st.download_button(
            "⬇️ Download Top Station + Top Part details (Excel — one sheet per department)",
            data=_ts_out,
            file_name=f"Top_Station_Part_By_Department_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="ts_download_btn",
        )


# Downloads
st.subheader("⬇️ Download Results")

# Add yield metrics to download
yield_by_ff_download = calculate_yield_by_group(df, "FF") if not df.empty else pd.DataFrame()

# Build filtered full dataset (PASS + FAIL) — only apply FF + date filters
# No column filters so ALL pass and fail records are included
f_all = df.copy()
f_all = f_all[f_all["FF"].isin(selected_ff)]

if date_range and len(date_range) == 2 and "Date Tested" in f_all.columns:
    _dl_start, _dl_end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    f_all = f_all[(f_all["Date Tested"].isna()) | ((f_all["Date Tested"] >= _dl_start) & (f_all["Date Tested"] <= _dl_end))]

download_dfs = {
    "All_Data": f_all,
    "Filtered_Failures": f,
    "Yield_By_FF": yield_by_ff_download,
    "FF_Overview": overview if 'overview' in locals() else pd.DataFrame(),
    "FF_Drivers_Long": drivers_long if 'drivers_long' in locals() else pd.DataFrame(),
    "FF_Combo_Drivers": combo_df if 'combo_df' in locals() else pd.DataFrame(),
    "Column_Mapping": colmap_df,
}

out = to_excel_download(download_dfs)
st.download_button(
    "Download filtered data + driver tables (Excel)",
    data=out,
    file_name=f"FF_Dashboard_Export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
