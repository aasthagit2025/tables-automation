"""
Tabulation Automation v4.3 — Per-question nets + WinCross banner layout
Save as tables01nov_v4_3.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import pyreadstat
import tempfile
from typing import Dict, Tuple, List
import xlsxwriter

st.set_page_config(page_title="Tabulation Automation v4.3", layout="wide")

# ----------------------
# Config
# ----------------------
DEFAULT_DK_CODES = {88, 99, -1, 98}
BLUE_HEADER = "#0070C0"
EXCLUDE_VARS = {"record", "uuid", "source", "date"}

# ----------------------
# File reader (labels applied + raw numeric backup)
# ----------------------
def read_file(uploaded_file) -> Tuple[pd.DataFrame, dict]:
    name = uploaded_file.name.lower()
    if name.endswith(".sav"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
            tmp.write(uploaded_file.getbuffer())
            tmp_path = tmp.name
        df_formatted, meta_fmt = pyreadstat.read_sav(tmp_path, apply_value_formats=True)
        df_raw, meta_raw = pyreadstat.read_sav(tmp_path, apply_value_formats=False)
        meta_info = {
            "format": "sav",
            "variable_labels": getattr(meta_raw, "variable_labels", {}),
            "value_labels": getattr(meta_raw, "value_labels", {}),
            "raw_df": df_raw
        }
        return df_formatted, meta_info
    elif name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
        return df, {"format": "csv", "variable_labels": {}, "value_labels": {}, "raw_df": df.copy()}
    elif name.endswith((".xls", ".xlsx")):
        df = pd.read_excel(uploaded_file)
        return df, {"format": "excel", "variable_labels": {}, "value_labels": {}, "raw_df": df.copy()}
    else:
        raise ValueError("Unsupported file type. Use .sav, .csv, or .xlsx")

# ----------------------
# Helpers
# ----------------------
def clean_title(text: str) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""
    junk = ["please select one", "select one", "tick one", "choose one", "please select"]
    txt = text.strip()
    for j in junk:
        if j in txt.lower():
            txt = txt.lower().replace(j, "")
    return txt.strip(" :;,-")

def get_label_for_variable(varname: str, meta: dict) -> str:
    vlabels = meta.get("variable_labels", {})
    # case-insensitive match
    if varname in vlabels:
        return clean_title(vlabels[varname])
    for k, v in vlabels.items():
        if k.strip().lower() == varname.strip().lower():
            return clean_title(v)
    return varname

def exclude_dk_mask(series: pd.Series, dk_codes:set):
    if pd.api.types.is_numeric_dtype(series):
        return ~series.isin(dk_codes)
    try:
        conv = pd.to_numeric(series, errors="coerce")
        return ~conv.isin(dk_codes)
    except Exception:
        return pd.Series(True, index=series.index)

# ----------------------
# Rating detection (same improved heuristic)
# ----------------------
def is_rating_variable(raw_series: pd.Series, varname: str = "", meta: dict = None) -> bool:
    if not pd.api.types.is_numeric_dtype(raw_series):
        return False
    nunique = raw_series.dropna().nunique()
    if nunique < 3 or nunique > 11:
        return False
    vlabel = ""
    if meta:
        vlabel = get_label_for_variable(varname, meta).lower()
    text = (varname + " " + vlabel).lower()
    rating_keywords = [
        "satisf", "likeli", "agree", "importance", "recommend",
        "happy", "quality", "performance", "trust", "confidence", "rating", "score"
    ]
    return any(kw in text for kw in rating_keywords)

# ----------------------
# Rating metrics
# ----------------------
def compute_rating_metrics(raw_series: pd.Series, base_mask: pd.Series) -> Dict:
    s = pd.to_numeric(raw_series[base_mask], errors="coerce").dropna()
    if s.empty:
        return {}
    mn = float(s.mean()); sd = float(s.std(ddof=0))
    scale_min = int(s.min()); scale_max = int(s.max())
    width = scale_max - scale_min + 1
    def pct_of(cond):
        return round(cond.sum() / len(s) * 100, 1)
    out = {"Base": len(s), "Mean": round(mn,2), "SD": round(sd,2)}
    if width >= 2:
        out["Top2%"] = pct_of(s >= scale_max - 1)
        out["Bottom2%"] = pct_of(s <= scale_min + 1)
    if width >= 3:
        out["Top3%"] = pct_of(s >= scale_max - 2)
        out["Bottom3%"] = pct_of(s <= scale_min + 2)
    if scale_min == 0 and scale_max == 10:
        prom = ((s>=9)&(s<=10)).sum()
        detr = ((s>=0)&(s<=6)).sum()
        out["NPS"] = round((prom-detr)/len(s)*100,1)
    return out

# ----------------------
# Count/Pct for Total; returns DataFrame with Stub | Total Count | Total %
# ----------------------
def compute_total_count_pct(formatted_series: pd.Series, base_mask: pd.Series, decimals:int=1, show_percent_sign:bool=False) -> pd.DataFrame:
    s = formatted_series[base_mask]
    counts = s.value_counts(dropna=False, sort=False)
    if counts.sum() == 0:
        return pd.DataFrame({"Stub": [], "Total Count": [], "Total %": []})
    pct = (counts / counts.sum() * 100).round(decimals)
    df = pd.DataFrame({"Stub": counts.index, "Total Count": counts.values, "Total %": pct.values})
    if show_percent_sign:
        df["Total %"] = df["Total %"].astype(str) + "%"
    return df.reset_index(drop=True)

# ----------------------
# Build question table: Total first, then for each banner category add Count + % (banner base)
# Also appends selected nets rows if requested (nets_selected list)
# ----------------------
def build_question_table(qvar: str,
                         df_formatted: pd.DataFrame,
                         df_raw: pd.DataFrame,
                         meta: dict,
                         settings: dict,
                         banner_vars: List[str],
                         nets_selected: List[str]) -> pd.DataFrame:

    dk_codes = set(settings.get("dk_codes", DEFAULT_DK_CODES))
    decimals = int(settings.get("decimals", 1))
    show_percent_sign = settings.get("show_percent_sign", False)

    # derive raw/ formatted series
    raw_series = df_raw[qvar] if qvar in df_raw else df_formatted[qvar]
    base_mask = exclude_dk_mask(raw_series, dk_codes)

    # Total columns
    total_df = compute_total_count_pct(df_formatted[qvar], base_mask, decimals, show_percent_sign)
    if total_df.empty:
        return total_df

    # For each banner, for each category, compute Count + % (percent relative to banner category base)
    for banner in banner_vars:
        if banner not in df_formatted.columns:
            continue
        cats = list(pd.unique(df_formatted[banner].dropna()))
        for cat in cats:
            bmask = (df_formatted[banner] == cat)
            mask_combined = base_mask & bmask
            counts = df_formatted[qvar][mask_combined].value_counts(dropna=False, sort=False)
            # align with main stub order
            aligned_counts = counts.reindex(total_df["Stub"]).fillna(0).astype(int)
            denom = aligned_counts.sum()
            if denom == 0:
                pct = pd.Series([0]*len(aligned_counts), index=aligned_counts.index)
            else:
                pct = (aligned_counts / denom * 100).round(decimals)
            # apply percent sign if requested
            pct_display = (pct.astype(str) + "%") if show_percent_sign else pct
            total_df[f"{banner} - {cat} Count"] = aligned_counts.values
            total_df[f"{banner} - {cat} %"] = pct_display.values

    # If nets selected for this question, and it's a rating variable, compute and append rows
    if nets_selected:
        # Ensure this is truly a rating var (use raw)
        if qvar in df_raw and is_rating_variable(df_raw[qvar], qvar, meta):
            metrics = compute_rating_metrics(df_raw[qvar], base_mask)
            extra_rows = []
            for metric in nets_selected:
                lbl = metric
                val = None
                if metric == "Top2":
                    val = metrics.get("Top2%", "")
                elif metric == "Bottom2":
                    val = metrics.get("Bottom2%", "")
                elif metric == "Top3":
                    val = metrics.get("Top3%", "")
                elif metric == "Bottom3":
                    val = metrics.get("Bottom3%", "")
                elif metric == "Mean":
                    val = metrics.get("Mean", "")
                elif metric == "SD":
                    val = metrics.get("SD", "")
                elif metric == "NPS":
                    val = metrics.get("NPS", "")
                # Build a row with same columns as total_df; put value in 'Total %' column for nets (for Mean use Total % too)
                row = {c: "" for c in total_df.columns}
                row["Stub"] = lbl
                # display numbers in Total % column (keeps consistent placement)
                row["Total %"] = val
                extra_rows.append(row)
            if extra_rows:
                extra_df = pd.DataFrame(extra_rows, columns=total_df.columns)
                total_df = pd.concat([total_df, extra_df], ignore_index=True)

    return total_df

# ----------------------
# Generate all worksheets: uses nets_config dict mapping var -> list of selected metrics
# ----------------------
def generate_all_worksheets(df_formatted: pd.DataFrame, meta: dict, settings: dict,
                            banner_vars: List[str], nets_config: Dict[str, List[str]]) -> Dict[str, Tuple[str, pd.DataFrame]]:
    dk_codes = set(settings.get("dk_codes", DEFAULT_DK_CODES))
    df_raw = meta.get("raw_df", df_formatted.copy())
    worksheets: Dict[str, Tuple[str, pd.DataFrame]] = {}
    rating_summary_rows = []

    vars_to_process = [v for v in df_formatted.columns if v.lower() not in EXCLUDE_VARS]

    for v in vars_to_process:
        qlabel = get_label_for_variable(v, meta)
        nets_selected = nets_config.get(v, [])  # list of metrics for this var
        table = build_question_table(v, df_formatted, df_raw, meta, settings, banner_vars, nets_selected)

        # collect rating summary (if rating)
        if v in df_raw and is_rating_variable(df_raw[v], v, meta):
            base_mask = exclude_dk_mask(df_raw[v], dk_codes)
            metrics = compute_rating_metrics(df_raw[v], base_mask)
            row = {"Question": qlabel}
            row.update(metrics)
            rating_summary_rows.append(row)

        worksheets[v] = (qlabel, table)

    # Add rating summaries sheets if any
    if rating_summary_rows:
        rs_df = pd.DataFrame(rating_summary_rows).fillna("")
        means_df = rs_df.reindex(columns=["Question", "Base", "Mean", "SD"]).fillna("")
        tb_df = rs_df.reindex(columns=["Question", "Top2%", "Bottom2%", "Top3%", "Bottom3%", "NPS"]).fillna("")
        worksheets["Means_Summary"] = ("Means Summary", means_df)
        worksheets["TopBottom_Summary"] = ("Top/Bottom Summary", tb_df)

    return worksheets

# ----------------------
# Excel writer
# ----------------------
def write_workbook(worksheets: Dict[str, Tuple[str, pd.DataFrame]]) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book
        title_fmt = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 11, "align": "left", "valign":"vcenter"})
        header_fmt = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 10, "align": "center", "valign":"vcenter", "font_color":"white", "bg_color":BLUE_HEADER})
        cell_fmt = wb.add_format({"font_name":"Calibri", "font_size":10, "align":"center", "valign":"vcenter"})
        left_fmt = wb.add_format({"font_name":"Calibri", "font_size":10, "align":"left", "valign":"vcenter"})

        for sheet, (qtext, df) in worksheets.items():
            safe = sheet[:31]
            if df is None or df.empty:
                ws = writer.book.add_worksheet(safe)
                ws.merge_range(0,0,0,2,qtext, title_fmt)
                continue
            df.to_excel(writer, sheet_name=safe, index=False, startrow=1)
            ws = writer.sheets[safe]
            ncols = len(df.columns)
            ws.merge_range(0, 0, 0, max(0, ncols-1), qtext, title_fmt)
            for i, col in enumerate(df.columns):
                ws.write(1, i, col, header_fmt)
                # wider first column (Stub)
                if i == 0:
                    ws.set_column(i, i, 30, left_fmt)
                else:
                    ws.set_column(i, i, 14, cell_fmt)
            ws.freeze_panes(2, 1)
            ws.hide_gridlines(2)
    out.seek(0)
    return out.read()

# ----------------------
# Streamlit UI
# ----------------------
st.title("Tabulation Automation — v4.3 (Per-question nets + Banner layout)")
st.markdown("Upload dataset (.sav, .csv, .xlsx). Configure per-question nets in sidebar, preview, then export Excel.")

# Sidebar controls
st.sidebar.header("Settings")
dk_text = st.sidebar.text_input("DK/Ref codes (comma separated)", value="88,99,-1,98")
dk_codes = set(int(x.strip()) for x in dk_text.split(",") if x.strip().lstrip('-').isdigit())
decimals = st.sidebar.number_input("Percent decimals", min_value=0, max_value=2, value=1)
show_percent_sign = st.sidebar.checkbox("Show % sign in Total and Banner % columns", value=True)
preview_n = st.sidebar.number_input("Number of tables to preview", min_value=1, max_value=50, value=5)

uploaded = st.file_uploader("Upload data (.sav, .csv, .xlsx)", type=["sav","csv","xls","xlsx"])

if uploaded:
    try:
        df_formatted, meta = read_file(uploaded)
    except Exception as e:
        st.error(f"Unable to read file: {e}")
        st.stop()

    # candidate banners: exclude meta vars
    banner_candidates = [c for c in df_formatted.columns if c.lower() not in EXCLUDE_VARS]
    st.sidebar.markdown("### Banner variables (simple)")
    banner_vars = st.sidebar.multiselect("Select banner variable(s) to split by", options=banner_candidates)

    st.sidebar.markdown("---")
    st.sidebar.markdown("### Nets / Means configuration (per question)")
    enable_nets = st.sidebar.checkbox("Enable nets/means configuration", value=False)

    # Build nets config interactively if enabled
    nets_config: Dict[str, List[str]] = {}
    if enable_nets:
        st.sidebar.markdown("Select questions (rating questions detected) and metrics to compute")
        # detect candidate rating vars
        df_raw = meta.get("raw_df", df_formatted.copy())
        rating_candidates = []
        for v in df_formatted.columns:
            if v.lower() in EXCLUDE_VARS:
                continue
            if v in df_raw and is_rating_variable(df_raw[v], v, meta):
                rating_candidates.append(v)
        if rating_candidates:
            # For each detected rating question, show a collapsible multi-select for metrics
            for v in rating_candidates:
                qlabel = get_label_for_variable(v, meta)
                with st.sidebar.expander(f"{qlabel} ({v})", expanded=False):
                    opts = st.multiselect(
                        "Select metrics to add (top/bottom/mean/sd/nps)",
                        options=["Top2", "Top3", "Bottom2", "Bottom3", "Mean", "SD", "NPS"],
                        key=f"nets_{v}"
                    )
                    if opts:
                        nets_config[v] = opts
        else:
            st.sidebar.markdown("No rating questions detected to configure.")

    # Show loaded data preview
    st.success(f"Loaded {df_formatted.shape[0]} rows × {df_formatted.shape[1]} columns")
    st.dataframe(df_formatted.head(8))

    # Generate tables
    settings = {"dk_codes": dk_codes, "decimals": decimals, "show_percent_sign": show_percent_sign}
    with st.spinner("Generating tables..."):
        worksheets = generate_all_worksheets(df_formatted, meta, settings, banner_vars, nets_config)

    st.success(f"Generated {len(worksheets)} sheets (includes summaries)")

    # Preview first N
    st.subheader("Preview generated tables")
    shown = 0
    for sheet, (qtext, df_tab) in worksheets.items():
        if shown >= preview_n:
            break
        st.markdown(f"### {shown+1}. {qtext} — {sheet}")
        if df_tab is None or df_tab.empty:
            st.write("No data.")
        else:
            st.dataframe(df_tab.head(60))
        shown += 1

    # Export
    if st.button("Export formatted Excel workbook"):
        with st.spinner("Writing Excel..."):
            excel_bytes = write_workbook(worksheets)
        st.download_button("Download Excel", data=excel_bytes,
                           file_name="tabulation_v4_3_export.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Upload your dataset to start.")
