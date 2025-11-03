"""
Tabulation Automation v4 — Top/Bottom/Mean + Simple Banners
- Builds on prior v3.3: correct SPSS labels, previews, Wincross-style export
- Adds Top2/Top3/Bottom2/Bottom3 + Mean/SD inline and summary sheets
- Simple banner variables: select existing variables to split by their categories
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import pyreadstat
import tempfile
from typing import Dict, Tuple, List
import xlsxwriter

st.set_page_config(page_title="Tabulation Automation v4", layout="wide")

# ----------------------
# Config
# ----------------------
DEFAULT_DK_CODES = {88, 99, -1, 98}
BLUE_HEADER = "#0070C0"
EXCLUDE_VARS = {"record", "uuid", "source", "date"}

# ----------------------
# File reader (labels applied + raw backup)
# ----------------------
def read_file(uploaded_file) -> Tuple[pd.DataFrame, dict]:
    name = uploaded_file.name.lower()
    if name.endswith(".sav"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
            tmp.write(uploaded_file.getbuffer())
            tmp_path = tmp.name
        # df has formatted labels (strings where labels exist)
        df, meta = pyreadstat.read_sav(tmp_path, apply_value_formats=True)
        # raw numeric-codes version
        df_raw, meta_raw = pyreadstat.read_sav(tmp_path, apply_value_formats=False)
        meta_info = {
            "format": "sav",
            "variable_labels": getattr(meta_raw, "variable_labels", {}),
            "value_labels": getattr(meta_raw, "value_labels", {}),
            "raw_df": df_raw
        }
        return df, meta_info
    elif name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
        return df, {"format": "csv", "variable_labels": {}, "value_labels": {}, "raw_df": df.copy()}
    elif name.endswith((".xls", ".xlsx")):
        df = pd.read_excel(uploaded_file)
        return df, {"format": "excel", "variable_labels": {}, "value_labels": {}, "raw_df": df.copy()}
    else:
        raise ValueError("Unsupported file type. Use .sav, .csv, or .xlsx")

# ----------------------
# Helpers for labels & DK
# ----------------------
def clean_title(text: str) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""
    remove_phrases = ["please select one", "select one", "tick one", "choose one", "please select"]
    txt = text.strip()
    for p in remove_phrases:
        if p in txt.lower():
            txt = txt.lower().replace(p, "")
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

def value_label_map_for_var(varname: str, meta: dict) -> Dict:
    all_maps = meta.get("value_labels", {})
    if varname in all_maps:
        return all_maps[varname]
    for k, mapping in all_maps.items():
        if k.strip().lower() == varname.strip().lower():
            return mapping
    return {}

def exclude_dk_mask(series: pd.Series, dk_codes:set):
    if pd.api.types.is_numeric_dtype(series):
        return ~series.isin(dk_codes)
    try:
        conv = pd.to_numeric(series, errors="coerce")
        return ~conv.isin(dk_codes)
    except Exception:
        return pd.Series(True, index=series.index)

# ----------------------
# Rating detection + metrics
# ----------------------
def is_rating_variable(raw_series: pd.Series) -> bool:
    # raw series should be numeric-coded (use meta raw_df)
    if not pd.api.types.is_numeric_dtype(raw_series):
        return False
    nunique = raw_series.dropna().nunique()
    return 3 <= nunique <= 11

def compute_rating_metrics(raw_series: pd.Series, base_mask: pd.Series) -> Dict:
    s = pd.to_numeric(raw_series[base_mask], errors="coerce").dropna()
    if s.empty:
        return {}
    mn = float(s.mean())
    sd = float(s.std(ddof=0))
    mn = round(mn, 2); sd = round(sd, 2)
    scale_min = int(s.min()); scale_max = int(s.max())
    width = scale_max - scale_min + 1

    def pct_of(cond):
        return round(cond.sum() / len(s) * 100, 1)

    metrics = {"Base": len(s), "Mean": mn, "SD": sd}
    # Top/Bottom 2 & 3 conditional on width
    if width >= 2:
        top2 = (s >= scale_max - 1)
        bottom2 = (s <= scale_min + 1)
        metrics["Top2%"] = pct_of(top2)
        metrics["Bottom2%"] = pct_of(bottom2)
    if width >= 3:
        top3 = (s >= scale_max - 2)
        bottom3 = (s <= scale_min + 2)
        metrics["Top3%"] = pct_of(top3)
        metrics["Bottom3%"] = pct_of(bottom3)
    # Special NPS (0-10)
    if scale_min == 0 and scale_max == 10:
        prom = ((s >= 9) & (s <= 10))
        detr = ((s >= 0) & (s <= 6))
        metrics["NPS"] = round((prom.sum() - detr.sum())/len(s) * 100, 1)
    return metrics

# ----------------------
# Count / Percent (with labels already applied in formatted df)
# ----------------------
def compute_count_pct(formatted_series: pd.Series, base_mask: pd.Series, decimals:int=1, show_percent_sign:bool=False) -> pd.DataFrame:
    s = formatted_series[base_mask]
    counts = s.value_counts(dropna=False, sort=False)
    if counts.sum() == 0:
        df = pd.DataFrame({"Stub": [], "Count": [], "Percent": []})
        return df
    pct = (counts / counts.sum() * 100).round(decimals)
    df = pd.DataFrame({"Stub": counts.index, "Count": counts.values, "Percent": pct.values})
    if show_percent_sign:
        df["Percent"] = df["Percent"].astype(str) + "%"
    return df.reset_index(drop=True)

# ----------------------
# Build table for one question (Total + banners)
# ----------------------
def build_question_table(qvar: str, df_formatted: pd.DataFrame, df_raw: pd.DataFrame,
                         meta: dict, settings: dict, banner_vars: List[str]) -> pd.DataFrame:
    dk_codes = set(settings.get("dk_codes", DEFAULT_DK_CODES))
    decimals = settings.get("decimals", 1)
    show_percent_sign = settings.get("show_percent_sign", False)

    # base mask on raw data (exclude DK codes)
    raw_series = df_raw[qvar] if qvar in df_raw else df_formatted[qvar]
    base_mask = exclude_dk_mask(raw_series, dk_codes)

    # start with Total columns (Stub, Count, Percent)
    total_df = compute_count_pct(df_formatted[qvar], base_mask, decimals, show_percent_sign)

    # If no categories, return empty
    if total_df.empty:
        return total_df

    # For each banner variable, add columns: <Banner> - <Category> Count & Percent
    # We'll reshape: for each banner, for each category, compute count & pct and add columns named Banner:Category Count / Banner:Category Percent
    for banner in banner_vars:
        if banner not in df_formatted.columns:
            continue
        # categories are already label-applied in formatted df
        cats = list(pd.unique(df_formatted[banner].dropna()))
        # For stable column order, use value_labels mapping if exists in meta to get labels; else unique()
        for cat in cats:
            # mask for banner category using formatted df (labels)
            bmask = (df_formatted[banner] == cat)
            mask_combined = base_mask & bmask
            counts = df_formatted[qvar][mask_combined].value_counts(dropna=False, sort=False)
            total_counts = counts.reindex(total_df["Stub"]).fillna(0).astype(int)
            pct = (total_counts / total_counts.sum().replace(0, np.nan) * 100).fillna(0).round(decimals)  # per category base
            if show_percent_sign:
                pct = pct.astype(str) + "%"
            # add Count column for this banner category, aligned to stubs order
            total_df[f"{banner} - {cat} Count"] = total_counts.values
            total_df[f"{banner} - {cat} Percent"] = pct.values

    return total_df

# ----------------------
# Generate all worksheets
# ----------------------
def generate_all_worksheets(df_formatted: pd.DataFrame, meta: dict, settings: dict, banner_vars: List[str]) -> Dict[str, Tuple[str, pd.DataFrame]]:
    dk_codes = set(settings.get("dk_codes", DEFAULT_DK_CODES))
    decimals = settings.get("decimals", 1)
    show_percent_sign = settings.get("show_percent_sign", False)

    df_raw = meta.get("raw_df", df_formatted.copy())
    worksheets = {}
    rating_rows = []  # for summary

    # filter variables
    vars_to_process = [v for v in df_formatted.columns if v.lower() not in EXCLUDE_VARS]

    for v in vars_to_process:
        qlabel = get_label_for_variable(v, meta)
        # build table (includes banner splits)
        table = build_question_table(v, df_formatted, df_raw, meta, settings, banner_vars)

        # compute rating metrics using raw series (for mean/net)
        if v in df_raw and is_rating_variable(df_raw[v]):
            base_mask = exclude_dk_mask(df_raw[v], dk_codes)
            metrics = compute_rating_metrics(df_raw[v], base_mask)
            # append metrics as additional rows to the table (inline)
            # we'll append human-friendly rows at bottom
            add_rows = []
            # Top2/Bottom2
            if "Top2%" in metrics:
                add_rows.append(("Top2_Box", "", metrics.get("Top2%","")))
                add_rows.append(("Bottom2_Box", "", metrics.get("Bottom2%","")))
            if "Top3%" in metrics:
                add_rows.append(("Top3_Box", "", metrics.get("Top3%","")))
                add_rows.append(("Bottom3_Box", "", metrics.get("Bottom3%","")))
            # Mean and SD
            add_rows.append(("Mean", "", metrics.get("Mean","")))
            add_rows.append(("SD", "", metrics.get("SD","")))
            # NPS if present
            if "NPS" in metrics:
                add_rows.append(("NPS", "", metrics.get("NPS","")))

            # convert add_rows into df with same columns; put counts blank and percent in Percent column for nets (or value in Count for Mean)
            # We'll add columns required by table
            extra_df = pd.DataFrame(columns=table.columns)
            for label, _, val in add_rows:
                row = {c: "" for c in table.columns}
                # place label in Stub column
                if "Stub" in row:
                    row["Stub"] = label
                # where to place metric? put numeric metrics in Percent (for nets) or Count for Mean? place numbers in Percent for consistency
                if "Percent" in row:
                    row["Percent"] = f"{val}" if not isinstance(val, (int,float)) else val
                extra_df = extra_df.append(row, ignore_index=True)
            table = pd.concat([table, extra_df], ignore_index=True)

            # Add to rating summary
            metrics_row = {"Question": qlabel}
            metrics_row.update(metrics)
            rating_rows.append(metrics_row)

        worksheets[v] = (qlabel, table)

    # create summary sheets
    if rating_rows:
        summary_df = pd.DataFrame(rating_rows)
        # Means_Summary
        means_cols = ["Question", "Base", "Mean", "SD"]
        means_df = summary_df.reindex(columns=means_cols).fillna("")
        worksheets["Means_Summary"] = ("Means Summary", means_df)
        # TopBottom summary
        tb_cols = ["Question", "Top2%", "Bottom2%", "Top3%", "Bottom3%", "NPS"]
        tb_df = summary_df.reindex(columns=tb_cols).fillna("")
        worksheets["TopBottom_Summary"] = ("Top/Bottom Summary", tb_df)

    # If banners selected, build per-banner summaries: for each banner var, compute metrics across banner categories for rating variables
    for banner in banner_vars:
        if banner not in df_formatted.columns:
            continue
        # get banner categories (formatted)
        cats = list(pd.unique(df_formatted[banner].dropna()))
        banner_rows = []
        for v in vars_to_process:
            if v not in df_raw or not is_rating_variable(df_raw[v]):
                continue
            qlabel = get_label_for_variable(v, meta)
            row = {"Question": qlabel}
            for cat in cats:
                # mask: base & banner==cat
                base_mask = exclude_dk_mask(df_raw[v], dk_codes) & (df_formatted[banner] == cat)
                metrics = compute_rating_metrics(df_raw[v], base_mask)
                # store Top2% in banner column
                row[f"{cat} Top2%"] = metrics.get("Top2%", "")
                row[f"{cat} Mean"] = metrics.get("Mean", "")
            banner_rows.append(row)
        if banner_rows:
            bdf = pd.DataFrame(banner_rows).fillna("")
            worksheets[f"Summary_by_{banner}"] = (f"Summary by {banner}", bdf)

    return worksheets

# ----------------------
# Excel writer (Wincross style)
# ----------------------
def write_workbook(worksheets: Dict[str, Tuple[str, pd.DataFrame]]) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book
        title_fmt = wb.add_format({"bold": True, "font_name":"Calibri", "font_size":11, "align":"left", "valign":"vcenter"})
        header_fmt = wb.add_format({"bold": True, "font_name":"Calibri", "font_size":10, "align":"center", "valign":"vcenter", "font_color":"white", "bg_color":BLUE_HEADER})
        cell_fmt = wb.add_format({"font_name":"Calibri","font_size":10,"align":"center","valign":"vcenter"})
        left_fmt = wb.add_format({"font_name":"Calibri","font_size":10,"align":"left","valign":"vcenter"})

        for sheet, (qtext, df) in worksheets.items():
            safe = sheet[:31]
            # prevent empty df error
            if df is None or df.empty:
                # write empty sheet with title only
                ws = writer.book.add_worksheet(safe)
                ws.merge_range(0,0,0,2, qtext, title_fmt)
                continue
            df.to_excel(writer, sheet_name=safe, index=False, startrow=1)
            ws = writer.sheets[safe]
            ncols = len(df.columns)
            ws.merge_range(0, 0, 0, max(0, ncols-1), qtext, title_fmt)
            # header row styling and column widths
            for i, col in enumerate(df.columns):
                ws.write(1, i, col, header_fmt)
                # left align first column (Stub) else center
                if i == 0:
                    ws.set_column(i, i, 30, left_fmt)
                else:
                    ws.set_column(i, i, 15, cell_fmt)
            ws.freeze_panes(2, 1)
            ws.hide_gridlines(2)

    out.seek(0)
    return out.read()

# ----------------------
# Streamlit UI
# ----------------------
st.title("Tabulation Automation — v4 (Top/Bottom/Mean + Banners)")
st.markdown("Upload dataset (.sav/.csv/.xlsx). After preview, export Excel with Top/Bottom and summary sheets.")

# Sidebar controls
st.sidebar.header("Settings")
dk_text = st.sidebar.text_input("DK/Ref codes (comma separated)", value="88,99,-1,98")
dk_codes = set(int(x.strip()) for x in dk_text.split(",") if x.strip().lstrip('-').isdigit())
decimals = st.sidebar.number_input("Percent decimals", min_value=0, max_value=2, value=1)
show_percent_sign = st.sidebar.checkbox("Show % symbol in Percent column", value=True)
preview_n = st.sidebar.number_input("Number of tables to preview", 1, 50, value=5)

# banner selection - allow multiple simple variables
uploaded = st.file_uploader("Upload data (.sav, .csv, .xlsx)", type=["sav","csv","xls","xlsx"])

if uploaded:
    try:
        df_formatted, meta = read_file(uploaded)
    except Exception as e:
        st.error(f"Unable to read file: {e}")
        st.stop()

    # list candidate banner variables (exclude meta)
    candidate_banners = [c for c in df_formatted.columns if c.lower() not in EXCLUDE_VARS]
    st.sidebar.markdown("### Banner variables (simple)")
    banner_vars = st.sidebar.multiselect("Choose banner variable(s) to split by (optional)", options=candidate_banners)

    st.success(f"Loaded {df_formatted.shape[0]} rows × {df_formatted.shape[1]} columns")
    st.dataframe(df_formatted.head(8))

    settings = {"dk_codes": dk_codes, "decimals": decimals, "show_percent_sign": show_percent_sign}

    with st.spinner("Generating tables and summaries..."):
        worksheets = generate_all_worksheets(df_formatted, meta, settings, banner_vars)

    st.success(f"Generated {len(worksheets)} sheets (preview below)")

    # Preview first N sheets
    st.subheader("Preview generated tables")
    shown = 0
    for sheet, (qtext, df_tab) in worksheets.items():
        if shown >= preview_n: break
        st.markdown(f"### {sheet} — {qtext}")
        if df_tab is None or df_tab.empty:
            st.write("No data.")
        else:
            st.dataframe(df_tab.head(40))
        shown += 1

    # Export option
    if st.button("Export formatted Excel workbook"):
        with st.spinner("Writing Excel..."):
            excel_bytes = write_workbook(worksheets)
        st.download_button("Download Excel", data=excel_bytes,
                           file_name="tabulation_v4_export.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Upload a dataset to start.")
