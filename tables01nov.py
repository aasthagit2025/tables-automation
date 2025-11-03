"""
Tabulation Automation v4.2 — Final Rating + Banners App
-------------------------------------------------------
✔ Correct SPSS value labels & question text
✔ Top2, Top3, Bottom2, Bottom3 & Mean only for rating scales
✔ Nominal questions (e.g., Gender, Ethnicity) show count/% only
✔ Banner columns show counts only (no %)
✔ Wincross-style Excel export (blue header, merged title)
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import pyreadstat
import tempfile
from typing import Dict, Tuple, List
import xlsxwriter

st.set_page_config(page_title="Tabulation Automation v4.2 (Final)", layout="wide")

# ----------------------
# Config
# ----------------------
DEFAULT_DK_CODES = {88, 99, -1, 98}
BLUE_HEADER = "#0070C0"
EXCLUDE_VARS = {"record", "uuid", "source", "date"}

# ----------------------
# File reader (SPSS + raw)
# ----------------------
def read_file(uploaded_file) -> Tuple[pd.DataFrame, dict]:
    name = uploaded_file.name.lower()
    if name.endswith(".sav"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
            tmp.write(uploaded_file.getbuffer())
            tmp_path = tmp.name
        df, meta = pyreadstat.read_sav(tmp_path, apply_value_formats=True)
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
# Rating detection logic
# ----------------------
def is_rating_variable(raw_series: pd.Series, varname: str = "", meta: dict = None) -> bool:
    """Identify rating questions only (1–5, 1–7, or 0–10 numeric scales)."""
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
# Rating metrics (Top/Bottom/Mean)
# ----------------------
def compute_rating_metrics(raw_series: pd.Series, base_mask: pd.Series) -> Dict:
    s = pd.to_numeric(raw_series[base_mask], errors="coerce").dropna()
    if s.empty:
        return {}
    mn = float(s.mean()); sd = float(s.std(ddof=0))
    scale_min = int(s.min()); scale_max = int(s.max())
    width = scale_max - scale_min + 1

    def pct(cond): return round(cond.sum() / len(s) * 100, 1)

    metrics = {"Base": len(s), "Mean": round(mn, 2), "SD": round(sd, 2)}
    if width >= 2:
        metrics["Top2%"] = pct(s >= scale_max - 1)
        metrics["Bottom2%"] = pct(s <= scale_min + 1)
    if width >= 3:
        metrics["Top3%"] = pct(s >= scale_max - 2)
        metrics["Bottom3%"] = pct(s <= scale_min + 2)
    if scale_min == 0 and scale_max == 10:
        prom = ((s >= 9) & (s <= 10)).sum()
        detr = ((s >= 0) & (s <= 6)).sum()
        metrics["NPS"] = round((prom - detr) / len(s) * 100, 1)
    return metrics

# ----------------------
# Count / Percent (Total only)
# ----------------------
def compute_count_pct(series: pd.Series, base_mask: pd.Series, decimals:int=1, show_percent_sign:bool=False) -> pd.DataFrame:
    s = series[base_mask]
    counts = s.value_counts(dropna=False, sort=False)
    if counts.sum() == 0:
        return pd.DataFrame({"Stub": [], "Count": [], "Percent": []})
    pct = (counts / counts.sum() * 100).round(decimals)
    df = pd.DataFrame({"Stub": counts.index, "Count": counts.values, "Percent": pct.values})
    if show_percent_sign:
        df["Percent"] = df["Percent"].astype(str) + "%"
    return df.reset_index(drop=True)

# ----------------------
# Question table (Total + Banners)
# ----------------------
def build_question_table(qvar: str, df_formatted: pd.DataFrame, df_raw: pd.DataFrame,
                         meta: dict, settings: dict, banner_vars: List[str]) -> pd.DataFrame:
    dk_codes = set(settings.get("dk_codes", DEFAULT_DK_CODES))
    decimals = int(settings.get("decimals", 1))
    show_percent_sign = settings.get("show_percent_sign", False)
    raw_series = df_raw[qvar] if qvar in df_raw else df_formatted[qvar]
    base_mask = exclude_dk_mask(raw_series, dk_codes)

    total_df = compute_count_pct(df_formatted[qvar], base_mask, decimals, show_percent_sign)
    if total_df.empty:
        return total_df

    # Banners: counts only (no %)
    for banner in banner_vars:
        if banner not in df_formatted.columns:
            continue
        cats = list(pd.unique(df_formatted[banner].dropna()))
        for cat in cats:
            bmask = (df_formatted[banner] == cat)
            mask_combined = base_mask & bmask
            counts = df_formatted[qvar][mask_combined].value_counts(dropna=False, sort=False)
            total_counts = counts.reindex(total_df["Stub"]).fillna(0).astype(int)
            total_df[f"{banner} - {cat}"] = total_counts.values

    return total_df

# ----------------------
# Worksheet generation
# ----------------------
def generate_all_worksheets(df_formatted: pd.DataFrame, meta: dict, settings: dict, banner_vars: List[str]) -> Dict[str, Tuple[str, pd.DataFrame]]:
    dk_codes = set(settings.get("dk_codes", DEFAULT_DK_CODES))
    df_raw = meta.get("raw_df", df_formatted.copy())
    worksheets = {}
    rating_rows = []

    vars_to_process = [v for v in df_formatted.columns if v.lower() not in EXCLUDE_VARS]

    for v in vars_to_process:
        qlabel = get_label_for_variable(v, meta)
        table = build_question_table(v, df_formatted, df_raw, meta, settings, banner_vars)

        # only compute metrics for rating scales
        if v in df_raw and is_rating_variable(df_raw[v], v, meta):
            base_mask = exclude_dk_mask(df_raw[v], dk_codes)
            metrics = compute_rating_metrics(df_raw[v], base_mask)
            extra_rows = []
            if "Top2%" in metrics:
                extra_rows += [("Top2_Box", metrics["Top2%"]), ("Bottom2_Box", metrics["Bottom2%"])]
            if "Top3%" in metrics:
                extra_rows += [("Top3_Box", metrics["Top3%"]), ("Bottom3_Box", metrics["Bottom3%"])]
            extra_rows += [("Mean", metrics["Mean"]), ("SD", metrics["SD"])]
            if "NPS" in metrics:
                extra_rows.append(("NPS", metrics["NPS"]))

            if extra_rows:
                extra_df = pd.DataFrame([{**{c: "" for c in table.columns}, "Stub": lbl, "Percent": val}
                                         for lbl, val in extra_rows], columns=table.columns)
                table = pd.concat([table, extra_df], ignore_index=True)

            metrics_row = {"Question": qlabel}; metrics_row.update(metrics)
            rating_rows.append(metrics_row)

        worksheets[v] = (qlabel, table)

    if rating_rows:
        summary_df = pd.DataFrame(rating_rows).fillna("")
        means_df = summary_df.reindex(columns=["Question", "Base", "Mean", "SD"]).fillna("")
        worksheets["Means_Summary"] = ("Means Summary", means_df)
        tb_df = summary_df.reindex(columns=["Question", "Top2%", "Bottom2%", "Top3%", "Bottom3%", "NPS"]).fillna("")
        worksheets["TopBottom_Summary"] = ("Top/Bottom Summary", tb_df)

    return worksheets

# ----------------------
# Excel export
# ----------------------
def write_workbook(worksheets: Dict[str, Tuple[str, pd.DataFrame]]) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book
        title_fmt = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 11, "align": "left"})
        header_fmt = wb.add_format({"bold": True, "font_name": "Calibri", "font_size": 10, "align": "center",
                                    "font_color": "white", "bg_color": BLUE_HEADER})
        cell_fmt = wb.add_format({"font_name": "Calibri", "font_size": 10, "align": "center"})
        left_fmt = wb.add_format({"font_name": "Calibri", "font_size": 10, "align": "left"})

        for sheet, (qtext, df) in worksheets.items():
            safe = sheet[:31]
            if df.empty:
                ws = writer.book.add_worksheet(safe)
                ws.merge_range(0, 0, 0, 2, qtext, title_fmt)
                continue
            df.to_excel(writer, sheet_name=safe, index=False, startrow=1)
            ws = writer.sheets[safe]
            ncols = len(df.columns)
            ws.merge_range(0, 0, 0, ncols - 1, qtext, title_fmt)
            for i, col in enumerate(df.columns):
                ws.write(1, i, col, header_fmt)
                ws.set_column(i, i, 18, left_fmt if i == 0 else cell_fmt)
            ws.freeze_panes(2, 1)
            ws.hide_gridlines(2)
    out.seek(0)
    return out.read()

# ----------------------
# Streamlit app
# ----------------------
st.title("Tabulation Automation — v4.2 (Final)")
st.markdown("Upload dataset (.sav/.csv/.xlsx) to generate formatted tables with Top/Bottom nets for rating questions only.")

st.sidebar.header("Settings")
dk_text = st.sidebar.text_input("DK/Ref codes (comma separated)", value="88,99,-1,98")
dk_codes = set(int(x.strip()) for x in dk_text.split(",") if x.strip().lstrip('-').isdigit())
decimals = st.sidebar.number_input("Percent decimals", min_value=0, max_value=2, value=1)
show_percent_sign = st.sidebar.checkbox("Show % symbol in Percent column", value=True)
preview_n = st.sidebar.number_input("Number of tables to preview", 1, 50, value=5)

uploaded = st.file_uploader("Upload data (.sav, .csv, .xlsx)", type=["sav", "csv", "xls", "xlsx"])

if uploaded:
    try:
        df_formatted, meta = read_file(uploaded)
    except Exception as e:
        st.error(f"Unable to read file: {e}")
        st.stop()

    banner_candidates = [c for c in df_formatted.columns if c.lower() not in EXCLUDE_VARS]
    st.sidebar.markdown("### Banner variables (simple)")
    banner_vars = st.sidebar.multiselect("Select banner variable(s) (optional)", options=banner_candidates)

    st.success(f"Loaded {df_formatted.shape[0]} rows × {df_formatted.shape[1]} columns")
    st.dataframe(df_formatted.head(8))

    settings = {"dk_codes": dk_codes, "decimals": decimals, "show_percent_sign": show_percent_sign}

    with st.spinner("Generating tables..."):
        worksheets = generate_all_worksheets(df_formatted, meta, settings, banner_vars)

    st.success(f"Generated {len(worksheets)} sheets (preview below)")

    st.subheader("Preview (first few tables)")
    for i, (sheet, (qtext, df_tab)) in enumerate(worksheets.items()):
        if i >= preview_n:
            break
        st.markdown(f"### {i+1}. {qtext}")
        st.dataframe(df_tab.head(40))

    if st.button("Export formatted Excel workbook"):
        with st.spinner("Writing Excel..."):
            excel_bytes = write_workbook(worksheets)
        st.download_button("Download Excel", data=excel_bytes,
                           file_name="tabulation_total_tables_v4_2.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Upload your dataset to start.")
