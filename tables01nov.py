"""
Tabulation Automation v3 — Wincross-style Total Tables
------------------------------------------------------
Features:
- Reads SPSS (.sav), Excel, CSV
- Uses variable labels (Question) & value labels (Stub)
- Excludes DK/Ref codes
- Adds Mean, Top2/Bottom2, NPS summaries
- Toggle for showing % symbol
- Exports formatted Excel workbook (merged header)
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import pyreadstat
import tempfile
import xlsxwriter
from typing import Dict, Tuple

st.set_page_config(page_title="Tabulation Automation v3", layout="wide")

# --------------------------
# Config
# --------------------------
DEFAULT_DK_CODES = {88, 99, -1, 98}
BLUE_HEADER = "#0070C0"

# --------------------------
# File Reader
# --------------------------
def read_file(uploaded_file) -> Tuple[pd.DataFrame, dict]:
    name = uploaded_file.name.lower()

    if name.endswith(".sav"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
            tmp.write(uploaded_file.getbuffer())
            tmp_path = tmp.name

        df, meta = pyreadstat.read_sav(tmp_path, apply_value_formats=False)
        meta_info = {
            "format": "sav",
            "variable_labels": getattr(meta, "variable_labels", {}),
            "value_labels": getattr(meta, "value_labels", {})
        }
        return df, meta_info

    elif name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
        return df, {"format": "csv", "variable_labels": {}, "value_labels": {}}

    elif name.endswith((".xls", ".xlsx")):
        df = pd.read_excel(uploaded_file)
        return df, {"format": "excel", "variable_labels": {}, "value_labels": {}}

    else:
        raise ValueError("Unsupported file type. Use .sav, .csv, or .xlsx")

# --------------------------
# Helpers
# --------------------------
def clean_title(text: str) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""
    remove_phrases = ["please select one", "select one", "tick one", "choose one", "please select"]
    txt = text.strip()
    for p in remove_phrases:
        txt = txt.replace(p, "", 1) if p in txt.lower() else txt
    return txt.strip(" :;,-")

def get_label_for_variable(varname: str, meta: dict) -> str:
    vlabels = meta.get("variable_labels", {})
    return clean_title(vlabels.get(varname, varname))

def value_label_map_for_var(varname: str, meta: dict) -> Dict:
    all_maps = meta.get("value_labels", {})
    if varname in all_maps:
        return all_maps[varname]
    for k, mm in all_maps.items():
        if k.lower() == varname.lower():
            return mm
    return {}

def exclude_dk(series: pd.Series, dk_codes:set):
    if pd.api.types.is_numeric_dtype(series):
        return ~series.isin(dk_codes)
    try:
        conv = pd.to_numeric(series, errors="coerce")
        return ~conv.isin(dk_codes)
    except Exception:
        return pd.Series(True, index=series.index)

# --------------------------
# Count/Percent Table
# --------------------------
def compute_count_pct(series: pd.Series, base_mask: pd.Series, decimals:int=0,
                      value_labels:dict=None, show_percent_sign:bool=False) -> pd.DataFrame:
    s = series[base_mask]
    counts = s.value_counts(dropna=False, sort=False)
    total = counts.sum()
    pct = (counts / total * 100).round(decimals)

    df = pd.DataFrame({
        "Stub": counts.index,
        "Count": counts.values,
        "Percent": pct.values
    })

    def label_stub(val):
        if pd.isna(val): return "<No Answer>"
        if not value_labels: return val
        try:
            val_num = int(float(val))
            if val_num in value_labels:
                return value_labels[val_num]
        except Exception:
            pass
        val_str = str(val).strip()
        if val_str in value_labels:
            return value_labels[val_str]
        for k, lbl in value_labels.items():
            if str(k).strip().lower() == val_str.lower():
                return lbl
        return val

    df["Stub"] = df["Stub"].apply(label_stub)
    if show_percent_sign:
        df["Percent"] = df["Percent"].astype(str) + "%"
    return df

# --------------------------
# Rating / NPS Summary
# --------------------------
def rating_summary(series: pd.Series, base_mask: pd.Series) -> dict:
    s = pd.to_numeric(series[base_mask], errors="coerce").dropna()
    if s.empty:
        return {}
    mn, sd = s.mean(), s.std(ddof=0)
    scale_min, scale_max = int(s.min()), int(s.max())
    width = scale_max - scale_min + 1
    top2 = (s >= scale_max-1).sum()
    bottom2 = (s <= scale_min+1).sum()
    out = {
        "Base": len(s),
        "Mean": round(mn, 2),
        "SD": round(sd, 2),
        "Top2%": round(top2/len(s)*100, 1),
        "Bottom2%": round(bottom2/len(s)*100, 1)
    }
    if scale_min == 0 and scale_max == 10:
        prom = ((s>=9)&(s<=10)).sum()
        detr = ((s>=0)&(s<=6)).sum()
        nps = round((prom-detr)/len(s)*100,1)
        out["NPS"] = nps
    return out

# --------------------------
# Table Generator
# --------------------------
def generate_tabulation(df: pd.DataFrame, meta: dict, settings: dict) -> Dict[str, pd.DataFrame]:
    dk_codes = set(settings.get("dk_codes", DEFAULT_DK_CODES))
    decimals = settings.get("decimals", 0)
    show_percent_sign = settings.get("show_percent_sign", False)
    worksheets, rating_summaries = {}, []

    for v in df.columns:
        s = df[v]
        base_mask = exclude_dk(s, dk_codes)
        vmap = value_label_map_for_var(v, meta)
        qtext = get_label_for_variable(v, meta)

        if pd.api.types.is_numeric_dtype(s) and s.dropna().nunique() >= 3 and s.dropna().nunique() <= 11:
            rating_summaries.append({
                "Question": qtext,
                **rating_summary(s, base_mask)
            })

        t = compute_count_pct(s, base_mask, decimals, vmap, show_percent_sign)
        worksheets[v] = (qtext, t)

    if rating_summaries:
        worksheets["Mean_Top2_Bottom2_Summary"] = ("Summary", pd.DataFrame(rating_summaries))
    return worksheets

# --------------------------
# Excel Export
# --------------------------
def write_workbook(worksheets: Dict[str, Tuple[str, pd.DataFrame]]) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book
        title_fmt = wb.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 11,
            "align": "left", "valign": "vcenter"
        })
        header_fmt = wb.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 10,
            "align": "center", "valign": "vcenter",
            "font_color": "white", "bg_color": BLUE_HEADER
        })
        cell_fmt = wb.add_format({
            "font_name": "Calibri", "font_size": 10,
            "align": "center", "valign": "vcenter"
        })

        for sheet, (qtext, df) in worksheets.items():
            safe = sheet[:31]
            df.to_excel(writer, sheet_name=safe, index=False, startrow=1)
            ws = writer.sheets[safe]
            ncols = len(df.columns)
            # merged title row
            ws.merge_range(0, 0, 0, ncols - 1, qtext, title_fmt)
            for i, col in enumerate(df.columns):
                ws.write(1, i, col, header_fmt)
                ws.set_column(i, i, 20, cell_fmt)
            ws.freeze_panes(2, 1)
            ws.hide_gridlines(2)

    out.seek(0)
    return out.read()

# --------------------------
# Streamlit UI
# --------------------------
st.title("Tabulation Automation — v3 (Wincross Style)")
st.markdown("Upload your dataset (.sav, .csv, .xlsx) to generate formatted Total tables with Mean/Top2/NPS summaries.")

# Sidebar
st.sidebar.header("Settings")
dk_text = st.sidebar.text_input("DK/Ref codes (comma separated)", value="88,99,-1,98")
dk_codes = set(int(x.strip()) for x in dk_text.split(",") if x.strip().lstrip('-').isdigit())
decimals = st.sidebar.number_input("Percent decimals", min_value=0, max_value=2, value=0)
show_percent_sign = st.sidebar.checkbox("Show % symbol in Percent column", value=True)
preview_n = st.sidebar.number_input("Preview tables", min_value=1, max_value=20, value=3)

uploaded = st.file_uploader("Upload data (.sav, .csv, .xlsx)", type=["sav","csv","xls","xlsx"])

if uploaded:
    try:
        df, meta = read_file(uploaded)
    except Exception as e:
        st.error(f"Unable to read file: {e}")
        st.stop()

    st.success(f"Loaded {df.shape[0]} rows × {df.shape[1]} columns")
    st.dataframe(df.head(8))

    settings = {
        "dk_codes": dk_codes,
        "decimals": decimals,
        "show_percent_sign": show_percent_sign
    }

    with st.spinner("Generating tables..."):
        worksheets = generate_tabulation(df, meta, settings)

    st.success(f"Generated {len(worksheets)} sheets (includes summaries)")
    for i, (sheet, (qtext, df_tab)) in enumerate(worksheets.items()):
        if i >= preview_n: break
        st.subheader(f"Sheet: {sheet}")
        st.markdown(f"**{qtext}**")
        st.dataframe(df_tab.head(20))

    if st.button("Export formatted Excel workbook"):
        with st.spinner("Exporting Excel..."):
            excel_bytes = write_workbook(worksheets)
        st.download_button(
            "Download workbook",
            data=excel_bytes,
            file_name="tabulation_total_tables_v3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload your dataset to start.")
