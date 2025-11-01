"""
Streamlit Tabulation Automation v2 (Final SAV Fix)
- Reads .sav, .csv, .xlsx correctly
- Uses SPSS variable/value labels (via pyreadstat)
- Implements Top2/Bottom2 & NPS logic
- Exports formatted Excel workbook (blue header, merged title, freeze panes)
- Fully compatible with Streamlit Cloud (Python 3.12)
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import pyreadstat
from scipy import stats
import xlsxwriter
from typing import Dict, Any, List, Tuple

st.set_page_config(page_title="Tabulation Automation v2", layout="wide")

# -------------------------
# Configuration
# -------------------------
DEFAULT_DK_CODES = {88, 99, -1, 98}
BLUE_HEADER = "#0070C0"
BANNER_SHADE_1 = "#D3D3D3"
BANNER_SHADE_2 = "#E9E9E9"

# -------------------------
# File reader (FINAL FIX)
# -------------------------
def read_file(uploaded_file) -> Tuple[pd.DataFrame, dict]:
    """
    Reads uploaded dataset (.sav, .csv, .xlsx).
    Final fix for Streamlit Cloud + pyreadstat compatibility.
    """
    name = uploaded_file.name.lower()

    if name.endswith(".sav"):
        # ✅ Save temporarily to disk (works universally)
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
            tmp.write(uploaded_file.getbuffer())
            tmp_path = tmp.name

        # read_sav works reliably with a temp file path
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

# -------------------------
# Helper functions
# -------------------------
def clean_title(text: str) -> str:
    """Cleans question text for table titles (removes RI phrases)."""
    if not isinstance(text, str) or not text.strip():
        return ""
    ri_phrases = ["please select one", "select one", "tick one", "choose one", "please select"]
    txt = text.strip()
    for p in ri_phrases:
        if p in txt.lower():
            txt = txt.lower().replace(p, "")
    return txt.strip(" :;,-")

def get_label_for_variable(varname: str, meta: dict) -> str:
    vlabels = meta.get("variable_labels", {})
    return clean_title(vlabels.get(varname, varname))

def value_label_map_for_var(varname: str, meta: dict) -> Dict[Any, str]:
    vmap = meta.get("value_labels", {})
    if varname in vmap:
        return vmap[varname]
    for k, mm in vmap.items():
        try:
            if varname == k or varname.lower() == k.lower():
                return mm
        except Exception:
            continue
    return {}

def is_binary_flag(series: pd.Series) -> bool:
    uniques = set(series.dropna().astype(str).str.strip().str.lower().unique())
    return uniques <= {"0","1","true","false","checked","unchecked","yes","no","1.0","0.0"}

def exclude_dk(series: pd.Series, dk_codes:set):
    if pd.api.types.is_numeric_dtype(series):
        mask = ~series.isin(dk_codes)
    else:
        try:
            conv = pd.to_numeric(series, errors="coerce")
            mask = ~conv.isin(dk_codes)
        except Exception:
            mask = pd.Series(True, index=series.index)
    return mask

# -------------------------
# Tabulation Logic
# -------------------------
def detect_question_type(series: pd.Series, varname: str, meta: dict) -> str:
    s = series.dropna()
    nunique = s.nunique(dropna=True)
    is_num = pd.api.types.is_numeric_dtype(series)
    if "_" in varname and is_binary_flag(series):
        return "multi_response"
    if is_num and 3 <= nunique <= 11:
        return "rating"
    if is_num and nunique > 11:
        return "numeric_open"
    if not is_num:
        if nunique > 20:
            return "text_open"
        return "single"
    return "single"

def compute_count_pct(series: pd.Series, base_mask: pd.Series, decimals:int=0, value_labels:dict=None) -> pd.DataFrame:
    s = series[base_mask]
    counts = s.value_counts(dropna=False, sort=False)
    total = counts.sum()
    pct = (counts / total * 100).round(decimals)
    df = pd.DataFrame({"Count": counts.astype(int), "Percent": pct})

    def label_idx(val):
        if pd.isna(val):
            return "<No Answer>"
        if value_labels:
            try:
                k = int(val)
                return value_labels.get(k, val)
            except Exception:
                return value_labels.get(str(val), val) if str(val) in value_labels else val
        return val

    df.index = [label_idx(i) for i in df.index]
    return df

def compute_rating_summary(series: pd.Series, base_mask: pd.Series, scale_min:int=None, scale_max:int=None) -> dict:
    s = pd.to_numeric(series[base_mask], errors='coerce').dropna()
    if s.empty:
        return {"Base":0,"Mean":None,"Std":None}
    mn, sd = s.mean(), s.std(ddof=0)
    if scale_min is None: scale_min = int(s.min())
    if scale_max is None: scale_max = int(s.max())
    width = scale_max - scale_min + 1

    def net_count(top_n:int):
        top_cut = scale_max - top_n + 1
        return (s >= top_cut).sum(), (s <= (scale_min + top_n - 1)).sum()

    summary = {"Base": len(s), "Mean": round(mn,2), "Std": round(sd,2)}

    if width >= 5:
        t2,b2 = net_count(2)
        summary.update({
            "Top2_Count":int(t2),"Top2_Pct":round(t2/len(s)*100,2),
            "Bottom2_Count":int(b2),"Bottom2_Pct":round(b2/len(s)*100,2)
        })
    if width >= 7:
        t3,b3 = net_count(3)
        summary.update({
            "Top3_Count":int(t3),"Top3_Pct":round(t3/len(s)*100,2),
            "Bottom3_Count":int(b3),"Bottom3_Pct":round(b3/len(s)*100,2)
        })
    if scale_min == 0 and scale_max == 10:
        prom = ((s >= 9) & (s <= 10)).sum()
        passive = ((s >= 7) & (s <= 8)).sum()
        detr = ((s >= 0) & (s <= 6)).sum()
        nps = round((prom - detr)/len(s) * 100, 2)
        summary.update({"Promoters":int(prom),"Passive":int(passive),"Detractors":int(detr),"NPS":nps})
    return summary

# -------------------------
# Excel Export
# -------------------------
def write_workbook(worksheets: Dict[str, pd.DataFrame]) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({
            "bold":True,"font_name":"Calibri","font_size":10,
            "align":"center","valign":"vcenter","font_color":"white","bg_color":BLUE_HEADER
        })
        title_fmt = workbook.add_format({"bold":True,"font_name":"Calibri","font_size":11,"align":"left"})

        for sheet, df in worksheets.items():
            safe = sheet[:31]
            df.to_excel(writer, sheet_name=safe, index=False, startrow=1)
            ws = writer.sheets[safe]
            ncols = df.shape[1]
            ws.merge_range(0, 0, 0, max(0, ncols-1), safe, title_fmt)
            for i, col in enumerate(df.columns):
                ws.write(1, i, col, header_fmt)
                ws.set_column(i, i, 15)
            ws.freeze_panes(2, 1)
            ws.hide_gridlines(2)
        writer.save()
        out.seek(0)
        return out.read()

# -------------------------
# Main generator
# -------------------------
def generate_tabulation(df: pd.DataFrame, meta: dict, settings: dict) -> Dict[str, pd.DataFrame]:
    dk_codes = set(settings.get("dk_codes", DEFAULT_DK_CODES))
    decimals = settings.get("decimals", 0)
    worksheets, rating_summaries = {}, []

    for v in df.columns:
        series = df[v]
        vlabel = get_label_for_variable(v, meta)
        vmap = value_label_map_for_var(v, meta)
        qtype = detect_question_type(series, v, meta)
        base_mask = exclude_dk(series, dk_codes)

        if qtype == "rating":
            table = compute_count_pct(series, base_mask, decimals, value_labels=vmap)
            worksheets[v] = table.reset_index().rename(columns={"index":"Option"})
            s = compute_rating_summary(series, base_mask)
            s["QuestionVar"], s["QuestionLabel"] = v, vlabel
            rating_summaries.append(s)
        elif qtype == "numeric_open":
            numeric = pd.to_numeric(series.dropna(), errors='coerce').dropna()
            if numeric.empty:
                df_tab = compute_count_pct(series, base_mask, decimals, vmap).reset_index().rename(columns={"index":"Value"})
            else:
                bins = pd.cut(numeric, bins=10)
                counts = bins.value_counts().sort_index()
                pct = (counts / counts.sum() * 100).round(decimals)
                df_tab = pd.DataFrame({"Range":counts.index.astype(str),"Count":counts.values,"Percent":pct.values})
            worksheets[v] = df_tab
        elif qtype == "text_open":
            top = series.dropna().astype(str).value_counts().head(100)
            pct = (top / top.sum() * 100).round(decimals)
            worksheets[v] = pd.DataFrame({"Response":top.index,"Count":top.values,"Percent":pct.values})
        else:
            table = compute_count_pct(series, base_mask, decimals, value_labels=vmap)
            worksheets[v] = table.reset_index().rename(columns={"index":"Option"})

    if rating_summaries:
        worksheets["Rating_Summary_TopBottom_NPS"] = pd.DataFrame(rating_summaries)

    index_rows = [{"Sheet":k,"Rows":v.shape[0],"Cols":v.shape[1]} for k,v in worksheets.items()]
    worksheets["INDEX"] = pd.DataFrame(index_rows)
    return worksheets

# -------------------------
# Streamlit UI
# -------------------------
st.title("Tabulation Automation — v2 (SPSS labels, Top2/Bottom2, NPS, Excel formatting)")
st.markdown("Upload a dataset (.sav, .csv, .xlsx) to auto-generate formatted tabulations.")

# Sidebar
st.sidebar.header("Settings")
dk_text = st.sidebar.text_input("DK/Ref codes (comma separated)", value="88,99,-1,98")
dk_codes = set(int(x.strip()) for x in dk_text.split(",") if x.strip().lstrip('-').isdigit())
decimals = st.sidebar.number_input("Percent decimals", min_value=0, max_value=2, value=0)
weight_col = st.sidebar.text_input("Weight column (optional)", value="")
sig_level = st.sidebar.selectbox("Significance level (hook - for later)", options=[95,90,99], index=0)
preview_n = st.sidebar.number_input("Preview tables (count)", min_value=1, max_value=20, value=3)

uploaded = st.file_uploader("Upload data (.sav, .csv, .xlsx)", type=["sav","csv","xls","xlsx"])

if uploaded:
    try:
        df, meta = read_file(uploaded)
    except Exception as e:
        st.error(f"Unable to read file: {e}")
        st.stop()

    st.success(f"Loaded {df.shape[0]} rows × {df.shape[1]} columns")
    st.dataframe(df.head(8))

    settings = {"dk_codes": dk_codes, "decimals": decimals, "weight_col": weight_col or None}
    with st.spinner("Generating tables..."):
        worksheets = generate_tabulation(df, meta, settings)

    st.success(f"Generated {len(worksheets)} sheets (includes INDEX and summaries)")

    st.subheader("Index of generated sheets")
    st.dataframe(worksheets.get("INDEX", pd.DataFrame()))

    st.subheader("Preview some outputs")
    for k in list(worksheets.keys())[:preview_n]:
        st.markdown(f"**Sheet: {k}**")
        st.dataframe(worksheets[k].head(20))

    if st.button("Export formatted Excel workbook"):
        with st.spinner("Writing Excel workbook..."):
            bytes_xl = write_workbook(worksheets)
        st.download_button(
            "Download workbook",
            data=bytes_xl,
            file_name="tabulation_v2_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload your .sav, .csv, or .xlsx data file to start.")
