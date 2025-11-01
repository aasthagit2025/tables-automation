"""
Streamlit Tabulation Automation — v2.2 (Value Label Fix)
- Reads SPSS (.sav), Excel, or CSV
- Question from SPSS variable label
- Stub from SPSS value labels (text, not numeric)
- Outputs Total-only tables (Count & Percent)
- Formatted Excel (blue header, merged title)
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import pyreadstat
import tempfile
import xlsxwriter
from typing import Dict, Tuple

st.set_page_config(page_title="Tabulation Automation v2.2", layout="wide")

# -------------------------
# Config
# -------------------------
DEFAULT_DK_CODES = {88, 99, -1, 98}
BLUE_HEADER = "#0070C0"

# -------------------------
# File Reader (universal fix)
# -------------------------
def read_file(uploaded_file) -> Tuple[pd.DataFrame, dict]:
    """
    Reads uploaded dataset (.sav, .csv, .xlsx)
    Handles SPSS via temp file (safe for Streamlit Cloud).
    """
    name = uploaded_file.name.lower()

    if name.endswith(".sav"):
        # Save to temp path
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

# -------------------------
# Helper Functions
# -------------------------
def clean_title(text: str) -> str:
    """Clean RI text like 'Please select one' from variable labels."""
    if not isinstance(text, str) or not text.strip():
        return ""
    ri_phrases = ["please select one", "select one", "tick one", "choose one", "please select"]
    txt = text.strip()
    for p in ri_phrases:
        txt = txt.replace(p, "", 1) if p in txt.lower() else txt
    return txt.strip(" :;,-")

def get_label_for_variable(varname: str, meta: dict) -> str:
    vlabels = meta.get("variable_labels", {})
    return clean_title(vlabels.get(varname, varname))

def value_label_map_for_var(varname: str, meta: dict) -> Dict:
    """Return dictionary of {code: label} for given SPSS variable."""
    all_maps = meta.get("value_labels", {})
    if varname in all_maps:
        return all_maps[varname]
    for k, mm in all_maps.items():
        if k.lower() == varname.lower():
            return mm
    return {}

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

def compute_count_pct(series: pd.Series, base_mask: pd.Series, decimals:int=0, value_labels:dict=None) -> pd.DataFrame:
    """Return Count & % table for one variable, mapping value labels."""
    s = series[base_mask]
    counts = s.value_counts(dropna=False, sort=False)
    total = counts.sum()
    pct = (counts / total * 100).round(decimals)
    df = pd.DataFrame({"Count": counts.astype(int), "Percent": pct})

    # Replace numeric codes with text labels from SPSS
    def label_idx(val):
        if pd.isna(val):
            return "<No Answer>"
        if value_labels:
            try:
                v_int = int(val)
                return value_labels.get(v_int, val)
            except Exception:
                return value_labels.get(str(val), val)
        return val

    df.index = [label_idx(i) for i in df.index]
    return df

# -------------------------
# Table Generator
# -------------------------
def generate_tabulation(df: pd.DataFrame, meta: dict, settings: dict) -> Dict[str, pd.DataFrame]:
    dk_codes = set(settings.get("dk_codes", DEFAULT_DK_CODES))
    decimals = settings.get("decimals", 0)
    worksheets = {}

    for v in df.columns:
        series = df[v]
        base_mask = exclude_dk(series, dk_codes)
        vmap = value_label_map_for_var(v, meta)
        qtext = get_label_for_variable(v, meta)

        table = compute_count_pct(series, base_mask, decimals, vmap)
        df_tab = table.reset_index().rename(columns={"index": "Stub"})
        df_tab.insert(0, "Question", qtext)
        worksheets[v] = df_tab

    # Index sheet
    index_df = pd.DataFrame(
        [{"Sheet": k, "Rows": v.shape[0], "Cols": v.shape[1]} for k, v in worksheets.items()]
    )
    worksheets["INDEX"] = index_df
    return worksheets

# -------------------------
# Excel Export (Wincross-style)
# -------------------------
def write_workbook(worksheets: Dict[str, pd.DataFrame]) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        workbook = writer.book
        title_fmt = workbook.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 11,
            "align": "left", "valign": "vcenter"
        })
        header_fmt = workbook.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 10,
            "align": "center", "valign": "vcenter",
            "font_color": "white", "bg_color": BLUE_HEADER
        })
        cell_fmt = workbook.add_format({
            "font_name": "Calibri", "font_size": 10,
            "align": "center", "valign": "vcenter"
        })

        for sheet_name, df in worksheets.items():
            safe = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe, index=False, startrow=1)
            ws = writer.sheets[safe]

            # Merge header title
            ncols = len(df.columns)
            ws.merge_range(0, 0, 0, ncols - 1, sheet_name, title_fmt)

            # Apply header format
            for i, col in enumerate(df.columns):
                ws.write(1, i, col, header_fmt)
                ws.set_column(i, i, 22, cell_fmt)

            ws.freeze_panes(2, 1)
            ws.hide_gridlines(2)

    out.seek(0)
    return out.read()

# -------------------------
# Streamlit UI
# -------------------------
st.title("Tabulation Automation — v2.2 (Total-only Wincross Style)")
st.markdown("Uploads SPSS / CSV / Excel and produces formatted Total tables using variable & value labels.")

# Sidebar
st.sidebar.header("Settings")
dk_text = st.sidebar.text_input("DK/Ref codes (comma separated)", value="88,99,-1,98")
dk_codes = set(int(x.strip()) for x in dk_text.split(",") if x.strip().lstrip('-').isdigit())
decimals = st.sidebar.number_input("Percent decimals", min_value=0, max_value=2, value=0)
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

    settings = {"dk_codes": dk_codes, "decimals": decimals}
    with st.spinner("Generating Total tables..."):
        worksheets = generate_tabulation(df, meta, settings)

    st.success(f"Generated {len(worksheets)} sheets (including INDEX)")

    st.subheader("Index of generated sheets")
    st.dataframe(worksheets.get("INDEX", pd.DataFrame()))

    st.subheader("Preview tables")
    for k in list(worksheets.keys())[:preview_n]:
        st.markdown(f"**Sheet: {k}**")
        st.dataframe(worksheets[k].head(20))

    if st.button("Export formatted Excel workbook"):
        with st.spinner("Exporting..."):
            bytes_xl = write_workbook(worksheets)
        st.download_button(
            "Download workbook",
            data=bytes_xl,
            file_name="tabulation_total_tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload a .sav, .csv, or .xlsx file to start.")
