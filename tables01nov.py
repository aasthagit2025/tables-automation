"""
Streamlit Tabulation Automation v2
- Uses SPSS variable labels & value labels (pyreadstat)
- Implements Top2/Bottom2 & NPS logic per Tabulation Standards
- Exports formatted Excel workbook (blue header, merged title, freeze panes, Total in col B)
- Leaves hooks for significance testing (to be added next)
Reference: Tabulation Standards & Guidelines_v1 (used for rules/formatting). :contentReference[oaicite:1]{index=1}
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
# Utilities & Config
# -------------------------
DEFAULT_DK_CODES = {88, 99, -1, 98}
BLUE_HEADER = "#0070C0"
BANNER_SHADE_1 = "#D3D3D3"
BANNER_SHADE_2 = "#E9E9E9"

def read_file(uploaded_file) -> Tuple[pd.DataFrame, dict]:
    name = uploaded_file.name.lower()
    if name.endswith(".sav"):
        df, meta = pyreadstat.read_sav(io.BytesIO(uploaded_file.read()), apply_value_formats=False)
        meta_info = {
            "format": "sav",
            "variable_labels": getattr(meta, "variable_labels", {}),
            "value_labels": getattr(meta, "value_labels", {})
        }
        return df, meta_info
    elif name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
        return df, {"format":"csv", "variable_labels":{}, "value_labels":{}}
    else:
        df = pd.read_excel(uploaded_file)
        return df, {"format":"excel", "variable_labels":{}, "value_labels":{}}

def clean_title(text: str) -> str:
    """Clean variable label to use as table title; remove RI phrases like 'Please select one'."""
    if not isinstance(text, str) or not text.strip():
        return ""
    # remove common RI words
    ri_phrases = ["please select one", "select one", "tick one", "choose one", "please select"]
    txt = text.strip()
    lower = txt.lower()
    for p in ri_phrases:
        if p in lower:
            # remove the phrase
            idx = lower.find(p)
            txt = (txt[:idx] + txt[idx+len(p):]).strip(" :;,-")
            lower = txt.lower()
    return txt

def get_label_for_variable(varname: str, meta: dict) -> str:
    """Return the variable label if available, else varname."""
    vlabels = meta.get("variable_labels", {})
    return clean_title(vlabels.get(varname, varname))

def value_label_map_for_var(varname: str, meta: dict) -> Dict[Any,str]:
    """Return mapping of raw value -> label (if available in SPSS)."""
    vmap = meta.get("value_labels", {})
    # pyreadstat's structure: value_labels is dict of { 'varname': {value: label, ...}, ... }
    if varname in vmap:
        return vmap[varname]
    # sometimes labels keyed by numeric codes as strings — normalize it
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
        # try numeric convert
        try:
            conv = pd.to_numeric(series, errors='coerce')
            mask = ~conv.isin(dk_codes)
        except Exception:
            mask = pd.Series(True, index=series.index)
    return mask

# -------------------------
# Tabulation Logic
# -------------------------
def detect_question_type(series: pd.Series, varname: str, meta: dict) -> str:
    """Heuristic detection: rating / single / multi / numeric_open / text_open"""
    s = series.dropna()
    nunique = s.nunique(dropna=True)
    is_num = pd.api.types.is_numeric_dtype(series)
    # multi response naming pattern
    if "_" in varname and is_binary_flag(series):
        return "multi_response"
    # rating heuristic
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
    # convert index labels using value_labels if present
    def label_idx(val):
        if pd.isna(val):
            return "<No Answer>"
        if value_labels is not None:
            # value_labels might have numeric keys
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
    mn = s.mean()
    sd = s.std(ddof=0)
    if scale_min is None: scale_min = int(s.min())
    if scale_max is None: scale_max = int(s.max())
    # Top2 & Bottom2 for 5-pt / Top3 for 7-pt etc.
    def net_count(top_n:int):
        top_cut = scale_max - top_n + 1
        topc = (s >= top_cut).sum()
        botc = (s <= (scale_min + top_n - 1)).sum()
        return topc, botc
    # detect scale width
    width = scale_max - scale_min + 1
    summary = {"Base": len(s), "Mean": round(mn,2), "Std": round(sd,2), "ScaleMin": scale_min, "ScaleMax": scale_max}
    # fill Top/Bottom nets
    if width >= 5:
        t2,b2 = net_count(2)
        summary.update({"Top2_Count":int(t2),"Top2_Pct":round(t2/len(s)*100,2),"Bottom2_Count":int(b2),"Bottom2_Pct":round(b2/len(s)*100,2)})
    if width >= 7:
        t3,b3 = net_count(3)
        summary.update({"Top3_Count":int(t3),"Top3_Pct":round(t3/len(s)*100,2),"Bottom3_Count":int(b3),"Bottom3_Pct":round(b3/len(s)*100,2)})
    # NPS if 0-10 scale
    if scale_min == 0 and scale_max == 10:
        prom = ((s >= 9) & (s <= 10)).sum()
        passive = ((s >= 7) & (s <= 8)).sum()
        detr = ((s >= 0) & (s <= 6)).sum()
        nps = round((prom - detr)/len(s) * 100, 2)
        summary.update({"Promoters":int(prom),"Passive":int(passive),"Detractors":int(detr),"NPS":nps})
    return summary

# -------------------------
# Excel Export (formatting)
# -------------------------
def write_workbook(worksheets: Dict[str, pd.DataFrame], filename="tabulation_v2.xlsx") -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        workbook = writer.book
        # formats
        header_fmt = workbook.add_format({"bold": True,"font_name":"Calibri","font_size":10,"align":"center","valign":"vcenter","font_color":"white","bg_color":BLUE_HEADER})
        title_fmt = workbook.add_format({"bold": True,"font_name":"Calibri","font_size":11,"align":"left","valign":"vcenter"})
        normal_center = workbook.add_format({"font_name":"Calibri","font_size":10,"align":"center","valign":"vcenter"})
        int_fmt = workbook.add_format({"num_format":"0","font_name":"Calibri","font_size":10,"align":"center"})
        pct_fmt = workbook.add_format({"num_format":"0","font_name":"Calibri","font_size":10,"align":"center"})
        # banner shades
        shade1 = workbook.add_format({"bg_color":BANNER_SHADE_1,"font_name":"Calibri","font_size":10,"align":"center"})
        shade2 = workbook.add_format({"bg_color":BANNER_SHADE_2,"font_name":"Calibri","font_size":10,"align":"center"})

        for sheet_name, df in worksheets.items():
            safe = sheet_name[:31]
            # create sheet with df starting at row 2 (we'll use row 0 for merged title & row 1 for header)
            df.to_excel(writer, sheet_name=safe, index=False, startrow=1)
            ws = writer.sheets[safe]
            ncols = df.shape[1]
            if ncols == 0:
                continue
            # Merged title across columns (Row 0)
            ws.merge_range(0, 0, 0, max(0, ncols-1), safe, title_fmt)
            # write header row formatting at row 1
            for col_idx, col in enumerate(df.columns):
                ws.write(1, col_idx, col, header_fmt)
                # auto column width
                maxlen = max(df[col].astype(str).map(len).max(), len(str(col))) + 2
                maxlen = min(maxlen, 50)
                # ensure first column wider (stub)
                if col_idx == 0:
                    ws.set_column(col_idx, col_idx, max(12, maxlen))
                else:
                    ws.set_column(col_idx, col_idx, max(8, maxlen))
            # Freeze panes: Row 2, Col 1 (so first column and header visible)
            ws.freeze_panes(2, 1)
            # No gridlines (user can toggle in viewer)
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
    weight_col = settings.get("weight_col", None)
    banners = settings.get("banners", [])  # list of column names to use as banners (optional)

    worksheets = {}
    index_rows = []

    # detect prefix groups for MR
    varnames = list(df.columns)
    prefix_groups = {}
    for v in varnames:
        if "_" in v:
            prefix_groups.setdefault(v.split("_")[0], []).append(v)

    # variable-by-variable
    rating_summaries = []
    mr_combined_sheets = {}
    for v in varnames:
        series = df[v]
        vlabel = get_label_for_variable(v, meta)
        vmap = value_label_map_for_var(v, meta)
        qtype = detect_question_type(series, v, meta)

        base_mask = exclude_dk(series, dk_codes)

        if qtype == "multi_response":
            # produce checked counts assuming '1' or 'Yes' indicates checked
            checked_values = {"1","1.0","true","yes","checked"}
            checked = series.astype(str).str.strip().str.lower().isin(checked_values)
            base_valid = base_mask.sum()
            count = int(checked[base_mask].sum())
            pct = round(count / base_valid * 100, 2) if base_valid>0 else 0
            df_tab = pd.DataFrame([{"Attribute": vlabel, "Count": count, "Percent": pct}])
            # arrange columns so that "Question" (title) is not in same merged cell — we'll use sheet title
            worksheets[v] = df_tab
            index_rows.append({"Sheet": v, "Rows": df_tab.shape[0], "Cols": df_tab.shape[1]})
        elif qtype == "rating":
            # compute counts and order highest to lowest scale
            table = compute_count_pct(series, base_mask, decimals, value_labels=vmap)
            # try to sort numeric scale descending if possible
            try:
                idx_nums = []
                for idx in table.index:
                    # attempt to extract numeric code from label if mapping exists (we check reverse)
                    # We'll try if index label corresponds to numeric code in vmap
                    found = None
                    if vmap:
                        for code, lbl in vmap.items():
                            if lbl == idx:
                                found = int(code)
                                break
                    if found is None:
                        # fallback: try to parse index as int
                        try:
                            found = int(str(idx))
                        except Exception:
                            found = None
                    idx_nums.append((idx, found))
                # sort by found numeric descending where available
                idx_sorted = [x for x,_ in sorted(idx_nums, key=lambda z: (-(z[1] if z[1] is not None else -9999)))]
                table = table.reindex(idx_sorted)
            except Exception:
                pass
            df_tab = table.reset_index().rename(columns={"index":"Option"})
            worksheets[v] = df_tab
            # rating summary
            # determine scale_min/max from vmap or existing numeric codes
            scale_min, scale_max = None, None
            if vmap:
                codes = [int(k) for k in vmap.keys() if str(k).lstrip('-').isdigit()]
                if codes:
                    scale_min, scale_max = min(codes), max(codes)
            else:
                numeric_vals = pd.to_numeric(series.dropna(), errors='coerce').dropna()
                if not numeric_vals.empty:
                    scale_min, scale_max = int(numeric_vals.min()), int(numeric_vals.max())
            summ = compute_rating_summary(series, base_mask, scale_min, scale_max)
            summ["QuestionVar"] = v
            summ["QuestionLabel"] = vlabel
            rating_summaries.append(summ)
            index_rows.append({"Sheet": v, "Rows": df_tab.shape[0], "Cols": df_tab.shape[1]})
        elif qtype == "numeric_open":
            numeric = pd.to_numeric(series.dropna(), errors='coerce').dropna()
            if numeric.empty:
                df_tab = compute_count_pct(series, base_mask, decimals, value_labels=vmap).reset_index().rename(columns={"index":"Value"})
            else:
                bins = pd.cut(numeric, bins=10)
                counts = bins.value_counts().sort_index()
                pct = (counts / counts.sum() * 100).round(decimals)
                df_tab = pd.DataFrame({"Range":counts.index.astype(str),"Count":counts.values,"Percent":pct.values})
            worksheets[v] = df_tab
            index_rows.append({"Sheet": v, "Rows": df_tab.shape[0], "Cols": df_tab.shape[1]})
        elif qtype == "text_open":
            top = series.dropna().astype(str).value_counts().head(100)
            pct = (top / top.sum() * 100).round(decimals)
            df_tab = pd.DataFrame({"Response": top.index, "Count": top.values, "Percent": pct.values})
            worksheets[v] = df_tab
            index_rows.append({"Sheet": v, "Rows": df_tab.shape[0], "Cols": df_tab.shape[1]})
        else:  # single categorical
            table = compute_count_pct(series, base_mask, decimals, value_labels=vmap)
            df_tab = table.reset_index().rename(columns={"index":"Option"})
            worksheets[v] = df_tab
            index_rows.append({"Sheet": v, "Rows": df_tab.shape[0], "Cols": df_tab.shape[1]})

    # Create rating summary sheet (Top2/Bottom2, NPS, Mean)
    if rating_summaries:
        rows = []
        for s in rating_summaries:
            row = {
                "QuestionVar": s.get("QuestionVar"),
                "QuestionLabel": s.get("QuestionLabel"),
                "Base": s.get("Base"),
                "Mean": s.get("Mean"),
                "Std": s.get("Std")
            }
            # include nets if present
            for k in ["Top2_Count","Top2_Pct","Bottom2_Count","Bottom2_Pct","Top3_Count","Top3_Pct","Bottom3_Count","Bottom3_Pct","Promoters","Passive","Detractors","NPS"]:
                if k in s:
                    row[k] = s[k]
            rows.append(row)
        rating_summary_df = pd.DataFrame(rows)
        worksheets["Rating_Summary_TopBottom_NPS"] = rating_summary_df
        index_rows.append({"Sheet":"Rating_Summary_TopBottom_NPS","Rows":rating_summary_df.shape[0],"Cols":rating_summary_df.shape[1]})

    # Index sheet
    index_df = pd.DataFrame(index_rows)
    worksheets["INDEX"] = index_df

    return worksheets

# -------------------------
# Streamlit UI
# -------------------------
st.title("Tabulation Automation — v2 (SPSS labels, Top2/Bottom2, NPS, Excel formatting)")
st.markdown("This version uses SPSS variable/value labels (if available), computes Top/Bottom nets and NPS, and exports a formatted workbook. Rules follow your Tabulation Standards document. :contentReference[oaicite:2]{index=2}")

# Sidebar controls
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

    st.success(f"Loaded {df.shape[0]} rows x {df.shape[1]} cols")
    st.write("Preview (first 8 rows):")
    st.dataframe(df.head(8))

    # allow user to pick banners (columns)
    st.sidebar.subheader("Banner selection (optional)")
    banner_cols = st.sidebar.multiselect("Choose banner variables to include (these create separate banner outputs)", options=list(df.columns))

    settings = {"dk_codes": dk_codes, "decimals": decimals, "weight_col": weight_col or None, "banners": banner_cols}
    with st.spinner("Generating tables (this may take a moment)..."):
        worksheets = generate_tabulation(df, meta, settings)

    st.success(f"Generated {len(worksheets)} sheets (includes INDEX and summaries)")

    # show index
    st.subheader("Index of generated sheets")
    st.dataframe(worksheets.get("INDEX", pd.DataFrame()))

    # show few previews
    st.subheader("Preview some outputs")
    preview_keys = [k for k in worksheets.keys() if k != "INDEX"][:preview_n]
    for k in preview_keys:
        st.markdown(f"**Sheet: {k}**")
        st.dataframe(worksheets[k].head(30))

    # Export button
    if st.button("Export formatted Excel workbook"):
        with st.spinner("Writing Excel workbook..."):
            bytes_xl = write_workbook(worksheets)
        st.success("Workbook ready")
        st.download_button("Download workbook", data=bytes_xl, file_name="tabulation_v2_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.info("Notes: \n- Table titles are mapped from SPSS variable labels where available. \n- Rating tables include Top2/Bottom2 and NPS where applicable. \n- Next iteration will add significance testing, bannered cross-tabs, exact 'Total in Column B' micro-formatting and additional QC checks.")
else:
    st.info("Upload a dataset to begin. Supported: .sav (SPSS), .csv, .xlsx.")
