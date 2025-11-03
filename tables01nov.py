"""
Tabulation Automation v4.3b — counts on one row, percents on the next; manual rating selection + nets per question.
Save as tables01nov_v4_3b.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import pyreadstat
import tempfile
from typing import Dict, Tuple, List
import xlsxwriter

st.set_page_config(page_title="Tabulation Automation v4.3b", layout="wide")

# Config
DEFAULT_DK_CODES = {88, 99, -1, 98}
BLUE_HEADER = "#0070C0"
EXCLUDE_VARS = {"record", "uuid", "source", "date"}

# ------- File reader -------
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

# ------- Helpers -------
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

# ------- Rating detection (candidate) -------
def is_rating_candidate(raw_series: pd.Series) -> bool:
    """Simple candidate detection: numeric & 3-11 unique values"""
    if not pd.api.types.is_numeric_dtype(raw_series):
        return False
    nunique = raw_series.dropna().nunique()
    return 3 <= nunique <= 11

# ------- Rating metrics -------
def compute_rating_metrics(raw_series: pd.Series, base_mask: pd.Series) -> Dict:
    s = pd.to_numeric(raw_series[base_mask], errors="coerce").dropna()
    if s.empty:
        return {}
    mn = float(s.mean()); sd = float(s.std(ddof=0))
    scale_min = int(s.min()); scale_max = int(s.max())
    width = scale_max - scale_min + 1
    def pct_of(cond): return round(cond.sum() / len(s) * 100, 1)
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

# ------- Compute Total Count & Percent (for formatted df) -------
def compute_total_count_pct(formatted_series: pd.Series, base_mask: pd.Series, decimals:int=1) -> Tuple[pd.Series, pd.Series]:
    """Return (counts_series, percent_series) aligned by index order (unique categories in appearance order)."""
    s = formatted_series[base_mask]
    counts = s.value_counts(dropna=False, sort=False)
    if counts.sum() == 0:
        return pd.Series([], dtype=object), pd.Series([], dtype=float)
    pct = (counts / counts.sum() * 100).round(decimals)
    return counts, pct

# ------- Build question table with double-row format -------
def build_question_table_double_row(qvar: str,
                                    df_formatted: pd.DataFrame,
                                    df_raw: pd.DataFrame,
                                    meta: dict,
                                    settings: dict,
                                    banner_vars: List[str],
                                    nets_selected: List[str]) -> pd.DataFrame:
    dk_codes = set(settings.get("dk_codes", DEFAULT_DK_CODES))
    decimals = int(settings.get("decimals", 1))
    show_percent_sign = settings.get("show_percent_sign", True)

    raw_series = df_raw[qvar] if qvar in df_raw else df_formatted[qvar]
    base_mask = exclude_dk_mask(raw_series, dk_codes)

    counts, pct = compute_total_count_pct(df_formatted[qvar], base_mask, decimals)
    if counts.empty:
        return pd.DataFrame()

    # Use order of counts index as stub order
    stubs = list(counts.index)

    # Build DataFrame rows: for each stub create two dicts (count-row and percent-row)
    rows = []
    # Header columns we'll create later; for now build a dict keyed by column name
    # Start with Total Count and Total % columns
    for stub in stubs:
        # count row
        row_count = {"Stub": stub, "Total Count": int(counts.get(stub, 0))}
        # percent row: Stub blank
        row_pct = {"Stub": ""}
        val_pct = pct.get(stub, 0.0)
        row_pct["Total %"] = f"{val_pct}%" if show_percent_sign else val_pct
        # initialize banner columns later (fill zeros) - use placeholders
        rows.append(row_count)
        rows.append(row_pct)

    # Convert to DataFrame to ensure consistent columns order, then add banner columns
    df_out = pd.DataFrame(rows)

    # For each banner and its categories, compute counts + % (percent relative to banner category base)
    for banner in banner_vars:
        if banner not in df_formatted.columns:
            continue
        cats = list(pd.unique(df_formatted[banner].dropna()))
        for cat in cats:
            # compute counts aligned to stubs
            bmask = (df_formatted[banner] == cat)
            mask_combined = base_mask & bmask
            counts_b = df_formatted[qvar][mask_combined].value_counts(dropna=False, sort=False)
            # compute percent within banner category base
            denom = counts_b.sum()
            if denom == 0:
                pct_b = pd.Series({s:0 for s in stubs})
            else:
                pct_b = (counts_b / denom * 100).reindex(stubs).fillna(0).round(decimals)
            # Now insert two columns: banner_count_col and banner_pct_col
            count_col = f"{banner} - {cat} Count"
            pct_col = f"{banner} - {cat} %"
            # fill column values for each stub-row pair
            count_vals = []
            pct_vals = []
            for stub in stubs:
                c = int(counts_b.get(stub, 0)) if stub in counts_b.index else int(0)
                count_vals.append(c)
                # percent goes on the percent-row (so second of the pair)
                pct_display = f"{pct_b.get(stub,0)}%" if show_percent_sign else pct_b.get(stub,0)
                # For count-row append value, for percent-row append percent
                count_vals.append("")  # placeholder for the percent row aligned under counts column
                pct_vals.append("")    # for count row (we'll place percent separately)
                pct_vals.append(pct_display)  # this will be in percent-row
            # Now interleave properly: we must build a column length equal to df_out rows
            # We'll create the final column by iterating stubs and building pair [count, percent]
            col_values = []
            for stub in stubs:
                col_values.append(int(counts_b.get(stub, 0)))
                pct_display = f"{pct_b.get(stub,0)}%" if show_percent_sign else pct_b.get(stub,0)
                col_values.append(pct_display)
            df_out[count_col] = col_values
            df_out[pct_col] = col_values.copy()  # percent same position; but we will override count places below
            # The above temporarily sets both columns to same; overwrite count column to have counts on odd rows
            # Actually simpler: set count column with counts on count rows and empty on percent rows; pct column vice versa
            count_col_vals = []
            pct_col_vals = []
            for stub in stubs:
                count_col_vals.append(int(counts_b.get(stub, 0)))
                count_col_vals.append("")  # percent row
                pct_display = f"{pct_b.get(stub,0)}%" if show_percent_sign else pct_b.get(stub,0)
                pct_col_vals.append("")  # count row
                pct_col_vals.append(pct_display)
            df_out[count_col] = count_col_vals
            df_out[pct_col] = pct_col_vals

    # Ensure Total % column exists on percent rows (we created only for stubs earlier)
    # For alignment: Total Count already exists on count rows; Total % exists on percent rows
    # Add any missing banner columns (if no banner chosen nothing else is needed)

    # Reorder columns: Stub, Total Count, Total %, then banner pairs in selected order
    cols = ["Stub", "Total Count", "Total %"]
    # find banner columns added
    banner_cols = [c for c in df_out.columns if c not in cols]
    cols += banner_cols
    df_out = df_out.reindex(columns=cols)

    # Finally append nets rows (if selected and qvar is rating and nets selected)
    if nets_selected:
        # ensure it's a rating variable (we let user mark manually; here compute metrics anyway)
        if qvar in df_raw and is_rating_candidate(df_raw[qvar]):
            metrics = compute_rating_metrics(df_raw[qvar], exclude_dk_mask(df_raw[qvar], set(settings.get("dk_codes", DEFAULT_DK_CODES))))
            extra_rows = []
            for metric in nets_selected:
                lbl = metric
                val = ""
                if metric == "Top2":
                    val = metrics.get("Top2%", "")
                elif metric == "Top3":
                    val = metrics.get("Top3%", "")
                elif metric == "Bottom2":
                    val = metrics.get("Bottom2%", "")
                elif metric == "Bottom3":
                    val = metrics.get("Bottom3%", "")
                elif metric == "Mean":
                    val = metrics.get("Mean", "")
                elif metric == "SD":
                    val = metrics.get("SD", "")
                elif metric == "NPS":
                    val = metrics.get("NPS", "")
                # make a row that matches df_out columns, put val in 'Total %' column for nets (consistent)
                row = {c: "" for c in df_out.columns}
                row["Stub"] = lbl
                row["Total %"] = f"{val}%" if isinstance(val, (int,float)) and "Top" in metric else (val if val!= "" else "")
                extra_rows.append(row)
            if extra_rows:
                df_out = pd.concat([df_out, pd.DataFrame(extra_rows, columns=df_out.columns)], ignore_index=True)

    return df_out

# ------- Generate worksheets -------
def generate_all_worksheets(df_formatted: pd.DataFrame, meta: dict, settings: dict,
                            banner_vars: List[str], nets_config: Dict[str, List[str]]) -> Dict[str, Tuple[str, pd.DataFrame]]:
    df_raw = meta.get("raw_df", df_formatted.copy())
    worksheets = {}
    rating_summary_rows = []

    vars_to_process = [v for v in df_formatted.columns if v.lower() not in EXCLUDE_VARS]

    for v in vars_to_process:
        qlabel = get_label_for_variable(v, meta)
        nets_selected = nets_config.get(v, [])
        table = build_question_table_double_row(v, df_formatted, df_raw, meta, settings, banner_vars, nets_selected)
        worksheets[v] = (qlabel, table)

        # rating summaries if requested (or candidate)
        if v in df_raw and is_rating_candidate(df_raw[v]):
            base_mask = exclude_dk_mask(df_raw[v], set(settings.get("dk_codes", DEFAULT_DK_CODES)))
            metrics = compute_rating_metrics(df_raw[v], base_mask)
            row = {"Question": qlabel}; row.update(metrics)
            rating_summary_rows.append(row)

    if rating_summary_rows:
        rs_df = pd.DataFrame(rating_summary_rows).fillna("")
        means_df = rs_df.reindex(columns=["Question", "Base", "Mean", "SD"]).fillna("")
        tb_df = rs_df.reindex(columns=["Question", "Top2%", "Bottom2%", "Top3%", "Bottom3%", "NPS"]).fillna("")
        worksheets["Means_Summary"] = ("Means Summary", means_df)
        worksheets["TopBottom_Summary"] = ("Top/Bottom Summary", tb_df)

    return worksheets

# ------- Excel writer (keeps merged title, blue header) -------
def write_workbook(worksheets: Dict[str, Tuple[str, pd.DataFrame]]) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book
        title_fmt = wb.add_format({"bold": True, "font_name":"Calibri", "font_size":11, "align":"left", "valign":"vcenter"})
        header_fmt = wb.add_format({"bold": True, "font_name":"Calibri", "font_size":10, "align":"center", "valign":"vcenter", "font_color":"white", "bg_color":BLUE_HEADER})
        left_fmt = wb.add_format({"font_name":"Calibri", "font_size":10, "align":"left"})
        center_fmt = wb.add_format({"font_name":"Calibri", "font_size":10, "align":"center"})

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
                if i == 0:
                    ws.set_column(i, i, 30, left_fmt)
                else:
                    ws.set_column(i, i, 14, center_fmt)
            ws.freeze_panes(2, 1)
            ws.hide_gridlines(2)
    out.seek(0)
    return out.read()

# ------- Streamlit UI -------
st.title("Tabulation Automation — v4.3b (Counts row + Percent row; manual rating selection)")
st.markdown("Upload dataset (.sav/.csv/.xlsx). Choose banners and optionally mark rating variables and select nets/means per question.")

# Sidebar controls
st.sidebar.header("Settings")
dk_text = st.sidebar.text_input("DK/Ref codes (comma separated)", value="88,99,-1,98")
dk_codes = set(int(x.strip()) for x in dk_text.split(",") if x.strip().lstrip('-').isdigit())
decimals = st.sidebar.number_input("Percent decimals", min_value=0, max_value=2, value=1)
show_percent_sign = st.sidebar.checkbox("Show % sign (in percent row)", value=True)
preview_n = st.sidebar.number_input("Number of tables to preview", min_value=1, max_value=50, value=5)

uploaded = st.file_uploader("Upload data (.sav, .csv, .xlsx)", type=["sav","csv","xls","xlsx"])

if uploaded:
    try:
        df_formatted, meta = read_file(uploaded)
    except Exception as e:
        st.error(f"Unable to read file: {e}")
        st.stop()

    # Banner selector
    banner_candidates = [c for c in df_formatted.columns if c.lower() not in EXCLUDE_VARS]
    st.sidebar.markdown("### Banner variables (simple)")
    banner_vars = st.sidebar.multiselect("Select banner variable(s)", options=banner_candidates)

    st.sidebar.markdown("---")
    st.sidebar.markdown("### Rating variables & nets (manual override)")
    st.sidebar.markdown("Auto-detected numeric candidates are shown; tick those that ARE rating scales and choose metrics.")
    # detect numeric candidates
    df_raw = meta.get("raw_df", df_formatted.copy())
    rating_candidates = [v for v in df_formatted.columns if v in df_raw and is_rating_candidate(df_raw[v]) and v.lower() not in EXCLUDE_VARS]
    nets_config: Dict[str, List[str]] = {}
    if rating_candidates:
        for v in rating_candidates:
            qlabel = get_label_for_variable(v, meta)
            with st.sidebar.expander(f"{qlabel} ({v})", expanded=False):
                is_rating_checked = st.checkbox("Treat as rating question (enable nets)", key=f"chk_{v}")
                if is_rating_checked:
                    opts = st.multiselect("Select metrics to compute for this question",
                                           options=["Top2", "Top3", "Bottom2", "Bottom3", "Mean", "SD", "NPS"],
                                           key=f"metrics_{v}")
                    if opts:
                        nets_config[v] = opts
    else:
        st.sidebar.markdown("No numeric candidates detected.")

    # loaded preview
    st.success(f"Loaded {df_formatted.shape[0]} rows × {df_formatted.shape[1]} columns")
    st.dataframe(df_formatted.head(8))

    # generate
    settings = {"dk_codes": dk_codes, "decimals": decimals, "show_percent_sign": show_percent_sign}
    with st.spinner("Generating tables..."):
        worksheets = generate_all_worksheets(df_formatted, meta, settings, banner_vars, nets_config)

    st.success(f"Generated {len(worksheets)} sheets (preview below)")

    # preview
    st.subheader("Preview generated tables")
    shown = 0
    for sheet, (qtext, df_tab) in worksheets.items():
        if shown >= preview_n: break
        st.markdown(f"### {shown+1}. {qtext} — {sheet}")
        if df_tab is None or df_tab.empty:
            st.write("No data.")
        else:
            st.dataframe(df_tab.head(80))
        shown += 1

    # export
    if st.button("Export formatted Excel workbook"):
        with st.spinner("Writing Excel..."):
            excel_bytes = write_workbook(worksheets)
        st.download_button("Download Excel", data=excel_bytes,
                           file_name="tabulation_v4_3b_export.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Upload your dataset to start.")
