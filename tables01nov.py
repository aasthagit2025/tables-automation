# Full file: tables01nov_v4_4.py
# Implements Wincross matrix layout (first selected banner as row groups) + per-question nets config
# Save/overwrite your app file and restart Streamlit Cloud.

import streamlit as st
import pandas as pd
import numpy as np
import io
import pyreadstat
import tempfile
from typing import Dict, Tuple, List
import xlsxwriter

st.set_page_config(page_title="Tabulation Automation v4.4 (Wincross matrix)", layout="wide")

# ----------------------
# Config
# ----------------------
DEFAULT_DK_CODES = {88, 99, -1, 98}
BLUE_HEADER = "#0070C0"
EXCLUDE_VARS = {"record", "uuid", "source", "date"}

# ----------------------
# File reader
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

# Candidate rating detection (simple numeric 3-11 unique)
def is_rating_candidate(raw_series: pd.Series) -> bool:
    if not pd.api.types.is_numeric_dtype(raw_series):
        return False
    nunique = raw_series.dropna().nunique()
    return 3 <= nunique <= 11

# Rating metrics
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

# ----------------------
# Wincross matrix builder
# - first_banner: used as row groups (Base + each category)
# - other_banners: used as column groups (Count + % per category)
# - percent values are numeric (Excel will keep them as numbers)
# ----------------------
def build_wincross_matrix(qvar: str,
                          df_formatted: pd.DataFrame,
                          df_raw: pd.DataFrame,
                          meta: dict,
                          settings: dict,
                          first_banner: str,
                          other_banners: List[str],
                          nets_selected: List[str]) -> pd.DataFrame:

    dk_codes = set(settings.get("dk_codes", DEFAULT_DK_CODES))
    decimals = int(settings.get("decimals", 1))
    show_percent_sign = settings.get("show_percent_sign", True)

    raw_series = df_raw[qvar] if qvar in df_raw else df_formatted[qvar]
    base_mask_all = exclude_dk_mask(raw_series, dk_codes)

    # prepare row groups: Base row + categories of first_banner
    row_groups = []
    # Base row label
    row_groups.append(("Base: All Respondents", base_mask_all))
    # categories of first banner as rows (if banner exists)
    if first_banner and first_banner in df_formatted.columns:
        cats = list(pd.unique(df_formatted[first_banner].dropna()))
        for cat in cats:
            mask = base_mask_all & (df_formatted[first_banner] == cat)
            row_groups.append((str(cat), mask))

    # Build columns: Total Count, Total % (numeric), then for each other banner -> for each category -> Count / Percent
    columns = ["Row Group", "Total Count", "Total %"]
    # for each other banner create pairs like Banner - <cat> Count and Banner - <cat> %
    banner_columns = []
    for banner in other_banners:
        if banner not in df_formatted.columns:
            continue
        cats = list(pd.unique(df_formatted[banner].dropna()))
        for cat in cats:
            banner_columns.append((banner, cat))
            columns.append(f"{banner} - {cat} Count")
            columns.append(f"{banner} - {cat} %")

    # Build rows (one row per group)
    rows = []
    for group_label, mask in row_groups:
        # totals relative to this row group (Total Count = count of qvar values within mask)
        counts = df_formatted[qvar][mask].value_counts(dropna=False, sort=False)
        # For totals we want the count of each response category - but WinCross view typically shows counts for each response category across column axis.
        # Here your requested matrix shows counts of the attribute categories across columns per row group.
        # To keep format consistent with your previous request, we will show counts for each response category in the same order as value labels,
        # but because the matrix spread expects Count and % per column category, we show aggregate total for the question overall in Total Count/Total %,
        # and banner columns reflect counts of each selected banner category within this row group for the question's response value presence.
        # Implementation below:
        total_count = int(df_formatted[qvar][mask].notna().sum())
        # compute total percent relative to full base? Your WinCross example shows Total % relative to total base (not row base).
        # We'll provide Total % as percent of total base (consistent with earlier behavior).
        # Total base = df_formatted[qvar][base_mask_all].notna().sum()
        total_base = int(df_formatted[qvar][base_mask_all].notna().sum())
        total_pct = round((total_count / total_base) * 100, decimals) if total_base > 0 else 0.0

        row = {"Row Group": group_label, "Total Count": total_count, "Total %": total_pct / 100 if not show_percent_sign else f"{total_pct}%"}
        # For each banner column (banner, cat), compute count of qvar responses where banner==cat within this row group mask
        # But that would be nested and perhaps redundant. The typical WinCross shows for each column (a banner category) the count of the question's response values that fall into that column *for the row group*. For simplicity, we compute:
        for banner, cat in banner_columns:
            # within this row group mask, how many respondents have banner==cat and qvar non-missing?
            mask_cat = mask & (df_formatted[banner] == cat)
            ccount = int(df_formatted[qvar][mask_cat].notna().sum())
            # percent relative to the banner category base (not row!). To get that we need denom = count of qvar non-missing where banner==cat across all rows (i.e., banner category base)
            denom = int(df_formatted[qvar][exclude_dk_mask(df_raw[qvar] if qvar in df_raw else df_formatted[qvar], dk_codes) & (df_formatted[banner] == cat)].notna().sum())
            if denom == 0:
                pct_val = 0.0
            else:
                pct_val = round(ccount / denom * 100, decimals)
            # put numeric percent (so Excel number format can be applied)
            if show_percent_sign:
                pct_display = f"{pct_val}%"
            else:
                pct_display = pct_val / 100.0
            row[f"{banner} - {cat} Count"] = ccount
            row[f"{banner} - {cat} %"] = pct_display
        rows.append(row)

    df_matrix = pd.DataFrame(rows, columns=columns)
    # If nets selected and this qvar is rating candidate, append rows for nets (we place them as extra Row Group rows)
    if nets_selected and qvar in df_raw and is_rating_candidate(df_raw[qvar]):
        metrics = compute_rating_metrics(df_raw[qvar], exclude_dk_mask(df_raw[qvar], dk_codes))
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
            # create a net row with value in Total % (match your example where nets show percent)
            net_row = {c: "" for c in df_matrix.columns}
            net_row["Row Group"] = lbl
            # show value in Total % (as percent if it's percent)
            if isinstance(val, (int, float)) and "Top" in metric or "Bottom" in metric or metric == "NPS":
                net_row["Total %"] = f"{val}%"
            else:
                net_row["Total %"] = val
            df_matrix = pd.concat([df_matrix, pd.DataFrame([net_row])], ignore_index=True)
    return df_matrix

# ----------------------
# Standard (vertical) builder — keep for fallback (we will not paste full here, using prior approach)
# For brevity we reuse prior vertical builder (simplified) if user selects Standard layout.
# ----------------------
def build_vertical_table(qvar: str, df_formatted: pd.DataFrame, df_raw: pd.DataFrame,
                         meta: dict, settings: dict, banner_vars: List[str], nets_selected: List[str]) -> pd.DataFrame:
    # build simple vertical table as before: Stub, Total Count, Total %, then banner pairs columns
    dk_codes = set(settings.get("dk_codes", DEFAULT_DK_CODES))
    decimals = int(settings.get("decimals", 1))
    show_percent_sign = settings.get("show_percent_sign", True)
    raw_series = df_raw[qvar] if qvar in df_raw else df_formatted[qvar]
    base_mask = exclude_dk_mask(raw_series, dk_codes)
    s = df_formatted[qvar][base_mask]
    counts = s.value_counts(dropna=False, sort=False)
    pct = (counts / counts.sum() * 100).round(decimals)
    df = pd.DataFrame({"Stub": counts.index, "Total Count": counts.values, "Total %": pct.values})
    if show_percent_sign:
        df["Total %"] = df["Total %"].astype(str) + "%"
    # banner columns: counts only or count + percent depending on earlier choices — for compatibility produce both
    for banner in banner_vars:
        if banner not in df_formatted.columns:
            continue
        cats = list(pd.unique(df_formatted[banner].dropna()))
        for cat in cats:
            mask = base_mask & (df_formatted[banner] == cat)
            counts_b = df_formatted[qvar][mask].value_counts(dropna=False, sort=False)
            aligned = [int(counts_b.get(stub, 0)) for stub in df["Stub"]]
            df[f"{banner} - {cat} Count"] = aligned
            # percent relative to banner base:
            denom = int(df_formatted[qvar][exclude_dk_mask(df_raw[qvar] if qvar in df_raw else df_formatted[qvar], dk_codes) & (df_formatted[banner] == cat)].notna().sum())
            pct_list = []
            for stub in df["Stub"]:
                c = counts_b.get(stub, 0)
                if denom == 0:
                    pct_list.append(0 if not show_percent_sign else "0%")
                else:
                    pval = round(int(c) / denom * 100, decimals)
                    pct_list.append(pval if not show_percent_sign else f"{pval}%")
            df[f"{banner} - {cat} %"] = pct_list
    # append nets rows if requested and rating
    if nets_selected and qvar in df_raw and is_rating_candidate(df_raw[qvar]):
        metrics = compute_rating_metrics(df_raw[qvar], exclude_dk_mask(df_raw[qvar], dk_codes))
        extra = []
        for m in nets_selected:
            row = {c: "" for c in df.columns}
            row["Stub"] = m
            if m == "Top2":
                row["Total %"] = metrics.get("Top2%", "")
            elif m == "Top3":
                row["Total %"] = metrics.get("Top3%", "")
            elif m == "Bottom2":
                row["Total %"] = metrics.get("Bottom2%", "")
            elif m == "Bottom3":
                row["Total %"] = metrics.get("Bottom3%", "")
            elif m == "Mean":
                row["Total %"] = metrics.get("Mean", "")
            elif m == "SD":
                row["Total %"] = metrics.get("SD", "")
            elif m == "NPS":
                row["Total %"] = metrics.get("NPS", "")
            extra.append(row)
        if extra:
            df = pd.concat([df, pd.DataFrame(extra, columns=df.columns)], ignore_index=True)
    return df

# ----------------------
# Generate worksheets (two layout modes)
# ----------------------
def generate_all_worksheets(df_formatted: pd.DataFrame, meta: dict, settings: dict,
                            banner_vars: List[str], nets_config: Dict[str, List[str]], layout_mode: str) -> Dict[str, Tuple[str, pd.DataFrame]]:
    df_raw = meta.get("raw_df", df_formatted.copy())
    worksheets = {}
    vars_to_process = [v for v in df_formatted.columns if v.lower() not in EXCLUDE_VARS]

    # first banner chosen will be used as row axis in wincross matrix
    first_banner = banner_vars[0] if banner_vars else None
    other_banners = banner_vars[1:] if len(banner_vars) > 1 else []

    for v in vars_to_process:
        qlabel = get_label_for_variable(v, meta)
        nets_selected = nets_config.get(v, [])
        if layout_mode == "Wincross Matrix":
            table = build_wincross_matrix(v, df_formatted, df_raw, meta, settings, first_banner, other_banners, nets_selected)
        else:
            table = build_vertical_table(v, df_formatted, df_raw, meta, settings, banner_vars, nets_selected)
        worksheets[v] = (qlabel, table)

    # rating summaries (same as before)
    rating_summary_rows = []
    for v in vars_to_process:
        if v in df_raw and is_rating_candidate(df_raw[v]):
            base_mask = exclude_dk_mask(df_raw[v], set(settings.get("dk_codes", DEFAULT_DK_CODES)))
            metrics = compute_rating_metrics(df_raw[v], base_mask)
            row = {"Question": get_label_for_variable(v, meta)}; row.update(metrics)
            rating_summary_rows.append(row)
    if rating_summary_rows:
        rs = pd.DataFrame(rating_summary_rows).fillna("")
        means_df = rs.reindex(columns=["Question", "Base", "Mean", "SD"]).fillna("")
        tb_df = rs.reindex(columns=["Question", "Top2%", "Bottom2%", "Top3%", "Bottom3%", "NPS"]).fillna("")
        worksheets["Means_Summary"] = ("Means Summary", means_df)
        worksheets["TopBottom_Summary"] = ("Top/Bottom Summary", tb_df)

    return worksheets

# ----------------------
# Excel writer (numeric percent formatting)
# ----------------------
def write_workbook(worksheets: Dict[str, Tuple[str, pd.DataFrame]]) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book
        title_fmt = wb.add_format({"bold": True, "font_name":"Calibri", "font_size":11, "align":"left", "valign":"vcenter"})
        header_fmt = wb.add_format({"bold": True, "font_name":"Calibri", "font_size":10, "align":"center", "valign":"vcenter", "font_color":"white", "bg_color":BLUE_HEADER})
        left_fmt = wb.add_format({"font_name":"Calibri", "font_size":10, "align":"left"})
        center_fmt = wb.add_format({"font_name":"Calibri", "font_size":10, "align":"center"})
        pct_fmt = wb.add_format({"num_format": "0.0%", "align": "center"})

        for sheet, (qtext, df) in worksheets.items():
            safe = sheet[:31]
            if df is None or df.empty:
                ws = writer.book.add_worksheet(safe)
                ws.merge_range(0,0,0,2,qtext, title_fmt)
                continue
            # write using pandas (startrow=1 for header row)
            df.to_excel(writer, sheet_name=safe, index=False, startrow=1)
            ws = writer.sheets[safe]
            ncols = len(df.columns)
            ws.merge_range(0, 0, 0, max(0, ncols-1), qtext, title_fmt)
            # write header row formatting
            for i, col in enumerate(df.columns):
                ws.write(1, i, col, header_fmt)
                if i == 0:
                    ws.set_column(i, i, 30, left_fmt)
                else:
                    # try to detect percent columns by name
                    if str(col).strip().endswith("%") or " %" in str(col):
                        ws.set_column(i, i, 12, pct_fmt)
                    else:
                        ws.set_column(i, i, 14, center_fmt)
            ws.freeze_panes(2, 1)
            ws.hide_gridlines(2)
    out.seek(0)
    return out.read()

# ----------------------
# Streamlit UI
# ----------------------
st.title("Tabulation Automation — v4.4 (Wincross matrix)")
st.markdown("Upload dataset (.sav/.csv/.xlsx). Use 'Wincross Matrix' layout to pivot data (first selected banner used as rows).")

# Sidebar controls
st.sidebar.header("Settings")
dk_text = st.sidebar.text_input("DK/Ref codes (comma separated)", value="88,99,-1,98")
dk_codes = set(int(x.strip()) for x in dk_text.split(",") if x.strip().lstrip('-').isdigit())
decimals = st.sidebar.number_input("Percent decimals", min_value=0, max_value=2, value=1)
show_percent_sign = st.sidebar.checkbox("Show % sign in percent columns", value=True)
preview_n = st.sidebar.number_input("Number of tables to preview", min_value=1, max_value=50, value=5)
layout_mode = st.sidebar.selectbox("Layout mode", options=["Standard", "Wincross Matrix"], index=1,
                                   help="Wincross Matrix uses first selected banner as row groups and pivots other banners across columns.")

uploaded = st.file_uploader("Upload data (.sav, .csv, .xlsx)", type=["sav","csv","xls","xlsx"])

if uploaded:
    try:
        df_formatted, meta = read_file(uploaded)
    except Exception as e:
        st.error(f"Unable to read file: {e}")
        st.stop()

    # banner selector (multiple)
    banner_candidates = [c for c in df_formatted.columns if c.lower() not in EXCLUDE_VARS]
    st.sidebar.markdown("### Banner variables (simple)")
    banner_vars = st.sidebar.multiselect("Select banner variable(s) to split by (first selected used as row axis in Wincross Matrix)", options=banner_candidates)

    st.sidebar.markdown("---")
    st.sidebar.markdown("### Nets / Means configuration (per question)")
    st.sidebar.markdown("Toggle candidate numeric questions and mark as rating then select metrics.")
    df_raw = meta.get("raw_df", df_formatted.copy())
    nets_config: Dict[str, List[str]] = {}
    rating_candidates = [v for v in df_formatted.columns if v in df_raw and is_rating_candidate(df_raw[v]) and v.lower() not in EXCLUDE_VARS]
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

    st.success(f"Loaded {df_formatted.shape[0]} rows × {df_formatted.shape[1]} columns")
    st.dataframe(df_formatted.head(8))

    settings = {"dk_codes": dk_codes, "decimals": decimals, "show_percent_sign": show_percent_sign}

    with st.spinner("Generating tables..."):
        worksheets = generate_all_worksheets(df_formatted, meta, settings, banner_vars, nets_config, layout_mode)

    st.success(f"Generated {len(worksheets)} sheets (preview below)")

    # Preview
    shown = 0
    for sheet, (qtext, df_tab) in worksheets.items():
        if shown >= preview_n:
            break
        st.markdown(f"### {shown+1}. {qtext}")
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
                           file_name="tabulation_v4_4_export.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Upload your dataset to start.")
