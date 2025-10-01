import streamlit as st
import pandas as pd
import io
import re

# --- Helper Function to combine counts and percentages ---
def format_cell(count, total):
    """Formats count and percentage into a single string (Count (P.P%))."""
    if total == 0:
        return "0 (0.0%)"
    percentage = (count / total) * 100
    return f"{int(count)} ({percentage:.1f}%)"

# --- Main App Function ---
def run_app():
    st.set_page_config(page_title="Advanced Table Automation", layout="wide")
    st.title("📊 Advanced Market Research Table Automation")
    st.write("This tool automatically generates crosstabulations based on raw data, value labels, and defined banner cuts.")

    # --- File Uploaders ---
    st.sidebar.header("1. Upload Your Files")
    
    # NOTE: You MUST upload the original multi-sheet XLSX files, NOT the separate CSV files.
    # The code is designed to read multiple sheets from a single XLSX workbook.
    st.warning("⚠️ Please ensure you are uploading the **original multi-sheet XLSX files**, not the individual CSV files.")
    raw_data_file = st.sidebar.file_uploader("Upload Raw Data (XLSX)", type=["xlsx"])
    banner_file = st.sidebar.file_uploader("Upload Banner Cuts (XLSX)", type=["xlsx"])

    if not raw_data_file or not banner_file:
        st.info("Please upload both the raw data and banner cuts XLSX files to begin.")
        return 

    # --- Data Processing Logic ---
    try:
        # Load All Necessary Data Sheets
        # CRITICAL FIX: The error suggests the sheets are not being read correctly.
        # We will attempt to read the specified sheets within the single file uploads.
        # If this fails, it means the sheet names are still incorrect, or the upload environment
        # is confusing the multi-sheet file.

        # --- Attempting to load files by sheet name (as per your file structure) ---
        df_raw = pd.read_excel(raw_data_file, sheet_name="Raw Data")
        # Header=1 is used to skip the first row (e.g., 'Val labels,,')
        df_val_labels = pd.read_excel(raw_data_file, sheet_name="Val labels", header=1)
        df_banners = pd.read_excel(banner_file, sheet_name="Banners", header=1)


        # --- FIX 1: Robust Column Naming for Val Labels (Positional Indexing) ---
        if len(df_val_labels.columns) >= 3:
            # Rename columns using positional indexing, which is highly robust against 'Unnamed' headers
            df_val_labels.rename(columns={
                df_val_labels.columns[0]: 'Variable Values', 
                df_val_labels.columns[1]: 'Value',           
                df_val_labels.columns[2]: 'Label'            
            }, inplace=True)
        else:
            st.error("The 'Val labels' sheet does not have the expected 3 columns. Please check its format.")
            return 
        
        # --- Data Pre-processing: Apply Value Labels to Raw Data ---
        df_labeled = df_raw.copy()
        df_val_labels['Variable Values'] = df_val_labels['Variable Values'].ffill()

        for var_name in df_val_labels['Variable Values'].unique():
            if isinstance(var_name, str) and var_name in df_labeled.columns:
                mapping = df_val_labels[df_val_labels['Variable Values'] == var_name].set_index('Value')['Label'].to_dict()
                df_labeled[var_name] = df_raw[var_name].map(mapping).fillna(df_raw[var_name])

        st.sidebar.header("2. Select Questions")
        all_columns = df_labeled.columns.tolist()
        questions_to_tabulate = st.sidebar.multiselect(
            "Choose questions to create tables for:",
            options=all_columns
        )

        st.sidebar.header("3. Generate Report")
        if st.sidebar.button("Generate Tables"):
            if not questions_to_tabulate:
                st.warning("Please select at least one question to tabulate.")
            else:
                with st.spinner('Processing... This might take a moment.'):
                    output_buffer = io.BytesIO()
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        for question in questions_to_tabulate:
                            # Total column calculation
                            total_counts = df_labeled[question].value_counts().sort_index()
                            grand_total = total_counts.sum()
                            
                            final_table = pd.DataFrame(index=total_counts.index.astype(str))
                            final_table['Total'] = total_counts.apply(lambda x: format_cell(x, grand_total))
                            
                            # Process each banner
                            for _, banner_row in df_banners.iterrows():
                                var_label = banner_row['var labels']
                                val_label = banner_row['Val labels']
                                banner_name = val_label 
                                
                                if pd.notna(var_label) and pd.notna(val_label) and var_label in df_labeled.columns:
                                    subgroup_data = df_labeled[df_labeled[var_label] == val_label]
                                    banner_counts = subgroup_data[question].value_counts()
                                    banner_total = banner_counts.sum()
                                    
                                    formatted_counts = banner_counts.reindex(final_table.index, fill_value=0)
                                    final_table[banner_name] = formatted_counts.apply(lambda x: format_cell(x, banner_total))

                            final_table = final_table.fillna("0 (0.0%)")
                            
                            # --- FIX 2: Robust Sheet Name Sanitization (Crucial for 'At least one sheet must be visible') ---
                            # Invalid Excel sheet characters: \ / * [ ] : ?
                            invalid_chars = r'[\\/*?\[\]:]'
                            
                            # CRITICAL CHANGE: Replace invalid chars with '_' instead of removing them
                            sheet_name = re.sub(invalid_chars, '_', str(question))
                            sheet_name = sheet_name[:31].strip() # Truncate to 31 chars
                            
                            # Fallback check for cases where the name is still empty
                            if not sheet_name:
                                sheet_name = f"Table_{questions_to_tabulate.index(question) + 1}"

                            final_table.to_excel(writer, sheet_name=sheet_name)

                    st.success("✅ Success! Your tables are ready.")
                    st.download_button(
                        label="📥 Download Tables as Excel File",
                        data=output_buffer.getvalue(),
                        file_name="automated_tables_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        # If the failure is still related to sheet reading, it's likely a file upload issue.
        # We provide a highly specific message to guide the user.
        st.error(f"A critical processing error occurred: {e}")
        st.error("### 🛑 **FINAL TROUBLESHOOTING STEP** 🛑")
        st.error(
            "The error persists because the app cannot find the sheets. This is usually due to one of two reasons:"
            "\n\n1.  **Incorrect File Type:** You must upload the **original multi-sheet XLSX file** for both 'Raw Data' and 'Banner Cuts', NOT separate CSV files."
            "\n2.  **Incorrect Sheet Names:** The sheet names in your XLSX files **must exactly match** `Raw Data`, `Val labels`, and `Banners`."
        )
        st.error("Please re-upload the original XLSX files and ensure your sheet names are correct.")

# --- Run the App ---
if __name__ == "__main__":
    run_app()