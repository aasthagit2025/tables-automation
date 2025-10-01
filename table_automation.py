import streamlit as st
import pandas as pd
import io
import re

# --- Helper Function to combine counts and percentages ---
def format_cell(count, total):
    if total == 0:
        return "0 (0.0%)"
    percentage = (count / total) * 100
    return f"{int(count)} ({percentage:.1f}%)"

# --- Main App Function ---
def run_app():
    st.set_page_config(page_title="Advanced Table Automation", layout="wide")
    st.title("📊 Advanced Market Research Table Automation")
    st.write("This version is designed to read multi-sheet Excel files and map value labels to raw data before creating tables.")

    # --- File Uploaders ---
    st.sidebar.header("1. Upload Your Files")
    raw_data_file = st.sidebar.file_uploader("Upload Raw Data (XLSX)", type=["xlsx"])
    banner_file = st.sidebar.file_uploader("Upload Banner Cuts (XLSX)", type=["xlsx"])

    if raw_data_file and banner_file:
        try:
            # --- Load All Necessary Data Sheets ---
            df_raw = pd.read_excel(raw_data_file, sheet_name="Raw Data")
            df_val_labels = pd.read_excel(raw_data_file, sheet_name="Val labels", header=1)
            df_banners = pd.read_excel(banner_file, sheet_name="Banners", header=1)

            # --- FIX: Ensure Correct Column Names for Val Labels ---
            if len(df_val_labels.columns) >= 3:
                # Rename columns based on the multi-row header structure
                df_val_labels.columns = ['Variable Values', 'Value', 'Label']
            else:
                st.error("The 'Val labels' sheet does not have the expected 3 columns. Check your file structure.")
                return 
            
            # --- Data Pre-processing: Apply Value Labels to Raw Data ---
            df_labeled = df_raw.copy()
            df_val_labels['Variable Values'] = df_val_labels['Variable Values'].ffill()

            for var_name in df_val_labels['Variable Values'].unique():
                if var_name in df_labeled.columns:
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
                                # Start with the Total column for the current question
                                total_counts = df_labeled[question].value_counts().sort_index()
                                grand_total = total_counts.sum()
                                
                                final_table = pd.DataFrame(index=total_counts.index)
                                final_table['Total'] = total_counts.apply(lambda x: format_cell(x, grand_total))
                                
                                # Process each banner defined in the banner file
                                for _, banner_row in df_banners.iterrows():
                                    var_label = banner_row['var labels']
                                    val_label = banner_row['Val labels']
                                    banner_name = val_label 
                                    
                                    if pd.notna(var_label) and pd.notna(val_label) and var_label in df_labeled.columns:
                                        subgroup_data = df_labeled[df_labeled[var_label] == val_label]
                                        banner_counts = subgroup_data[question].value_counts()
                                        banner_total = banner_counts.sum()
                                        
                                        final_table[banner_name] = banner_counts.apply(lambda x: format_cell(x, banner_total))

                                final_table = final_table.fillna("0 (0.0%)")
                                
                                # --- CRITICAL FIX: Sanitize Sheet Name More Thoroughly ---
                                # Remove all invalid Excel sheet characters: \ / * [ ] : ?
                                # The re.sub function will replace all occurrences of these characters with an empty string.
                                invalid_chars = r'[\\/*?\[\]:]'
                                sheet_name = re.sub(invalid_chars, '', question)
                                sheet_name = sheet_name[:31].strip() # Truncate to 31 chars and remove leading/trailing whitespace
                                
                                # Handle case where sanitization leaves an empty or invalid name
                                if not sheet_name:
                                    sheet_name = f"Table_{questions_to_tabulate.index(question) + 1}"
                                # --------------------------------------------------------

                                final_table.to_excel(writer, sheet_name=sheet_name)

                        st.success("✅ Success! Your tables are ready.")
                        st.download_button(
                            label="📥 Download Tables as Excel File",
                            data=output_buffer.getvalue(),
                            file_name="automated_tables_report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

        except Exception as e:
            st.error(f"An error occurred: {e}")
            st.error("Please check that your sheet names ('Raw Data', 'Val labels', 'Banners') and column names are correct.")
    else:
        st.info("Please upload both the raw data and banner cuts files to begin.")

# --- Run the App ---
if __name__ == "__main__":
    run_app()