import streamlit as st
import pandas as pd
import io
import re

# ... (format_cell function remains the same) ...

# --- Main App Function ---
def run_app():
    # ... (st.set_page_config, st.title, st.write, st.sidebar.header, file_uploader remain the same) ...

    if raw_data_file and banner_file:
        try:
            # --- Load All Necessary Data Sheets ---
            # NOTE: If the user uploads CSV files, this will read the CSV content.
            df_raw = pd.read_excel(raw_data_file, sheet_name="Raw Data")
            df_val_labels = pd.read_excel(raw_data_file, sheet_name="Val labels", header=1)
            df_banners = pd.read_excel(banner_file, sheet_name="Banners", header=1)

            # --- CRITICAL FIX: Ensure Correct Column Names for Val Labels ---
            # Use positional index [0] to get the column containing the variable names.
            # This is robust against 'Variable Values' being read as 'Unnamed: 0' in CSVs.
            if len(df_val_labels.columns) >= 3:
                # Assuming the structure is: [Variable Name Column], [Value Column], [Label Column]
                df_val_labels.rename(columns={
                    df_val_labels.columns[0]: 'Variable Values', # This handles the 'Unnamed: 0' or 'Val labels' issue
                    df_val_labels.columns[1]: 'Value',
                    df_val_labels.columns[2]: 'Label'
                }, inplace=True)
            else:
                st.error("The 'Val labels' sheet does not have the expected 3 columns. Check your file structure.")
                return 
            
            # --- Data Pre-processing: Apply Value Labels to Raw Data ---
            df_labeled = df_raw.copy()
            # Forward fill the 'Variable Values' column
            df_val_labels['Variable Values'] = df_val_labels['Variable Values'].ffill()

            for var_name in df_val_labels['Variable Values'].unique():
                # Ensure var_name is a string before checking membership
                if isinstance(var_name, str) and var_name in df_labeled.columns:
                    # Create a mapping dictionary for the current variable
                    mapping = df_val_labels[df_val_labels['Variable Values'] == var_name].set_index('Value')['Label'].to_dict()
                    # Apply the mapping to the raw data
                    df_labeled[var_name] = df_raw[var_name].map(mapping).fillna(df_raw[var_name])

            # ... (Sections 2 and 3 remain the same, including the sheet name sanitization fix from before) ...

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
                                total_counts = df_labeled[question].value_counts().sort_index()
                                grand_total = total_counts.sum()
                                
                                # Use .astype(str) for the index to prevent potential non-string issues
                                final_table = pd.DataFrame(index=total_counts.index.astype(str))
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
                                
                                # Sheet Name Sanitization (using replacement for safety)
                                invalid_chars = r'[\\/*?\[\]:]'
                                sheet_name = re.sub(invalid_chars, '_', question)
                                sheet_name = sheet_name[:31].strip() 
                                
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
            st.error(f"An error occurred: {e}")
            st.error("Please check that your sheet names ('Raw Data', 'Val labels', 'Banners') and column names are correct.")
    else:
        st.info("Please upload both the raw data and banner cuts files to begin.")

# --- Run the App ---
if __name__ == "__main__":
    run_app()