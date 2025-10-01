import streamlit as st
import pandas as pd
import io

# --- Helper Function (No changes needed here) ---
def format_table(crosstab_counts, crosstab_percents):
    """
    Combines counts and percentages into a single DataFrame with the
    format 'Count (Percentage%)'.
    """
    # Ensure percentages are rounded before converting to string
    crosstab_percents_rounded = crosstab_percents.round(1).astype(str)
    formatted_table = crosstab_counts.astype(str) + " (" + crosstab_percents_rounded + "%)"
    return formatted_table

# --- Set up the Streamlit page ---
st.set_page_config(page_title="Market Research Table Automation", layout="wide")
st.title("📊 Market Research Table Automation")
st.write("This app automates the creation of cross-tabulation tables from raw survey data and banner definitions.")

# --- File Uploader Widgets ---
st.sidebar.header("1. Upload Your Files")
raw_data_file = st.sidebar.file_uploader("Upload Raw Data (XLSX)", type=["xlsx"])
banner_file = st.sidebar.file_uploader("Upload Banner Cuts (XLSX)", type=["xlsx"])

# --- Main App Logic ---
if raw_data_file is not None and banner_file is not None:
    try:
        df_raw = pd.read_excel(raw_data_file)
        df_banners = pd.read_excel(banner_file)

        st.sidebar.header("2. Select Questions")
        # Let user select which questions (columns) to tabulate
        all_columns = df_raw.columns.tolist()
        questions_to_tabulate = st.sidebar.multiselect(
            "Choose questions to create tables for:",
            options=all_columns
        )

        st.sidebar.header("3. Generate Report")
        if st.sidebar.button("Generate Tables"):
            if not questions_to_tabulate:
                st.warning("Please select at least one question to tabulate.")
            else:
                with st.spinner('Processing tables... This may take a moment.'):
                    # Use an in-memory buffer to store the Excel file
                    output_buffer = io.BytesIO()
                    
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        for question in questions_to_tabulate:
                            
                            # --- Create 'Total' Column ---
                            total_counts = df_raw[question].value_counts()
                            total_percents = df_raw[question].value_counts(normalize=True) * 100
                            final_table = pd.DataFrame({
                                'Total_Count': total_counts,
                                'Total_Percent': total_percents
                            })
                            final_table['Total'] = final_table.apply(
                                lambda row: f"{int(row['Total_Count'])} ({row['Total_Percent']:.1f}%)", axis=1
                            )
                            final_table = final_table[['Total']]

                            # --- Loop through each banner ---
                            for _, row in df_banners.iterrows():
                                banner_variable = row['BannerVariable']
                                banner_name = row['BannerName']
                                
                                # Crosstab for counts and percentages
                                crosstab_counts = pd.crosstab(df_raw[question], df_raw[banner_variable])
                                crosstab_percents = pd.crosstab(df_raw[question], df_raw[banner_variable], normalize='columns') * 100
                                
                                # Format and combine
                                formatted_banner_table = format_table(crosstab_counts, crosstab_percents)
                                formatted_banner_table.columns = pd.MultiIndex.from_product([[banner_name], formatted_banner_table.columns])
                                
                                final_table = final_table.join(formatted_banner_table, how='outer')

                            final_table = final_table.fillna('-')
                            
                            # Write the completed table to a sheet in the in-memory Excel file
                            sheet_name = question[:31] # Excel sheet names are max 31 chars
                            final_table.to_excel(writer, sheet_name=sheet_name)
                    
                    # After the loop, prepare the buffer for download
                    output_buffer.seek(0)
                    
                    st.success("✅ Success! Your tables are ready for download.")
                    
                    # --- Download Button ---
                    st.download_button(
                        label="📥 Download Tables as Excel File",
                        data=output_buffer,
                        file_name="automated_tables.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.error("Please check if your file formats are correct and column names in the banner file match the data file.")
else:
    st.info("Please upload both the raw data and banner cuts files to begin.")