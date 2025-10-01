import pandas as pd

def format_table(crosstab_counts, crosstab_percents):
    """
    Combines counts and percentages into a single DataFrame with the
    format 'Count (Percentage%)'.
    """
    formatted_table = crosstab_counts.astype(str) + " (" + crosstab_percents.round(1).astype(str) + "%)"
    return formatted_table

def run_automation():
    """
    Main function to run the table automation process.
    """
    print("--- Market Research Table Automation ---")

    # --- 1. GET USER INPUTS ---
    try:
        raw_data_file = input("Enter the name of your raw data file (e.g., raw_data.xlsx): ")
        banner_file = input("Enter the name of your banner cuts file (e.g., banner_cuts.xlsx): ")
        output_file = input("Enter the desired name for your output Excel file (e.g., output_tables.xlsx): ")

        # Load the data using pandas
        df_raw = pd.read_excel(raw_data_file)
        df_banners = pd.read_excel(banner_file)
    except FileNotFoundError as e:
        print(f"\nERROR: File not found -> {e}. Please make sure the files are in the same folder as the script.")
        return
    except Exception as e:
        print(f"\nAn error occurred while reading the files: {e}")
        return

    # --- 2. DEFINE QUESTIONS TO TABULATE ---
    # List the column names from your raw data that you want to create tables for.
    questions_to_tabulate = [
        'Q1_Awareness',
        'Q2_Satisfaction'
        # Add more question columns here
    ]

    # --- 3. CREATE EXCEL WRITER TO SAVE OUTPUT ---
    # This allows us to write multiple tables to different sheets in one file.
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        
        # --- 4. MAIN LOOP TO GENERATE TABLES ---
        for question in questions_to_tabulate:
            print(f"Processing table for: {question}...")

            # --- Create the 'Total' Column first ---
            total_counts = df_raw[question].value_counts()
            total_percents = df_raw[question].value_counts(normalize=True) * 100
            
            # Combine into a single formatted DataFrame
            final_table = pd.DataFrame({
                'Total_Count': total_counts,
                'Total_Percent': total_percents
            })
            final_table['Total'] = final_table['Total_Count'].astype(str) + " (" + final_table['Total_Percent'].round(1).astype(str) + "%)"
            final_table = final_table[['Total']] # Keep only the formatted column

            # --- Loop through each banner defined in the banner_cuts file ---
            for index, row in df_banners.iterrows():
                banner_variable = row['BannerVariable']
                banner_name = row['BannerName']

                # Create the cross-tabulation for counts
                crosstab_counts = pd.crosstab(df_raw[question], df_raw[banner_variable])
                
                # Create the cross-tabulation for column percentages
                crosstab_percents = pd.crosstab(df_raw[question], df_raw[banner_variable], normalize='columns') * 100

                # Format the table by combining counts and percentages
                formatted_banner_table = format_table(crosstab_counts, crosstab_percents)
                
                # Add a header for the banner group
                formatted_banner_table.columns = pd.MultiIndex.from_product([[banner_name], formatted_banner_table.columns])

                # Join this banner's results to the main table for the question
                final_table = final_table.join(formatted_banner_table, how='outer')

            # Fill any missing values with a dash for clarity
            final_table = final_table.fillna('-')

            # --- 5. SAVE THE COMPLETED TABLE TO A SHEET ---
            # Use the question name as the sheet name
            final_table.to_excel(writer, sheet_name=question[:31]) # Excel sheet names have a 31-char limit

    print(f"\n✅ Success! All tables have been generated and saved to '{output_file}'.")


# --- Run the main function ---
if __name__ == "__main__":
    run_automation()