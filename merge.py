import pandas as pd
import os

def combine_cost_data_excel(mcr_filepath, categorized_filepath, template_filepath, output_filepath):
    """
    Combines MCR and categorized ALM data from Excel files based on a template.
    This version fixes the MultiIndex export error.
    """
    try:
        print("--- Starting Data Combination ---")
        print(f"Loading MCR data from: {mcr_filepath}")
        mcr_df = pd.read_excel(mcr_filepath)
        mcr_df.columns = mcr_df.columns.str.strip()
        mcr_df.rename(columns={'BM ID': 'BM_ID'}, inplace=True)
        print("MCR data loaded successfully.")

        print(f"Loading ALM data from: {categorized_filepath}")
        alm_df = pd.read_excel(categorized_filepath, sheet_name='Aggregated Data', skiprows=1)
        print("ALM data loaded successfully.")
        
        alm_df.rename(columns={
            'PM Interface Element ID': 'BM_ID',
            'Actual (ALM)': 'Actual (ALM) NET',
            'Unnamed: 18': 'Actual (ALM) DCOM&DSM',
            'Development Cost (ALM)': 'Development Cost (ALM) NET',
            'Unnamed: 20': 'Development Cost (ALM) DCOM&DSM',
            'Rework Cost (ALM)': 'Rework Cost (ALM) NET',
            'Unnamed: 22': 'Rework Cost (ALM) DCOM&DSM'
        }, inplace=True)
        
        if 'BM_ID' in alm_df.columns and alm_df['BM_ID'].dtype == 'object':
             alm_df['BM_ID'] = alm_df['BM_ID'].str.replace(r'\|.*', '', regex=True).str.strip()

        print("Merging MCR and ALM dataframes...")
        merged_df = pd.merge(alm_df, mcr_df, on='BM_ID', how='left')
        print("Merge complete.")

        final_columns = [
            'BM_ID', 'Category', 'OEM', 'Region', 'Budget (MCR)-NET', 'Budget (MCR)-DCOM&DSM',
            'Actual (MCR)-NET', 'Actual (MCR)-DCOM&DSM', 'Actual (ALM) NET', 'Actual (ALM) DCOM&DSM',
            'Development Cost (ALM) NET', 'Development Cost (ALM) DCOM&DSM',
            'Rework Cost (ALM) NET', 'Rework Cost (ALM) DCOM&DSM'
        ]
        
        for col in final_columns:
            if col not in merged_df.columns:
                merged_df[col] = None
        
        final_df = merged_df[final_columns]
        
        final_df.columns = pd.MultiIndex.from_tuples([
            ('BM ID', ''), ('Category', ''), ('OEM', ''), ('Region', ''),
            ('Budget (MCR)', 'NET'), ('Budget (MCR)', 'DCOM&DSM'),
            ('Actual (MCR)', 'NET'), ('Actual (MCR)', 'DCOM&DSM'),
            ('Actual (ALM)', 'NET'), ('Actual (ALM)', 'DCOM&DSM'),
            ('Development Cost (ALM)', 'NET'), ('Development Cost (ALM)', 'DCOM&DSM'),
            ('Rework Cost (ALM)', 'NET'), ('Rework Cost (ALM)', 'DCOM&DSM')
        ])

        # CORRECTED: Removed `index=False` to resolve the MultiIndex implementation error.
        final_df.to_excel(output_filepath)
        print(f"\n🎉 Success! Combined data and saved to '{output_filepath}'")
        
    except FileNotFoundError as e:
        print(f"\n❌ Error: {e}. Please ensure file names are spelled correctly and located in the same folder as the script.")
    except Exception as e:
        print(f"\n❌ An unexpected error occurred: {e}")

# --- Main Execution Block ---
if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # --- File Configuration ---
    mcr_data_filename = os.path.join(script_dir, 'cleaned_mcr_data.xlsx')
    categorized_data_filename = os.path.join(script_dir, 'output_categorized.xlsx')
    template_filename = os.path.join(script_dir, 'Template_Cost Catalogue.xlsx')
    output_filename = os.path.join(script_dir, 'final_cost_catalogue.xlsx')

    # --- Execute the Function ---
    combine_cost_data_excel(mcr_data_filename, categorized_data_filename, template_filename, output_filename)
