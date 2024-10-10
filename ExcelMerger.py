import pandas as pd
import os

def update_main_excel_with_month_files(main_file_path, month_files_folder):
    # Load the main Excel file
    main_df = pd.read_excel(main_file_path)
    
    # Ensure that the 'Profile' column exists in the main file
    if 'Profile' not in main_df.columns:
        raise ValueError("Main file must have a 'Profile' column")
    
    # Get the list of month files (e.g., '2024-04', '2024-05', etc.)
    month_files = [f for f in os.listdir(month_files_folder) if f.endswith('.xlsx')]
    
    for month_file in month_files:
        # Load each month file
        month_df = pd.read_excel(os.path.join(month_files_folder, month_file))
        
        # Ensure 'Profile' and the value column exist in the month file
        if 'Profile' not in month_df.columns:
            raise ValueError(f"File {month_file} must have a 'Profile' column")
        if 'Value' not in month_df.columns:  # Assuming 'value' is the column with 1 or 0
            raise ValueError(f"File {month_file} must have a 'Value' column (1 or 0)")

        # Extract year and month from the filename (assuming format is like '04.2024.xlsx')
        month_year_filename = os.path.splitext(month_file)[0]

        # Merge the main file with the current month file on 'Profile'
        merged_df = main_df.merge(month_df[['Profile', 'Value']], on='Profile', how='left')
        
        # Replace the 1/0 values with 'X' and '-' and fill NaN with '?'
        merged_df[month_year_filename] = merged_df['Value'].apply(lambda x: 'x' if x == 1 else '-')
        merged_df.fillna({month_year_filename: "?"}, inplace=True)
        
        # Drop the temporary 'Value' column after merging
        main_df = merged_df.drop(columns=['Value'])

    # Save the updated main file
    updated_file_path = os.path.splitext(main_file_path)[0] + '_updated.xlsx'
    main_df.to_excel(updated_file_path, index=False)
    print(f"Updated file saved as {updated_file_path}")

# Usage:
main_file_path = 'Excel Files\\main\\KPI_Inactive_Unused_Profiles.xlsx'
month_files_folder = 'Excel Files\\to_merge'
update_main_excel_with_month_files(main_file_path, month_files_folder)
