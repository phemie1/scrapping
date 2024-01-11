import pandas as pd
from pathlib import Path

def merge_excel_files(input_folder, output_file):
    # Create a Path object for input_folder
    input_path = Path(input_folder)

    # Get a list of all Excel files in the input folder
    excel_files = list(input_path.glob("*.xlsx"))

    # Check if there are any Excel files in the folder
    if not excel_files:
        print("No Excel files found in the specified folder.")
        return

    # Initialize an empty DataFrame to store the merged data
    merged_data = pd.DataFrame()

    # Loop through each Excel file
    for file_path in excel_files:
        try:
            # Read the 'TOTAL' row based on the 'PARISH' column
            df_total_row = pd.read_excel(file_path, engine='openpyxl')
            df_total_row = df_total_row[df_total_row['PARISH'] == 'TOTAL']

            # Replace 'TOTAL' with the Excel file name in the 'PARISH' column
            df_total_row['PARISH'] = file_path.stem
           
            # Concatenate the 'TOTAL' row to the merged data
            merged_data = pd.concat([merged_data, df_total_row], ignore_index=True)

        except Exception as e:
            print(f"Error reading file {file_path}: {e}")

    # Write the merged data to a new Excel file
    output_path = input_path / output_file  # Creating output file path
    merged_data.to_excel(output_path, index=False)
    print(f"Merged data saved to {output_path}")

# Example usage:
input_folder = r'C:\Users\PMD - FEMI\Desktop\provinces attendance\REGION 4'
output_file = r"C:\Users\PMD - FEMI\Desktop\provinces attendance\REGION4_merged.xlsx"
merge_excel_files(input_folder, output_file)
