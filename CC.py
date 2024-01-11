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

    # Loop through each Excel file and merge its data into the DataFrame
    for file_path in excel_files:
        df = pd.read_excel(file_path)
        merged_data = pd.concat([merged_data, df], ignore_index=True)

        # Add an empty row after each file
        if not merged_data.empty:
            merged_data = pd.concat([merged_data, pd.DataFrame(index=[None])], ignore_index=True)
    
    # Write the merged data to a new Excel file
    output_path = input_path / output_file  # Creating output file path
    merged_data.to_excel(output_path, index=False)
    print(f"Merged data saved to {output_path}")

# Example usage:
input_folder = r"C:\Users\PMD - FEMI\Desktop\REGION 1"
output_file = r"C:\Users\PMD - FEMI\Desktop\REGION 1\"REGION1.xlsx"
merge_excel_files(input_folder, output_file)

