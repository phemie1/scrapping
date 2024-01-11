import os
import pandas as pd

def extract_region_name(filename):
    # Extract the region name from the filename
    region_name = os.path.splitext(filename)[0].replace("_pivot", "").replace(",", "").strip()
    return region_name

def copy_pivot_tables(folder_path, output_excel_path):
    # Create an empty DataFrame to store the combined pivot table data
    combined_pivot_table = pd.DataFrame()

    # Loop through all files in the specified folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(folder_path, filename)

            # Read the Excel file and extract the pivot table (assuming it's on 'Sheet1')
            try:
                df = pd.read_excel(file_path, sheet_name="Sheet1")
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
                continue

            # Assuming the pivot table columns are 'Province', 'SEP-22/AUG-23', 'NOV-23', 'DIFF'
            if set(['Province', 'SEP-22/AUG-23', 'NOV-23', 'DIFF']).issubset(df.columns):
                # Include the region name as a new column in the DataFrame
                df['Region'] = extract_region_name(filename)

                # Leave a row before pasting the next table
                if not combined_pivot_table.empty:
                    combined_pivot_table = pd.concat([combined_pivot_table, pd.DataFrame(index=[None])])

                # Append the pivot table data to the combined DataFrame
                combined_pivot_table = pd.concat([combined_pivot_table, df[['Region', 'Province', 'SEP-22/AUG-23', 'NOV-23', 'DIFF']]])

    # Save the combined pivot table to a new Excel file
    combined_pivot_table.to_excel(os.path.join(output_excel_path, 'Combinedregionpivots.xlsx'), index=False)
    print(f"Combined pivot table saved to regionpivot.xlsx")

if __name__ == "__main__":
    folder_path = r'C:\Users\PMD - FEMI\Desktop\pivot\MergedRegions_pivot'  # Change this to the actual folder path
    output_excel_path = r'C:\Users\PMD - FEMI\Desktop\pivot\MergedRegions_pivot'  # Change this to the desired output Excel file path
    copy_pivot_tables(folder_path, output_excel_path)
