import os
import pandas as pd

# Replace 'your_folder_path' with the actual path to your folder containing Excel files
folder_path = r'C:\Users\PMD - FEMI\Desktop\PROVINCIAL REPORTS\Regions CA\cleanedNOVEMBER'

# Iterate through all Excel files in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        # Construct the full path to the Excel file
        excel_file = os.path.join(folder_path, file_name)

        # Read the data from Sheet1 into a DataFrame
        df = pd.read_excel(excel_file, sheet_name='Sheet1')

        # Pivot the data to get the desired format
        pivot_table = pd.pivot_table(df, values='ATTENDANCE', index=['REGION', 'PROVINCE'], aggfunc='count').reset_index()
        pivot_table.rename(columns={'ATTENDANCE': 'NOV. < 100'}, inplace=True)

        # Create a new Excel writer object
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            # Write the pivoted result to Sheet2 with the desired format
            pivot_table.to_excel(writer, sheet_name='Sheet2', index=False)

        # Print a success message
        print(f"Successfully pivoted and saved '{excel_file}'")
