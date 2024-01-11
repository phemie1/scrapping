import os
import pandas as pd

# Input and output folder paths
input_folder_path = r'C:\Users\PMD - FEMI\Desktop\PROVINCIAL REPORTS\Regions CA\Nov'
output_folder_path = r'C:\Users\PMD - FEMI\Desktop\PROVINCIAL REPORTS\Regions CA\cleanedNOVEMBER'

# Create the output folder if it doesn't exist
os.makedirs(output_folder_path, exist_ok=True)

# Iterate over all Excel files in the input folder
for file_name in os.listdir(input_folder_path):
    if file_name.endswith('.xlsx'):
        # Construct the full path for each file
        file_path = os.path.join(input_folder_path, file_name)

        # Read the Excel file
        df = pd.read_excel(file_path)

        # Delete row 1
        df = df.drop(0)

        # List of column indices to be deleted
        columns_to_delete = [0, 1, 3, 6, 7, 9, 10, 11]

        # Delete the specified columns
        df_cleaned = df.drop(df.columns[columns_to_delete], axis=1)

        # Set the specified heading
        df_cleaned.columns = ['SN', 'PROVINCE', 'REGION', 'ATTENDANCE']

        # Delete rows where 'ATTENDANCE' is 'AVG. ATT'
        df_cleaned = df_cleaned[df_cleaned['ATTENDANCE'] != 'AVG. ATT']

        # Save the cleaned data to a new Excel file in the output folder
        output_file_path = os.path.join(output_folder_path, f'cleaned_{file_name}')
        df_cleaned.to_excel(output_file_path, index=False)

        print(f"The cleaned data has been saved to {output_file_path}")
