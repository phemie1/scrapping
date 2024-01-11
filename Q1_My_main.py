import subprocess

def run_main_script(script_path):
    try:
        subprocess.run(["python", script_path], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running {script_path}: {e}")

def main():
 
    # Clean the first three sheets; i.e attendance, hf, and converts
    threesheets_script_path = "Q1_3_sheets_Expected.py"
    # For Church Analysis, MRR, CSRs
    foursheets_script_path = "Q1_4_sheets.py"
   
    # For Adding Charts
    charts_script_path = "Q1_charts.py"
    # For Populating the Powerpoint for province
    pptx_script_path = "Q1_pptx.py"

    # For Populating the Powerpoint for region
    #pptx_script_path = "Q1_regions_pptx.py"
   
    # Run the first cleaning script (Q1_3_sheets.py)
    print("Cleaning Attendance, House Fellowship And Converts' sheets...")
    run_main_script(threesheets_script_path)

    # Run the second cleaning script (Q1_4_sheets.py) if the first script succeeds
    print("\nCleaning Church Analysis, MRR, CSRs' sheets...")
    run_main_script(foursheets_script_path)
    
    # Adding charts to the specified sheets
    print("\nAdding charts to Excel sheet...")
    run_main_script(charts_script_path)
    
    # Populating the PPTX
    print("\nPopulating the PPTX...")
    run_main_script(pptx_script_path)
    
if __name__ == "__main__":
    main()
