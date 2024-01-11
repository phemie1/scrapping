import time
import os
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pandas as pd

# Constants
LOGIN_URL = "https://portal.rccg.org/home.php"
PERFORMANCE_URL = "https://portal.rccg.org/main_official_report_spread.php?financial=0"
provinces = {"REGION 1": 7, "REGION 2": 10, "REGION 3": 11, "REGION 4": 12, "REGION 5": 13, "REGION 6": 14, "REGION 7": 15, "REGION 8": 16,
    "REGION 9": 17, "REGION 10": 8, "REGION 11": 5, "REGION 12": 9, "REGION 13": 18, "REGION 14": 19, "REGION 15": 20, "REGION 16": 21,
    "REGION 17": 22, "REGION 18": 23, "REGION 19": 24, "REGION 20": 2, "REGION 21": 25, "REGION 22": 26, "REGION 23": 27, "REGION 24": 28,
    "REGION 25": 30, "REGION 26": 31, "REGION 27": 32, "REGION 28": 33, "REGION 29": 34, "REGION 30": 35, "REGION 31": 36,"REGION 32": 37,
    "REGION 33": 38, "REGION 34": 39, "REGION 35": 40, "REGION 36": 41, "REGION 37": 42, "REGION 38": 43, "REGION 39": 47,"REGION 40": 46,
    "REGION 41": 45, "REGION 42": 44, "REGION 43": 56, "REGION 44": 57, "REGION 45": 58, "REGION 46": 59,"REGION 47": 60, "REGION 48": 61,
    "REGION 49": 62, "REGION 50": 63, "REGION 51": 64, "REDEMPTION CITY REGION": 48
}

# Credentials
USERNAME = '08145045108'
PASSWORD = '@d3m0l@000'

def login(driver, username, password):
    driver.get(LOGIN_URL)
    username_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "login_username")))
    password_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "login_password")))
    submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "login")))
    
    username_field.send_keys(username)
    password_field.send_keys(password)
    driver.execute_script("arguments[0].scrollIntoView();", submit_button)
    submit_button.click()

    # Check for a successful login
    if driver.current_url != LOGIN_URL:
        print("Login Successful!")
        click_ict_link(driver)  # Click on the ICT link after successful login
    else:
        print("Login failed.")

def click_ict_link(driver):
    link_text = "NATIONAL INFORMATION TECHNOLOGY DEPARTMENT"
    try:
        link = driver.find_element(By.PARTIAL_LINK_TEXT, link_text)
        link.click()
        print("Clicked on the ICT link.")
    except Exception as e:
        print(f"An error occurred while clicking the ICT link: {e}")

def select_report(driver, province_name, report_type):
    province_url = f"https://portal.rccg.org/main_official_report_spread.php?month1=Jan&month2=Dec&year1=2023&year2=2023&prov={province_name.replace(' ', '%20')}&quarter_or_annual=1&hide_benchmark=0&hide_position=1"
    driver.get(province_url)
    # Select the desired report type
    report_type_dropdown = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//select[@name='report_type']")))
    report_type_select = Select(report_type_dropdown)
    report_type_select.select_by_visible_text(report_type)

    # Set the default values for year and month (modify as needed)
    driver.execute_script("document.getElementsByName('month1')[0].value = 'Sep';")
    driver.execute_script("document.getElementsByName('year1')[0].value = '2023';")
    driver.execute_script("document.getElementsByName('month2')[0].value = 'Nov';")
    driver.execute_script("document.getElementsByName('year2')[0].value = '2023';")

    try:
        quarter_or_annual_radio = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//input[@name='quarter_or_annual' and @value='1']")))
        # Use JavaScript to force the click
        driver.execute_script("arguments[0].click();", quarter_or_annual_radio)
        print("Selected quarter or annual.")
    except Exception as e:
        print(f"An error occurred while selecting quarter or annual: {e}")

    # Click the "Generate" button directly
    try:
        generate_button = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "bigsubmit")))
        #print(f"Found the 'Generate' button with XPath: {generate_button}")
        driver.execute_script("arguments[0].click();", generate_button)
        print("Clicked on Generate button.")
    except Exception as e:
        print(f"An error occurred while clicking the Generate button: {e}")
    # Wait for the report to load
    time.sleep(10)  

def extract_table_data(driver, table_id='DataTables_Table_0'):
    print("Extracting the data...")
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    table = soup.find('table', {'id': table_id})
    
    headers = table.find_all(['th', 'td'])  # Find both headers and data cells
    header_texts = [header.text.strip() for header in headers if header.text.strip()]
    
    rows = table.find_all('tr', {'role': 'row'})
    data = []
    for row in rows:
        cols = row.find_all(['th', 'td'])
        cols = [col.text.strip() for col in cols]
        data.append(cols)
    
    # Determine the number of columns dynamically
    num_columns = max(len(row) for row in data)
    
    # Adjust header_texts if necessary
    if len(header_texts) != num_columns:
        if len(header_texts) < num_columns:
            header_texts += [''] * (num_columns - len(header_texts))
        else:
            header_texts = header_texts[:num_columns]
    
    # Pad rows with empty strings if the number of columns is less than the maximum number of columns
    for i in range(len(data)):
        data[i] += [''] * (num_columns - len(data[i]))
    
    return pd.DataFrame(data, columns=header_texts)  

def church_analysis(driver, provinces):
    for province_name, region_id in provinces.items():
        # Use JavaScript to click on the "View breakdown" link directly
        breakdown_url = f"https://portal.rccg.org/parish_report_breakdown.php?month=Nov&year=2023&from=1&to=99&region={region_id}"
        driver.get(breakdown_url)

def extract_ca_data(driver, provinces):
    print("Navigating to CHURCH ANALYSIS Page...")
    church_analysis(driver, provinces)
    print("Presently on CHURCH ANALYSIS Page...")
    data = extract_table_data(driver, 'DataTables_Table_0')
    return data

def save_to_excel(data, sheet_name, province_name, sheet_number=None):
    download_directory = r'C:\Users\PMD - FEMI\Desktop\PROVINCIAL REPORTS\Regions CA\Nov'
    excel_filename = os.path.join(download_directory, f"{province_name}.xlsx")
    
    if not os.path.exists(excel_filename):
        # Create a new Excel file and directory if they don't exist
        os.makedirs(os.path.dirname(excel_filename), exist_ok=True)

    # Check if the file already exists
    if os.path.exists(excel_filename):
        with pd.ExcelWriter(excel_filename, engine='openpyxl', mode='a') as writer:
            # Check if the sheet already exists
            if sheet_name in writer.sheets:
                # Replace the existing sheet
                writer.book.remove(writer.sheets[sheet_name])
                data.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # Create a new sheet if it doesn't exist
                data.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        # Create a new Excel file and sheet if the file doesn't exist
        with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
            data.to_excel(writer, sheet_name=sheet_name, index=False)

def main():
    # Create a Chrome WebDriver instance
    driver = webdriver.Chrome()

    try:
        login(driver, USERNAME, PASSWORD)
        for province_name, region_id in provinces.items():
            # Extract and save CHURCH ANALYSIS data for the current province
            ca_data = extract_ca_data(driver, {province_name: region_id})
            save_to_excel(ca_data, "CHURCH ANALYSIS", province_name)
                       
            print(f"Successfully downloaded all the reports for {province_name}.")
   
    except Exception as e:
        print(f"An error occurred: {e}")

    """finally:
        # Close the browser
        driver.quit()"""

if __name__ == "__main__":
    main()