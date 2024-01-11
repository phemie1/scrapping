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
provinces = ["OSUN PROVINCE 1","OSUN PROVINCE 4","OSUN PROVINCE 6","OSUN PROVINCE 10", "OSUN PROVINCE 13", "OSUN PROVINCE 15", "REGION 43"]

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
    driver.execute_script("document.getElementsByName('year1')[0].value = '2022';")
    driver.execute_script("document.getElementsByName('month2')[0].value = 'Aug';")
    driver.execute_script("document.getElementsByName('year2')[0].value = '2023';")

    """# Select radio buttons
    try:
        display_type_radio = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "//input[@name='display_type' and @value='1']")))
        # Use JavaScript to set the radio button as checked
        js_code = 'arguments[0].checked = true;'
        driver.execute_script(js_code, display_type_radio)
        print("Selected display type.")
    except Exception as e:
        print(f"An error occurred while selecting display type: {e}")
    """
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

#MRR
def select_mrr_options(driver, provinces):
    mrr_url = "https://portal.rccg.org/mrr.php"
    driver.get(mrr_url)
    type_dropdown = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.NAME, "the_type")))
    type_dropdown_select = Select(type_dropdown)
    type_dropdown_select.select_by_value("0")

    month1_dropdown = Select(driver.find_element(By.NAME, "month1"))
    month1_dropdown.select_by_visible_text("Sep")

    year1_dropdown = Select(driver.find_element(By.NAME, "year1"))
    year1_dropdown.select_by_visible_text("2022")

    month2_dropdown = Select(driver.find_element(By.NAME, "month2"))
    month2_dropdown.select_by_visible_text("Aug")

    year2_dropdown = Select(driver.find_element(By.NAME, "year2"))
    year2_dropdown.select_by_visible_text("2023")

    for selected_province in provinces:
        province_dropdown = Select(driver.find_element(By.NAME, "sel_prov"))
        province_dropdown.select_by_visible_text(selected_province)

        javascript_code = 'document.querySelector(\'input[type="submit"][name="submit"][value="Generate"].bigsubmit\').click();'
        driver.execute_script(javascript_code)

        province_dropdown = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "sel_prov")))
        province_select = Select(province_dropdown)

        province_name = province_select.first_selected_option.text

        province_header_locator = (By.XPATH, f'//h2[contains(text(), "{province_name}")]')
        province_header = WebDriverWait(driver, 30).until(EC.presence_of_element_located(province_header_locator))

def extract_mrr_data(driver, provinces):
    print("Navigating to MRR Page...")
    select_mrr_options(driver, provinces)
    print("Presently on MRR Page...")
    data = extract_table_data(driver, 'DataTables_Table_1')
    return data

# CSR Functions
def select_csr_options(driver, province_name):
    csr_url = f"https://portal.rccg.org/csr_multiple_years.php?cmonth=Oct&kyear=2023&cprovince={province_name}"
    driver.get(csr_url)
    
    # Wait for the select elements to be present and interactable
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'select[name="month1"]'))
    )
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'select[name="year1"]'))
    )
    
    # Set default values using JavaScript
    driver.execute_script('''
        document.querySelector('select[name="month1"]').value = 'Sep';
        document.querySelector('select[name="year1"]').value = '2022';
        document.querySelector('select[name="month2"]').value = 'Aug';
        document.querySelector('select[name="year2"]').value = '2023';
    ''')

    # Wait for the submit button to be present and clickable
    submit_button = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="submit"][name="submit"][value="submit"]'))
    )

    # Click the submit button
    submit_button.click()

    # Wait for the report to load
    time.sleep(10)

def extract_csr_data(driver, province_name):
    print("Navigating to CSR Page...")
    select_csr_options(driver, province_name)
    print("Presently on CSR Page...")
    csr_table1 = extract_table_data(driver, 'DataTables_Table_0')
    save_to_excel(csr_table1, 'CSR', province_name)
    time.sleep(5)
    csr_table2 = extract_table_data(driver, 'DataTables_Table_1')
    save_to_excel(csr_table2, 'CSR DISTRIBUTION', province_name)
    return csr_table1, csr_table2
   
def save_to_excel(data, sheet_name, province_name, sheet_number=None):
    download_directory = r'C:\Users\PMD - FEMI\Desktop\PROVINCIAL REPORTS\REGION 43'
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
        for province_name in provinces:
            for report_type in ["AVG. ATTENDANCE", "HOUSE FELLOWSHIP", "CONVERTS"]:
                select_report(driver, province_name, report_type)
                data = extract_table_data(driver)
                save_to_excel(data, report_type, province_name)
           
            # Extract and save MRR data for the current province
            mrr_data = extract_mrr_data(driver, [province_name])
            save_to_excel(mrr_data, "MRR", province_name)

            # Extract and save CSR tables for the current province
            csr_data = extract_csr_data(driver, province_name)
            if csr_data:
                # Separate CSR data into two tables
                table1, table2 = csr_data

            save_to_excel(table1, "CSR", province_name)
            save_to_excel(table2, "CSR DISTRIBUTION", province_name)
                       
            print(f"Successfully downloaded all the reports for {province_name}.")
   
    except Exception as e:
        print(f"An error occurred: {e}")

    """finally:
        # Close the browser
        driver.quit()"""

if __name__ == "__main__":
    main()