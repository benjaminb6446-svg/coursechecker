import streamlit as st
import pandas as pd
import time
import io
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from openpyxl import load_workbook

# --- UI SETUP ---
st.set_page_config(page_title="UChicago Course Checker", page_icon="🎓")
st.title("🎓 UChicago Course Scheduler Checker")

st.markdown("""
### How to use:
1. **Upload** your Excel file (.xlsx). 
2. **Auto-Detection**: The program checks Row 1 for headers.
3. **Format**: Department in **Column A**, Course Number in **Column B**.
4. **Results**: Found courses are marked with a **'Y' in Column C**.
""")

# --- USER INPUTS ---
uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")

col1, col2 = st.columns(2)
with col1:
    quarter = st.selectbox("Select Quarter", ["Spring", "Autumn", "Winter", "Summer"])
with col2:
    year = st.selectbox("Select Year", [2025, 2026, 2027, 2028])

target_term = f"{quarter} {year}"

def setup_headless_driver():
    """
    Configures Selenium for Streamlit Cloud (Linux).
    Forces use of system Chromium via paths set in packages.txt.
    """
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    
    # Path to Chromium installed by packages.txt
    options.binary_location = "/usr/bin/chromium"
    service = Service(executable_path="/usr/bin/chromedriver")
    
    try:
        driver = webdriver.Chrome(service=service, options=options)
        driver.set_page_load_timeout(60)
        return driver
    except Exception:
        # Fallback path if necessary
        alt_service = Service(executable_path="/usr/lib/chromium-browser/chromedriver")
        return webdriver.Chrome(service=alt_service, options=options)

if uploaded_file and st.button("🔍 Run Availability Check"):
    # Read Excel into memory
    file_bytes = uploaded_file.read()
    wb = load_workbook(filename=io.BytesIO(file_bytes))
    ws = wb.active 
    
    # --- INTELLIGENT HEADER DETECTION ---
    # Check if Row 1, Col B is a number or text
    first_cell_b = ws.cell(row=1, column=2).value
    start_row = 1
    
    try:
        # If this succeeds, Row 1 is a course number (No Header)
        if first_cell_b is not None:
            int(str(first_cell_b).strip().split('-')[0])
            start_row = 1
            st.info("No header detected. Processing starting from Row 1.")
        else:
            start_row = 2 # Empty cell, assume header or skip
    except (ValueError, TypeError):
        # Conversion failed, Row 1 is likely a text header
        start_row = 2
        st.info("Header detected. Processing starting from Row 2.")
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        status_text.text("Launching browser...")
        driver = setup_headless_driver()
        wait = WebDriverWait(driver, 20)
        
        driver.get("http://coursesearch92.ais.uchicago.edu/psc/prd92guest/EMPLOYEE/HRMS/c/UC_STUDENT_RECORDS_FL.UC_CLASS_SEARCH_FL.GBL")
        
        # Select target Term
        status_text.text(f"Selecting term: {target_term}...")
        term_dropdown = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[id*='STRM']")))
        Select(term_dropdown).select_by_visible_text(target_term)
        time.sleep(2)
        
        total_rows = ws.max_row
        found_count = 0

        # Loop through rows
        for row in range(start_row, total_rows + 1):
            subj = ws.cell(row=row, column=1).value # Dept (Col A)
            num = ws.cell(row=row, column=2).value  # Number (Col B)
            
            if not subj or not num:
                continue
            
            # Formatting query
            clean_num = str(num).strip().split('-')[0]
            query = f"{str(subj).strip()} {clean_num}"
            
            status_text.text(f"Searching: {query} (Row {row}/{total_rows})")
            
            # Perform Search
            search_bar = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input.ps-edit")))
            search_bar.click()
            search_bar.send_keys(Keys.CONTROL + "a")
            search_bar.send_keys(Keys.DELETE)
            search_bar.send_keys(query + Keys.ENTER)
            
            time.sleep(2)
            page_content = driver.find_element(By.TAG_NAME, "body").text.lower()
            
            # Place 'Y' in Column C (Index 3)
            if "no results found" not in page_content and clean_num in page_content:
                ws.cell(row=row, column=3).value = "Y"
                found_count += 1
            else:
                ws.cell(row=row, column=3).value = None

            # Update progress
            progress_bar.progress((row - start_row + 1) / (total_rows - start_row + 1))

        status_text.text(f"✅ Finished! {found_count} courses found.")
        
        # Prepare for download
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="💾 Download Updated Excel",
            data=output,
            file_name=f"Checked_{target_term.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        driver.quit()

    except Exception as e:
        st.error(f"Error: {str(e)}")
        if 'driver' in locals():
            driver.quit()