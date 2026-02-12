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
This tool checks course availability on the UChicago Course Search website.
1. **Upload** an Excel file (.xlsx).
2. Ensure **Department** is in Column A and **Course Number** is in Column B.
3. The program will mark **Column L** with a 'Y' if the course is found.
""")

# --- INPUTS ---
uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")

col1, col2 = st.columns(2)
with col1:
    quarter = st.selectbox("Select Quarter", ["Spring", "Autumn", "Winter", "Summer"])
with col2:
    year = st.selectbox("Select Year", [2025, 2026, 2027])

target_term = f"{quarter} {year}"

def setup_headless_driver():
    """
    Configures Selenium to run on Streamlit Cloud's Linux environment.
    Requires 'chromium' and 'chromium-driver' in packages.txt.
    """
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    
    # Standard paths for Chromium on Streamlit Cloud
    options.binary_location = "/usr/bin/chromium"
    service = Service("/usr/bin/chromedriver")
    
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(60)
    return driver

if uploaded_file and st.button("🔍 Run Availability Check"):
    # Load the Excel file from the upload widget
    file_bytes = uploaded_file.read()
    wb = load_workbook(filename=io.BytesIO(file_bytes))
    ws = wb.active 
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    status_text.text("Starting browser...")
    driver = setup_headless_driver()
    wait = WebDriverWait(driver, 20)
    
    try:
        # Load the search page
        driver.get("http://coursesearch92.ais.uchicago.edu/psc/prd92guest/EMPLOYEE/HRMS/c/UC_STUDENT_RECORDS_FL.UC_CLASS_SEARCH_FL.GBL")
        
        # Set the Term
        status_text.text(f"Selecting term: {target_term}...")
        term_dropdown = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[id*='STRM']")))
        Select(term_dropdown).select_by_visible_text(target_term)
        
        # Wait for PeopleSoft spinner
        time.sleep(2)
        
        total_rows = ws.max_row
        found_count = 0

        # Loop through rows starting at Row 2
        for row in range(2, total_rows + 1):
            subj = ws.cell(row=row, column=1).value # Column A
            num = ws.cell(row=row, column=2).value  # Column B
            
            if not subj or not num:
                continue
            
            # Clean course number (remove section suffix like -01)
            clean_num = str(num).strip().split('-')[0]
            query = f"{str(subj).strip()} {clean_num}"
            
            status_text.text(f"Checking {query} (Row {row}/{total_rows})")
            
            # Find and clear search bar
            search_bar = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input.ps-edit")))
            search_bar.click()
            search_bar.send_keys(Keys.CONTROL + "a")
            search_bar.send_keys(Keys.DELETE)
            search_bar.send_keys(query + Keys.ENTER)
            
            # Wait for results to load
            time.sleep(2)
            page_text = driver.find_element(By.TAG_NAME, "body").text.lower()
            
            # Validation logic
            if "no results found" not in page_text and clean_num in page_text:
                ws.cell(row=row, column=12).value = "Y" # Column L
                found_count += 1
            else:
                ws.cell(row=row, column=12).value = None

            # Update progress
            progress_bar.progress((row - 1) / (total_rows - 1))

        status_text.text(f"✅ Complete! Found {found_count} courses.")
        
        # Save to a buffer for download
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="💾 Download Updated Excel File",
            data=output,
            file_name=f"UChicago_Courses_{target_term.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error during execution: {str(e)}")
    finally:
        driver.quit()