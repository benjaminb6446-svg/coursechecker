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
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook

# --- UI SETUP ---
st.title("🎓 UChicago Course Scheduler Checker")
st.markdown("""
1. **Upload** your course list (.xlsx)
2. **Select** the target Quarter/Year
3. **Download** the updated file with 'Y' markers
""")

# --- USER INPUTS ---
uploaded_file = st.file_uploader("Upload Excel File (Columns: A=Dept, B=Number)", type="xlsx")

col1, col2 = st.columns(2)
with col1:
    quarter = st.selectbox("Select Quarter", ["Spring", "Autumn", "Winter", "Summer"])
with col2:
    year = st.selectbox("Select Year", [2025, 2026, 2027])

target_term = f"{quarter} {year}"

def setup_headless_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    
    # This installs Chrome on the Streamlit Server
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    
    # CORRECT PLACE FOR TIMEOUT
    driver.set_page_load_timeout(60) 
    return driver

if uploaded_file and st.button("🔍 Check Courses"):
    wb = load_workbook(filename=io.BytesIO(uploaded_file.read()))
    ws = wb.active 
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    driver = setup_headless_driver()
    wait = WebDriverWait(driver, 15)
    
    try:
        driver.get("http://coursesearch92.ais.uchicago.edu/psc/prd92guest/EMPLOYEE/HRMS/c/UC_STUDENT_RECORDS_FL.UC_CLASS_SEARCH_FL.GBL")
        
        status_text.text(f"Setting term to {target_term}...")
        term_dropdown = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[id*='STRM']")))
        Select(term_dropdown).select_by_visible_text(target_term)
        time.sleep(2)

        rows_to_process = ws.max_row
        for row in range(2, rows_to_process + 1):
            subj = ws.cell(row=row, column=1).value
            num = ws.cell(row=row, column=2).value
            
            if not subj or not num:
                continue
            
            query = f"{str(subj).strip()} {str(num).strip().split('-')[0]}"
            status_text.text(f"Searching: {query} ({row-1}/{rows_to_process-1})")
            
            # Use ActionChains style or direct send_keys for headless
            search_bar = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input.ps-edit")))
            search_bar.click()
            search_bar.send_keys(Keys.CONTROL + "a") # Use Control for Linux servers
            search_bar.send_keys(Keys.DELETE)
            search_bar.send_keys(query + Keys.ENTER)
            
            time.sleep(2)
            page_text = driver.find_element(By.TAG_NAME, "body").text.lower()
            
            if "no results found" not in page_text and str(num).strip().split('-')[0] in page_text:
                ws.cell(row=row, column=12).value = "Y"
            else:
                ws.cell(row=row, column=12).value = None

            progress_bar.progress((row - 1) / (rows_to_process - 1))

        status_text.text("✅ Finished!")
        
        output = io.BytesIO()
        wb.save(output)
        
        st.download_button(
            label="💾 Download Updated Excel",
            data=output.getvalue(),
            file_name=f"Courses_{target_term.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
    finally:
        driver.quit()