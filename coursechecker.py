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

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="Course Availability Tool",
    page_icon="🔍",
    layout="centered"
)

# --- CUSTOM CSS ---
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        width: 100%;
        border-radius: 4px;
        height: 3em;
        background-color: #800000;
        color: white;
        font-weight: bold;
        border: none;
    }
    .stButton>button:hover {
        background-color: #a00000;
        color: white;
    }
    .status-box {
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #800000;
        background-color: #ffffff;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- HEADER ---
st.title("UChicago Course Scheduler Checker")
st.markdown("Please use required formatting: Department in **Column A**, Course Number in **Column B**.")
st.divider()

# --- MAIN INPUT AREA ---
uploaded_file = st.file_uploader("Upload course list (.xlsx)", type="xlsx")

input_col1, input_col2 = st.columns(2)
with input_col1:
    quarter = st.selectbox("Select Quarter", ["Spring", "Autumn", "Winter", "Summer"])
with input_col2:
    year = st.selectbox("Select Year", [2025, 2026, 2027, 2028])

target_term = f"{quarter} {year}"

st.markdown(" ") # Spacer
run_button = st.button("Run Availability Check")
st.divider()

# --- FUNCTIONS ---
def setup_headless_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    
    options.binary_location = "/usr/bin/chromium"
    service = Service(executable_path="/usr/bin/chromedriver")
    
    try:
        driver = webdriver.Chrome(service=service, options=options)
        driver.set_page_load_timeout(60)
        return driver
    except Exception:
        alt_service = Service(executable_path="/usr/lib/chromium-browser/chromedriver")
        return webdriver.Chrome(service=alt_service, options=options)

# --- EXECUTION LOGIC ---
if uploaded_file and run_button:
    file_bytes = uploaded_file.read()
    wb = load_workbook(filename=io.BytesIO(file_bytes))
    ws = wb.active 
    
    # Intelligent Header Detection
    first_cell_b = ws.cell(row=1, column=2).value
    start_row = 1
    try:
        if first_cell_b is not None:
            int(str(first_cell_b).strip().split('-')[0])
            start_row = 1
        else:
            start_row = 2 
    except (ValueError, TypeError):
        start_row = 2

    # Status Displays
    status_card = st.empty()
    progress_bar = st.progress(0)
    metric_col1, metric_col2 = st.columns(2)
    found_metric = metric_col1.empty()
    
    try:
        status_card.markdown('<div class="status-box">Launching server-side browser...</div>', unsafe_allow_html=True)
        driver = setup_headless_driver()
        wait = WebDriverWait(driver, 20)
        
        driver.get("http://coursesearch92.ais.uchicago.edu/psc/prd92guest/EMPLOYEE/HRMS/c/UC_STUDENT_RECORDS_FL.UC_CLASS_SEARCH_FL.GBL")
        
        status_card.markdown(f'<div class="status-box">Setting search term to: <b>{target_term}</b></div>', unsafe_allow_html=True)
        term_dropdown = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[id*='STRM']")))
        Select(term_dropdown).select_by_visible_text(target_term)
        time.sleep(2)
        
        total_rows = ws.max_row
        found_count = 0

        for row in range(start_row, total_rows + 1):
            subj = ws.cell(row=row, column=1).value 
            num = ws.cell(row=row, column=2).value  
            
            if not subj or not num:
                continue
            
            clean_num = str(num).strip().split('-')[0]
            query = f"{str(subj).strip()} {clean_num}"
            
            status_card.markdown(f'<div class="status-box">Checking: <b>{query}</b> (Row {row} of {total_rows})</div>', unsafe_allow_html=True)
            found_metric.metric("Courses Found", found_count)
            
            # PeopleSoft interaction
            search_bar = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input.ps-edit")))
            search_bar.click()
            search_bar.send_keys(Keys.CONTROL + "a")
            search_bar.send_keys(Keys.DELETE)
            search_bar.send_keys(query + Keys.ENTER)
            
            time.sleep(2)
            page_content = driver.find_element(By.TAG_NAME, "body").text.lower()
            
            # Logic for Column C (index 3)
            if "no results found" not in page_content and clean_num in page_content:
                ws.cell(row=row, column=3).value = "Y"
                found_count += 1
            else:
                ws.cell(row=row, column=3).value = None

            progress_val = (row - start_row + 1) / (total_rows - start_row + 1)
            progress_bar.progress(min(progress_val, 1.0))

        status_card.success(f"Check complete. {found_count} matches identified.")
        
        # Download preparation
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="Download Results",
            data=output,
            file_name=f"Checked_Schedule_{target_term.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        driver.quit()

    except Exception as e:
        st.error(f"Operation failed: {str(e)}")
        if 'driver' in locals():
            driver.quit()

elif not uploaded_file and run_button:
    st.warning("Please upload an Excel file to begin.")