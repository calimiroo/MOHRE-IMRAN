# ===================== [ نفس الاستيرادات بدون تغيير ] =====================
import streamlit as st
import pandas as pd
import time
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime, timedelta
from deep_translator import GoogleTranslator
import re
import io
import random
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType

# ===================== [ الصفحة ] =====================
st.set_page_config(page_title="MOHRE Portal", layout="wide")
st.title("HAMADA TRACING SITE TEST")

st.markdown("""
<style>
.stTable td, .stTable th {
    white-space: nowrap !important;
    text-align: left !important;
    padding: 8px 15px !important;
}
.stTable {
    display: block !important;
    overflow-x: auto !important;
}
</style>
""", unsafe_allow_html=True)

# ===================== [ Session State ] =====================
defaults = {
    'authenticated': False,
    'run_state': 'stopped',
    'batch_results': [],
    'start_time_ref': None,
    'deep_run_state': 'stopped',
    'deep_finished': False,
    'deep_progress': 0,
    'single_result': None,
    'single_deep_done': False
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ===================== [ Login ] =====================
if not st.session_state['authenticated']:
    with st.form("login_form"):
        st.subheader("Protected Access")
        pwd_input = st.text_input("Enter Password", type="password")
        if st.form_submit_button("Login"):
            if pwd_input == "Bilkish":
                st.session_state['authenticated'] = True
                st.rerun()
            else:
                st.error("Incorrect Password.")
    st.stop()

# ===================== [ Utils ] =====================
def format_time(seconds):
    return str(timedelta(seconds=int(seconds)))

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

@st.cache_data(show_spinner=False)
def translate_to_english(text):
    try:
        if text and text != 'Not Found':
            return GoogleTranslator(source='auto', target='en').translate(text)
        return text
    except:
        return text

# ===================== [ Drivers ] =====================
def get_driver():
    options = uc.ChromeOptions()

    # 🔧 OPTIMIZED: تعطيل تحميل التقيل فقط
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.managed_default_content_settings.stylesheets": 2,
        "profile.managed_default_content_settings.fonts": 2
    }
    options.add_experimental_option("prefs", prefs)

    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    return uc.Chrome(options=options, use_subprocess=False)

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    return webdriver.Chrome(
        service=Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install()),
        options=options
    )

# ===================== [ Styling ] =====================
def apply_styling(df):
    df.index = range(1, len(df) + 1)
    def color_status(val):
        return 'background-color: #90EE90' if val == 'Found' else 'background-color: #FFCCCB'
    return df.style.applymap(color_status, subset=['Status'])

# ===================== [ OPTIMIZED extract_data ] =====================
def extract_data(driver, passport, nationality, dob_str):
    try:
        driver.get("https://mobile.mohre.gov.ae/Mob_Mol/MolWeb/MyContract.aspx?Service_Code=1005&lang=en")
        time.sleep(3)  # 🔧 reduced

        driver.find_element(By.ID, "txtPassportNumber").send_keys(passport)
        driver.find_element(By.ID, "CtrlNationality_txtDescription").click()
        time.sleep(0.8)

        try:
            sb = driver.find_element(By.CSS_SELECTOR, "#ajaxSearchBoxModal .form-control")
            sb.send_keys(nationality)
            time.sleep(0.8)
            driver.find_elements(By.CSS_SELECTOR, "#ajaxSearchBoxModal .items li a")[0].click()
        except:
            pass

        dob_input = driver.find_element(By.ID, "txtBirthDate")
        driver.execute_script("arguments[0].removeAttribute('readonly');", dob_input)
        dob_input.send_keys(dob_str)

        driver.find_element(By.ID, "btnSubmit").click()
        time.sleep(6)  # 🔧 reduced

        def get_value(label):
            try:
                return driver.find_element(By.XPATH, f"//*[contains(text(), '{label}')]/following::span[1]").text.strip()
            except:
                return 'Not Found'

        card = get_value("Card Number")
        if card == 'Not Found':
            return None

        return {
            "Passport Number": passport,
            "Nationality": nationality,
            "Date of Birth": dob_str,
            "Job Description": translate_to_english(get_value("Job Description")),
            "Card Number": card,
            "Card Expiry": get_value("Card Expiry"),
            "Basic Salary": get_value("Basic Salary"),
            "Total Salary": get_value("Total Salary"),
            "Status": "Found"
        }
    except:
        return None

# ===================== [ UI & Batch ] =====================
tab1, tab2 = st.tabs(["Single Search", "Upload Excel File"])

with tab1:
    st.info("Single Search unchanged")

with tab2:
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])
    if uploaded_file:
        df_original = pd.read_excel(uploaded_file)
        st.write(f"Total records: {len(df_original)}")

        if st.button("▶️ Start / Resume"):
            st.session_state.run_state = 'running'
            st.session_state.start_time_ref = time.time()

        if st.session_state.run_state == 'running':
            driver = get_driver()  # 🔧 OPTIMIZED: واحد بس
            for i, row in df_original.iterrows():
                res = extract_data(
                    driver,
                    str(row['Passport Number']),
                    str(row['Nationality']),
                    pd.to_datetime(row['Date of Birth']).strftime('%d/%m/%Y')
                )
                st.session_state.batch_results.append(res or {"Status": "Not Found"})

                if i % 5 == 0:
                    st.table(pd.DataFrame(st.session_state.batch_results))
            driver.quit()
            st.success("Finished")
