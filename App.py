import streamlit as st
import pandas as pd
import time
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
from deep_translator import GoogleTranslator
import re
import io
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType

# ================= PAGE CONFIG =================
st.set_page_config(page_title="MOHRE Portal", layout="wide")
st.title("HAMADA TRACING SITE TEST")

# ================= STYLE =================
st.markdown("""
<style>
.stTable td, .stTable th {
    white-space: nowrap !important;
    padding: 8px 15px !important;
}
.stTable {
    overflow-x: auto !important;
}
</style>
""", unsafe_allow_html=True)

# ================= SESSION =================
for k, v in {
    "authenticated": False,
    "run_state": "stopped",
    "batch_results": [],
    "start_time_ref": None,
    "deep_run_state": "stopped",
    "deep_finished": False,
    "single_result": None,
    "single_deep_done": False
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ================= LOGIN =================
if not st.session_state.authenticated:
    with st.form("login"):
        pwd = st.text_input("Enter Password", type="password")
        if st.form_submit_button("Login"):
            if pwd == "Bilkish":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Wrong password")
    st.stop()

# ================= UTILS =================
def format_time(sec): return str(timedelta(seconds=int(sec)))

def to_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

@st.cache_data(show_spinner=False)
def translate_cached(text):
    try:
        return GoogleTranslator(source='auto', target='en').translate(text)
    except:
        return text

# ================= DRIVER =================
def get_driver():
    options = uc.ChromeOptions()
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.managed_default_content_settings.stylesheets": 2,
        "profile.managed_default_content_settings.fonts": 2
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    return uc.Chrome(options=options, use_subprocess=False)

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    return webdriver.Chrome(
        service=Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install()),
        options=options
    )

# ================= MAIN EXTRACT =================
def extract_data(driver, passport, nationality, dob):
    wait = WebDriverWait(driver, 15)
    driver.get("https://mobile.mohre.gov.ae/Mob_Mol/MolWeb/MyContract.aspx?Service_Code=1005&lang=en")

    wait.until(EC.presence_of_element_located((By.ID, "txtPassportNumber"))).send_keys(passport)
    driver.find_element(By.ID, "CtrlNationality_txtDescription").click()

    try:
        sb = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#ajaxSearchBoxModal .form-control")))
        sb.send_keys(nationality)
        time.sleep(0.8)
        driver.find_elements(By.CSS_SELECTOR, "#ajaxSearchBoxModal .items li a")[0].click()
    except:
        pass

    dob_el = driver.find_element(By.ID, "txtBirthDate")
    driver.execute_script("arguments[0].removeAttribute('readonly');", dob_el)
    dob_el.send_keys(dob)

    driver.find_element(By.ID, "btnSubmit").click()
    time.sleep(4)

    def get_val(label):
        try:
            return driver.find_element(By.XPATH, f"//*[contains(text(),'{label}')]/following::span[1]").text
        except:
            return "Not Found"

    card = get_val("Card Number")
    if card == "Not Found":
        return None

    return {
        "Passport Number": passport,
        "Nationality": nationality,
        "Date of Birth": dob,
        "Job Description": translate_cached(get_val("Job Description")),
        "Card Number": card,
        "Card Expiry": get_val("Card Expiry"),
        "Basic Salary": get_val("Basic Salary"),
        "Total Salary": get_val("Total Salary"),
        "Status": "Found"
    }

# ================= UI =================
tab1, tab2 = st.tabs(["Single Search", "Upload Excel File"])

with tab1:
    p = st.text_input("Passport Number")
    n = st.text_input("Nationality")
    d = st.date_input("DOB", format="DD/MM/YYYY")

    if st.button("Search"):
        driver = get_driver()
        with st.spinner("Searching..."):
            res = extract_data(driver, p, n, d.strftime("%d/%m/%Y"))
        driver.quit()

        if res:
            st.table(pd.DataFrame([res]))
        else:
            st.error("Not Found")

with tab2:
    file = st.file_uploader("Upload Excel", type=["xlsx"])
    if file:
        df = pd.read_excel(file)
        st.write(f"Records: {len(df)}")

        if st.button("▶️ Start"):
            st.session_state.run_state = "running"
            st.session_state.start_time_ref = time.time()

        if st.session_state.run_state == "running":
            driver = get_driver()
            for i, r in df.iterrows():
                res = extract_data(
                    driver,
                    str(r["Passport Number"]),
                    str(r["Nationality"]),
                    pd.to_datetime(r["Date of Birth"]).strftime("%d/%m/%Y")
                )
                st.session_state.batch_results.append(
                    res if res else {
                        "Passport Number": r["Passport Number"],
                        "Status": "Not Found"
                    }
                )

                if i % 5 == 0:
                    st.table(pd.DataFrame(st.session_state.batch_results))

            driver.quit()
            st.success("Finished")
            st.download_button(
                "Download",
                to_excel(pd.DataFrame(st.session_state.batch_results)),
                "results.xlsx"
            )
