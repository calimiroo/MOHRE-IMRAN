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

# --- إعداد الصفحة --- 
st.set_page_config(page_title="MOHRE Portal", layout="wide") 
st.title("HAMADA TRACING SITE TEST") 

# --- تحسين مظهر الجدول وجعله سطر واحد (No Wrap) ---
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

# --- إدارة جلسة العمل (Session State) ---
if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False
if 'run_state' not in st.session_state:
    st.session_state['run_state'] = 'stopped'
if 'batch_results' not in st.session_state:
    st.session_state['batch_results'] = []
if 'start_time_ref' not in st.session_state:
    st.session_state['start_time_ref'] = None
if 'deep_run_state' not in st.session_state:
    st.session_state['deep_run_state'] = 'stopped'
if 'deep_finished' not in st.session_state:
    st.session_state['deep_finished'] = False
if 'deep_progress' not in st.session_state:
    st.session_state['deep_progress'] = 0
if 'single_result' not in st.session_state:
    st.session_state['single_result'] = None
if 'single_deep_done' not in st.session_state:
    st.session_state['single_deep_done'] = False

# قائمة الجنسيات
countries_list = ["Select Nationality", "Afghanistan", "Albania", "Algeria", "Andorra", "Angola", "Antigua and Barbuda", "Argentina", "Armenia", "Australia", "Austria", "Azerbaijan", "Bahamas", "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium", "Belize", "Benin", "Bhutan", "Bolivia", "Bosnia and Herzegovina", "Botswana", "Brazil", "Brunei", "Bulgaria", "Burkina Faso", "Burundi", "Cabo Verde", "Cambodia", "Cameroon", "Canada", "Central African Republic", "Chad", "Chile", "China", "Colombia", "Comoros", "Congo (Congo-Brazzaville)", "Costa Rica", "Côte d'Ivoire", "Croatia", "Cuba", "Cyprus", "Czechia (Czech Republic)", "Democratic Republic of the Congo", "Denmark", "Djibouti", "Dominica", "Dominican Republic", "Ecuador", "Egypt", "El Salvador", "Equatorial Guinea", "Eritrea", "Estonia", "Eswatini", "Ethiopia", "Fiji", "Finland", "France", "Gabon", "Gambia", "Georgia", "Germany", "Ghana", "Greece", "Grenada", "Guatemala", "Guinea", "Guinea-Bissau", "Guyana", "Haiti", "Holy See", "Honduras", "Hungary", "Iceland", "India", "Indonesia", "Iran", "Iraq", "Ireland", "Israel", "Italy", "Jamaica", "Japan", "Jordan", "Kazakhstan", "Kenya", "Kiribati", "Kuwait", "Kyrgyzstan", "Laos", "Latvia", "Lebanon", "Lesotho", "Liberia", "Libya", "Liechtenstein", "Lithuania", "Luxembourg", "Madagascar", "Malawi", "Malaysia", "Maldives", "Mali", "Malta", "Marshall Islands", "Mauritania", "Mauritius", "Mexico", "Micronesia", "Moldova", "Monaco", "Mongolia", "Montenegro", "Morocco", "Mozambique", "Myanmar", "Namibia", "Nauru", "Nepal", "Netherlands", "New Zealand", "Nicaragua", "Niger", "Nigeria", "North Korea", "North Macedonia", "Norway", "Oman", "Pakistan", "Palau", "Palestine State", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines", "Poland", "Portugal", "Qatar", "Romania", "Russia", "Rwanda", "Saint Kitts and Nevis", "Saint Lucia", "Saint Vincent and the Grenadines", "Samoa", "San Marino", "Sao Tome and Principe", "Saudi Arabia", "Senegal", "Serbia", "Seychelles", "Sierra Leone", "Singapore", "Slovakia", "Slovenia", "Solomon Islands", "Somalia", "South Africa", "South Korea", "South Sudan", "Spain", "Sri Lanka", "Sudan", "Suriname", "Sweden", "Switzerland", "Syria", "Tajikistan", "Tanzania", "Thailand", "Timor-Leste", "Togo", "Tonga", "Trinidad and Tobago", "Tunisia", "Turkey", "Turkmenistan", "Tuvalu", "Uganda", "Ukraine", "United Arab Emirates", "United Kingdom", "United States of America", "Uruguay", "Uzbekistan", "Vanuatu", "Venezuela", "Vietnam", "Yemen", "Zambia", "Zimbabwe"] 

# --- تسجيل الدخول ---
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

# --- دالة تحويل الوقت ---
def format_time(seconds):
    return str(timedelta(seconds=int(seconds)))

# --- دالة تحويل البيانات لملف اكسل للتحميل ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# --- وظائف الاستخراج والترجمة ---
def translate_to_english(text):
    try:
        if text and text != 'Not Found':
            return GoogleTranslator(source='auto', target='en').translate(text)
        return text
    except:
        return text

def get_driver():
    options = uc.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    return uc.Chrome(options=options, headless=True, use_subprocess=False)

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--lang=en-US")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install()), options=options)
    return driver

def apply_styling(df):
    df.index = range(1, len(df) + 1)
    def color_status(val):
        color = '#90EE90' if val == 'Found' else '#FFCCCB'
        return f'background-color: {color}'
    def color_expiry(val):
        try:
            expiry_date = datetime.strptime(str(val), '%d/%m/%Y')
            if expiry_date < datetime.now():
                return 'color: red'
        except:
            pass
        return ''
    return df.style.applymap(color_status, subset=['Status']).applymap(color_expiry, subset=['Card Expiry'])

def extract_data(passport, nationality, dob_str):
    driver = get_driver()
    try:
        driver.get("https://mobile.mohre.gov.ae/Mob_Mol/MolWeb/MyContract.aspx?Service_Code=1005&lang=en")
        time.sleep(4)
        driver.find_element(By.ID, "txtPassportNumber").send_keys(passport)
        driver.find_element(By.ID, "CtrlNationality_txtDescription").click()
        time.sleep(1)
        try:
            search_box = driver.find_element(By.CSS_SELECTOR, "#ajaxSearchBoxModal .form-control")
            search_box.send_keys(nationality)
            time.sleep(1)
            items = driver.find_elements(By.CSS_SELECTOR, "#ajaxSearchBoxModal .items li a")
            if items:
                items[0].click()
        except:
            pass
        dob_input = driver.find_element(By.ID, "txtBirthDate")
        driver.execute_script("arguments[0].removeAttribute('readonly');", dob_input)
        dob_input.clear()
        dob_input.send_keys(dob_str)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", dob_input)
        driver.find_element(By.ID, "btnSubmit").click()
        time.sleep(8)
        def get_value(label):
            try:
                xpath = f"//span[contains(text(), '{label}')]/following::span[1] | //label[contains(text(), '{label}')]/following-sibling::div"
                val = driver.find_element(By.XPATH, xpath).text.strip()
                return val if val else 'Not Found'
            except:
                return 'Not Found'
        card_num = get_value("Card Number")
        if card_num == 'Not Found':
            return None
        return {
            "Passport Number": passport,
            "Nationality": nationality,
            "Date of Birth": dob_str,
            "Job Description": translate_to_english(get_value("Job Description")),
            "Card Number": card_num,
            "Card Expiry": get_value("Card Expiry"),
            "Basic Salary": get_value("Basic Salary"),
            "Total Salary": get_value("Total Salary"),
            "Status": "Found"
        }
    except:
        return None
    finally:
        try:
            driver.quit()
        except:
            pass

def deep_extract_by_card(card_number):
    driver = setup_driver()
    try:
        driver.get("https://inquiry.mohre.gov.ae/")
        force_english(driver) 
        wait = WebDriverWait(driver, 20)
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Please select the service')]")))
        dropdown.click()
        time.sleep(1.5)
        service_opt = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Electronic Work Permit Information')]")))
        service_opt.click()
        input_box = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@placeholder='Enter data here']")))
        input_box.clear()
        input_box.send_keys(card_number)
        captcha_code = solve_captcha_using_your_script(driver)
        if captcha_code:
            captcha_field = driver.find_element(By.XPATH, "//input[contains(@placeholder, 'captcha')]")
            captcha_field.clear()
            captcha_field.send_keys(captcha_code)
        else:
            return None
        driver.find_element(By.XPATH, "//button[contains(text(), 'Search')]").click()
        time.sleep(5)
        if "No Data Found" in driver.page_source:
            return None
        else:
            comp_name = driver.find_element(By.XPATH, "//*[contains(text(), 'Est Name')]/..").text.replace("Est Name", "").replace(":", "").strip()
            cust_name = driver.find_element(By.XPATH, "//*[contains(text(), 'Name') and not(contains(text(), 'Est Name'))]/..").text.replace("Name", "").replace(":", "").strip()
            designation = driver.find_element(By.XPATH, "//*[contains(text(), 'Designation')]/..").text.replace("Designation", "").replace(":", "").strip()
            try:
                company_code = driver.find_element(By.XPATH, "//*[contains(text(), 'Company Code')]/..").text.replace("Company Code", "").replace(":", "").strip()
            except:
                company_code = 'Not Found'
            return {
                'Name': cust_name if cust_name else 'Not Found',
                'Est Name': comp_name if comp_name else 'Not Found',
                'Company Code': company_code if company_code else 'Not Found',
                'Job_Deep': designation 
            }
    except:
        return None
    finally:
        try:
            driver.quit()
        except:
            pass

def force_english(driver):
    try:
        wait = WebDriverWait(driver, 10)
        lang_btn = wait.until(EC.presence_of_element_located((By.ID, "btnlanguage")))
        if "English" in lang_btn.text:
            driver.execute_script("arguments[0].click();", lang_btn)
            wait.until(EC.text_to_be_present_in_element((By.TAG_NAME, "body"), "Please select the service"))
            time.sleep(2)
    except:
        pass

def solve_captcha_using_your_script(driver):
    try:
        elements = driver.find_elements(By.XPATH, "//div | //span | //b | //strong | //p")
        for el in elements:
            text = el.text.strip()
            if re.match(r'^\d{4}$', text):
                return text
    except:
        pass
    return None

# --- واجهة المستخدم ---
tab1, tab2 = st.tabs(["Single Search", "Upload Excel File"]) 

with tab1:
    st.subheader("Single Person Search")
    c1, c2, c3 = st.columns(3)
    p_in = c1.text_input("Passport Number", key="s_p")
    n_in = c2.selectbox("Nationality", countries_list, key="s_n")
    d_in = c3.date_input("Date of Birth", value=None, min_value=datetime(1900,1,1), format="DD/MM/YYYY", key="s_d")
    
    if st.button("Search Now", key="single_search_button"):
        if p_in and n_in != "Select Nationality" and d_in:
            with st.spinner("Searching..."):
                res = extract_data(p_in, n_in, d_in.strftime("%d/%m/%Y"))
                if res:
                    st.session_state.single_result = res
                    st.session_state.single_deep_done = False
                else:
                    st.error("No data found.")
                    st.session_state.single_result = None

    single_table_area = st.empty()

    if st.session_state.single_result:
        current_df = pd.DataFrame([st.session_state.single_result])
        for col in ['Name', 'Est Name', 'Company Code']:
            if col not in current_df.columns: current_df[col] = ''
        styled_df = apply_styling(current_df)
        single_table_area.table(styled_df)

        if st.session_state.single_result.get('Status') == 'Found' and not st.session_state.single_deep_done:
            if st.button("Run Deep Search"):
                with st.spinner("Deep Searching..."):
                    deep_res = deep_extract_by_card(st.session_state.single_result['Card Number'])
                    if deep_res:
                        st.session_state.single_result['Job Description'] = deep_res.get('Job_Deep', 'Not Found')
                        st.session_state.single_result['Name'] = deep_res.get('Name', 'Not Found')
                        st.session_state.single_result['Est Name'] = deep_res.get('Est Name', 'Not Found')
                        st.session_state.single_result['Company Code'] = deep_res.get('Company Code', 'Not Found')
                    else:
                        st.session_state.single_result['Name'] = 'Not Found'
                    st.session_state.single_deep_done = True
                current_df = pd.DataFrame([st.session_state.single_result])
                styled_df = apply_styling(current_df)
                single_table_area.table(styled_df)

with tab2:
    st.subheader("Batch Processing Control")
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"]) 
    
    if uploaded_file:
        df_original = pd.read_excel(uploaded_file)
        df_show = df_original.copy()
        df_show.index = range(1, len(df_show) + 1)
        st.write(f"Total records: {len(df_original)}")
        st.dataframe(df_show, height=150, use_container_width=True)
        
        col_ctrl1, col_ctrl2, col_ctrl3 = st.columns(3)
        if col_ctrl1.button("▶️ Start / Resume"):
            st.session_state.run_state = 'running'
            if st.session_state.start_time_ref is None: st.session_state.start_time_ref = time.time()
        if col_ctrl2.button("⏸️ Pause"): st.session_state.run_state = 'paused'
        if col_ctrl3.button("⏹️ Stop & Reset"):
            st.session_state.run_state = 'stopped'
            st.session_state.batch_results = []
            st.session_state.start_time_ref = None
            st.session_state.deep_finished = False
            st.rerun()

        progress_bar = st.progress(0)
        status_text = st.empty()
        stats_area = st.empty()
        live_table_area = st.empty()
        deep_status_area = st.empty()
        deep_progress_bar = st.empty()
        actual_success = 0
        
        for i, row in df_original.iterrows():
            while st.session_state.run_state == 'paused':
                status_text.warning("Paused...")
                time.sleep(1)
            if st.session_state.run_state == 'stopped': break
            
            if i < len(st.session_state.batch_results):
                if st.session_state.batch_results[i].get("Status") == "Found": actual_success += 1
                current_df = pd.DataFrame(st.session_state.batch_results)
                live_table_area.table(apply_styling(current_df))
                progress_bar.progress((i + 1) / len(df_original))
                continue

            p_num = str(row.get('Passport Number', '')).strip()
            nat = str(row.get('Nationality', 'Egypt')).strip()
            try: dob = pd.to_datetime(row.get('Date of Birth')).strftime('%d/%m/%Y')
            except: dob = str(row.get('Date of Birth', ''))
                
            status_text.info(f"Processing {i+1}/{len(df_original)}: {p_num}")
            res = extract_data(p_num, nat, dob)
            
            if res:
                actual_success += 1
                st.session_state.batch_results.append(res)
            else:
                st.session_state.batch_results.append({
                    "Passport Number": p_num, "Nationality": nat, "Date of Birth": dob,
                    "Job Description": "N/A", "Card Number": "N/A", "Card Expiry": "N/A",
                    "Basic Salary": "N/A", "Total Salary": "N/A", "Status": "Not Found"
                })
                
            elapsed = time.time() - st.session_state.start_time_ref if st.session_state.start_time_ref else 0
            stats_area.markdown(f"✅ **Success:** {actual_success} | ⏱️ **Time:** {format_time(elapsed)}")
            
            current_df = pd.DataFrame(st.session_state.batch_results)
            live_table_area.table(apply_styling(current_df))
            progress_bar.progress((i + 1) / len(df_original))

        # --- زر تحميل المرحلة الأولى ---
        if len(st.session_state.batch_results) == len(df_original) and len(df_original) > 0:
            st.success("Stage 1 Finished!")
            
            df_stage1 = pd.DataFrame(st.session_state.batch_results)
            excel_stage1 = to_excel(df_stage1)
            st.download_button(
                label="📥 Download Stage 1 Results",
                data=excel_stage1,
                file_name=f"stage1_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_stage1"
            )

            if not st.session_state.deep_finished:
                if st.button("🚀 Run Deep Search (Stage 2)"):
                    st.session_state.deep_run_state = 'running'
            
            if st.session_state.deep_run_state == 'running':
                deep_recs = [r for r in st.session_state.batch_results if r.get('Status') == 'Found']
                deep_total = len(deep_recs)
                deep_idx = 0
                for idx, rec in enumerate(st.session_state.batch_results):
                    if rec.get('Status') != 'Found': continue
                    card = rec.get('Card Number')
                    deep_status_area.info(f"Deep Searching {deep_idx+1}/{deep_total}: {card}")
                    deep_res = deep_extract_by_card(card)
                    if deep_res:
                        st.session_state.batch_results[idx]['Name'] = deep_res.get('Name')
                        st.session_state.batch_results[idx]['Est Name'] = deep_res.get('Est Name')
                        st.session_state.batch_results[idx]['Company Code'] = deep_res.get('Company Code')
                        st.session_state.batch_results[idx]['Job Description'] = deep_res.get('Job_Deep')
                    deep_idx += 1
                    deep_progress_bar.progress(deep_idx / deep_total)
                    live_table_area.table(apply_styling(pd.DataFrame(st.session_state.batch_results)))
                    time.sleep(2)
                st.success("Deep Search Completed!")
                st.session_state.deep_run_state = 'stopped'
                st.session_state.deep_finished = True
                st.rerun() # تحديث الصفحة لإظهار الزر الثابت

            # --- زر تحميل المرحلة الثانية (البحث العميق) الثابت ---
            if st.session_state.deep_finished:
                df_stage2 = pd.DataFrame(st.session_state.batch_results)
                excel_stage2 = to_excel(df_stage2)
                st.download_button(
                    label="📥 Download Deep Search Results",
                    data=excel_stage2,
                    file_name=f"deep_search_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_stage2"
                )
