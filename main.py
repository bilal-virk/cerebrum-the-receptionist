import os, sys, time, logging, csv
import time, os, re
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from openai import OpenAI
import traceback
import psutil
import subprocess
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import unicodedata
from dotenv import load_dotenv

load_dotenv()
senders_email = os.getenv("SENDERS_EMAIL")
port = int(os.getenv("PORT"))
server = os.getenv("SERVER")
password = os.getenv("PASSWORD")

def patient_not_found(name, dob, check_in_time):
    receivers = ['receptionist@drsingla.ca', 'Mohit.singla.md@gmail.com']
    for receiver in receivers:
        msg = MIMEMultipart()
        msg['From'] = senders_email
        msg['To'] = receiver
        msg['Subject'] = f'{name} not found in Cerebrum'
        body = f'Patient {name} with Date of birth: {dob} not found in Cerebrum. Alert from Python Bot.'
        msg.attach(MIMEText(body, 'plain'))
        try:
            with smtplib.SMTP(server, port) as smtp:
                smtp.starttls()
                smtp.login(senders_email, password)
                smtp.send_message(msg)
                print("Email sent successfully!")
        except Exception as e:
            print(f"Failed to send email: {e}")
        save_record(name, dob)
        save_record(name, check_in_time)
script_directory = os.getcwd()
downloads_folder = os.path.join(script_directory, "Downloads")
RECORDS_FILE = os.path.join(script_directory, "records.txt")
RECORDS_FILE_LATE = os.path.join(script_directory, "late_records.txt")
opennai_api_key = os.getenv("OPENAI_API_KEY")
logger = logging.getLogger("customLogger")
logger.setLevel(logging.INFO)
file_handler = logging.FileHandler(os.path.join(script_directory,'App.log'), encoding="utf-8")
file_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)
def pwrite(*args, p=True):
    message = " ".join(str(arg) for arg in args)
    message = f'{message}'
    logger.info(message)
    if p:
        print(message)
def ai_address_match(address_input, postal_input, address_table, city_expected, postal_expected):
    system_prompt = """
    You are an expert Canadian address verifier.
    All addresses are in Ontario, Canada.
    Decide if two addresses refer to the same physical location.
    """

    user_prompt = f"""
    Compare the following two addresses.

    Guidelines:
    - Priority: STREET (house number + street name) → most important for matching.
    - CITY: If present, use it. If missing in the patient’s input, do not reject. 
    - POSTAL: Can be invalid, missing, or mistyped. If it matches, it helps confirm, 
    but if it does not match, do not automatically reject.
    - Abbreviations (St. = Street, Rd = Road, Ave = Avenue, Apt = Apartment) 
    and minor typos should still match.
    - If STREET is clearly the same and CITY is compatible (or absent in input), lean toward YES.
    - Only return NO if the street or city clearly refers to a different place.
    - Reduce confidence threshold: if it’s a close call, lean toward YES.

    Output strictly one word: YES or NO. Do not explain.

    Typed by patient:
    Address: {address_input}
    Postal: {postal_input}

    Reference (staff entry):
    Address: {address_table}
    City: {city_expected}
    Postal: {postal_expected}
    """
    client = OpenAI(api_key=opennai_api_key)
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0,
        max_tokens=10
    )
    pwrite(f"AI response: {response.choices[0].message.content.strip()}", p=True)
    if response.choices[0].message.content.strip().upper() == "YES":
        return True
    else:
        return False
def normalize_text(text: str) -> str:
    if not text:
        return ""
    
    text = unicodedata.normalize("NFKC", text)
    text = text.strip().replace("\u200b", "").replace("\ufeff", "")  # remove BOM & zero-width space
    return text.lower()
def record_exists(name: str, dob: str) -> bool:
    if not os.path.exists(RECORDS_FILE):
        pwrite('Record file not found')
        return False

    name = str(normalize_text(name)).strip().lower()
    dob = str(normalize_text(dob)).strip()

    with open(RECORDS_FILE, newline='', encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row["name"].strip().lower() == name and row["dob"].strip() == dob:
                return True
    return False


def save_record(name: str, dob: str):
    file_exists = os.path.exists(RECORDS_FILE)

    # ✅ Normalize before saving
    name = str(normalize_text(name)).strip().lower()
    dob = str(normalize_text(dob)).strip()

    with open(RECORDS_FILE, "a", newline='', encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["name", "dob"])
        if not file_exists:
            writer.writeheader()
        writer.writerow({"name": name, "dob": dob})

def record_exists_late(name: str, dob: str) -> bool:
    if not os.path.exists(RECORDS_FILE_LATE):
        pwrite('Record file not found')
        return False

    name = str(normalize_text(name)).strip().lower()
    dob = str(normalize_text(dob)).strip()

    with open(RECORDS_FILE_LATE, newline='', encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row["name"].strip().lower() == name and row["dob"].strip() == dob:
                return True
    return False


def save_record_late(name: str, dob: str):
    file_exists = os.path.exists(RECORDS_FILE_LATE)

    # ✅ Normalize before saving
    name = str(normalize_text(name)).strip().lower()
    dob = str(normalize_text(dob)).strip()

    with open(RECORDS_FILE_LATE, "a", newline='', encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["name", "dob"])
        if not file_exists:
            writer.writeheader()
        writer.writerow({"name": name, "dob": dob})

chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
script_folder = os.path.dirname(os.path.abspath(__file__))
user_data_dir = os.path.join(script_folder, "chrome")
custom_port = 9233

cmd = [
    chrome_path,
    f"--remote-debugging-port={custom_port}",
    f"--user-data-dir={user_data_dir}",
    "--disable-popup-blocking"
]
def main(daytime=True, test=False):
    def is_chrome_running(port):
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                if proc.info['name'] and "chrome" in proc.info['name'].lower():
                    if any(f"--remote-debugging-port={port}" in arg for arg in proc.info['cmdline']):
                        return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return False

    if not is_chrome_running(custom_port):
        print("Starting Chrome...")
        subprocess.Popen(cmd)
    

    thereceptionist_username = os.getenv("THERECEPTIONIST_USERNAME")
    thereceptionist_password = os.getenv("THERECEPTIONIST_PASSWORD")
    cerebrum_username = os.getenv("CEREBRUM_USERNAME")
    cerebrum_password = os.getenv("CEREBRUM_PASSWORD")
    o_day_m = os.getenv("CONFIG_MONDAY")
    o_day_t = os.getenv("CONFIG_TUESDAY")
    o_day_w = os.getenv("CONFIG_WEDNESDAY")
    o_day_th = os.getenv("CONFIG_THURSDAY")
    o_day_f = os.getenv("CONFIG_FRIDAY")


    


                
    ## Writing Text In Element
    def write_t(xpathe, text, t=10, sleep_time=None, p=False):
        WebDriverWait(driver, t).until(EC.presence_of_element_located((By.XPATH, xpathe)))
        if sleep_time is not None:
                time.sleep(sleep_time)
        element = WebDriverWait(driver, t).until(EC.presence_of_element_located((By.XPATH, xpathe)))
        element.clear()
        if p:
            element.send_keys(text)
        else:
            for char in text:
                element.send_keys(char)
                time.sleep(0.05)  # Adjust typing speed here if needed
        
    ## Extracting Text
    def extract_text(xpathe, t=10):
        element_text = WebDriverWait(driver, t).until(EC.presence_of_element_located((By.XPATH, xpathe))).text
        return element_text.strip()
    def do_check_in(name, check_in_date):
        pass

    

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option('debuggerAddress', 'localhost:9233')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    if len(driver.window_handles) > 2:
        try:
            target_domains = ["cerebrum", "thereceptionist"]
            kept_tabs = {}  # store the first tab found for each domain

            # Go through all open tabs
            for handle in driver.window_handles[:]:
                driver.switch_to.window(handle)
                current_url = driver.current_url.lower()

                matched_domain = None
                for domain in target_domains:
                    if domain in current_url:
                        matched_domain = domain
                        break

                # If the tab matches one of the target domains
                if matched_domain:
                    # Keep only the first tab for this domain
                    if matched_domain not in kept_tabs:
                        kept_tabs[matched_domain] = handle
                    else:
                        driver.close()
                else:
                    # Close all other irrelevant tabs
                    driver.close()
        except:
            pass
    def make_click(xpathe, driver=driver, t=10, sleep_time=None):

        element =WebDriverWait(driver, t).until(EC.presence_of_element_located((By.XPATH, xpathe)))
        if sleep_time is not None:
                time.sleep(sleep_time)
        try:
            element =WebDriverWait(driver, t).until(EC.element_to_be_clickable((By.XPATH, xpathe)))
        except:
            driver.execute_script("arguments[0].click();", element)
        try:
            try:
                element.click()            
            except:                
                element.send_keys(Keys.ENTER)
        except:
            try:
                driver.execute_script("arguments[0].scrollIntoView({ behavior: 'auto', block: 'center', inline: 'center' });", element)
                time.sleep(1)
                element.click()
            except:
                driver.execute_script("arguments[0].click();", element)

    found = False
    try:
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            if "cerebrum" in driver.current_url:
                driver.refresh()
                found = True
                break
    except:
        pass
    if not found:
        # driver.execute_script("window.open('');")
        # driver.switch_to.window(driver.window_handles[-1])
        driver.get("https://cerebrum.mycerebrum.com/Account/Login?")

    try:
        write_t('//*[@id="Email"]', cerebrum_username, sleep_time=2, t=3)
        write_t('//*[@id="Password"]', cerebrum_password, sleep_time=2)
        make_click('//*[@type="submit"]', t=10, sleep_time=2)
        #time.sleep(5)
        try:
            make_click('//*[@value="Return to main"]')
        except:
            pass
    except:
        pwrite("Cerebrum Already Logged In")



    found = False
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if "thereceptionist" in driver.current_url:
            driver.refresh()
            found = True
            break

    if not found:
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])
        driver.get("https://app.thereceptionist.com/sign_in")
    try:
        driver.get('https://app.thereceptionist.com/visits')
        write_t('//*[@type="email"]', thereceptionist_username, sleep_time=2, t=3)
        write_t('//*[@type="password"]', thereceptionist_password, sleep_time=2)
        make_click('//*[@type="submit"]', t=10, sleep_time=2)
        time.sleep(5)
    except:
        pwrite("Thereceptionist Already Logged In")
    if daytime:
        if not test:
            make_click('//*[@id="date-range"]', t=10)
            make_click('//*[@data-range-key="Today"]', t=10, sleep_time=2)
        Select(driver.find_element(By.XPATH, '(//select)[1]')).select_by_visible_text('Check In')
        make_click('//button[@type="submit"]')
        
        time.sleep(5)

        try:
            leads = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="visits"]//tr')))
        except:
            leads = []
        nowe = datetime.now()
        today_date = nowe.strftime("%Y-%m-%d")
        for lead in leads[:5]:
            try:
                for handle in driver.window_handles:
                    driver.switch_to.window(handle)
                    if "thereceptionist" in driver.current_url:
                        break
                try:
                    make_click('//*[@class="close"]', t=1)
                except:
                    pass
                #data_timestamp = WebDriverWait(lead, 5).until(EC.presence_of_element_located((By.XPATH, './/*[contains(text(), " at ") and contains(text(), "-")]'))).text.strip()
                try:
                    status = WebDriverWait(lead, 5).until(EC.presence_of_element_located((By.XPATH, './/td[3]//div[@title]//span')))
                except:
                    try:
                        make_click('//*[@class="close"]', t=1)
                    except:
                        pass
                check_in_time = lead.find_element(By.XPATH, './/td[6]//div[@title]//span').text.strip()
                status = lead.find_element(By.XPATH, './/td[3]//div[@title]//span').text.strip()
                name = lead.find_element(By.XPATH, './/*[@class="ellipsis text-bold"]').text.strip()
                if 'Sachleen kaur' in name:
                    continue
                check_in_time = f'{check_in_time}||{today_date}'
                if record_exists(name, check_in_time):
                    pwrite(f'Skipping already processed patient: {name} {check_in_time}')
                    continue
                
                last_name= name.split()[-1]
                try:
                    make_click('.//*[@class="ellipsis text-bold"]',driver=lead, t=10)
                except:
                    element = lead.find_element(By.XPATH, './/*[@class="ellipsis text-bold"]')
                    driver.execute_script("arguments[0].click();", element)
                time.sleep(1)
                
                #if not record_exists_late(name, check_in_time):
                if True:
                    if status == "Check In":
                        dob = extract_text('//*[contains(text(), "Date of Birth")]/../..//*[@class="text-right"]')
                        if record_exists(name, check_in_time) or record_exists(name, dob):
                            pwrite(f'Skipping already processed patient: {name} {dob}')
                            continue
                        last_two = extract_text('//*[contains(text(), "Last two")]/../..//*[@class="text-right"]')
                        #phone = extract_text('//*[contains(text(), "hone number")]/../..//*[@class="text-right"]')
                        try:
                            phone = extract_text('//a[contains(@href, "tel")]')
                        except:
                            phone = None
                        address = extract_text('//*[contains(text(), "street address")]/../..//*[@class="text-right"]')
                        city = extract_text('//*[contains(text(), "City/Town")]/../..//*[@class="text-right"]')
                        family_doc = extract_text('//*[contains(text(), "FULL Name of family doctor")]/../..//*[@class="text-right"]')
                        postal_code_thereceptionist = extract_text('//*[contains(text(), "postal code")]/../..//*[@class="text-right"]')
                        email = extract_text('//a[contains(@href, "mailto")]')
                        make_click('(//*[@class="close" and @data-dismiss="modal"])[1]', t=10)

                        # Months mapping
                        months = {
                            "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
                            "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
                            "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"
                        }
                        parts = dob.split("-")
                        newDate = f"{months[parts[1]]}-{parts[0]}-{parts[2]}"
                        
                        for handle in driver.window_handles:
                            driver.switch_to.window(handle)
                            if "cerebrum" in driver.current_url:
                                break
                        # write_t('//*[@placeholder="Patient Search"]', last_name)
                        # time.sleep(3)
                        # make_click('//*[@id="btn_search"]')
                        #driver.get(f'https://cerebrum.mycerebrum.com/patients/patientsearch?searchType=0&patientSearch={last_name}')
                        DAY_ENV_MAP = {
                            "Monday": "CONFIG_MONDAY",
                            "Tuesday": "CONFIG_TUESDAY",
                            "Wednesday": "CONFIG_WEDNESDAY",
                            "Thursday": "CONFIG_THURSDAY",
                            "Friday": "CONFIG_FRIDAY",
                        }
                        dt = datetime.now()
                        check_in_date = dt.strftime("%m/%d/%Y")
                        check_in_day = dt.strftime("%A")
                        env_key = DAY_ENV_MAP.get(check_in_day)
                        clinic_name = os.getenv(env_key) if env_key else None
                        d_t = check_in_date.replace("/", "%2F")
                        if test:
                            d_t= "12%2F19%2F2025"
                        for handle in driver.window_handles:
                                driver.switch_to.window(handle)
                                if "cerebrum" in driver.current_url:
                                    break

                        #c_url = f'https://cerebrum.mycerebrum.com/Schedule/daysheet?Date={d_t}&FilterPatient={last_name}'
                        c_url = f'https://cerebrum.mycerebrum.com/Schedule/daysheet?OfficeId=30&Date={d_t}&AppointmentStatusId=-1&Expected=False&ExcludeTestOnly=False&ExcludeCancelled=True&OnlyActionOnAbnormal=False&FilterPatient={last_name}&ShowOrders=False&Page=1&PageSize=25'
                        driver.get(c_url)
                        # make_click('//input[@name="Date" and @type="text"]', sleep_time=2)
                        # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//input[@name="Date" and @type="text"]'))).clear()
                        # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//input[@name="Date" and @type="text"]'))).send_keys(check_in_date)
                        if clinic_name:
                            dropdown = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, "OfficeId"))
                            )
                            Select(dropdown).select_by_visible_text(clinic_name)
                        #write_t('//*[@placeholder="Filter by patient"]', last_name)
                        #WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@placeholder="Filter by patient"]'))).send_keys(last_name)
                        first_name = name.split()[0]
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//*[@class="td-patient"]//*[contains(text(), "{last_name.upper()}") and contains(text(), "{first_name.upper()}")]')))
                            try:
                                make_click(f'//*[@class="td-patient"]//*[contains(text(), "{last_name.upper()}") and contains(text(), "{first_name.upper()}")]/../../../..//*[@data-cb-tp-title="Set Arrival Time"]')
                                time.sleep(5)
                            except:
                                patients = driver.find_elements(By.XPATH, f'//*[@class="td-patient"]//*[contains(text(), "{last_name.upper()}") and contains(text(), "{first_name.upper()}")]/../../../..//*[@data-cb-tp-title="Set Arrival Time"]')
                                # if len(patients) > 0:

                                pwrite(f'Error While clicking Arrived Button for {name} -- DOB: {dob} -- Patients Found: {len(patients)}\n {traceback.format_exc()}')
                                driver.find_element(By.TAG_NAME, 'body').screenshot(os.path.join(script_directory, 'Images', f'{name}.png'))
                                pass
                            save_record(name, check_in_time)
                        except:
                            continue
                            save_record_late(name, check_in_time)
                        
                        d_t = check_in_date.replace("/", "%2F")
                        # c_url = f'https://cerebrum.mycerebrum.com/Schedule/daysheet?OfficeId=30&Date={d_t}&AppointmentStatusId=-1&Expected=False&ExcludeTestOnly=False&ExcludeCancelled=True&OnlyActionOnAbnormal=False&FilterPatient={last_name}&ShowOrders=False&Page=1&PageSize=25'
                        # driver.get(c_url)

                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//*[@class="td-patient"]//*[contains(text(), "{last_name.upper()}") and contains(text(), "{first_name.upper()}")]')))
                        except:
                            # try:
                            #     driver.get(f'https://cerebrum.mycerebrum.com/patients/patientsearch?searchType=0&patientSearch={first_name}')
                            #     WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//*[contains(text(), "{newDate}")]')))
                            # except:
                            pwrite(f'Patient Not Found {name} {dob} {address} {last_two} {phone} {email}')
                            continue
                        make_click(f'//*[@class="td-patient"]//*[contains(text(), "{last_name.upper()}") and contains(text(), "{first_name.upper()}")]')
                        make_click(f'//*[@class="td-patient"]//*[contains(text(), "{last_name.upper()}") and contains(text(), "{first_name.upper()}")]/..//*[@class="btn-edit-patient"]')
                        # try:
                        #     element = WebDriverWait(driver, 10).until(
                        #         EC.presence_of_element_located(
                        #             (By.XPATH, f'//*[contains(text(), "{newDate}")]/..//*[@class="btn-edit-patient"]')
                        #         )
                        #     )
                        #     driver.execute_script("arguments[0].click();", element)
                        # except:
                        #     pwrite(f'Patient Not Found {name} {dob} {address} {last_two} {phone} {email}')
                        #     patient_not_found(name, dob, check_in_time)
                        #     save_record(name, dob)
                        #     save_record(name, check_in_time)
                        #     continue
                        try:
                            phlist = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@name="phList"]/tbody//tr[position() < last()]/td[1]')))
                        except:
                            phlist = []
                        def digits_only(s):
                            return re.sub(r'\D', '', s) if s else ""
                        try:
                            phlists = [ph.text.strip().replace(')', '').replace('(', '').replace(' ', '') for ph in phlist]
                        except:
                            phlists = []
                        try:phlists = [digits_only(ph.text) for ph in phlist]
                        except: 
                            phlists = []
                            pwrite(traceback.format_exc(), p=False)
                        if digits_only(phone) not in phlists:
                            make_click('//*[@title="Add New Phone"]')
                            #write_t('//input[@id="p_phone_"]', phone, sleep_time=2, p=True)
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@id="p_phone_"]'))).send_keys(phone)
                            try:
                                Select(driver.find_element(By.XPATH, '//*[@name="P_PhoneTypeId"]')).select_by_value('1')
                            except:
                                pass
                            try:
                                make_click('//*[@id="checkBoxPrimId"]')
                            except:
                                pass
                            make_click('//*[@class="p_editRow"]//*[text()="Save"]')


                            
                        else:
                            pwrite(f'Phone Number Already Exists {phlists} {phone}')
                        write_t('//*[@id="Email"]', email, p=True)
                        write_t('//textarea[@id="Notes" and  @rows="3" and @class="form-control "]', family_doc)
                        
                        # address_ = extract_text('(//*[contains(@id, "addressid")]//td)[1]')
                        # city_ = extract_text('(//*[contains(@id, "addressid")]//td)[2]')
                        # address_type = extract_text('(//*[contains(@id, "addressid")]//td)[3]')
                        # postal_code = extract_text('(//*[contains(@id, "addressid")]//td)[4]')

                        matched = False
                        addresss = driver.find_elements(By.XPATH, '//*[contains(@id, "addressid")]')
                        
                        for row in addresss:
                                tds = row.find_elements(By.XPATH, ".//td")
                                
                                # if len(tds) < 4:
                                #     continue

                                # Staff-entered / database values
                                table_address = tds[0].text.strip()
                                table_city = tds[1].text.strip()
                                table_type = tds[2].text.strip()
                                table_postal = tds[3].text.strip()

                                # Patient-typed values
                                patient_address = address  # from your earlier variable
                                patient_postal = postal_code_thereceptionist
                                # if address.lower().strip() in table_address.lower().strip() or table_address.lower().strip() in address.lower().strip():
                                #     matched = True
                                #     break 
                                # if (postal_code_thereceptionist.lower().strip() in table_postal.lower().strip() or table_postal.lower().strip() in postal_code_thereceptionist.lower().strip()) and (table_city.lower().strip() in city.lower().strip() or city.lower().strip() in table_city.lower().strip()):
                                #     matched = True
                                #     break

                                # AI address match check                                        
                                is_match = ai_address_match(
                                    patient_address,             # typed address
                                    patient_postal,              # typed postal (may be invalid/missing)
                                    table_address,               # reference address
                                    table_city,                  # reference city
                                    table_postal                 # reference postal
                                )

                                if is_match:
                                    matched = True
                                    break
                        

                        # result
                        if matched:
                            pwrite(f"Matched address-- Patient: {address}, {postal_code_thereceptionist} | Table: {table_address}, {table_postal}")
                        else:
                            pwrite("No match found")
                            make_click('//*[@title="Add New / Edit"]')
                            # try:
                            #     make_click('//*[@id="newDemAddr"]', sleep_time=2)
                            # except:
                            #     pass
                            write_t('//*[@name="dem_addressLine1"]', address, sleep_time=2, p=True)
                            write_t('//*[@name="dem_city"]', city, sleep_time=2)
                            write_t('//*[@name="dem_postalCode"]', postal_code_thereceptionist, sleep_time=2)
                            make_click('//*[@id="dem_addr_submit"]')
                            pwrite(f'Address Updated')
                            try:
                                make_click('//*[@id="dem_address_id"]//*[@class="close"]')            
                            except:
                                pass
                        time.sleep(2)
                        if not any(ch.isdigit() for ch in last_two.strip()):
                            write_t('//input[@name="Version"]', last_two, p=True)
                            make_click('//*[@data-cb-tp-title="OHIP CHECK"]')
                            time.sleep(5)
                            try:
                                make_click('//*[@id="ohip_check_id"]//*[@data-dismiss="modal"]')
                            except:
                                save_record(name, check_in_time)
                                pwrite(f'OHIP Check Failed for {name} {dob} {address} {last_two} {phone} {email}')
                                continue
                        #make_click('//*[@type="submit" and text()="Update Patient"]')
                        submit_button = driver.find_element(By.XPATH, '//*[@type="submit" and text()="Update Patient"]')
                        driver.execute_script("arguments[0].click();", submit_button)
                        time.sleep(5)
                        pwrite(f'Patient Updated {name} {dob} {address} {last_two} {phone} {email}')
                        try:
                            make_click('//*[@action="/Patients/Edit"]//*[@class="close"]', t=5)
                        except:
                            pass
                    elif status == "Holter 72 hr":
                        dob = extract_text('//*[contains(text(), "Patient")]/../..//*[@class="text-right"]')

                        if record_exists(name, dob) or record_exists(name, check_in_time):
                            pwrite(f'Skipping already processed patient: {name} {dob}')
                            continue
                        try:
                            make_click('//*[@data-balloon="View PDF Agreement"]')
                        except:
                            make_click('.//*[@class="ellipsis text-bold"]',driver=lead, t=10)
                        
                        
                        months = {
                            "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
                            "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
                            "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"
                        }
                        parts = dob.split("-")
                        newDate = f"{months[parts[1]]}-{parts[0]}-{parts[2]}"
                        timeout = 20  # seconds
                        max_age = 60  # seconds (file must be <= 60 seconds old)

                        end_time = time.time() + timeout
                        pdf_path = None

                        while time.time() < end_time:
                            now = datetime.now()
                            for fname in os.listdir(downloads_folder):
                                if fname.lower().endswith(".pdf"):
                                    fpath = os.path.join(downloads_folder, fname)
                                    mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
                                    if now - mtime <= timedelta(seconds=max_age):
                                        pdf_path = fpath
                                        break
                            if pdf_path:
                                break
                            time.sleep(1)
                        for handle in driver.window_handles:
                            driver.switch_to.window(handle)
                            if "cerebrum" in driver.current_url:
                                break
                        # write_t('//*[@placeholder="Patient Search"]', last_name)
                        # time.sleep(3)
                        # make_click('//*[@id="btn_search"]')
                        driver.get(f'https://cerebrum.mycerebrum.com/patients/patientsearch?searchType=0&patientSearch={last_name}')
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//*[contains(text(), "{newDate}")]')))
                        except:
                            try:
                                driver.get(f'https://cerebrum.mycerebrum.com/patients/patientsearch?searchType=0&patientSearch={first_name}')
                                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//*[contains(text(), "{newDate}")]')))
                            except:

                                pwrite(f'Patient Not Found {name}')
                                patient_not_found(name, dob, check_in_time)
                                save_record(name, dob)
                                save_record(name, check_in_time)

                        try:
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH, f'//*[contains(text(), "{newDate}")]/..//*[@class="btn-loose-report-upload"]')
                                )
                            )
                            driver.execute_script("arguments[0].click();", element)
                        except:
                            pwrite(f'Patient Not Found {name} {dob} ')
                            continue
                        if pdf_path:
                            file_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="File"]')))
                            file_input.send_keys(pdf_path)
                        try:
                            Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="PracticeDoctorId"]')))).select_by_visible_text('Singla Mohit')
                        except:
                            Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="PracticeDoctorId"]')))).select_by_index(1)
                        time.sleep(2)
                        try:
                            Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="ReportClassId"]')))).select_by_visible_text('Consent Form')
                        except:
                            Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="ReportClassId"]')))).select_by_value('53')
                        time.sleep(2)



                        try:
                            Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="CategoryId"]')))).select_by_visible_text('Holter Consent')
                        except:
                            Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="CategoryId"]')))).select_by_index(4)
                        time.sleep(2)
                        make_click('//*[@value="save" and text()="Upload File"]')
                        time.sleep(5)
                        os.remove(pdf_path)
                        pwrite(f'Holter Consent Uploaded {name} {dob}')
                        try:
                            make_click('//*[@action="/Documents/uploads/UploadLooseReport"]//*[@class="close"]', t=5)
                        except:
                            pass
                        save_record(name, check_in_time)
                        save_record(name, dob)
                
            
                
                # if check_in_day in ["Wednesday", "Thursday", "Friday"]:
                #     clinic_name = "Singla_BT"
                # elif check_in_day == "Monday":
                #     clinic_name = "Singla_NM"
                # else:
                #     clinic_name = None
                
                    
                
            except: 
                pwrite(f"Error occurred: {traceback.format_exc()}")

    else:
        if not test:
            make_click('//*[@id="date-range"]', t=10)
            make_click('//*[@data-range-key="Today"]', t=10, sleep_time=2)        
            make_click('//button[@type="submit"]')
        
        time.sleep(5)
        for l in range(100):
            try:
                make_click('//*[@id="load-more"]', t=5)
            except:
                time.sleep(5)
                try:
                    make_click('//*[@id="load-more"]', t=5)
                except:
                    break
            

        try:
            leads = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="visits"]//tr')))
        except:
            leads = []
        nowe = datetime.now()
        today_date = nowe.strftime("%Y-%m-%d")
        for lead in leads:
            try:
                for handle in driver.window_handles:
                    driver.switch_to.window(handle)
                    if "thereceptionist" in driver.current_url:
                        break
                try:
                    make_click('//*[@class="close"]', t=1)
                except:
                    pass
                #data_timestamp = WebDriverWait(lead, 5).until(EC.presence_of_element_located((By.XPATH, './/*[contains(text(), " at ") and contains(text(), "-")]'))).text.strip()
                try:
                    status = WebDriverWait(lead, 5).until(EC.presence_of_element_located((By.XPATH, './/td[3]//div[@title]//span')))
                except:
                    try:
                        make_click('//*[@class="close"]', t=1)
                    except:
                        pass
                check_in_time = lead.find_element(By.XPATH, './/td[6]//div[@title]//span').text.strip()
                status = lead.find_element(By.XPATH, './/td[3]//div[@title]//span').text.strip()
                name = lead.find_element(By.XPATH, './/*[@class="ellipsis text-bold"]').text.strip()
                if 'Sachleen kaur' in name:
                    continue
                check_in_time = f'{check_in_time}||{today_date}'
                if record_exists(name, check_in_time):
                    pwrite(f'Skipping already processed patient: {name} {check_in_time}')
                    continue
                
                last_name= name.split()[-1]
                try:
                    make_click('.//*[@class="ellipsis text-bold"]',driver=lead, t=10)
                except:
                    element = lead.find_element(By.XPATH, './/*[@class="ellipsis text-bold"]')
                    driver.execute_script("arguments[0].click();", element)
                time.sleep(1)
                
                #if not record_exists_late(name, check_in_time):
                if status == "Check In":
                        continue
                # Holter 74, Holter 14, stress tilt content,  ROI release - all can be uploaded end of day just like already you have made for Holter 72
                # if status == "Holter 72 hr" or status == "Holter 74" or status=="Stress/ Tilt Consent" or status == "Holter 14 days" or status == "ROI Release" or status == "ROI Request":
                #     ...
                dob = extract_text('//*[contains(text(), "Patient")]/../..//*[@class="text-right"]')

                if record_exists(name, dob) or record_exists(name, check_in_time):
                    pwrite(f'Skipping already processed patient: {name} {dob}')
                    continue
                try:
                    make_click('//*[@data-balloon="View PDF Agreement"]')
                except:
                    make_click('.//*[@class="ellipsis text-bold"]',driver=lead, t=10)
                
                
                months = {
                    "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
                    "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
                    "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"
                }
                parts = dob.split("-")
                newDate = f"{months[parts[1]]}-{parts[0]}-{parts[2]}"
                timeout = 20  # seconds
                max_age = 60  # seconds (file must be <= 60 seconds old)

                end_time = time.time() + timeout
                pdf_path = None

                while time.time() < end_time:
                    now = datetime.now()
                    for fname in os.listdir(downloads_folder):
                        if fname.lower().endswith(".pdf"):
                            fpath = os.path.join(downloads_folder, fname)
                            mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
                            if now - mtime <= timedelta(seconds=max_age):
                                pdf_path = fpath
                                break
                    if pdf_path:
                        break
                    time.sleep(1)
                for handle in driver.window_handles:
                    driver.switch_to.window(handle)
                    if "cerebrum" in driver.current_url:
                        break
                # write_t('//*[@placeholder="Patient Search"]', last_name)
                # time.sleep(3)
                # make_click('//*[@id="btn_search"]')
                dt = datetime.now()
                check_in_date = dt.strftime("%m/%d/%Y")
                check_in_day = dt.strftime("%A")
                env_key = DAY_ENV_MAP.get(check_in_day)
                clinic_name = os.getenv(env_key) if env_key else None
                d_t = check_in_date.replace("/", "%2F")
                if test:
                    d_t= "12%2F19%2F2025"
                for handle in driver.window_handles:
                        driver.switch_to.window(handle)
                        if "cerebrum" in driver.current_url:
                            break

                #c_url = f'https://cerebrum.mycerebrum.com/Schedule/daysheet?Date={d_t}&FilterPatient={last_name}'
                c_url = f'https://cerebrum.mycerebrum.com/Schedule/daysheet?OfficeId=30&Date={d_t}&AppointmentStatusId=-1&Expected=False&ExcludeTestOnly=False&ExcludeCancelled=True&OnlyActionOnAbnormal=False&FilterPatient={last_name}&ShowOrders=False&Page=1&PageSize=25'
                driver.get(c_url)
                #driver.get(f'https://cerebrum.mycerebrum.com/patients/patientsearch?searchType=0&patientSearch={last_name}')
                if clinic_name:
                    dropdown = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "OfficeId"))
                    )
                    Select(dropdown).select_by_visible_text(clinic_name)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//*[@class="td-patient"]//*[contains(text(), "{last_name.upper()}") and contains(text(), "{first_name.upper()}")]')))
                    make_click(f'//*[@class="td-patient"]//*[contains(text(), "{last_name.upper()}") and contains(text(), "{first_name.upper()}")]')
                    make_click(f'//*[@class="td-patient"]//*[contains(text(), "{last_name.upper()}") and contains(text(), "{first_name.upper()}")]/..//*[@class="btn-loose-report-upload"]')
                except:
                    try:
                        driver.get(f'https://cerebrum.mycerebrum.com/patients/patientsearch?searchType=0&patientSearch={first_name}')
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//*[contains(text(), "{newDate}")]')))
                        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//*[contains(text(), "{newDate}")]/..//*[@class="btn-loose-report-upload"]')))
                        driver.execute_script("arguments[0].click();", element)
                    except:
                        continue
                        pwrite(f'Patient Not Found {name}')
                        patient_not_found(name, dob, check_in_time)
                        save_record(name, dob)
                        save_record(name, check_in_time)

                
                if pdf_path:
                    file_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="File"]')))
                    file_input.send_keys(pdf_path)
                try:
                    Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="PracticeDoctorId"]')))).select_by_visible_text('Singla Mohit')
                except:
                    Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="PracticeDoctorId"]')))).select_by_index(1)
                time.sleep(2)
                try:
                    Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="ReportClassId"]')))).select_by_visible_text('Consent Form')
                except:
                    Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="ReportClassId"]')))).select_by_value('53')
                time.sleep(2)



                try:
                    Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="CategoryId"]')))).select_by_visible_text('Holter Consent')
                except:
                    Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@name="CategoryId"]')))).select_by_index(4)
                time.sleep(2)
                make_click('//*[@value="save" and text()="Upload File"]')
                time.sleep(5)
                os.remove(pdf_path)
                pwrite(f'Holter Consent Uploaded {name} {dob}')
                try:
                    make_click('//*[@action="/Documents/uploads/UploadLooseReport"]//*[@class="close"]', t=5)
                except:
                    pass
                save_record(name, check_in_time)
                save_record(name, dob)
                
            except:
                pwrite(f"Error occurred: {traceback.format_exc()}")
main(daytime=False, test=True)

end_of_day_executed = False
while True:
    now = datetime.now()
    weekday = now.weekday()  # Monday = 0, Sunday = 6
    hour = now.hour
    

    if 0 <= weekday <= 4 and 7 <= hour < 18:  # Mon–Fri, 7AM–6PM
        
        end_of_day_executed = False
        try:


            main()
        except Exception as e:
            pwrite(f"Error occurred: {traceback.format_exc()}")
        pwrite("Waiting for 20 seconds before next check...")
        time.sleep(20)

    elif not end_of_day_executed and 0 <= weekday <= 4 and hour >= 18:
        # Execute end-of-day tasks
        pwrite("Executing end-of-day tasks...")
        try:
            # Placeholder for actual end-of-day tasks
            main(daytime=False)
        except Exception as e:
            pwrite(f"Error during end-of-day tasks: {traceback.format_exc()}")
        end_of_day_executed = True
    else:
        # Reason logging
        if weekday >= 5:
            reason = "It's weekend (Saturday/Sunday)."
            # Move to next Monday 7 AM
            days_until_monday = (7 - weekday) % 7
            next_time = (datetime.now() + timedelta(days=days_until_monday)).replace(hour=7, minute=0, second=0, microsecond=0)
        elif hour < 7:
            reason = f"It's too early ({hour}:00). Waiting until 7 AM."
            next_time = datetime.now().replace(hour=7, minute=0, second=0, microsecond=0)
        else:
            reason = f"It's too late ({hour}:00). Waiting until 7 AM tomorrow."
            next_time = (datetime.now() + timedelta(days=1)).replace(hour=7, minute=0, second=0, microsecond=0)

        seconds_to_sleep = (next_time - datetime.now()).total_seconds()

        pwrite(f"Outside working hours. {reason}")
        pwrite(f"Sleeping for {seconds_to_sleep/3600:.2f} hours...")
        time.sleep(seconds_to_sleep)