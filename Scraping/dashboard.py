'''
    Things to change when we switch between India, GLobal, KSA
    1) Below Client inside init_google_sheet function change the sheet name
    2) In the step 0 below the Driver varibale change between driver.get 
    3) In the step 2 while direct url for publisher change between driver.get
    4) In the step E we have urls to directly go to publisher after data is fetching
    5) This is pretty main we have to change the LC based on the global, india, ksa
'''

'''
    For above steps we changed it now things are simple
    1) Change REGION
    2) If required change the license_codes inside the REGION_CONFIG
    3) Check if the json file of credential file path is right - This is for me to remember if in future i change the system
    4) Check if the sheet name is right we associated with the region
'''
import os
import time
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from google.oauth2.service_account import Credentials
from selenium.webdriver.common.action_chains import ActionChains

# =======================
# REGION CONFIG
# =======================
REGION = "KSA"   # options: "INDIA", "GLOBAL", "KSA"

REGION_CONFIG = {
    "INDIA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com",
        "sheet_name": "Dashboard India",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "in~~134106156","in~aa131896","in~58adcd07","in~~47b66667","in~14507c784","in~11b564357","in~~c2ab364c","in~~c2ab3662","in~~47b66639","in~58adcc70","in~~c2ab363b","in~11b56430c","in~~47b66678","in~~1341061ac","in~~15ba205d1","in~~1341061b5","in~~10a5cbb1d","in~~991991c8","in~311c4742","in~311c4724","in~~134106208","in~~991991c4","in~~134106220","in~311c488b","in~11b564274","in~14507c728","in~~15ba20670","in~~15ba2068a","in~76aa24d","in~~47b666d5","in~76aa1d8","in~76aa247","in~~99199258","in~~71680b65","in~311c4744","in~~2024c207","in~76aa2b4","in~~10a5cbb42","in~~c2ab36a2","in~aa1318ab","in~d3a49bac","in~~c2ab36d5","in~~c2ab3671","in~~47b66716","in~~991991d0","in~~10a5cbb34","in~76aa273","in~~99199217","in~76aa298","in~~1341061c6","in~~10a5cba3a","in~d3a49b75","in~58adcb79","in~~15ba20672","in~~71680bbb","in~58adcb94","in~11b5641d0","in~d3a49b43","in~~10a5cbb79","in~826171c3","in~58adcb40","in~76aa22a","in~~15ba20749","in~311c467c","in~~47b66689","in~~2024c1d7","in~~2024c218","in~~71680b3c","in~~47b6668a","in~~99199244","in~~2024c246","in~11b564276","in~~134106213","in~~15ba206a8","in~~10a5cbb61","in~11b564256","in~~c2ab36ad","in~76aa268","in~~9919921b","in~~134106216","in~~71680b92","in~76aa245","in~311c46d4","in~311c46d3","in~58adcb85","in~~2024c249","in~76aa1c0","in~11b564332","in~~71680c19","in~~15ba20752","in~~2024c085","in~~2024c2a0","in~14507c681","in~~2024c2c1","in~76aa206","in~~71680bb9","in~aa131675","in~14507c65b","in~11b5641a9","in~~71680c2b","in~~10a5cbc25","in~aa13163a","in~11b56418d","in~311c4663","in~~c2ab3781","in~~1341062bb","in~~991992c4","in~~10a5cbc2d","in~~1341062c1","in~14507c641","in~~71680c30","in~aa13164b","in~~15ba20759","in~~15ba205c0","in~~2024c231","in~~47b6677d","in~58adcb36","in~aa13166b","in~~991992d1","in~~1341062c2","in~~99199081","in~14507c63b","in~11b564246","in~11b56417b","in~aa131665","in~~71680b90","in~14507c666","in~aa131632","in~76aa20d","in~311c464b","in~311c4766","in~11b564172","in~~47b66782","in~~c2ab36a7","in~~47b6665b","in~~c2ab3794","in~~2024c262","in~~47b66752","in~aa1317ba","in~~99199213","in~58adcb46","old-in~~99199240","in~aa13168a","in~311c4685","in~58adcb73","in~~134106255","in~~c2ab36cb"
        ]
    },
    "GLOBAL": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com",
        "sheet_name": "Dashboard Global",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "~2024b5d8","d3a4ac1c","14507cc77","99198968","~15ba1da68","aa13266b","~c2ab3108","~13410604b","~1341056a0","~134105965","d3a4a32d","76aa124","~9919868d","76a9c30","~311c4b76","58adca91","~2024bada","~99198a29","~2024bad5","~c2ab3033","82618089","~1341061bb","d3a4a69c","58add2da","~10a5cb6b0","aa132703","~aa1321c5","~c2ab2c0c","~9919871c","~47b66614","~991981d3","8261786b","11b564b69","aa132225","311c4c11","311c4bbb","14507cba8","76aac69","~71680627","~15ba20042","~oldetmoney","~c2ab275a","~15ba1d70a","~c2ab2ba2","~47b6665c","~991989d1","~old2024c085","~oldmagma1","~oldmagma2","~oldmagma3","11b56527d","~c2ab2c08","~11b5646b8","~d3a4a286","76aa7c6","76aa85d","~47b65b94","~7168053d","~d3a49c4c-old","76aa868","~7168069d","~oldrangde","~10a5cb636","~99198226","11b56470b","~99198a18","311c4bc4","d3a4a420","~15ba1db98","~14507ccb9","~11b564836","~14507ccc0","~47b661c8","~7168071b","~10a5cb278","~oldUPES","d3a4a72a","~47b65a1c","~134105732","~9919837c","~311c4dc3","~2024bb26","~76aab32","old~2024c1a3","826182a0","~10a5cb24b","11b56488a","~15ba2063d","~47b661ab","99198624","14507cd97","d3a4ab04","~ c2ab260c","~aa131bd4","aa131752","11b564bc3","~14507d169","~d3a4a64d","~aa1321a1","~716802d0","~7168026d","~15ba1db9b","~7168028a","~47b65bd8","~99199107","~oldid","~10a5cb319","~10a5cb2a0","~11b564252","~47b65c27","~aa131717","~10a5cb33d","~82616dd3","~14507c905","~58adc916","~11b563c29","~311c4664","~c2ab3c28","~7167db9d","~99198b2b","~oldedelwiess","~2024b1c1","~15ba20147","old~10a5cbb38","58addb9c","58adcba4","~oldsasai","~ c2ab36a7","~9919922b","~15ba1dda1","~71680426","~15ba206bc","~teamGreatLearning","old~7167d2da","~c2ab3713","old~76aab14","~c2ab2d61","old~1341061b2","~2024b742","~47b65ca0","~82617225","76abb05","~15ba1cc68","~82617957","d3a49b72","~991988b2","d3a4a4aa","~oldaccount","~47b66522","~15ba1cc70","~15ba20218","76aa239","14507d0c5","11b565a86","~10a5cb29d","~10a5cbb86","aa1318a4","~15ba20712","aa133154","~58add514","~82617a45","311c6080","~7167d365","~9919921c","~716806cb","~47b666aa","~71680bc5","76aa235","~76aa276","~13410606b","~2024c07d","~47b66726","1450800cc","82618a78","~99197808","~99197879","~7167dd04","~15ba20518","~11b565a1","~c2ab1d64","~99197854","d3a4a339","~c2ab3737","~134105770","~82618284","11b564400","~47b64d14","~145080099","~2024a898","~99197908","~10a5cbbc3","~145080038","d3a4b4bc","~10a5cbb7a","311c617c","~1341058da","76ab933","76aa777","76ab97b","d3a4b5a8","~2024a887","145080030","11b565a3b","d3a4b4a3","~2024ba16","~99197905","~10a5cb448","~2024a920","~2024a932","~99197925","~7167d473","d3a4b49c","~2024a938","~amplicom","~15ba1cda8","82618947.0","~TapsiMarket","~GreenlifeEwaste","76ab953","~7167d478","11b565ab8","aa133008","~10a5ca413","~c2ab1d61","~47b64d69","aa13306d","~c2ab3a30","58b0047d","76aab90","311c4c52","~71680307","~99198b33","d3a4a33c","~15ba1d007","~134104952","76ab929","~c2ab1dc3","76ab940","58add629","~2024a96d","311c6107","~c2ab2d96","~7167d4b3","76ab912","~76aa7a9","82617869","99198624","76ab910","~organicmandya","~noveltywealth","~megafinance","~srcbeauty","~71680b83","~15ba1cc59","~47b660c0","71680330","11b5648c7","~134105acb","11b56488d","~c2ab3684","~47b65c7c","~2024b7ca","~99198ad5","d3a4aa4b","~76aaaa1","~82617a90","d3a4a3cc","~7167d340","~d3a4a2c1","~716802d7","~1341059d8","~10a5cad25","~aa1326d5","~58add611","58add2dc","~99198a33","99198274","~7167dccb","d3a4a621","71680076","311c5233","311c4bbd","~47b6578d","11b5650ac","~47b657c3","aa133160","d3a4a49a","134106264","82617258","11b564972","~10a5cb6c3","~15ba1dcbd","~71680b9c","~15ba1cc80","~10a5cb3d1","aa133114","~oldmelooha","~7167d349","d3a4b591","~aa132104","~c2ab1cba","76aa228","d3a4a3b3","134104863","~old2","~15ba1cd11","~2024a93b","~47b6673b","aa131d37","d3a4a667","~11b56421","11b565abc","~c2ab1c88","old~11b564403","~c2ab1cdb","~bit24newtempold","11b565a43","134104829","82617855","~2024c07c","~991978c7","311c60ad","82618978","oldin~~c2ab3761","~99198b68","145080023","11b565971","11b565961","134104919","old~76ab96b","~2024a939","82618947","~7167d4a4","~organicmandya","~noveltywealth","~megafinance","~srcbeauty"
        ]
    },
    "KSA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com",
        "sheet_name": "Dashboard KSA",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "ksa~~2024c070","ksa~58adcd4c","ksa~11b564406","ksa~~15ba2051c","ksa~aa131897","ksa~82617408","ksa~d3a49d49","ksa~11b5643db","ksa~d3a49d46","ksa~~134106074","ksa~~2024c07d","ksa~~2024c085","ksa~82617402","ksa~~10a5cb9c4","ksa~~15ba20518","ksa~~134106069","ksa~311c4892","ksa~~99199083","ksa~~13410606b","ksa~11b5643d5","ksa~~134106080","ksa~58adcd44","ksa~826173db","ksa~d3a49d41","ksa~~99199087","ksa~aa131893","ksa~76aa3da","ksa~826173dc","ksa~aa131890","ksa~~10a5cb9d1","ksa~~2024c08d","ksa~~716809d4","ksa~~99199088","ksa~~2024c083","ksa~~2024c084"
        ]
    }
}

def init_google_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(
        "/Users/admin/Desktop/Code Directory/Product_Adaption_Data/Credential File/mycred-googlesheet.json",
        scopes=scopes
    )
    client = gspread.authorize(creds)

    sheet_name = REGION_CONFIG[REGION]["sheet_name"]
    sheet = client.open_by_key(
        "1D0-O3OX3TOmZRMRZmyYoubRxBtekAB05b50lcaIOweM"
    ).worksheet(sheet_name)

    return sheet

try:
    sheet = init_google_sheet()
    print("✅ Connected to Google Sheets")
except Exception as e:
    print(f"❌ Failed to connect to Google Sheets: {e}")
    exit()

# ========== We need no change from here till the edit button clicked ========== #

# ---------- STEP 0: SETUP PERSISTENT PROFILE ----------
script_dir = os.path.dirname(os.path.abspath(__file__))
user_data_dir = os.path.join(script_dir, "selenium_profile")

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument(f"--user-data-dir={user_data_dir}")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 120) 

# Step 1: Open admin dashboard
BASE_URL = REGION_CONFIG[REGION]["base_url"]
driver.get(f"{BASE_URL}/admin")

# --- GIVE IT A MOMENT TO REDIRECT ---
print("⏳ Waiting for page to settle...")
time.sleep(3) 

# ---------- STEP 2: SMART PAGE DETECTION ----------
print(f"🔍 Current URL: {driver.current_url}")

# Give it one refresh if we aren't where we expect to be
if "publisher.html" not in driver.current_url:
    print("🔄 Refreshing to ensure state...")
    driver.refresh()
    time.sleep(2)

if "publisher.html" in driver.current_url:
    print("✅ Already on Publishers page via Cookies.")

else:
    try:
        print("🔍 Checking if Publishers link is visible...")
        # Try a more generic XPath that finds the link even if text is weird
        publisher_xpath = "//a[contains(@href, 'publisher.html')]"
        wait_short = WebDriverWait(driver, 15) # Increased to 15
        
        link = wait_short.until(EC.element_to_be_clickable((By.XPATH, publisher_xpath)))
        link.click()
        print("✅ Clicked Publishers from sidebar.")

    except Exception:
        print("📍 Sidebar link not found. Starting Deep Navigation...")
        
        # FIX: The profile head might be nested. Let's use a simpler selector.
        try:
            profile_xpath = "//div[contains(@class,'pop-over__head')] | //div[contains(@class,'noselect')]"
            dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, profile_xpath)))
            driver.execute_script("arguments[0].click();", dropdown) # JS click is safer here
            print("✅ Opened Profile Dropdown")
            time.sleep(1.5)

            # Click Super Admin - Use a more flexible text match
            super_admin_xpath = "//a[normalize-space()='Super Admin' or contains(text(),'Super Admin')]"
            sa_btn = wait.until(EC.element_to_be_clickable((By.XPATH, super_admin_xpath)))
            sa_btn.click()
            print("✅ Clicked Super Admin")

            # Final check for Publisher link after landing in Super Admin
            final_publisher_xpath = "//a[contains(@href,'publisher.html')]"
            wait.until(EC.element_to_be_clickable((By.XPATH, final_publisher_xpath))).click()
        except Exception as e:
            print(f"❌ Deep Navigation failed: {e}")
            driver.save_screenshot("nav_failure.png")
            # If everything fails, try going to the URL directly as a last resort
            print("🚀 Attempting direct URL navigation...")
            PUBLISHER_URL = REGION_CONFIG[REGION]["publisher_url"]
            driver.get(PUBLISHER_URL)


print("🎯 SUCCESS: You are now on the Publishers page.")

def search_by_license(driver, wait, license_code):
    license_input = wait.until(
        EC.presence_of_element_located((By.NAME, "licenseCode"))
    )
    license_input.clear()
    license_input.send_keys(license_code)

    search_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[normalize-space()='Search']")
        )
    )
    search_btn.click()
    print("✅ Succesfully entered the LC and clicked search button")

def should_skip_account(driver):

    # ❌ Case 1: Oops page
    if driver.find_elements(By.XPATH, "//h2[contains(text(),'Oops')]"):
        print("⛔ Oops page → Skipping")
        return True, "Region mismatch"

    # ❌ Case 2: Request Demo (ROBUST FIX)
    if (
        driver.find_elements(By.XPATH, "//*[contains(normalize-space(.), 'Request a Demo')]")
        or "Request a Demo" in driver.page_source
    ):
        print("⛔ Request Demo → Skipping")
        return True, "Service stopped"

    return False, None

def check_if_result_exists(driver, license_code, timeout=5):
    """
    Returns False if LC does not appear in publisher table
    (wrong region / invalid LC)
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                d.find_elements(By.XPATH, f"//tr[contains(., '{license_code}')]")
                or "No data" in d.page_source
            )
        )
    except:
        return False

    rows = driver.find_elements(By.XPATH, f"//tr[contains(., '{license_code}')]")
    return len(rows) > 0


def open_actions_dropdown(driver, wait, license_code):
    print(f"⏳ Opening Actions dropdown for {license_code}...")
    
    # Improved XPath: Finds the specific TR containing the LC, then the toggle button inside it
    toggle_xpath = f"//tr[contains(., '{license_code}')]//button[contains(@class, 'dropdown-toggle')]"
    
    try:
        dropdown_btn = wait.until(EC.element_to_be_clickable((By.XPATH, toggle_xpath)))
        # Force scroll and use JS click to bypass any overlapping UI elements
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", dropdown_btn)
        print("✅ Successfully clicked the specific dropdown toggle")
    except Exception as e:
        print(f"⚠️ Could not click dropdown for {license_code}: {e}")
        # If it fails, we try a refresh as a fallback
        driver.refresh()
        time.sleep(2)

def click_request_access(wait):
    request_access = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@class,'requestAccess')]")
        )
    )
    request_access.click()
    print("✅ Succesfully opened the access dialog box")

def handle_request_modal(wait, driver):
    print("⏳ Handling request modal...")
    
    # Give the modal a second to fully animate and render
    time.sleep(2)

    # 1. SMART IFRAME DETECTION
    # We only switch if an iframe actually exists AND is visible
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    switched = False
    for frame in iframes:
        if frame.is_displayed():
            driver.switch_to.frame(frame)
            switched = True
            print("↔️ Switched to active iframe")
            break

    try:
        # 2. WAIT FOR DROPDOWN
        # We use a very broad XPath to find the select by its ID or Name
        dropdown_xpath = "//select[@id='roleIdField' or @name='roleEId']"
        role_dropdown = wait.until(
            EC.presence_of_element_located((By.XPATH, dropdown_xpath))
        )
        
        # 3. FORCE SELECTION VIA JAVASCRIPT
        # This is the most reliable way to select 'Viewer' regardless of UI quirks
        print("🚀 # React-safe role selection (required only during first-time access request)")
        Select(role_dropdown).select_by_visible_text("Viewer")
        
        # 4. FILL COMMENT
        comment_box = wait.until(EC.presence_of_element_located((By.ID, "commentText")))
        comment_box.clear()
        comment_box.send_keys("access request")

        # 5. CLICK REQUEST
        request_btn = driver.find_element(By.XPATH, "//button[contains(.,'Request')]")
        driver.execute_script("arguments[0].click();", request_btn)
        print("✅ Successfully submitted access request")

    except Exception as e:
        print(f"❌ Error inside modal: {e}")
        driver.save_screenshot("error_modal_view.png")
    
    finally:
        if switched:
            driver.switch_to.default_content()
            print("↔️ Switched back to main content")

def close_modal_if_exists(driver):
    """Closes the modal only if it is still visible."""
    try:
        # Check for the close button with a very short timeout
        close_btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.ID, "cboxClose"))
        )
        close_btn.click()
        print("✅ Manually closed the dialog box")
        time.sleep(1)
    except:
        print("ℹ️ Modal already closed or close button not found (which is fine)")

def click_edit(driver, wait, license_code):
    print(f"⏳ Clicking Edit for {license_code}...")
    
    # Find the Edit link inside the row that matches the License Code
    edit_xpath = f"//tr[contains(., '{license_code}')]//a[contains(@href, '/publisher/edit')]"
    
    edit_btn = wait.until(
        EC.element_to_be_clickable((By.XPATH, edit_xpath))
    )
    
    edit_btn.click()
    print("✅ Clicked the specific edit button")

# ========== Till here no change things will be same ========== #

def click_dashboards(wait):
    print("⏳ Clicking Dashboards...")

    dashboard_xpath = (
        "//a[contains(@href,'/custom-dashboard/list') and .//span[text()='Dashboards']]"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, dashboard_xpath))
    ).click()

    print("✅ Dashboards opened")

def wait_for_dashboard_table_or_empty(driver, timeout=8):
    """
    Waits for either:
    - dashboard rows
    - empty dashboard state
    - pagination (even if empty)
    Returns True if rows exist, False if NO_DATA
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                d.find_elements(By.CLASS_NAME, "table__row")
                or d.find_elements(By.CLASS_NAME, "pagination")
                or "No dashboards" in d.page_source
            )
        )
    except:
        return False

    rows = driver.find_elements(By.CLASS_NAME, "table__row")
    return len(rows) > 0

def extract_dashboard_page(driver, license_code):
    rows_data = []

    rows = driver.find_elements(
        By.XPATH,
        "//tbody/tr[contains(@class,'table__row')]"
    )

    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, "./td")
            if len(cells) < 5:
                continue

            # 1️⃣ Dashboard Name
            dashboard_name = cells[0].find_element(
                By.XPATH, ".//a"
            ).get_attribute("title").strip()

            # 2️⃣ Cards Count
            cards = cells[1].text.strip()

            # 3️⃣ Last Updated
            last_updated = cells[2].find_element(
                By.XPATH, ".//span"
            ).get_attribute("title").strip()

            # 4️⃣ Tags
            try:
                tags = ",".join([
                    t.text.strip()
                    for t in cells[3].find_elements(By.CLASS_NAME, "pill-text")
                ])
            except:
                tags = ""

            rows_data.append([
                license_code,
                dashboard_name,
                cards,
                last_updated,
                tags
            ])

        except Exception as e:
            print(f"⚠️ Skipping dashboard row: {e}")
            continue

    return rows_data


def extract_all_dashboards(driver, license_code):
    print("📥 Extracting Dashboards (all pages)...")

    has_data = wait_for_dashboard_table_or_empty(driver)

    # 🔹 NO DATA FAST EXIT
    if not has_data:
        print("⚠️ No dashboards found — fast skipping")
        return [[
            license_code,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA"
        ]]

    all_data = []
    page_number = 1

    while True:
        print(f"📄 Scraping Dashboard Page {page_number}...")
        page_data = extract_dashboard_page(driver, license_code)
        all_data.extend(page_data)

        try:
            next_li = driver.find_element(
                By.XPATH,
                "//li[contains(@class,'pagination__next')]"
            )

            if "is-disabled" in next_li.get_attribute("class"):
                break

            next_link = next_li.find_element(By.TAG_NAME, "a")
            driver.execute_script("arguments[0].click();", next_link)
            time.sleep(1.5)
            page_number += 1

        except:
            break

    if not all_data:
        return [[
            license_code,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA"
        ]]

    print(f"✅ Total Dashboards Extracted: {len(all_data)}")
    return all_data


def append_to_sheet(sheet, rows):
    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
        print("📤 Data pushed to Google Sheets")
    else:
        print("⚠️ No data to push")

def log_error_to_sheet(sheet, license_code, stage, error_reason):
    print(f"📝 Logging error for {license_code} at stage: {stage}")

    row = [
        license_code,
        "ERROR",
        stage,
        error_reason[:300],  # keep it readable
        time.strftime("%Y-%m-%d %H:%M:%S")
    ]

    sheet.append_row(row, value_input_option="USER_ENTERED")

LICENSE_CODES = REGION_CONFIG[REGION]["license_codes"]

for code in LICENSE_CODES:
    print(f"\n▶ Processing {code}")
    try:
        # Step A: Search and land on result
        search_by_license(driver, wait, code)

        # 🔥 MUST ADD THIS BLOCK
        skip, reason = should_skip_account(driver)

        if skip:
            print(f"⏭ Skipping {code} → {reason}")

            log_error_to_sheet(
                sheet,
                code,
                stage="SKIPPED",
                error_reason=reason
            )

            driver.get(REGION_CONFIG[REGION]["publisher_url"])
            wait.until(EC.presence_of_element_located((By.NAME, "licenseCode")))
            continue

        # 🔥 DASHBOARD ERROR GUARD: LC belongs to another region
        if not check_if_result_exists(driver, code):
            log_error_to_sheet(
                sheet,
                code,
                stage="REGION_MISMATCH",
                error_reason="License code not found in this region"
            )
            continue
        
        # Step B: Try to get access (only if needed)
        try:
            open_actions_dropdown(driver, wait, code) # Added driver and code
            
            request_btn_xpath = f"//tr[contains(., '{code}')]//a[contains(@class,'requestAccess')]"
            if len(driver.find_elements(By.XPATH, request_btn_xpath)) > 0:
                click_request_access(wait)
                handle_request_modal(wait, driver)
                time.sleep(2)
                close_modal_if_exists(driver)
            else:
                print("ℹ️ Access already available. Moving to Edit.")
                # If dropdown is open but we don't need it, refresh or Esc
                ActionChains(driver).send_keys(Keys.ESCAPE).perform() 
        except Exception as e:
            print(f"⚠️ Access step skipped: {e}")

        # --- Updated Step C: Pass 'code' to the edit function ---
        main_window = driver.current_window_handle
        click_edit(driver, wait, code)

        # SWITCH TO NEW TAB
        for window_handle in driver.window_handles:
            if window_handle != main_window:
                driver.switch_to.window(window_handle)
                break
        print(f"↔️ Switched to Edit tab for {code}")

        # Step D: Extract Data
        try:
            click_dashboards(wait)

            # Wait for dashboard table to load
            dashboard_rows = extract_all_dashboards(driver, code)
            append_to_sheet(sheet, dashboard_rows)

            time.sleep(1)

            dashboard_rows = extract_all_dashboards(driver, code)

            append_to_sheet(sheet, dashboard_rows)

        except Exception as e:
            error_msg = str(e)
            print(f"❌ Dashboard extraction failed for {code}: {error_msg}")

            log_error_to_sheet(
                sheet,
                code,
                stage="DASHBOARD_EXTRACTION",
                error_reason=error_msg
            )
    
    finally:
        # Step E: Cleanup for next iteration
        # Close extra tabs and go back to the list
        if len(driver.window_handles) > 1:
            driver.close() # Closes current (Edit) tab
            driver.switch_to.window(main_window)
        
        PUBLISHER_URL = REGION_CONFIG[REGION]["publisher_url"]
        driver.get(PUBLISHER_URL)

        time.sleep(2)
