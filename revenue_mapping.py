'''
    Things to change when we switch between India, GLobal, KSA
    1) Below Client inside init_google_sheet function change the sheet name
    2) In the step 0 below the Driver varibale change between driver.get 
    3) In the step 2 while direct url for publisher change between driver.get
    4) Inside the go_to_revenue_mapping function we switch between the url
    5) In the step 5 we have urls to directly go to publisher after data is fetching
    6) This is pretty main we have to change the LC based on the global, india, ksa
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
        "base_url": "https://dashboard.in.webengage.com",
        "sheet_name": "Revenue Mapping India",
        "publisher_url": "https://dashboard.in.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "in~58adcc11"
        ]
    },
    "GLOBAL": {
        "base_url": "https://dashboard.webengage.com",
        "sheet_name": "Revenue Mapping Global",
        "publisher_url": "https://dashboard.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "~99198968","82617894","82617822","82617757","~71680543","~71680588","~134105251","~71680655","82617754","~134105965","~71680577","82618089","~71680627","~99198226","82617775","~82617869","~71680632","82617779","~134105732","~99198624","~134105353","~99199107","~71680426","~82617225","~82617957","82618240","~71680374","~134104786","~99197808","~99197879","~134105365","~99197854","~134105770","~82618284","~145080099","~99197908","~145080038","145080012","~134104915","145080010","145080030","~99197905","~99197925","82618947.0","~134104929","~99197942","~134104910","~99198826","~71680307","~134104942","~134104952","~134104968"
        ]
    },
    "KSA": {
        "base_url": "https://dashboard.ksa.webengage.com",
        "sheet_name": "Revenue Mapping KSA",
        "publisher_url": "https://dashboard.ksa.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "ksa~~47b66537"
        ]
    }
}

def init_google_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(
        "/Users/admin/Desktop/Python Script/agreement_file_pasting/mycred-googlesheet.json",
        scopes=scopes
    )
    client = gspread.authorize(creds)

    sheet_name = REGION_CONFIG[REGION]["sheet_name"]
    sheet = client.open_by_key(
        "1sathL7caATX3PnV2urKhp8UpCLOcwdTIxzkJkA4Kar4"
    ).worksheet(sheet_name)

    return sheet

try:
    sheet = init_google_sheet()
    print("✅ Connected to Google Sheets")
except Exception as e:
    print(f"❌ Failed to connect to Google Sheets: {e}")
    exit()

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
        print("🚀 Forcing selection via JavaScript...")
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


def open_data_platform(driver, wait):
    print("⏳ Ensuring Data Platform menu is open...")

    data_platform_li = wait.until(
        EC.presence_of_element_located((By.ID, "nav-data-platform"))
    )

    class_attr = data_platform_li.get_attribute("class")

    if "menu__group--is-active" not in class_attr:
        print("🔓 Opening Data Platform sidebar via JS")
        driver.execute_script(
            "arguments[0].classList.add('menu__group--is-active');",
            data_platform_li
        )
        time.sleep(0.5)
    else:
        print("✅ Data Platform already open")

def go_to_revenue_mapping(driver, wait, account_id):
    base_url = REGION_CONFIG[REGION]["base_url"]
    url = f"{base_url}/accounts/{account_id}/data-management/events/revenue"

    driver.get(url)

    # wait for rows OR empty state
    wait_for_revenue_or_empty(driver)
    print("✅ Landed on Revenue Mapping (via URL)")


def click_data_management(wait):
    print("⏳ Clicking Data Management...")

    data_management_xpath = (
        "//a[contains(@href,'/data-management/system/attributes') and .//span[text()='Data Management']]"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, data_management_xpath))
    ).click()

    print("✅ Clicked Data Management")

def click_revenue_mapping(wait):
    print("⏳ Waiting for Revenue Mapping tab...")

    revenue_tab_xpath = (
        "//a[contains(@href,'/data-management/events/revenue') and normalize-space()='Revenue Mapping']"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, revenue_tab_xpath))
    ).click()

    print("✅ Revenue Mapping opened")

def extract_revenue_mapping_data(driver, wait, license_code):
    print("📥 Extracting Revenue Mapping data...")
    data_rows = []

    account_name = wait.until(
        EC.presence_of_element_located((By.ID, "we-account-name"))
    ).text.strip()

    try:
        currency = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//div[@aria-label='Select a currency']//span[contains(@class,'handle-text-overflow')]"
            ))
        ).text.strip()
    except:
        currency = "NULL"

    # ⏳ Wait briefly for rows OR empty state
    has_data = wait_for_revenue_or_empty(driver)

    if not has_data:
        print("⚠️ No Revenue Mapping found — inserting NO_DATA row")
        return [[
            license_code,
            account_name,
            currency,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            ""
        ]]

    rows = driver.find_elements(
        By.XPATH,
        "//div[contains(@class,'row') and .//i[contains(@class,'fl-delete')]]"
    )

    for row in rows:
        try:
            dropdowns = row.find_elements(By.XPATH, ".//div[contains(@class,'r-ss-trigger')]")
            if len(dropdowns) >= 2:
                event_name = dropdowns[0].text.strip()
                attribute_name = dropdowns[1].text.strip()

                data_rows.append([
                    license_code,
                    account_name,
                    currency,
                    event_name,
                    attribute_name,
                    "SUCCESS",
                    ""
                ])
        except:
            continue

    if not data_rows:
        return [[
            license_code,
            account_name,
            currency,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            ""
        ]]

    print(f"✅ Extracted {len(data_rows)} revenue mappings")
    return data_rows


def append_to_sheet(sheet, rows):
    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
        print("📤 Data pushed to Google Sheets")
    else:
        print("⚠️ No data to push")

def log_error_to_sheet(sheet, license_code, error_reason):
    print(f"📝 Logging error for {license_code}")

    row = [
        license_code,
        "NULL",
        "NULL",
        "NULL",
        "NULL",
        "ERROR",
        error_reason[:300]  # prevent huge stack traces
    ]

    sheet.append_row(row, value_input_option="USER_ENTERED")

def wait_for_revenue_or_empty(driver, timeout=6):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                d.find_elements(By.XPATH, "//i[contains(@class,'fl-delete')]")
                or "No data" in d.page_source
            )
        )
    except:
        return False

    rows = driver.find_elements(By.XPATH, "//i[contains(@class,'fl-delete')]")
    return len(rows) > 0

# Here we start to loop in the LC in the site 

LICENSE_CODES = REGION_CONFIG[REGION]["license_codes"]

for code in LICENSE_CODES:
    print(f"\n▶ Processing {code}")
    try:
        # Step A: Search and land on result
        search_by_license(driver, wait, code)
        
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
        # open_data_platform(driver, wait)
        # click_data_management(wait)
        # click_revenue_mapping(wait)

        # We are not using above three as im unable to open sidebar so we will directly use the url 
        account_id = code

        go_to_revenue_mapping(driver, wait, account_id)

        WebDriverWait(driver, 10).until(
            lambda d: (
                d.find_elements(By.XPATH, "//i[contains(@class,'fl-delete')]")
                or "No data" in d.page_source
            )
        )

        rows = extract_revenue_mapping_data(driver, wait, code)
        append_to_sheet(sheet, rows)

        print(f"✅ Success for {code}")


    except Exception as e:
        error_msg = str(e)
        print(f"❌ Failed to process {code}. Error: {error_msg}")

        log_error_to_sheet(
            sheet,
            code,
            error_msg
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