'''
    Things to change when we switch between India, GLobal, KSA
    1) Below Client inside init_google_sheet function change the sheet name
    2) In the step 0 below the Driver varibale change between driver.get 
    3) In the step 2 while direct url for publisher change between driver.get
    4) Inside the go_to_custom_events function we switch between the url
    5) In the step 5 we have urls to directly go to publisher after data is fetching
    6) This is pretty main we have to change the LC based on the global, india, ksa
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

def init_google_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(
        "/Users/admin/Desktop/Product_Adaption_Data/Credential File/mycred-googlesheet.json", scopes=scopes
    )
    client = gspread.authorize(creds)
    # India
    # sheet = client.open_by_key("1sathL7caATX3PnV2urKhp8UpCLOcwdTIxzkJkA4Kar4").worksheet("Custome Event Script Data 3")
    # Global
    sheet = client.open_by_key("1sathL7caATX3PnV2urKhp8UpCLOcwdTIxzkJkA4Kar4").worksheet("Custome Event Script Data 3")
    # KSA
    # sheet = client.open_by_key("1sathL7caATX3PnV2urKhp8UpCLOcwdTIxzkJkA4Kar4").worksheet("Custome Event Script Data 3")
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
# India
driver.get("https://dashboard.in.webengage.com/admin")
# Global
driver.get("https://dashboard.webengage.com/admin")
# KSA
driver.get("https://dashboard.ksa.webengage.com/admin")

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
            # India
            # driver.get("https://dashboard.in.webengage.com/admin/publisher.html?action=list")
            # Global
            driver.get("https://dashboard.webengage.com/admin/publisher.html?action=list")
            # KSA
            # driver.get("https://dashboard.ksa.webengage.com/admin/publisher.html?action=list")

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
        driver.execute_script(
            "arguments[0].value = '~32537i7'; arguments[0].dispatchEvent(new Event('change'));", 
            role_dropdown
        )
        
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
    print("⏳ Opening Data Platform sidebar (hybrid React-safe)...")

    li = wait.until(
        EC.presence_of_element_located((By.ID, "nav-data-platform"))
    )

    head = li.find_element(By.XPATH, ".//div[contains(@class,'menu__group__head')]")

    # Always scroll first
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", head)
    time.sleep(0.2)

    # 1️⃣ Try normal JS click (React listener)
    driver.execute_script("arguments[0].click();", head)
    time.sleep(0.5)

    # 2️⃣ Check if React actually opened it
    class_now = li.get_attribute("class")
    if "menu__group--is-active" in class_now:
        print("✅ Data Platform opened via click")
        return

    print("⚠️ Click did not open menu, forcing state...")

    # 3️⃣ Force class (DOM state)
    driver.execute_script("""
        arguments[0].classList.add('menu__group--is-active');
    """, li)

    # 4️⃣ Dispatch synthetic click event (React fallback)
    driver.execute_script("""
        arguments[0].dispatchEvent(
            new MouseEvent('click', { bubbles: true })
        );
    """, head)

    time.sleep(0.5)

    # Final confirmation
    if "menu__group--is-active" in li.get_attribute("class"):
        print("✅ Data Platform opened via hybrid fallback")
    else:
        raise Exception("❌ Failed to open Data Platform sidebar")
    
def go_to_custom_events(driver, wait, account_id):
    # India
    # url = f"https://dashboard.in.webengage.com/accounts/{account_id}/data-management/events/attributes"
    # Global
    url = f"https://dashboard.webengage.com/accounts/{account_id}/data-management/events/attributes"
    # KSA
    # url = f"https://dashboard.ksa.webengage.com/accounts/{account_id}/data-management/events/attributes"
    driver.get(url)
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "table__row")))
    print("✅ Landed on Custom Events")



def click_data_management(wait):
    print("⏳ Clicking Data Management...")

    data_management_xpath = (
        "//a[contains(@href,'/data-management/system/attributes') and .//span[text()='Data Management']]"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, data_management_xpath))
    ).click()

    print("✅ Clicked Data Management")

def click_custom_events(wait):
    print("⏳ Clicking Custom Events tab...")

    custom_events_xpath = (
        "//a[contains(@href,'/data-management/events/attributes') and normalize-space()='Custom Events']"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, custom_events_xpath))
    ).click()

    print("✅ Custom Events opened")

def extract_custom_events_page(driver, license_code):
    rows_data = []
    # Wait for at least one row to be present before scraping
    has_data = wait_for_table_or_empty(driver)

    if not has_data:
        return []

    rows = driver.find_elements(By.XPATH, "//tr[contains(@class,'table__row')]")

    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, "./td")
            if len(cells) < 8: continue 

            # 1. Event Name 
            event_name = cells[0].find_element(By.XPATH, ".//span[contains(@class,'text-ellipsis')]").get_attribute("title").strip()

            # 2. Is PII / Data Type / Personalization
            is_pii = cells[1].text.strip() or "NO"
            data_type = cells[2].text.strip() or "NULL"
            personalization = cells[3].text.strip() or "Disabled"

            # 3. Status Columns (Website, Android, iOS, Others)
            def get_status(cell):
                try:
                    return cell.find_element(By.CLASS_NAME, "status-label").text.strip()
                except:
                    return "NULL"

            website = get_status(cells[4])
            android = get_status(cells[5])
            ios = get_status(cells[6])
            others = get_status(cells[7])

            rows_data.append([
                license_code, event_name, is_pii, data_type, 
                personalization, website, android, ios, others
            ])
        except Exception as e:
            print(f"⚠️ Skipping a row due to error: {e}")
            continue

    return rows_data

def extract_all_custom_events(driver, license_code):
    print("📥 Extracting Custom Events (all pages)...")
    all_data = []
    page_number = 1

    while True:
        print(f"📄 Scraping Page {page_number}...")
        page_data = extract_custom_events_page(driver, license_code)
        all_data.extend(page_data)

        try:
            # Check for the "Next" button link
            next_parent = driver.find_element(By.XPATH, "//li[contains(@class,'pagination__next')]")
            
            # If the parent <li> has 'is-disabled', we are on the last page
            if "is-disabled" in next_parent.get_attribute("class"):
                print("🏁 Reached last page.")
                break

            next_link = next_parent.find_element(By.TAG_NAME, "a")
            driver.execute_script("arguments[0].click();", next_link)
            
            # Short wait for the table to update
            time.sleep(2)
            page_number += 1
            
        except Exception as e:
            print(f"ℹ️ Pagination ended or failed: {e}")
            break

    print(f"✅ Total Extracted: {len(all_data)} events")

    if not all_data:
        print("⚠️ No custom events found — inserting NO_DATA row")
        return [[
            license_code,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA"
        ]]

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

def wait_for_table_or_empty(driver, timeout=6):
    """
    Waits for:
    - table rows
    - empty state
    Returns True if rows exist, False otherwise
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                d.find_elements(By.XPATH, "//tr[contains(@class,'table__row')]")
                or "No data" in d.page_source
            )
        )
    except:
        return False

    rows = driver.find_elements(By.XPATH, "//tr[contains(@class,'table__row')]")
    return len(rows) > 0



LICENSE_CODES = [
    
    "in~~15ba205d1","in~~10a5cbb1d","in~311c4742"
]

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
            # open_data_platform(driver, wait)
            # click_data_management(wait)
            # click_custom_events(wait)
            
            # We are not using above three as im unable to open sidebar so we will directly use the url 
            account_id = code 

            go_to_custom_events(driver, wait, account_id)

            custom_event_rows = extract_all_custom_events(driver, code)
            append_to_sheet(sheet, custom_event_rows)

        except Exception as e:
            error_msg = str(e)
            print(f"❌ Failed to process {code}. Error: {error_msg}")

            log_error_to_sheet(
                sheet,
                code,
                stage="CUSTOM_EVENTS",
                error_reason=error_msg
            )
    
    finally:
        # Step E: Cleanup for next iteration
        # Close extra tabs and go back to the list
        if len(driver.window_handles) > 1:
            driver.close() # Closes current (Edit) tab
            driver.switch_to.window(main_window)
        
        # India
        # driver.get("https://dashboard.in.webengage.com/admin/publisher.html?action=list")
        # Global
        driver.get("https://dashboard.webengage.com/admin/publisher.html?action=list")
        # KSA
        # driver.get("https://dashboard.ksa.webengage.com/admin/publisher.html?action=list")
        time.sleep(2)
