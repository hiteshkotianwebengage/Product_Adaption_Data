'''
    Things to change when we switch between India, GLobal, KSA
    1) Below Client inside init_google_sheet function change the sheet name
    2) In the step 0 below the Driver varibale change between driver.get 
    3) In the step 2 while direct url for publisher change between driver.get
    4) Inside the go_to_Alert_events function we switch between the url
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
REGION = "INDIA"   # options: "INDIA", "GLOBAL", "KSA"

REGION_CONFIG = {
    "INDIA": {
        "base_url": "https://dashboard.in.webengage.com",
        "sheet_name": "Alert Event India",
        "publisher_url": "https://dashboard.in.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "in~~71680c2b","in~~1341062c9","in~14507c647","in~82617199","in~~71680c38","in~58adcb50","in~~991992aa","in~76aa201","in~58adcb08","in~~991992c7","in~~47b66733","in~~10a5cbc25","in~aa131650","in~aa13163a","in~11b56418d","in~11b564191","in~~2024c2b8","in~311c4663","in~76aa1a2","in~~15ba2074d","in~~c2ab3781","in~~1341062bb","in~~991992c4","in~~10a5cbc2d","in~~1341062c1","in~~99199240","in~~71680b76","in~~991992cc","in~~47b6678c","in~311c4664","in~14507c641","in~~71680c30","in~aa13164b","in~~991992a4","in~~15ba20759","in~~15ba205c0","in~~2024c231","in~76aa1ac","in~11b5641b1","in~~47b6677d","in~58adcb36","in~aa13166b","in~~991992d1","in~~1341062c2","in~~99199081","in~14507c63b","in~~c2ab3786","in~11b564246","in~~99199278","in~11b56417b","in~aa131665","in~~71680b90","in~14507c666","in~aa131632","in~76aa20d","in~311c464b","in~311c4766","in~~c2ab3761","in~~71680c4c","in~11b564177","in~11b564172","in~d3a49ad8","in~~47b66782","in~11b56417c","in~11b564181","in~~c2ab3789","in~11b5642d2","in~311c4646"
        ]
    },
    "GLOBAL": {
        "base_url": "https://dashboard.webengage.com",
        "sheet_name": "Alert Event Global",
        "publisher_url": "https://dashboard.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "~47b6574d","d3a4ab38","~47b66864","76aac96","8261829c","~10a5cad0b","58add307","~10a5cabbb","58add216","~10a5cb2a9","~2024b5d8","d3a4a457","311c5625","~134105a60","58add423","311c5642","14507cd00","d3a4a301","~15ba20153","d3a4ac1c","~1341059b6","~10a5cab6a","14507cd51","~15ba1d846","~10a5cac40","~7167db84","14507cc77","58add283","99198968","~15ba1da68","14507cda4","58add2d9","~134105a52","~7167db54","~47b66064","76aa813","~716800b0","aa13266b","826174d0","~47b65b6c","82617c25","~47b6607a","~15ba2020b","82617894","82617822","76aa800","~c2ab3242","~15ba1d691","~10a5cb63c","~c2ab3108","~134105a8c","76ab0a5","~c2ab2dd0","~15ba20105","82617757","aa131c59","311c4b69","~2024bb2d","~134105a04","14507cd4d","~13410604b","76aa78b","~71680543","~2024b99a","~71680588","~9919839a","aa131c84","76aa858","~134105251","~99198a20","~1341056a0","~2024b6c6","~15ba20116","76aa833","~991983db","76aa7a3","14507cba6","14507d14d","~71680655","82617754","~2024bb10","~c2ab2851","~99198b61","~10a5cb5d1","~10a5cb677","~10a5cb557","~c2ab313b","~134105965","~1341059c5","58adc5c7","~7168057d","~99198ab8","~15ba1ddc6","~47b65848","d3a4a32d","311c4c4b","~47b66045","~15ba201a6","76aa124","~9919868d","76a9c30","~311c4b76","58adca91","76aa76b","82617b34","~2024bada","d3a4a403","8261827a","~71680577","~99198a29","~2024bad5","~2024b291","58add69d","~134105b84","d3a4a6dd","311c558d","~c2ab3033","~15ba1ddc2","82618089","58add346","~1341061bb","d3a4a69c","58add2da","~716805d8","~15ba20214","~10a5cb6b0","~1341056bb","aa132703","~aa1321c5","~c2ab2c0c","~9919871c","~47b66614","~991981d3","14507cc74","8261786b","11b564b69","aa132225","311c4c14","311c4c11","76aa762","311c4bbb","14507cba8","58add7aa","11b564830","76aac69","76aab88","~71680627","~15ba20042","~oldetmoney","11b564720","~c2ab275a","~c2ab2bc0","~15ba1d70a","~c2ab2ba2","8261812c","~47b6665c","~1341059cb","~991989d1","~old2024c085","d3a4a36a","~10a5cac99","~oldmagma1","~oldmagma2","~oldmagma3","11b56527d","11b5646ca","~c2ab2c08","~11b5646b8","~d3a4a286","76aa7c6","76aa85d","~2024ba8b","~47b65b94","~7168053d","~c2ab3083","58addc40","76aa844","~d3a49c4c-old","76aa868","d3a4a663","~15ba2019a","~7168069d","~oldrangde","~10a5cb533","~10a5cb636","~c2ab26b8","~99198226","~2024bbb6","~99198a14","82617775","~134105aac","11b56470b","aa13264a","~c2ab3042","~99198a18","11b5650c0","311c4bc4","d3a4a420","~10a5cb20c","14507cc0a","~15ba1db98","~14507ccb9","~134105a45","~15ba200d7","~11b564836","~14507ccc0","aa131c70","~47b661c8","~7168071b","~10a5cb278","~2024bb90","~82617869","~oldUPES","14507cba1","d3a4a72a","~71680632","~47b65a1c","82617779","~10a5cb53d","311c4d24","58add3a2","~47b660d4","~134105732","~9919837c","~311c4dc3","~2024bb26","~76aab32","~c2ab323a","old~2024c1a3","826182a0","~10a5cb24b","11b56488a","~15ba20234","~15ba2063d","~47b661ab","~47b65875","99198624","14507cd97","~134105353","d3a4ab04","~ c2ab260c","~47b66257","~c2ab3091","~c2ab30aa","~15ba20080","d3a4a3b7","~aa131bd4","~134105b82","~134105a36","58add667","~47b65c2a","311c4b7c","aa131752","11b564bc3","82617bba","~716802a1","~14507d169","~d3a4a64d","~aa1321a1","~716802d0","~7168026d","~15ba1db9b","~2024b725","~2024b72c","~7168028a","~47b65bd8","~99199107","~oldid","~2024b7a2","aa132182","~10a5cb319","~15ba1dbd6","~10a5cb2a0","~1341056c2","~2024b80a","aa13210d","~11b564252","~47b65c27","14507d197","14507d153","~aa131717","311c5166","~15ba1dc69","~10a5cb33d","~c2ab2c5d","~82616dd3","~14507c905","~58adc916","~11b563c29","~311c4664","~c2ab3c28","~7167db9d","~99198b2b","~oldedelwiess","~2024b1c1","~10a5cb621","~15ba20147","~15ba1dc2d","~10a5cb283","82617aac","old~10a5cbb38","58addb9c","58adcba4","~oldsasai","~ c2ab36a7","~9919922b","~15ba1dd60","~9919879c","~15ba1dda1","~716803a0","~10a5cb41b","82617a35","~71680426","~15ba206bc","~teamGreatLearning","old~7167d2da","~10a5cb41c","d3a4b631","~c2ab3713","d3a4a554","old~76aab14","58add639","~2024b6b1","~716802c1","~c2ab2d61","~47b65736","aa132108","~13410583d","~716806c3","58add571","old~1341061b2","~2024b742","~15ba2009b","~2024b7a6","~47b65ca0","d3a4a5c9","~82617225","76abb05","~15ba1cc68","~134105ba0","82617a1a","~2024b8a7","~82617957","76aab70","d3a49b72","~10a5cac62","~991988b2","14507d028","~10a5cb41c","d3a4a4aa","311c5128","~oldaccount","82618240","76abad2","~47b66522","~15ba1cc70","~15ba20218","~7167d2da","~47b66119","~716805b9","14507d167","76aa239","14507d0c5","11b565a86","aa133168","~15ba201dc","~1341047d4","~716806db","~47b65c58","~10a5cb29d","~10a5cbb86","aa1318a4","~15ba20712","~2024a7ad","aa133154","d3a4a4c6","311c561d","~716801dd","~58add514","~2024b7a3","~71680374","~9919835a","~82617a45","d3a4a292","~9919779b","311c6080","~7167d365","~9919921c","~716806cb","76aab70","~47b666aa","~2024a7c9","~134104786","~71680bc5","76aa235","~2024b77c","~1341047d6","311c6141","d3a4b607","11b565a2a","82618a47","~10a5ca344","11b565a4b","~76aa276","~13410606b","~2024c07d","58b005a5","76ab982","58add2d9","14507d153","~47b66726","1450800cc","82618a78","76aba82","~99197808","~9919786b","76ab099","~99197879","~c2ab1d22","~134105365","~7167dd04","aa133088","~15ba20518","~11b565a1","~c2ab1d64","~99197854","d3a4a339","~c2ab3737","~134105770","~82618284","11b564400","58b005b2","76aab2b","311c6114","~15ba1cd7b","~15ba1cd57","8261895b","~2024a872","~47b64d14","~145080099","~2024a898","~10a5cb63c","~2024bb10","~99197908","~10a5cbbc3","311c6137","~145080038","d3a4b4bc","11b56595a","~10a5cbb7a","d3a4a3d2","311c617c","~1341058da","~15ba1dd60","~15ba1cd06","311c6067","~10a5cb5d1","~7167d43b","11b56597b","76ab9b6","76ab933","76aa777","76aab14","~15ba1cc97","145080012","~47b64db3","76ab966","~134104915","~9919792b","76ab97b","d3a4b5a8","145080010","~2024a887","145080030","~47b64d1c","~47b64dc6","11b565a3b","d3a4b4a3","~2024ba16","~c2ab3054","~99197905","~10a5cb448","~2024a920","~c2ab1db9","311c6015","76ab975","~2024a932","~99197925","~7167d473","~10a5ca45d","d3a4b49c","58b00497","~2024a938","d3a4b48c","~7167d481","~amplicom","~15ba1cda8","82618947.0","~TapsiMarket","58b004a0","~15ba1cdc0","~134104929","~GreenlifeEwaste","311c5018","~15ba1cd29","76ab953","~7167d478","11b565ab8","d3a4b501","311c6036","aa133008","~2024a928","76ab919","~15ba1cdc7","311c6106","~2024a8d9","~7167d459","~47b64dcd","76ab94b","~10a5ca413","~c2ab1d61","~47b64d69","11b565933","aa13306d","14507ddc0","76ab948","~99197942","311c601a","311c5ddd","~15ba1cd71","11b56593a","14507ddd3","58add3d1","~c2ab3a30","58b00572","d3a4b47a","58b00482","~10a5ca49a","58b0047d","aa131b8b","~134104910","~10a5cb607","~7167d4a7","311c50c2","76ab96b","~13410487b","311c50ac","~99198826","76aab90","311c4c52","~71680307","~15ba1dcd4","~47b64d09","~99198b33","~c2ab1d11","d3a4a33c","~c2ab1dda","aa132dcb","76ab943","58b00487","11b565931","~134104942","aa132dc1","~15ba1d007","~134104952","76ab929","~1341053c6","d3a4b479","aa132dac","~c2ab1dc3","11b565915","76ab940","58add629","~47b64c57","~2024a96d","~7167d309","311c6107","~c2ab2d96","aa131c6a","~7167d4b3","~7168061a","76ab912","~2024a973","311c5dac","~134104968","~2024a7c4","58b00457","76ab922","~7167d4b4","~10a5ca4a2"
        ]
    },
    "KSA": {
        "base_url": "https://dashboard.ksa.webengage.com",
        "sheet_name": "Alert Event KSA",
        "publisher_url": "https://dashboard.ksa.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "ksa~~15ba20526","ksa~82617412","ksa~~716809b4","ksa~14507c890","ksa~~2024c070","ksa~76aa41c","ksa~~716809bd","ksa~11b564409","ksa~~47b6652a","ksa~aa1318a1","ksa~~47b6652c","ksa~58adcd4c","ksa~11b564406","ksa~~15ba2051c","ksa~aa131897","ksa~~134106071","ksa~82617408","ksa~58adcd55","ksa~~15ba20523","ksa~~10a5cb9bd","ksa~d3a49d49","ksa~11b5643db","ksa~d3a49d4a","ksa~d3a49d44","ksa~58adcd54","ksa~aa13189b","ksa~311c489a","ksa~82617404","ksa~~2024c08a","ksa~~134106076","ksa~14507c891","ksa~~716809c9","ksa~d3a49d46","ksa~~47b66537","ksa~~134106074","ksa~~2024c07d","ksa~~2024c085","ksa~82617402","ksa~~13410607a","ksa~11b564403","ksa~~716809ba","ksa~~10a5cb9c4","ksa~~99199078","ksa~~15ba20518","ksa~~134106069","ksa~311c4892","ksa~~99199083","ksa~aa1318a0","ksa~~13410606b","ksa~11b5643d5","ksa~~134106080","ksa~58adcd47","ksa~58adcd44","ksa~11b5643d3","ksa~826173db","ksa~d3a49d41","ksa~~134106084","ksa~~99199073","ksa~~99199087","ksa~aa131893","ksa~~15ba20531","ksa~76aa3da","ksa~~2024c091","ksa~826173dc","ksa~82617401","ksa~aa131890","ksa~14507c89c","ksa~~10a5cb9d1","ksa~~2024c08d","ksa~~47b66522"
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
    
def go_to_alerts(driver, wait, account_id):
    base_url = REGION_CONFIG[REGION]["base_url"]
    url = f"{base_url}/accounts/{account_id}/alerts"

    driver.get(url)
    time.sleep(2)  # critical for modal detection

def is_alerts_locked(driver):
    # modal visible
    if "Locked feature" in driver.page_source:
        return True

    # redirected silently (very common)
    if not driver.current_url.endswith("/alerts"):
        return True

    return False

def click_data_management(wait):
    print("⏳ Clicking Data Management...")

    data_management_xpath = (
        "//a[contains(@href,'/data-management/system/attributes') and .//span[text()='Data Management']]"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, data_management_xpath))
    ).click()

    print("✅ Clicked Data Management")

def click_Alert_events(wait):
    print("⏳ Clicking Alert Events tab...")

    Alert_events_xpath = (
        "//a[contains(@href,'/data-management/events/attributes') and normalize-space()='Alert Events']"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, Alert_events_xpath))
    ).click()

    print("✅ Alert Events opened")

def wait_for_alerts_ready(driver, timeout=8):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                "Locked feature" in d.page_source
                or "/alerts" in d.current_url
            )
        )
        return True
    except:
        return False

def extract_Alert_events_page(driver, license_code):
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

def extract_alerts(driver, license_code):
    print("📥 Extracting Alerts...")

    # hard wait to allow React to paint content
    time.sleep(2)

    rows = driver.find_elements(
        By.XPATH, "//tbody/tr[contains(@class,'table__row')]"
    )

    if not rows:
        return [[
            license_code,
            "NO_DATA",
            "",
            "",
            "",
            "",
            "",
            ""
        ]]

    rows_data = []

    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, "./td")
            if len(cells) < 6:
                continue

            # 1️⃣ Alert Name (use title – most stable)
            name_container = cells[0].find_element(
                By.XPATH, ".//div[@title]"
            )
            alert_name = name_container.get_attribute("title").strip()

            # Template flag
            template = "Template" if name_container.find_elements(
                By.XPATH, ".//span[contains(@class,'template-banner')]"
            ) else ""

            # 2️⃣ Frequency
            frequency = cells[1].text.strip() or "-"

            # 3️⃣ Status
            status = cells[2].text.strip() or "-"

            # 4️⃣ Subscribers (handles '-', emails, +N)
            subscribers = cells[3].text.strip() or "-"

            # 5️⃣ Last Check
            last_check = cells[4].text.strip() or "-"

            # 6️⃣ Created On
            created_on = cells[5].text.strip() or "-"

            rows_data.append([
                license_code,
                alert_name,
                template,
                frequency,
                status,
                subscribers,
                last_check,
                created_on
            ])

        except Exception as e:
            print(f"⚠️ Skipping alert row due to parse error: {e}")
            continue

    return rows_data

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

LICENSE_CODES = REGION_CONFIG[REGION]["license_codes"]


for code in LICENSE_CODES:
    print(f"\n▶ Processing {code}")
    try:
        # Step A: Search and land on result
        search_by_license(driver, wait, code)

        # 🚫 REGION MISMATCH GUARD (same as Users / Dashboards)
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

        # Step D: Extract Alerts
        try:
            account_id = code
            
            go_to_alerts(driver, wait, account_id)
            # wait for either alerts OR locked modal
            wait_for_alerts_ready(driver)

            if is_alerts_locked(driver):
                log_error_to_sheet(
                    sheet,
                    code,
                    stage="ALERTS_LOCKED",
                    error_reason="Alerts feature not enabled for this account"
                )
                continue

            alerts_rows = extract_alerts(driver, code)
            append_to_sheet(sheet, alerts_rows)

        except Exception as e:
            log_error_to_sheet(
                sheet,
                code,
                stage="ALERTS",
                error_reason=str(e)
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
