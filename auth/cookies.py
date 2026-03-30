# This solely capture the cookie 

def get_session_cookies(driver):
    cookies = driver.get_cookies()
    return {c['name']: c['value'] for c in cookies}