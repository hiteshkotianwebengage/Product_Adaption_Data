# open browser + login + return driver

def init_driver(profile_name):

    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager

    import os

    project_root = os.path.dirname(
        os.path.dirname(
            os.path.abspath(__file__)
        )
    )

    profile_path = os.path.join(
        project_root,
        profile_name
    )

    options = webdriver.ChromeOptions()

    options.add_argument("--start-maximized")
    options.add_argument(
        f"--user-data-dir={profile_path}"
    )

    driver = webdriver.Chrome(
        service=Service(
            ChromeDriverManager().install()
        ),
        options=options
    )

    return driver