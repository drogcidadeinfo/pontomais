import os
import time
import logging
import shutil
import tempfile
from datetime import datetime, timedelta

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException

# Logging setup
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Environment variables
user = os.getenv("user")
password = os.getenv("password")

if not user or not password:
    raise ValueError("Environment variables 'user' and/or 'password' not set.")

# Calculate date range
today = datetime.today()
report_date = today - timedelta(days=1)
start_date = report_date - timedelta(days=1) if report_date.weekday() == 6 else report_date
date_range_str = f"{start_date.strftime('%d/%m/%Y')} - {report_date.strftime('%d/%m/%Y')}"

# Helpers
def setup_driver(download_dir):
    user_data_dir = tempfile.mkdtemp(prefix="chrome-profile-")
    chrome_options = Options()
    chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(options=chrome_options)
    return driver, user_data_dir

def safe_click(driver, by, selector, timeout=10):
    element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, selector)))
    element.click()
    return element

def retry_click(driver, by, selector, retries=3, timeout=10):
    for attempt in range(retries):
        try:
            return safe_click(driver, by, selector, timeout)
        except ElementClickInterceptedException:
            time.sleep(1)
    raise TimeoutException(f"Failed to click {selector} after {retries} retries.")

def get_latest_downloaded_file(download_dir, extensions=(".xls", ".xlsx")):
    files = [f for f in os.listdir(download_dir)
             if f.endswith(extensions) and not f.endswith('.crdownload')]
    if not files:
        return None
    files.sort(key=lambda f: os.path.getmtime(os.path.join(download_dir, f)))
    return os.path.join(download_dir, files[-1])

# Main automation
def run_download():
    download_dir = os.getcwd()
    driver, user_data_dir = setup_driver(download_dir)

    try:
        logging.info("Accessing login page")
        driver.get("https://app2.pontomais.com.br/login")
        WebDriverWait(driver, 10).until(
            lambda x: x.execute_script("return document.readyState === 'complete'")
        )

        logging.info("Filling login credentials")
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#container-login > div.left-content > div > div:nth-child(4) > div:nth-child(1) > login-form > pm-form > form > div > div > div:nth-child(1) > pm-input > div > div > pm-text > div > input"))
        ).send_keys(user)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#container-login > div.left-content > div > div:nth-child(4) > div:nth-child(1) > login-form > pm-form > form > div > div > div.password-field.pm-input-form.col-sm-18.col-xs-18 > pm-input > div > div > pm-password > div > input"))
        ).send_keys(password)
        safe_click(driver, By.CSS_SELECTOR, "#container-login > div.left-content > div > div:nth-child(4) > div:nth-child(1) > login-form > pm-button.size.mt-3.pm-spining-btn > button > span > span")

        logging.info("Waiting for dashboard")
        time.sleep(5)

        logging.info("Navigating to 'Relatórios'")
        safe_click(driver, By.CSS_SELECTOR, "body > app-mfe-remote > app-side-nav-outer-toolbar > dx-drawer > div > div.dx-drawer-panel-content > app-side-navigation-menu > div > dx-tree-view:nth-child(1) > div > div > div > div.dx-scrollable-content > ul > li:nth-child(9) > div > div.dx-template-wrapper.dx-item-content.dx-treeview-item-content > a")

        logging.info("Selecting 'Auditoria'")
        safe_click(driver, By.XPATH, "/html/body/app-mfe-remote/app-side-nav-outer-toolbar/dx-drawer/div/div[2]/dx-scroll-view/div[1]/div/div[1]/div[2]/div[1]/app-container/reports/div/div[1]/div/pm-card/div/div[2]/pm-form/form/div[2]/div/div[1]/pm-input/div/div/pm-select/div/ng-select/div/span")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/ng-dropdown-panel/div[1]/div/input"))).send_keys("Auditoria")
        time.sleep(1)
        retry_click(driver, By.XPATH, "/html/body/ng-dropdown-panel/div[2]/div[2]/div/div/div/div[2]/span")

        logging.info("Selecting information to include in the report")
        safe_click(driver, By.XPATH, "/html/body/app-mfe-remote/app-side-nav-outer-toolbar/dx-drawer/div/div[2]/dx-scroll-view/div[1]/div/div[1]/div[2]/div[1]/app-container/reports/div/div[2]/div[1]/div/div[1]/pm-button/button/span")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/ngb-modal-window/div/div/pm-modal-multi-select-modal/div[2]/div/div/div[1]/pm-form/form/div[2]/div/div/pm-input/div/div/pm-checkbox/ul/li/label/input"))
        )
        safe_click(driver, By.XPATH, "/html/body/ngb-modal-window/div/div/pm-modal-multi-select-modal/div[2]/div/div/div[1]/pm-form/form/div[2]/div/div/pm-input/div/div/pm-checkbox/ul/li/label/input")
        safe_click(driver, By.XPATH, "/html/body/ngb-modal-window/div/div/pm-modal-multi-select-modal/div[2]/div/div/div[2]/pm-button/button/span")
        time.sleep(2)

        logging.info("Setting date range")
        date_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "pm-date-range.pm-input > div:nth-child(2) > input:nth-child(1)"))
        )
        date_input.clear()
        date_input.send_keys(date_range_str)

        logging.info("Applying filter")
        safe_click(driver, By.CSS_SELECTOR, ".pm-dropdown > pm-button:nth-child(1) > button:nth-child(1) > span:nth-child(1)")

        logging.info("Downloading report")
        retry_click(driver, By.ID, "relatorios-baixar-xls")
        time.sleep(30)

        downloaded_file = get_latest_downloaded_file(download_dir)
        if downloaded_file:
            file_size = os.path.getsize(downloaded_file)
            logging.info(f"✅ Download complete: {downloaded_file} ({file_size} bytes)")
        else:
            logging.error("❌ Download failed: no file found.")

    except Exception as e:
        logging.exception(f"Error occurred: {e}")

    finally:
        driver.quit()
        shutil.rmtree(user_data_dir, ignore_errors=True)

if __name__ == "__main__":
    run_download()
