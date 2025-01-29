import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
load_dotenv()

email = os.getenv("SQUARE_EMAIL")
password = os.getenv("SQUARE_PASSWORD")

if not email or not password:
    print("[ERROR] Environment variables SQUARE_EMAIL and SQUARE_PASSWORD are not set.")
    exit(1)

# Set up custom download directory
download_directory = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(download_directory):
    os.makedirs(download_directory)

# Configure Chrome options
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_directory,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Set up ChromeDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    # Debug: Start the script
    print("[INFO] Starting the script...")

    # Step 1: Open the login page
    driver.get("https://app.squareup.com/dashboard/items")
    print("[INFO] Opened the login page.")

    # Step 2: Wait for the email input field to become visible
    email_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "mpui-combo-field-input"))
    )
    print("[DEBUG] Email input field located.")

    # Step 3: Enter email
    email_field.send_keys(email)
    print(f"[INFO] Entered email: {email}")
    email_field.send_keys(Keys.RETURN)

    # Step 4: Wait for the password field and enter password
    password_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "password"))
    )
    print("[DEBUG] Password input field located.")

    password_field.send_keys(password)
    print("[INFO] Entered password.")
    driver.find_element(By.NAME, "sign-in-button").click()
    print("[INFO] Clicked 'Sign In' button.")

    # Step 5: Handle "Remind me next time" button
    remind_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "2fa-post-login-promo-sms-remind-me-btn"))
    )
    remind_button.click()
    print("[INFO] Clicked 'Remind me next time' button.")

    # Step 6: Handle "Continue to Square" button
    continue_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "2fa-post-login-promo-opt-out-modal-continue"))
    )
    continue_button.click()
    print("[INFO] Clicked 'Continue to Square' button.")

    # Step 7: Navigate to the Square Dashboard items page
    dashboard_url = "https://app.squareup.com/dashboard/items/library"
    driver.get(dashboard_url)
    print(f"[INFO] Navigated to the dashboard page: {dashboard_url}")
    
    time.sleep(25)

    # Step 8: Click on the action button
    action_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "item-library-actions-dropdown-button-label"))
    )
    action_button.click()
    print("[INFO] Clicked on the action button.")

    # Step 9: Click on the export library button
    export_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "item-library-actions-export-row-label"))
    )
    export_button.click()
    print("[INFO] Clicked on the export library button.")

    # Step 10: Click on the export button in the modal
    final_export_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "market-button[data-test-catalog-export-modal-export]"))
    )
    final_export_button.click()
    print("[INFO] Clicked on the final export button.")

    # Step 11: Wait for the file to download
    def wait_for_download(directory, timeout=60):
        end_time = time.time() + timeout
        while True:
            files = os.listdir(directory)
            if any(file.endswith(".xlsx") for file in files):  # Look for Excel files
                print("[INFO] Excel file downloaded successfully.")
                return True
            if time.time() > end_time:
                print("[ERROR] File download timed out.")
                return False
            time.sleep(1)

    if wait_for_download(download_directory):
        print(f"[INFO] File downloaded to: {download_directory}")
    else:
        print("[ERROR] File download did not complete successfully.")

except Exception as e:
    print(f"[ERROR] An error occurred: {e}")

finally:
    print("[INFO] Closing the browser.")
    driver.quit()
       