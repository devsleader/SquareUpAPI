import os
import time
import csv
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Load environment variables
load_dotenv()

email = os.getenv("SQUARE_EMAIL")
password = os.getenv("SQUARE_PASSWORD")

if not email or not password:
    print("[ERROR] Environment variables SQUARE_EMAIL and SQUARE_PASSWORD are not set.")
    exit(1)

# Google Sheets setup
def setup_google_sheets():
    """
    Authenticate and connect to Google Sheets.
    """
    print("[DEBUG] Setting up Google Sheets API connection...")
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(os.getenv("CREDENTIALS_JSON"), scope)
    client = gspread.authorize(creds)
    print("[DEBUG] Google Sheets API setup complete.")
    return client

def upload_csv_to_drive(download_directory):
    """
    Upload the downloaded CSV file to Google Drive with the same name (without extension).
    """
    print("[DEBUG] Searching for CSV files in the download directory...")
    files = os.listdir(download_directory)
    csv_files = [f for f in files if f.endswith(".csv")]
    
    if not csv_files:
        print("[ERROR] No CSV file found in the download directory.")
        return None
    
    # Get the most recent CSV file by creation time
    csv_file = max(csv_files, key=lambda f: os.path.getctime(os.path.join(download_directory, f)))
    file_path = os.path.join(download_directory, csv_file)
    print(f"[DEBUG] Found CSV file: {file_path}")

    # Remove the `.csv` extension from the file name
    file_name_without_extension = os.path.splitext(csv_file)[0]
    print(f"[DEBUG] Using file name without extension: {file_name_without_extension}")

    # Upload to Google Drive with the same name as the file (without extension)
    service = build('drive', 'v3', credentials=ServiceAccountCredentials.from_json_keyfile_name(os.getenv("CREDENTIALS_JSON"), ["https://www.googleapis.com/auth/drive.file"]))
    file_metadata = {
        'name': file_name_without_extension,  # Use the downloaded file's name without extension
        'mimeType': 'application/vnd.ms-excel'
    }
    media = MediaFileUpload(file_path, mimetype='text/csv', resumable=True)
    uploaded_file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()

    print(f"[DEBUG] File uploaded to Google Drive. File ID: {uploaded_file['id']}")
    return uploaded_file['id'], file_path, file_name_without_extension  # Return file ID, path, and file name without extension

def create_new_sheet_and_import_csv(file_path, sheet_name):
    """
    Create a new sheet and import the CSV data into it, after deleting the last sheet if it exists.
    """
    print(f"[DEBUG] Creating new sheet and importing CSV data...")
    
    # Connect to Google Sheets API
    client = setup_google_sheets()
    spreadsheet = client.open('Admin1')  # Replace with your actual Google Sheet name

    # Delete the last sheet before proceeding
    delete_last_sheet(spreadsheet)

    # Read the CSV data into memory and determine the number of rows and columns
    with open(file_path, 'r', newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        data = list(reader)
    
    num_rows = len(data)
    num_cols = len(data[0]) if num_rows > 0 else 0

    print(f"[DEBUG] Creating new worksheet with {num_rows} rows and {num_cols} columns.")
    
    # Create a new worksheet with the required number of rows and columns
    spreadsheet.add_worksheet(title=sheet_name, rows=str(num_rows), cols=str(num_cols))
    print(f"[INFO] Created new sheet '{sheet_name}' with {num_rows} rows and {num_cols} columns.")

    # Get the newly created worksheet
    worksheet = spreadsheet.worksheet(sheet_name)

    # Update the new sheet with the CSV data
    worksheet.update('A1', data)  # Start inserting data from A1
    print(f"[INFO] Data successfully imported into the new sheet: {sheet_name}")

def delete_last_sheet(spreadsheet):
    """
    Deletes the last sheet in the spreadsheet.
    """
    sheets = spreadsheet.worksheets()
    if sheets:
        last_sheet = sheets[-1]  # Get the last sheet
        spreadsheet.del_worksheet(last_sheet)  # Delete the last sheet
        print(f"[INFO] Deleted last sheet: {last_sheet.title}")
    else:
        print("[INFO] No sheets found to delete.")

# Set up Chrome options for Selenium
download_directory = os.path.join(os.getcwd(), "download Sales")
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

def wait_for_download(download_directory):
    """
    Wait until the CSV file is fully downloaded.
    """
    print("[DEBUG] Waiting for CSV file to be downloaded...")
    while True:
        # List files in the download directory
        files = os.listdir(download_directory)
        # Look for any CSV file that's downloading (Chrome's .crdownload extension)
        downloading = [f for f in files if f.endswith(".crdownload")]
        if not downloading:
            # Once the download is complete, check for the CSV file
            csv_files = [f for f in files if f.endswith(".csv")]
            if csv_files:
                print(f"[DEBUG] Download complete. File: {csv_files[-1]}")
                return os.path.join(download_directory, csv_files[-1])
        time.sleep(1)  # Check every second

try:
    # Step 1: Open the login page
    print("[DEBUG] Opening login page...")
    driver.get("https://app.squareup.com/dashboard/sales/reports/item-sales")
    print("[INFO] Opened the login page.")

    # Step 2: Wait for the email input field to become visible
    print("[DEBUG] Waiting for the email input field...")
    email_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "mpui-combo-field-input"))
    )
    email_field.send_keys(email)
    email_field.send_keys(Keys.RETURN)

    # Step 3: Wait for the password field and enter password
    print("[DEBUG] Waiting for the password field...")
    password_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "password"))
    )
    password_field.send_keys(password)
    driver.find_element(By.NAME, "sign-in-button").click()

    # Step 4: Handle "Remind me next time" and "Continue to Square"
    print("[DEBUG] Handling 'Remind me next time' and 'Continue to Square'...")
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "2fa-post-login-promo-sms-remind-me-btn"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "2fa-post-login-promo-opt-out-modal-continue"))).click()

    # Step 5: Navigate to the sales report page
    print("[DEBUG] Navigating to the sales report page...")
    driver.get("https://app.squareup.com/dashboard/sales/reports/item-sales")

    time.sleep(25)  # Wait for the page to load fully
    
    

    # Step 6: Click on the date selector and set the range
    print("[DEBUG] Clicking on the date selector and setting the range...")
    date_selector_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ember87")))
    date_selector_button.click()

    start_date = (datetime.now() - timedelta(days=30)).strftime("%m/%d/%Y")  # 30 days ago
    start_date_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ember137")))
    start_date_field.clear()
    start_date_field.send_keys(start_date)

    end_date = datetime.now().strftime("%m/%d/%Y")  # Today's date
    end_date_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ember139")))
    end_date_field.clear()
    end_date_field.send_keys(end_date)
    end_date_field.send_keys(Keys.RETURN)  # Submit

    time.sleep(8)  # Wait for the data to load

    # Step 7: Click on the "Export" button
    print("[DEBUG] Clicking on the Export button...")
    export_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ember283")))
    export_button.click()

    time.sleep(5)  # Wait for the export options to load

    # Step 8: Click on the "Detail CSV" button
    print("[DEBUG] Clicking on the 'Detail CSV' button...")
    detail_csv_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "market-row:nth-of-type(2) .market-export-link__label"))
)
    detail_csv_button.click()


    time.sleep(5)  # Wait for the CSV file to start downloading

    # Wait for the CSV file to finish downloading
    downloaded_file = wait_for_download(download_directory)

    # Step 9: Upload the downloaded CSV to Google Drive and import it to Google Sheets
    print("[DEBUG] Uploading the CSV file to Google Drive...")
    file_id, file_path, file_name_without_extension = upload_csv_to_drive(download_directory)  # Upload the file to Google Drive
    if file_id:
        print("[DEBUG] Importing the CSV data to Google Sheets...")
        # Create a new sheet with the same name as the CSV file and import the CSV data
        create_new_sheet_and_import_csv(file_path, file_name_without_extension)  # Use the CSV file name without extension as the new sheet name
        
except Exception as e:
    print(f"[ERROR] An error occurred: {e}")

finally:
    print("[INFO] Closing the browser.")
    driver.quit()
