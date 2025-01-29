import os
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# -------------------------------------------------------------------
# >>> GLOBALS & CONSTANTS <<<
# -------------------------------------------------------------------

load_dotenv()

# Define the list of sheet tab names to process
SHEET_TAB_NAMES = [
    "POSUPK",
    "PO-BPK",
    "PendingPOsKW",
    "PendingPOMarathon",
    "POMarco"
]

# Global variable to remember our current row in the sheet
search_start_row_global = 2  # We'll update this as we find matches

CREDENTIALS_JSON = os.getenv('CREDENTIALS_JSON')  
GOOGLE_SHEET_NAME = "Admin1"


# Define column indices (1-based)
SEARCH_COLUMN_INDEX = 13  # Column M for ORDER #
STATUS_COLUMN_INDEX = 14  # Column N
NOTES_COLUMN_INDEX = 6     # Column F for "Notes"

# Environment Variables
email = os.getenv("SQUARE_EMAIL")
password = os.getenv("SQUARE_PASSWORD")
if not email or not password:
    print("[ERROR] Environment variables SQUARE_EMAIL and SQUARE_PASSWORD are not set.")
    exit(1)

# -------------------------------------------------------------------
# >>> SELENIUM & SHEETS INIT <<<
# -------------------------------------------------------------------

def init_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option(
        "prefs",
        {
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
    )
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

def connect_to_google_sheet(sheet_name, sheet_tab_name):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_JSON, scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).worksheet(sheet_tab_name)
    return sheet

def login_to_square(driver, email, password):
    driver.get("https://app.squareup.com/dashboard/items/inventory/purchase-orders")
    print("[INFO] Opened the login page.")

    # Enter email
    email_field = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "mpui-combo-field-input"))
    )
    email_field.send_keys(email)
    email_field.send_keys(Keys.RETURN)
    print("[INFO] Entered email.")

    # Enter password
    password_field = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "password"))
    )
    password_field.send_keys(password)
    driver.find_element(By.NAME, "sign-in-button").click()
    print("[INFO] Clicked 'Sign In' button.")

    # Handle optional post-login prompts
    try:
        remind_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "2fa-post-login-promo-sms-remind-me-btn"))
        )
        remind_button.click()
        print("[INFO] Clicked 'Remind me next time' button.")

        continue_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "2fa-post-login-promo-opt-out-modal-continue"))
        )
        continue_button.click()
        print("[INFO] Clicked 'Continue to Square' button.")
        
        time.sleep(30)
    except Exception as e:
        print("[WARNING] Post-login prompts not encountered or skipped.")

# -------------------------------------------------------------------
# >>> MODAL HELPER FUNCTIONS <<<
# -------------------------------------------------------------------

def close_modal(driver, timeout=10):
    """
    Closes the currently open modal by clicking the "Close" button.

    Parameters:
    - driver: Selenium WebDriver instance.
    - timeout: Maximum time to wait for the "Close" button to be clickable.
    """
    wait = WebDriverWait(driver, timeout)
    
    try:
        # 1. Click the "Close" button
        close_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "[aria-label='Close']"))
        )
        close_button.click()
        print("[INFO] 'Close' button clicked successfully.")
        
        time.sleep(1) 
        
    except (NoSuchElementException, TimeoutException) as e:
        print("[ERROR] 'Close' button not found or not clickable. Details:", str(e))
        return  # Exit the function as "Close" is mandatory

def click_save_button(driver, timeout=5):
    """
    Clicks the "Save" button if it appears after closing the modal.

    Parameters:
    - driver: Selenium WebDriver instance.
    - timeout: Maximum time to wait for the "Save" button to be clickable.
    """
    save_wait = WebDriverWait(driver, timeout)
    try:
        save_button = save_wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-test-save-changes]"))
        )
        save_button.click()
        print("[INFO] 'Save' button clicked successfully.")
    except TimeoutException:
        # "Save" button did not appear within the specified timeout
        print("[DEBUG] 'Save' button did not appear. Proceeding without clicking it.")
    except NoSuchElementException:
        # "Save" button is not present in the DOM
        print("[DEBUG] 'Save' button not found in the DOM.")

# -------------------------------------------------------------------
# >>> DRIVER CLOSE FUNCTION <<<
# -------------------------------------------------------------------

def close_driver(driver):
    """
    Closes the Selenium WebDriver session.

    Parameters:
    - driver: Selenium WebDriver instance.
    """
    try:
        driver.quit()
        print("[INFO] WebDriver session closed successfully.")
    except Exception as e:
        print(f"[ERROR] Could not close WebDriver session. Error: {e}")

# -------------------------------------------------------------------
# >>> STATUS HANDLERS <<<
# -------------------------------------------------------------------

def process_status_pending(order_number, driver, sheet, row_index):
    """Handler for Pending status."""
    print(f"[INFO] Order {order_number}: Status is 'Pending'. Skipping.")
    # No further processing required for Pending orders
    return False  # Skip to the next order

def process_status_received(order_number, driver, sheet, row_index):
    """Handler for Received status."""
    print(f"[INFO] Order {order_number}: Status is 'Received'. Processing.")
    try:
        # --- Step 1: Verify if order_number is present at the bottom ---
        try:
            # Construct an XPath to locate the <td> with the specific classes and text
            order_xpath = (
                f"//td[contains(@class, 'table-cell--selectable') "
                f"and contains(@class, 'table-cell--link') "
                f"and text()='{order_number}']"
            )

            # Wait up to 10 seconds for the element to be present in the DOM
            bottom_order_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, order_xpath))
            )
            print(f"[DEBUG] Order number '{order_number}' found at the bottom.")

        except TimeoutException:
            # If the order_number is not found within the timeout, skip processing
            print(f"[INFO] Order number '{order_number}' not found at the bottom. Skipping 'Received' status.")
            return False  # Exit the function as the order number is not present

        # --- Step 2: Click the "Received" status ---
        try:
            status_css_selector = (
                "td.table-cell.table-cell--selectable.page-inventory-list-table__cell--status."
                "page-purchase-order-list__received-color"
            )
            status_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, status_css_selector))
            )
            time.sleep(5)  # Reduced sleep time for efficiency
            status_element.click()
            print(f"[DEBUG] Clicked 'Received' status for order {order_number}.")
            time.sleep(5)  # Reduced sleep time to wait for modal to load
        except TimeoutException:
            print(f"[WARNING] 'Received' status element not found for order {order_number}.")
            return False

        # --- Step 3: Wait for the modal to load ---
        try:
            close_button_selector = "[aria-label='Close']"
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, close_button_selector))
            )
            print(f"[DEBUG] Modal loaded for order {order_number}.")
            time.sleep(5)  # Brief pause to ensure modal is fully loaded
        except TimeoutException:
            print(f"[WARNING] Modal did not load in time for order {order_number}.")
            return False

        # --- Step 4: Gather line items ---
        line_item_rows = driver.find_elements(
            By.CSS_SELECTOR, "div[data-test-po-details-line-item]"
        )

        line_items = []
        # --- Step 5: Extract Name, Qty, and Status for each line item ---
        for idx, row_element in enumerate(line_item_rows, start=1):
            try:
                name_el = row_element.find_element(
                    By.CSS_SELECTOR, "p.po-detail-sheet-row__item-name"
                )
                name_value = name_el.text.strip()

                # Locate the Qty element
                qty_el = row_element.find_element(
                    By.CSS_SELECTOR, "[data-test-details-line-item-quantity]"
                )
                qty_value = qty_el.text.strip()

                # Try to locate an <a> with "Receive" text
                try:
                    receive_link = row_element.find_element(
                        By.CSS_SELECTOR, "a[data-test-details-line-item-receive-link]"
                    )
                    line_status = receive_link.text.strip()  # Usually "Receive"
                except NoSuchElementException:
                    # If there's no "Receive" link, look for a "Received" div
                    status_div = row_element.find_element(
                        By.CSS_SELECTOR, "div[data-test-details-line-item-status]"
                    )
                    line_status = status_div.text.strip()  # Usually "Received"

                print(f"[DEBUG] Line item #{idx}: name='{name_value}', qty='{qty_value}', status='{line_status}'")
                line_items.append((name_value, qty_value, line_status))

            except Exception as e:
                print(f"[DEBUG] Could not retrieve name/qty/status for line item #{idx}. Error: {e}")
                continue

        # --- Step 6: Build a Lookup Dictionary from Sheet Data ---
        lookup_dict = {}
        sheet_values = sheet.get_all_values()  # Fetch all data at once for efficiency
        for r_idx, row in enumerate(sheet_values, start=1):
            sheet_name_value = str(row[0]).strip()  # Column A for Name
            sheet_qty_value = str(row[6]).strip()  # Column G for Qty (0-based index)
            if sheet_name_value and sheet_qty_value:
                # Normalize quantities to integers for comparison
                try:
                    sheet_qty_normalized = int(float(sheet_qty_value))
                except ValueError:
                    print(f"[DEBUG] Invalid quantity '{sheet_qty_value}' at row {r_idx}. Skipping this row.")
                    continue
                lookup_key = (sheet_name_value.lower(), sheet_qty_normalized)
                lookup_dict.setdefault(lookup_key, []).append(r_idx)

        # --- Step 7: Update the Sheet based on Line Items ---
        for (name_value, qty_value, line_status) in line_items:
            # Skip lines where status == "Receive"
            if line_status.lower() == "receive":
                print(f"[DEBUG] Skipping Name='{name_value}', Qty='{qty_value}' because status='{line_status}'.")
                continue

            # Skip empty qty
            if not qty_value:
                print("[DEBUG] Skipping empty qty value.")
                continue

            # Normalize values for comparison
            name_normalized = name_value.lower()
            try:
                qty_normalized = int(float(qty_value))
            except ValueError:
                print(f"[WARNING] Invalid qty '{qty_value}' for Name='{name_value}'. Skipping this item.")
                continue

            lookup_key = (name_normalized, qty_normalized)
            matched_rows = lookup_dict.get(lookup_key, [])

            if matched_rows:
                for r_idx in matched_rows:
                    # Move value to the Notes column
                    sheet.update_cell(r_idx, NOTES_COLUMN_INDEX, qty_normalized)
                    # Clear the Qty in column G
                    sheet.update_cell(r_idx, 7, '')
                    print(f"[INFO] Moved qty '{qty_normalized}' from row {r_idx} to Notes column for Name='{name_value}'.")
                # Optionally, remove the matched rows to prevent duplicate processing
                del lookup_dict[lookup_key]
            else:
                print(f"[WARNING] Qty value '{qty_value}' with Name='{name_value}' from modal not found in sheet.")

        time.sleep(5)
        # --- Step 8: Close the modal ---
        close_modal(driver)

        # --- Step 9: Click the "Save" button after closing the modal ---
        try:
            click_save_button(driver)
            print("[INFO] 'Save' button clicked successfully.")
        except Exception as e:
            print(f"[DEBUG] 'Save' button did not appear or could not be clicked. Error: {e}. Proceeding without clicking it.")

        print(f"[INFO] Successfully processed 'Received' order {order_number}.")
        return True
    except Exception as e:
        print(f"[WARNING] Could not process 'Received' for order {order_number}. Error: {e}")
        return False

def process_status_partially_received(order_number, driver, sheet, row_index):
    """
    Handler for 'Partially Received' status.
    - Clicks the 'Partially Received' order row.
    - Waits for the modal to appear.
    - For each line item, attempts to read Name, Qty, and Status (either 'Receive' or 'Received').
    - SKIPS lines with 'Receive' status, only moves Qty if status is 'Received' and Name matches.
    """
    print(f"[INFO] Order {order_number}: Status is 'Partially Received'. Processing.")
    try:
        # 1) Click the "Partially Received" status on the main page
        status_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,
                 "//td[contains(@class, 'page-purchase-order-list__receiving-color') "
                 "and contains(., 'Partially Received')]"
                )
            )
        )
        time.sleep(5)  # Reduced sleep time for efficiency
        status_element.click()
        print(f"[DEBUG] Clicked 'Partially Received' status for order {order_number}.")
        time.sleep(5)  # Wait for modal to load

        # 2) Wait for the modal to appear
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "[aria-label='Close']"))
        )
        print(f"[DEBUG] Modal loaded for order {order_number}.")
        time.sleep(5)  # Brief pause

        # 3) Gather line items as <div data-test-po-details-line-item="...">
        line_item_rows = driver.find_elements(
            By.CSS_SELECTOR, "div[data-test-po-details-line-item]"
        )

        line_items = []
        # 4) For each line-item div, get Name, Qty + Status
        for idx, row_element in enumerate(line_item_rows, start=1):
            try:
                name_el = row_element.find_element(
                    By.CSS_SELECTOR, "p.po-detail-sheet-row__item-name"
                )
                name_value = name_el.text.strip()

                # Locate the Qty element
                qty_el = row_element.find_element(
                    By.CSS_SELECTOR, "[data-test-details-line-item-quantity]"
                )
                qty_value = qty_el.text.strip()

                # Try to locate an <a> with "Receive" text
                try:
                    receive_link = row_element.find_element(
                        By.CSS_SELECTOR, "a[data-test-details-line-item-receive-link]"
                    )
                    line_status = receive_link.text.strip()  # Usually "Receive"
                except NoSuchElementException:
                    # If there's no "Receive" link, look for a "Received" div
                    status_div = row_element.find_element(
                        By.CSS_SELECTOR, "div[data-test-details-line-item-status]"
                    )
                    line_status = status_div.text.strip()  # Usually "Received"

                print(f"[DEBUG] Line item #{idx}: name='{name_value}', qty='{qty_value}', status='{line_status}'")
                line_items.append((name_value, qty_value, line_status))

            except Exception as e:
                print(f"[DEBUG] Could not retrieve name/qty/status for line item #{idx}. Error: {e}")
                continue

        # >>> Global approach to track the row pointer <<<
        global search_start_row_global

        # Count how many total rows the sheet has in col_values(SEARCH_COLUMN_INDEX)
        max_rows = len(sheet.col_values(SEARCH_COLUMN_INDEX)) + 1

        # 5) Only move the Qty → Notes column if status is 'Received' and Name matches
        for (name_value, qty_value, line_status) in line_items:
            # Skip empty qty
            if not qty_value:
                print("[DEBUG] Skipping empty qty value.")
                continue

            # Skip lines where status == "Receive"
            if line_status.lower() == "receive":
                print(f"[DEBUG] Skipping Name='{name_value}', Qty='{qty_value}' because status='{line_status}'.")
                continue

            matched = False
            # Search from our global pointer forward
            for r_idx in range(search_start_row_global, max_rows + 1):
                sheet_qty_value = sheet.cell(r_idx, 7).value  # column G for Qty
                sheet_name_value = sheet.cell(r_idx, 1).value  # column A for Name

                # Normalize both Name and Qty for accurate comparison
                if (str(sheet_name_value).strip().lower() == name_value.lower()) and (str(sheet_qty_value).strip() == str(qty_value).strip()):
                    # Move value to the Notes column
                    sheet.update_cell(r_idx, NOTES_COLUMN_INDEX, qty_value)
                    # Clear the Qty in column G
                    sheet.update_cell(r_idx, 7, '')
                    print(f"[INFO] Moved qty '{qty_value}' from row {r_idx} to Notes column for Name='{name_value}'.")

                    # Update the global pointer so the next item won't start over
                    search_start_row_global = r_idx + 1
                    matched = True
                    break

            if not matched:
                print(f"[WARNING] Qty value '{qty_value}' with Name='{name_value}' from modal not found in sheet.")
        time.sleep(5)
        # 6) Close the modal
        close_modal(driver)

        # 7) Click the "Save" button after closing the modal
        click_save_button(driver)

        print(f"[INFO] Successfully processed 'Partially Received' order {order_number}.")
        return True
    except Exception as e:
        print(f"[WARNING] Could not process 'Partially Received' for order {order_number}. Error: {e}")
        return False

# -------------------------------------------------------------------
# >>> ORDER STATUS DISPATCH <<<
# -------------------------------------------------------------------

def handle_order_status(order_number, driver, sheet, row_index):
    """
    Determine the status of the order and process it accordingly.
    Returns True if processing was successful, False otherwise.
    """
    try:
        # Check the status of the order
        status_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,
                "td.table-cell.table-cell--selectable.page-inventory-list-table__cell--status"
            ))
        )
        status_text = status_element.text.strip()

        # Map statuses to processing functions
        status_handlers = {
            "Pending": process_status_pending,       # we skip "Pending"
            "Partially Received": process_status_partially_received,
            "Received": process_status_received,
        }

        # Call the appropriate handler
        if status_text in status_handlers:
            return status_handlers[status_text](order_number, driver, sheet, row_index)
        else:
            print(f"[WARNING] Order {order_number}: Unrecognized status '{status_text}'. Skipping.")
            return False
    except Exception as e:
        print(f"[ERROR] Could not determine status for order {order_number}. Error: {e}")
        return False

# -------------------------------------------------------------------
# >>> MAIN LOOP: Checking Orders <<<
# -------------------------------------------------------------------

def check_order_status(sheet, driver):
    """Main function to check and process all orders in the sheet."""
    order_numbers = sheet.col_values(SEARCH_COLUMN_INDEX)[1:]  # Skip header row

    for index, order_number in enumerate(order_numbers):
        if not order_number:
            continue  # Skip empty rows

        try:
            # Search for the order in the UI
            search_input = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Search Vendor or Order #']"))
            )
            search_input.clear()
            search_input.send_keys(order_number)
            search_input.send_keys(Keys.RETURN)
            time.sleep(3)  # Wait for results to load

            # Handle the order based on its status
            if not handle_order_status(order_number, driver, sheet, index + 1):
                print(f"[INFO] Order {order_number}: Processing skipped or failed.")
        except Exception as e:
            print(f"[ERROR] Could not process order {order_number}. Error: {e}")

        time.sleep(2)  # Pause before processing next order

# -------------------------------------------------------------------
# >>> MAIN ENTRY POINT <<<
# -------------------------------------------------------------------

if __name__ == "__main__":
    print("[INFO] Starting the script...")

    try:
        driver = init_driver()
        login_to_square(driver, email, password)

        # Iterate over each sheet tab name
        for sheet_tab_name in SHEET_TAB_NAMES:
            print(f"[INFO] Processing sheet: '{sheet_tab_name}'")
            sheet = connect_to_google_sheet(GOOGLE_SHEET_NAME, sheet_tab_name)
            
            # Reset the global search start row for each sheet
            search_start_row_global = 2

            # Process orders in the current sheet
            check_order_status(sheet, driver)
    except Exception as e:
        print(f"[ERROR] An error occurred: {e}")
    finally:
        # Ensure any pending save actions are handled and driver is closed
        try:
            click_save_button(driver)  # Ensure any pending save actions are handled
        except Exception as e:
            print(f"[DEBUG] 'Save' button could not be clicked during cleanup. Error: {e}")
        close_driver(driver)
        print("[INFO] Script completed.")
