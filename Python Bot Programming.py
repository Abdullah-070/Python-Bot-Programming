import time
import threading

# web driver to control automation:
from selenium import webdriver
# to locate elements on webpage:
from selenium.webdriver.common.by import By
# to run chrome browser window:
from selenium.webdriver.chrome.service import Service
# to wait for certain conditions to become true:
from selenium.webdriver.support.ui import WebDriverWait
# to define conditions:
from selenium.webdriver.support import expected_conditions as EC
# to give access to google api(s) using service account
from oauth2client.service_account import ServiceAccountCredentials
# to manage chromedriver:
from webdriver_manager.chrome import ChromeDriverManager
# to interact with sheets
import gspread

# google sheets setup
def setup_google_sheets(credentials_file):
    # taking permission to access drive and google sheets
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    # creating credentials using credentials file dowloaded and scope defined above
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
    # authorizing the client
    client = gspread.authorize(creds)
    # creating sheet element by using key of particular google sheet
    sheet = client.open_by_key('1ridua9WwGAFYavw2hsdOoTE7LD9V9XphbEBLHcj9kBQ').sheet1
    return sheet

# Function to fill the form using Selenium
def fill_form(name, email, sheet, row_number):
    driver = None  
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        driver.get("https://tally.so/r/waDMG2")

        # pause execution to load entire page
        time.sleep(2)

        # Wait for the name input field to be present
        name_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Name"]'))
        )
        name_input.send_keys(name)

        # Wait for the email input field to be present
        email_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Email"]'))  
        )
        email_input.send_keys(email)

        # Submitting the form
        submit_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[type="submit"]'))
        )
        submit_button.click()

        # Wait for a confirmation or form reload
        confirmation_xpath = '//*[@id="__next"]/div/main/section/div/h1'  
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, confirmation_xpath)))

        # Mark the row as done in the Google Sheet
        sheet.update_cell(row_number, 3, "Done")

    except Exception as e:
        print(f"Error processing row {row_number}: {e}")
    
    finally:
        if driver:  # Check if driver was initialized
            driver.quit()

# Main function to start processing rows
def process_rows(start_row, sheet, num_threads):
    rows = sheet.get_all_values()[start_row - 1:]  
    threads = []

    for row_num, row_data in enumerate(rows, start=start_row):
        name = row_data[0]
        email = row_data[1]
        
        # Check if name or email is missing
        if not name or not email:
            print(f"Skipping row {row_num}: missing name or email")
            continue

        if len(row_data) > 2 and row_data[2] == "Done":
            continue  # Skip if already marked as Done

        # Create a thread for each form submission
        t = threading.Thread(target=fill_form, args=(name, email, sheet, row_num))
        threads.append(t)
        t.start()

        # Limit the number of threads running at the same time
        if len(threads) >= num_threads:
            for t in threads:
                t.join()  # Wait for all threads to finish before starting new ones
            threads = []

    # Ensure all threads are finished before ending
    for t in threads:
        t.join()

# Main function
if __name__ == "__main__":
    
    # Google Sheets credentials and document
    credentials_file = 'credentials.json'  

    # Set up Google Sheets
    sheet = setup_google_sheets(credentials_file)

    # Ask user for input
    initial_row_number = int(input("Enter the starting row number: "))
    threads_number = int(input("Enter the number of threads to run: "))

    # Start processing the rows with multithreading
    process_rows(initial_row_number, sheet, threads_number)