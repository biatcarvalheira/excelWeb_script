from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from excel_parser.excel_parser import *
import sys
import threading


def make_request_firefox(url):
    driver = webdriver.Firefox()

    # Open the specified URL
    driver.get(url)

    # Sleep for 1 second
    time.sleep(1)

    # Set the success flag to True
    success = True

    # Print a message
    print('Keep Chrome Window Open')

    # Return the driver object and the success flag
    return driver, success


def make_request_headless(url):
    print('and Extracting data from website...')
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')  # Run Chrome in headless mode
    options.add_argument('--disable-gpu')  # Disable GPU usage (may prevent some issues)

    # Specify the binary location if needed (e.g., for Chrome in a specific folder)
    # options.binary_location = folder_path

    driver = webdriver.Chrome(options=options)
    driver.get(url)

    time.sleep(1)
    success = True
    return driver, success

def loading_bar(progress, total_symbols):
    percentage = int(progress / total_symbols * 100)
    sys.stdout.write('\r')
    sys.stdout.write(f"[{'=' * progress}{' ' * (total_symbols - progress)}] {percentage}%")
    sys.stdout.flush()

def load_and_click(driver, xpath):
    try:
        wait = WebDriverWait(driver, 30)
        element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        label_content = element.text
        time.sleep(2)
        element.click()
        print('Selecting... ', label_content)
    except Exception as e:
        print("Unable to click element.")


    time.sleep(2)


def insert_text(driver, xpath, input_text):
    wait = WebDriverWait(driver, 2)
    element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    element.send_keys(input_text)


underlying_price = []
mkt_beta_list = []


def scrape_for_data():
    global loading_thread, index
    error_list = []
    driver, success = make_request_headless('https://www.cnbc.com/quotes/AAPL')
    if success:
        total_symbols = len(underlying_symbol)

        for index, u in enumerate(underlying_symbol, start=1):
            symbol_url = 'https://www.cnbc.com/quotes/' + u

            #print(symbol_url)
            driver.get(symbol_url)
            driver.execute_script("window.scrollTo(0, -800);")

            # --- open value ---- #
            try:
                open_value = WebDriverWait(driver, 15).until(
                    EC.visibility_of_element_located((By.XPATH,
                                                      '/html/body/div[2]/div/div[1]/div[3]/div/div[2]/div[1]/div[5]/div[2]/section/div[1]/ul/li[1]/span[2]'))
                )
                underlying_price.append(open_value.text)
            except:
                underlying_price.append('Data not available')
                error_list.append('Open Value: ' + u)

            # --- mkt beta value ---- #
            try:
                mkt_beta_value = WebDriverWait(driver, 15).until(
                    EC.visibility_of_element_located((By.XPATH,
                                                      '/html/body/div[2]/div/div[1]/div[3]/div/div[2]/div[1]/div[5]/div[2]/section/div[1]/ul/li[14]/span[2]'))
                )
                mkt_beta_list.append(mkt_beta_value.text)
            except:
                mkt_beta_list.append('Data not available')
                error_list.append("Mkt Beta: " + u)
                continue

            # Update loading bar dynamically
            loading_bar(index, total_symbols)

        driver.quit()

        # Move the cursor to the next line after the loading bar
        print()

    print('Could not load the webpages for', error_list)







scrape_for_data()
