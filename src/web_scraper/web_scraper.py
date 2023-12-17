from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from excel_parser.excel_parser import *
import sys
import os
import threading
from selenium.common.exceptions import WebDriverException

# Get the absolute path to the script
script_path = os.path.abspath(sys.argv[0])

# Get the directory where the script is located (folder containing scripts)
script_directory = os.path.dirname(script_path)

# IDE
project_root = os.path.abspath(os.path.join(script_directory, ".."))

#Executable
#project_root = os.path.abspath(os.path.join(script_directory))
driver_directory = os.path.join(project_root, "config", "chromedriver")
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


def make_request_headless_firefox(url):
    options = webdriver.FirefoxOptions()
    options.add_argument('--headless')  # Run Firefox in headless mode
    try:
        driver = webdriver.Firefox(options=options)
        driver.get(url)
        time.sleep(1)
        success = True
        return driver, success
    except WebDriverException as e:
        print(f"WebDriverException: {e}")
        success = False
        return None, success

def make_request_headless(url, chromedriver_path):
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')  # Run Chrome in headless mode
    try:
        # Specify the path to ChromeDriver executable
        driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
        driver.get(url)
        time.sleep(1)
        success = True
        return driver, success
    except WebDriverException as e:
        print(f"WebDriverException: {e}")
        success = False
        return None, success

def loading_bar(progress, total_symbols):
    percentage = int(progress / total_symbols * 100)
    sys.stdout.write('\r')
    sys.stdout.write(f"[{'=' * progress}{' ' * (total_symbols - progress)}] {percentage}%")
    sys.stdout.flush()

def load_and_click(driver, xpath):
    try:
        wait = WebDriverWait(driver, 10)
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
    driver, success = make_request_headless(driver_directory,'https://www.cnbc.com/quotes/AAPL')
    if success:
        total_symbols = len(underlying_symbol)
        # if there's no xlsx file -- stop the program
        if total_symbols == 0:
            print('No .xlsx file was found')
            sys.exit()
        else:
            print('Scraping website...')

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
