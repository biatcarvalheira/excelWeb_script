import time
import random
import sys
import os
import requests
from bs4 import BeautifulSoup
#import for when exporting to executable
from excel_parser.excel_parser import underlying_symbol

#import for when using IDE
#from src.excel_parser.excel_parser import underlying_symbol


# Get the absolute path to the script
script_path = os.path.abspath(sys.argv[0])

# Get the directory where the script is located (folder containing scripts)
script_directory = os.path.dirname(script_path)

# IDE
project_root = os.path.abspath(os.path.join(script_directory, ".."))

#Executable
#project_root = os.path.abspath(os.path.join(script_directory))
driver_directory = os.path.join(project_root, "config", "chromedriver")


def make_request_bs4(url):
    print(f'Retrieving data from: {url}')
    response = requests.get(url)
    time.sleep(1 + 2 * random.random())

    # Check if the request was successful (status code 200)
    if response.status_code == 200:
        # Parse the HTML content using Beautiful Soup
        soup = BeautifulSoup(response.text, 'html.parser')

        # Now you can use Beautiful Soup to navigate and extract information
        # For example, let's extract all the links on the page
        return soup
    else:
        print(f"Failed to retrieve the page. Status code: {response.status_code}.Please restart the program")
        return None

def get_open_value(source):
    element_with_data_test = source.find(attrs={'data-test': 'OPEN-value'})
    open_value = element_with_data_test.text
    return open_value

def get_mkt_beta_value(source):
    element_with_data_test = source.find(attrs={'data-test': 'BETA_5Y-value'})
    mkt_value = element_with_data_test.text
    return mkt_value

underlying_price_at_time_of_trade = []
mkt_beta_list = []


for u in underlying_symbol:
    new_url = 'https://finance.yahoo.com/quote/'+u
    page_content = make_request_bs4(new_url)
    if page_content is not None:
        open_value = get_open_value(page_content)
        mkt_value = get_mkt_beta_value(page_content)
        underlying_price_at_time_of_trade.append(open_value)
        mkt_beta_list.append(mkt_value)
    else:
        break


