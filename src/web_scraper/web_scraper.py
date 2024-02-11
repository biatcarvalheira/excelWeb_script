import time
import random
import sys
import os
import requests
from bs4 import BeautifulSoup
# ---- to use when working in Executable ---- #
##from excel_parser.excel_parser import underlying_symbol




# Get the absolute path to the script
script_path = os.path.abspath(sys.argv[0])

# Get the directory where the script is located (folder containing scripts)
script_directory = os.path.dirname(script_path)

# ---- to use when working in IDE ---- #
#project_root = os.path.abspath(os.path.join(script_directory, ".."))

# ---- to use when working in Executable ---- #
project_root = os.path.abspath(os.path.join(script_directory))

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



def get_open_value_yahoo(source):
    element_with_data_test = source.find(attrs={'data-test': 'OPEN-value'})
    open_value = element_with_data_test.text
    return open_value

def get_mkt_beta_value_yahoo(source):
    element_with_data_test = source.find(attrs={'data-test': 'BETA_5Y-value'})
    mkt_value = element_with_data_test.text
    return mkt_value

def get_values_from_cnbc(source):
    item_label = source.find_all(class_="Summary-label")
    item_value = source.find_all(class_="Summary-value")
    open_val = ''
    beta_value = ''
    for index, i in enumerate(item_value):
        label_list_item = item_label[index].text
        if label_list_item == 'Open':
            open_val = i.text
        if label_list_item == 'Beta':
            beta_value = i.text

    return open_val, beta_value


def run_all_web(underlying_symbol_excel):
    print('UNDERLYING SYMBOL', underlying_symbol_excel)
    underlying_price_at_time_of_trade = []
    mkt_beta_list = []

    for u in underlying_symbol_excel:
        new_url = 'https://finance.yahoo.com/quote/' + u
        page_content = make_request_bs4(new_url)
        if page_content is not None:
            open_value = get_open_value_yahoo(page_content)
            mkt_value = get_mkt_beta_value_yahoo(page_content)
            underlying_price_at_time_of_trade.append(open_value)
            mkt_beta_list.append(mkt_value)
        else:
            new_url = 'https://www.cnbc.com/quotes/' + u
            page_content = make_request_bs4(new_url)
            time.sleep(1)
            if page_content is not None:
                open_value, mkt_value = get_values_from_cnbc(page_content)
                underlying_price_at_time_of_trade.append(open_value)
                mkt_beta_list.append(mkt_value)
            else:
                underlying_price_at_time_of_trade.append('N/A')
                mkt_beta_list.append('N/A')
                continue
    return underlying_price_at_time_of_trade, mkt_beta_list

