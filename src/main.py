from excel_parser.excel_parser import *
from web_scraper.web_scraper import mkt_beta_list
from web_scraper.web_scraper import underlying_price_at_time_of_trade
import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from pandas.tseries.offsets import BMonthBegin


def main():
    # print(project_root)
    # print(data_input_directory)
    df = format_data(column_headers)
    save_data(df, data_output_directory)
    insert_line_after(data_output_directory, 'Sheet1', 1, first_line_data)


def format_data(column_headers):
    number_of_headers = len(column_headers)

    # Create an empty DataFrame with 40 columns
    columns = [column_headers[i] for i in range(1, number_of_headers)]
    df = pd.DataFrame(columns=columns)

    # Fill specific columns with initial values
    df['option Expiration date'] = option_expiration_date
    df['Strike'] = strike
    df['underlying symbol'] = underlying_symbol
    df['Type'] = stock_type
    df['Qty'] = quantity
    df['Trade Price/premium'] = price
    df['premium'] = premium
    df['trade date-entered?'] = date_of_extraction
    df['days till exp (trade date)'] = days_till_exp_date
    df['days till exp (current)'] = days_till_exp_date_current
    df['underlying price at time of trade'] = underlying_price_at_time_of_trade
    df['mkt beta'] = mkt_beta_list
    df.insert(0, 'check date >>', '')  # or use an empty string: ''



    return df
    # Specify the Excel file path


def save_data(data_frame, saving_directory):
    # Save the DataFrame to an Excel file
    data_frame.to_excel(saving_directory, index=False)
    print(f'Data saved to {saving_directory}')
    print('************ Processes completed successfully! ************')

def insert_line_after(file_path, sheet_name, row_number, data):
    # Load the existing workbook
    wb = load_workbook(file_path)

    # Select the desired sheet
    sheet = wb[sheet_name]

    # Shift existing rows down to make space for the new line
    sheet.insert_rows(row_number)

    # Write the data to the new line
    for col_num, value in enumerate(data, start=1):
        sheet.cell(row=row_number, column=col_num, value=value)

    # Save the changes
    wb.save(file_path)

# --- lists and other data --- #
# Get the current date
current_date = datetime.now()

# Calculate the first business day of the current month and normalize to midnight
first_business_day_current_month = pd.date_range(start=current_date, periods=1, freq=BMonthBegin()).normalize()[0]

# Format the result to mm/dd/yy without the time
formatted_result = first_business_day_current_month.strftime('%m/%d/%y')

date_string = formatted_result
date_format = "%m/%d/%y"

# Convert string to date
date_object = datetime.strptime(date_string, date_format)

# Format date as a string
formatted_date = date_object.strftime('%m/%d/%y')

print(formatted_date)


timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
header_time_stamp = datetime.now().strftime('%m/%d/%Y')
name_with_web_scraper = 'TDA_YAHOO_DATA_'
name_without_web_scraper = 'TDA_DATA_'
excel_filename = f'{name_with_web_scraper}{timestamp}.xlsx'
data_output_directory = os.path.join(project_root, "data", "output", excel_filename)

column_headers = [
    'check date >>', header_time_stamp, 'trade date-entered?', 'option Expiration date', 'days till exp (trade date)',
    'days till exp (current)', 'order expiration date "time in force"', 'days till expiration (if an order)',
    'Strike', 'underlying symbol',
    'underlying price at time of trade', 'otm at time of trade', 'underlying price, current', 'otm, current.',
    '$ amount of stock itm can be called (-) or put (+)', 'weight', 'weighted otm', 'mkt beta', 'Type', 'mkt beta* mkt price*contracts', 'Qty',
    'mkt price *number of contracts', 'Trade Price/premium', 'trade price as percent of notional',
    'annual yield at strike at time of trade', 'yield on cost at time of trade', 'multiple on cost',
    'yield at current mkt price at time of trade', 'premium', 'contracted in august', 'contracted in september',
    'contracted in october', 'contracted in november', 'contracted in december', 'cash if exercised', 'days >>',
    '10', '17', '24', '31', '38', '45', '50', '73', '101',
]

# add first line
first_line_data = [header_time_stamp, formatted_date, 'Open Positions', "Total"] + [None] * (len(column_headers) - 3)




try:
    main()

except Exception as e:
    print(f'Error loading the program. {e}\nPlease try again.')
