from excel_parser.excel_parser import *
from web_scraper.web_scraper import mkt_beta_list
from web_scraper.web_scraper import underlying_price_at_time_of_trade
import os
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from pandas.tseries.offsets import BMonthBegin
from openpyxl.styles import NamedStyle



def main():
    # print(project_root)
    # print(data_input_directory)
    df = format_data(column_headers)
    save_data(df, data_output_directory)
    insert_line_after(data_output_directory, 'Sheet1', 1, first_line_data)
    # format percentage cells
    format_columns(data_output_directory, 'Sheet1', ['L', 'N', 'Q', 'X', 'Y'], 'percentage', '0.00%')

    # format currency cells
    format_columns(data_output_directory, 'Sheet1', ['T', 'V'], 'currency_format', '"$"#,##0.00')





def format_data(column_headers):
    number_of_headers = len(column_headers)

    # Create an empty DataFrame with 40 columns
    columns = [column_headers[i] for i in range(1, number_of_headers)]
    df = pd.DataFrame(columns=columns)

    percent_format = '{:.2%}'

    # Fill specific columns with initial values
    df['option Expiration date'] = option_expiration_date
    df['Strike'] = strike
    df['underlying symbol'] = underlying_symbol
    df['Type'] = stock_type
    df['mkt beta* mkt px*contracts'] = mkt_beta_px_contracts
    df['Qty'] = quantity
    df['mkt price *number of contracts'] = mkt_price_of_contracts
    df['Trade Price/premium'] = price
    df['trade price as percent of notional'] = trade_price_percent_notional
    df['annual yield at strike at time of trade'] = annual_yield_at_strike
    df['premium'] = premium
    df['trade date'] = date_of_extraction
    df['days till exp (trade date)'] = days_till_exp_date
    df['days till exp (current)'] = days_till_exp_date_current
    df['underlying price at time of trade'] = underlying_price_at_time_of_trade
    df['otm at time of trade'] = otm_at_time_of_trade
    df['underlying price, current'] = underlying_price_current
    df['otm, current'] = otm_current
    df['$ amount of stock itm can be called (-) or put (+)'] = amount_of_stock_itm_can_be_called
    df['weight'] = weight
    df['weighted otm'] = weighted_otm
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

def format_columns(file_path, sheet_name, column_letters, formatting_name, formatting_number):
    # Load the Excel file
    workbook = openpyxl.load_workbook(file_path)

    # Select the desired sheet
    sheet = workbook[sheet_name]

    # Define the starting row
    starting_row = 3  # Change this to the row where formatting should begin

    # Convert column letters to column indices
    column_indices = [openpyxl.utils.column_index_from_string(col) for col in column_letters]

    # Create a custom style with the specified formatting
    formatting_style = NamedStyle(name=formatting_name, number_format=formatting_number)

    # Apply the custom style to the specified range in each column
    for col_index in column_indices:
        for row in sheet.iter_rows(min_col=col_index, max_col=col_index, min_row=starting_row):
            for cell in row:
                cell.style = formatting_style

    # Save the changes to the Excel file
    workbook.save(file_path)

# --- Next 9 Fridays Dates --- #
def find_next_9_fridays():
    # Get today's date
    today = datetime.now().date()

    # Find the next Friday from today
    days_until_next_friday = (4 - today.weekday() + 7) % 7
    next_friday = today + timedelta(days=days_until_next_friday)

    # Calculate the dates for the next 9 Fridays
    next_fridays_primary_list = [next_friday + timedelta(weeks=i) for i in range(9)]
    next_fridays = []
    # Print the result in "mm/dd/yy" format
    for date in next_fridays_primary_list:
        formatted_date = date.strftime("%m/%d/%y")
        next_fridays.append(formatted_date)

    return next_fridays

# --- Last Five Months --- #
def last_five_months():
    today = datetime.now()
    last_five_months_dates = [today - timedelta(days=30*i) for i in range(5)]

    # Formatting the month names and printing in reverse order
    formatted_months = [date.strftime('%B') for date in last_five_months_dates][::-1]

    return formatted_months

# --- LISTS AND OTHER DATA --- #

# --- First Business Day (Cell2) --- #
# Get the current date
current_date = datetime.now()

# first business day
first_business_day_current_month = pd.date_range(start=current_date, periods=1, freq=BMonthBegin()).normalize()[0]

# Format the result to mm/dd/yy without the time
formatted_result = first_business_day_current_month.strftime('%m/%d/%y')

date_string = formatted_result
date_format = "%m/%d/%y"

# Convert string to date
date_object = datetime.strptime(date_string, date_format)

# Format date as a string
formatted_date = date_object.strftime('%m/%d/%y')

# --- Date of Extraction --- #
timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
header_time_stamp = datetime.now().strftime('%m/%d/%y')

# Todays date
today = datetime.now().date()

# --- Previous Months function usage --- #
previous_5_months = last_five_months()

# -- formatting styles --#
percentage_style = NamedStyle(name='percentage', number_format='0.00%')
date_style = NamedStyle(name='date', number_format='d-mmm-yyyy')

# ------------------#
name_with_web_scraper = 'TDA_YAHOO_DATA_'
name_without_web_scraper = 'TDA_DATA_'
excel_filename = f'{name_with_web_scraper}{timestamp}.xlsx'
data_output_directory = os.path.join(project_root, "data", "output", excel_filename)

column_headers = [
    'check date >>', header_time_stamp, 'trade date', 'option Expiration date', 'days till exp (trade date)',
    'days till exp (current)', 'order expiration date "time in force"', 'days till expiration (if an order)',
    'Strike', 'underlying symbol',
    'underlying price at time of trade', 'otm at time of trade', 'underlying price, current', 'otm, current',
    '$ amount of stock itm can be called (-) or put (+)', 'weight', 'weighted otm', 'mkt beta', 'Type',
    'mkt beta* mkt px*contracts', 'Qty',
    'mkt price *number of contracts', 'Trade Price/premium', 'trade price as percent of notional',
    'annual yield at strike at time of trade', 'yield on cost at time of trade', 'multiple on cost',
    'yield at current mkt price at time of trade', 'premium', f'contracted in {previous_5_months[0]}', f'contracted in {previous_5_months[1]}',
    f'contracted in {previous_5_months[2]}', f'contracted in {previous_5_months[3]}', f'contracted in {previous_5_months[4]}', 'cash if exercised', 'days >>',
    '=AK1-A1', '=AL1-A1', '=AM1-A1', '=AN1-A1', '=AO1-A1', '=AP1-A1', '=AQ1-A1', '=AR1-A1', '=AS1-A1',
]

# first line content
first_line_data = [header_time_stamp, formatted_date, 'Open Positions', "Total"] + [None] * (len(column_headers) - 3)
fridays_list = find_next_9_fridays()
start_position = 36
for f in fridays_list:
    first_line_data[start_position] = f
    start_position += 1

try:
    main()

except Exception as e:
    print(f'Error loading the program. {e}\nPlease try again.')
