from excel_parser.excel_parser import (
    option_expiration_date,
    strike, underlying_symbol,
    stock_type,
    mkt_beta_px_contracts,
    quantity,
    mkt_price_of_contracts,
    price,
    trade_price_percent_notional,
    annual_yield_at_strike,
    yield_at_current_mkt_price_at_trade,
    premium,
    month_5,
    month_1,
    month_2,
    month_3,
    month_4,
    trade_date,
    days_till_exp_date,
    days_till_exp_date_current,
    otm_at_time_of_trade,
    underlying_price_current,
    otm_current,
    amount_of_stock_itm_can_be_called,
    weight,
    weighted_otm,
    cash_if_exercised,
    week_1,
    week_2,
    week_3,
    week_4,
    week_5,
    week_6,
    week_7,
    week_8,
    week_9
)
import openpyxl
import sys

import os
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from pandas.tseries.offsets import BMonthBegin
from openpyxl.styles import NamedStyle

def display_menu():
    print("Choose an option:")
    print("1. Run Trade Mixer")
    print("2. Run Orders Mixer")
    print("3. Run Trade and Orders Mixer")
    print("0. Exit")


def trade_mixer_run_all():
    print('File validation ✓')
    print('Reading XLSX file...')
    print('XLSX file processed ✓')
    df = format_data(column_headers, option_expiration_date, strike, underlying_symbol, stock_type, mkt_beta_px_contracts, quantity, mkt_price_of_contracts, price, trade_price_percent_notional, annual_yield_at_strike, yield_at_current_mkt_price_at_trade, premium, month_5, month_1, month_2, month_3, month_4, trade_date, days_till_exp_date, days_till_exp_date_current, otm_at_time_of_trade, underlying_price_current, otm_current, amount_of_stock_itm_can_be_called, weight, weighted_otm, cash_if_exercised, week_1, week_2, week_3, week_4, week_5, week_6, week_7, week_8, week_9)
    save_data(df, data_output_directory)
    insert_line_after(data_output_directory, 'Sheet1', 1, first_line_data)
    # format percentage cells
    format_columns(data_output_directory, 'Sheet1', ['L', 'N', 'Q', 'X', 'Y', 'AB'], 'percentage', '0.00%')

    # format currency cells
    format_columns(data_output_directory, 'Sheet1',
                   ['T', 'V', 'AI', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS'], 'currency_format',
                   '"$"#,##0.00')


def format_data(column_headers, list1_excel, list2_excel, list3_excel, list4_excel, list5_excel, list6_excel,
                list7_excel, list8_excel, list9_excel, list10_excel, list11_excel, list12_excel,
                list13_excel, list14_excel, list15_excel, list16_excel, list17_excel, list18_excel,
                list19_excel, list20_excel, list21_excel, list22_excel, list23_excel, list24_excel,
                list25_excel, list26_excel, list27_excel, list28_excel, list29_excel, list30_excel,
                list31_excel, list32_excel, list33_excel, list34_excel, list35_excel, list36_excel):
    number_of_headers = len(column_headers)

    # Create an empty DataFrame with 40 columns
    columns = [column_headers[i] for i in range(1, number_of_headers)]
    df = pd.DataFrame(columns=columns)

    percent_format = '{:.2%}'

    # Fill specific columns with initial values
    df['option Expiration date'] = list1_excel
    df['Strike'] = list2_excel
    df['underlying symbol'] = list3_excel
    df['Type'] = list4_excel
    df['mkt beta* mkt px*contracts'] = list5_excel
    df['Qty'] = list6_excel
    df['mkt price *number of contracts'] = list7_excel
    df['Trade Price/premium'] = list8_excel
    df['trade price as percent of notional'] = list9_excel
    df['annual yield at strike at time of trade'] = list10_excel
    df['yield at current mkt price at time of trade'] = list11_excel
    df['premium'] = list12_excel
    df[f'contracted in {previous_5_months[4]}'] = list13_excel
    df[f'contracted in {previous_5_months[3]}'] = list14_excel
    df[f'contracted in {previous_5_months[2]}'] = list15_excel
    df[f'contracted in {previous_5_months[1]}'] = list16_excel
    df[f'contracted in {previous_5_months[0]}'] = list17_excel
    df['trade date'] = list18_excel
    df['days till exp (trade date)'] = list19_excel
    df['days till exp (current)'] = list20_excel
    df['underlying price at time of trade'] = underlying_price_at_time_of_trade
    df['otm at time of trade'] = list21_excel
    df['underlying price, current'] = list22_excel
    df['otm, current'] = list23_excel
    df['$ amount of stock itm can be called (-) or put (+)'] = list24_excel
    df['weight'] = list25_excel
    df['weighted otm'] = list26_excel
    df['mkt beta'] = mkt_beta_list
    df['cash if exercised'] = list27_excel
    df['=AK1-A1'] = list28_excel
    df['=AL1-A1'] = list29_excel
    df['=AM1-A1'] = list30_excel
    df['=AN1-A1'] = list31_excel
    df['=AO1-A1'] = list32_excel
    df['=AP1-A1'] = list33_excel
    df['=AQ1-A1'] = list34_excel
    df['=AR1-A1'] = list35_excel
    df['=AS1-A1'] = list36_excel
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
        formatted_date = date.strftime("%-m/%d/%y")
        next_fridays.append(formatted_date)

    return next_fridays

# --- Last Five Months --- #
def last_five_months_writing():
    today = datetime.now()
    last_five_months_dates = [today - timedelta(days=30 * i) for i in range(5)]

    # Formatting the month names and printing in reverse order
    formatted_months = [date.strftime('%B') for date in last_five_months_dates][::-1]

    return formatted_months


# --- LISTS AND OTHER DATA --- #

# Get the absolute path to the script
script_path = os.path.abspath(sys.argv[0])
script_directory = os.path.dirname(script_path)
project_root = os.path.abspath(os.path.join(script_directory))


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
previous_5_months = last_five_months_writing()

# -- formatting styles --#
percentage_style = NamedStyle(name='percentage', number_format='0.00%')
date_style = NamedStyle(name='date', number_format='d-mmm-yyyy')

# ------------------#
name_with_web_scraper = 'TDA_YAHOO_DATA_'
name_without_web_scraper = 'TDA_DATA_'
excel_filename = f'{name_with_web_scraper}{timestamp}.xlsx'
#IDE
#project_root_for_saving = os.path.join(project_root, "..")

#Executable
project_root_for_saving = os.path.join(project_root)
data_output_directory = os.path.join(project_root_for_saving, "data", "output", excel_filename)

column_headers = [
    'check date >>', header_time_stamp, 'trade date', 'option Expiration date', 'days till exp (trade date)',
    'days till exp (current)', 'order expiration date "time in force"', 'days till expiration (if an order)',
    'Strike', 'underlying symbol',
    'underlying price at time of trade', 'otm at time of trade', 'underlying price, current', 'otm, current',
    '$ amount of stock itm can be called (-) or put (+)', 'weight', 'weighted otm', 'mkt beta', 'Type',
    'mkt beta* mkt px*contracts', 'Qty',
    'mkt price *number of contracts', 'Trade Price/premium', 'trade price as percent of notional',
    'annual yield at strike at time of trade', 'yield on cost at time of trade', 'multiple on cost',
    'yield at current mkt price at time of trade', 'premium', f'contracted in {previous_5_months[0]}',
    f'contracted in {previous_5_months[1]}',
    f'contracted in {previous_5_months[2]}', f'contracted in {previous_5_months[3]}',
    f'contracted in {previous_5_months[4]}', 'cash if exercised', 'days >>',
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
    while True:

        display_menu()
        choice = input('Choose 0-3')
        if choice == '1':
            print(project_root)
            from web_scraper.web_scraper import mkt_beta_list
            from web_scraper.web_scraper import underlying_price_at_time_of_trade
            trade_mixer_run_all()
        if choice == '2':
            print('Orders')
        if choice == '3':
            print('Trade and Orders')
        if choice == '0':
            print('Program finalized')
            break



except Exception as e:
    print(f'Error loading the program. {e}\nPlease try again.')
