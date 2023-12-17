import os

import datetime as datetime
import openpyxl
import sys
import re
from datetime import datetime

# Get the absolute path to the script
script_path = os.path.abspath(sys.argv[0])

# Get the directory where the script is located (folder containing scripts)
script_directory = os.path.dirname(script_path)

# ---- to use when working in IDE ---- #
#project_root = os.path.abspath(os.path.join(script_directory, ".."))

# ---- to use when exporting as an executable --- #
project_root = os.path.abspath(os.path.join(script_directory))

# Specify the relative path to the data/input directory
data_input_directory = os.path.join(project_root, "data", "input")


def get_xlsx(directory_path):
    try:
        # Get a list of all files in the directory
        files = os.listdir(directory_path)
        # Check if any of the files has a .xlsx extension
        for file in files:
            if file.lower().endswith('.xlsx') and not file.startswith('~$'):
                file_path = os.path.join(directory_path, file)
                try:
                    wb = openpyxl.load_workbook(file_path)
                    sheet_names = wb.sheetnames
                    if sheet_names:
                        first_sheet_name = sheet_names[0]
                        sheet = wb[first_sheet_name]

                        # Get the letters of all active columns
                        active_columns_letters = [openpyxl.utils.get_column_letter(col) for col in
                                                  range(1, sheet.max_column + 1)]

                        # Get the last active row
                        last_active_row = sheet.max_row

                        return first_sheet_name, file_path, active_columns_letters, last_active_row
                except Exception as e:
                    print(f"Error while processing file '{file}': {e}")
        # If no XLSX file with a sheet is found
        return None
    except Exception as e:
        print(f"Error while listing directory '{directory_path}': {e}")


def get_data_from_range(file_path, sheet_name, column_letters, start_row, end_row):
    try:
        # Load the Excel workbook
        workbook = openpyxl.load_workbook(file_path)

        # Get the desired worksheet
        worksheet = workbook[sheet_name]

        # Initialize a list to store the separate column lists
        num_columns = len(column_letters)
        list_of_lists = [[] for _ in range(num_columns)]

        # Get header values excluding 'none'
        header_row = worksheet[1]
        header_values = [str(cell.value).lower().replace(' ', '_') for cell in header_row if str(cell.value).lower() != 'none']

        # Iterate over the rows within the specified range
        for row_number in range(start_row, end_row + 1):
            for col_idx, column_letter in enumerate(column_letters):
                cell_value = worksheet[column_letter + str(row_number)].value
                if str(cell_value).lower() != 'none':
                    list_of_lists[col_idx].append(cell_value)

        # Close the workbook
        workbook.close()

        return list_of_lists, header_values

    except Exception as e:
        print(f"Error while processing the file '{file_path}': {e}")
        return None, None

def removeNone_listOfLists(listName):
    for i, sublist in enumerate(listName):
        listName[i] = [value for value in sublist if value is not None]
    return listName


def removeEmptyList(listName):
    cleaned_list = [sublist for sublist in listName if sublist]
    return cleaned_list


def remove_stock_symbol(text):
    # Define a regular expression pattern to match stock symbols
    stock_pattern = re.compile(r'\b[A-Z]+\b')
    # Find all stock symbols in the text
    if stock_pattern is not None:
        stock_symbols = stock_pattern.findall(text)
    else:
        stock_symbols = 'not found'
    # return stock value
    return stock_symbols


def remove_and_extract_date(text):
    date_pattern = re.compile(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{1,2} \d{4}')
    match = date_pattern.search(text)
    if match:
        extracted_date = match.group()
        return extracted_date
    else:
        return 'not found'


def extract_strike_value(text):
    strike_pattern = re.compile(r'\b\d+\.\d+\b')
    match = strike_pattern.search(text)

    if match:
        return float(match.group())
    else:
        return 'not found'


def extract_stock_type(text):
    stock_type_pattern = re.compile(r'\b(Call|Put)\b', re.IGNORECASE)
    match = stock_type_pattern.search(text)

    if match:
        return match.group().capitalize()
    else:
        return 'not found'

def has_sequence_case_insensitive(my_list, sequence):
    # Convert both the list and the sequence to lowercase
    my_list_lower = [str(item).lower() for item in my_list]
    sequence_lower = [str(item).lower() for item in sequence]

    # Iterate through the modified list and check for the modified sequence
    for i in range(len(my_list_lower) - len(sequence_lower) + 1):
        if my_list_lower[i:i+len(sequence_lower)] == sequence_lower:
            return True
    return False

def convert_date_to_dd_mm_yyyy(input_format, output_format, date_item):
    # Parse the date string into a datetime object
    datetime_object = datetime.strptime(date_item, input_format)
    date_formatted = datetime_object.strftime(output_format)
    # Extract the date part and return it as datetime.date
    return date_formatted

def retrieve_numerical_value (input_text):
    numeric_part = re.search(r'\d+', input_text).group()
    numeric_value = int(numeric_part)
    return numeric_value

# - - - Lists to be used - - - - #
header_list_check = ['Date', 'Description', 'Quantity', 'Symbol', 'Price', 'Amount']
result_lists = []
underlying_symbol = []
option_expiration_date = []
strike = []
stock_type = []
quantity = []
premium = []
price = []
date_of_extraction = []
days_till_exp_date = []
days_till_exp_date_current = []

xlsx_file = get_xlsx(data_input_directory)
if xlsx_file is not None:
    first_sheet_name, file_path, columns, last_row = xlsx_file
    result_lists, headers = get_data_from_range(file_path, first_sheet_name, columns, 2, last_row)
    print(headers)
    if has_sequence_case_insensitive(headers, header_list_check):
        print('File validation ✓')
        print('Reading XLSX file...')
        # cleaning result list
        new_resultList = removeNone_listOfLists(result_lists)
        final_resultList = removeEmptyList(new_resultList)

        ### - - - - - separating the content of the first list => result: 4 lists - - - - - #

        # New Lists: option_expiration_date, strike, underlying_symbol, stock_type
        for n in final_resultList[1]:
            # --- underlying_symbol -- #
            remove_value = remove_stock_symbol(n)
            stock_symbol = ', '.join(remove_value)
            underlying_symbol.append(stock_symbol)

            # --- option_expiration_date -- #
            expiration_date = remove_and_extract_date(n)
            formatted_date = convert_date_to_dd_mm_yyyy("%b %d %Y", "%m/%d/%Y", expiration_date)
            option_expiration_date.append(formatted_date)

            # --- strike -- #
            strike_value = extract_strike_value(n)
            strike.append(strike_value)

            # --- stock_type -- #
            stock_type_v = extract_stock_type(n)
            stock_type.append(stock_type_v)

        # New list: price
        for y in final_resultList[4]:
            price.append(y)

        # New List: quantity
        for y in final_resultList[2]:
            quantity.append(y)

        # New List: premium
        for y in final_resultList[5]:
            premium.append(y)

        # data calculation
        length_of_data = len(final_resultList[1])
        for index in range(length_of_data):
            # get extraction date
            today = str(datetime.today().date())
            today_formatted = convert_date_to_dd_mm_yyyy("%Y-%m-%d", "%m/%d/%Y", today)
            date_of_extraction.append(today_formatted)

            # get days till exp
            string_index = str(index+2)
            formula_exp_date = '=C' + string_index + '-B'+ string_index
            days_till_exp_date.append(formula_exp_date)

            # get days till exp current
            formula_exp_date_current = '=C' + string_index + '-TODAY()'
            days_till_exp_date_current.append(formula_exp_date_current)













            # days till expiration date current








    else:
        print(f'INVALID INPUT FILE. The file should have headers with the following titles in this order:{header_list_check}')
        print('Please start the program again!')
        sys.exit()


