import pandas as pd
from src.excel_parser.excel_parser import *
from src.web_scraper.web_scraper import *
column_headers = [
    'check date >>', '01/09/23', 'trade date-entered?', 'option Expiration date', 'days till exp (trade date)', 'days till exp (current)', 'order expiration date "time in force"', 'days till expiration (if an order)', 'Strike', 'underlying symbol',
    'underlying price at time of trade', 'otm at time of trade', 'underlying price, current', 'otm, current.', 'weight', 'weighted otm', 'mkt beta', 'Type', 'mkt beta* mkt price*contracts', 'Qty',
    'mkt price *number of contracts', 'Trade Price/premium', 'trade price as percent of notional', 'annual yield at strike at time of trade', 'yield on cost at time of trade', 'multiple on cost', 'yield at current mkt price at time of trade', 'premium', 'contracted in august', 'contracted in september',
    'contracted in october', 'contracted in november', 'contracted in december', 'cash if exercised', 'days >>', '10', '17', '24', '31', '38', '45', '50', '73', '101',
]

number_of_headers = len(column_headers)

# Create an empty DataFrame with 40 columns
columns = [column_headers[i] for i in range(1, number_of_headers)]
df = pd.DataFrame(columns=columns)

# Fill specific columns with initial values
df['option Expiration date'] = option_expiration_date
df['Strike'] = strike
df['underlying symbol'] = underlying_symbol
df['underlying price at time of trade'] = underlying_price
df['Type'] = stock_type
df['mkt beta'] = mkt_beta_list
df['Qty'] = quantity
df['premium'] = premium


# Specify the Excel file path
data_output_directory = os.path.join(project_root, "data", "output", "output.xlsx")



# Save the DataFrame to an Excel file
df.to_excel(data_output_directory, index=False)

print(f'Data saved to {data_output_directory}')
