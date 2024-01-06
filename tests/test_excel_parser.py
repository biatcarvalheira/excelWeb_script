import pandas as pd
from datetime import datetime
from pandas.tseries.offsets import BMonthBegin

# Get the current date
current_date = datetime.now()

# Calculate the first business day of the current month and normalize to midnight
first_business_day_current_month = pd.date_range(start=current_date, periods=1, freq=BMonthBegin()).normalize()[0]

# Format the result to mm/dd/yy without the time
formatted_result = first_business_day_current_month.strftime('%m/%d/%y')

print(f"The first business day of the current month is: {formatted_result}")
date_string = '02/01/24'
date_format = "%m/%d/%y"

# Convert string to date
date_object = datetime.strptime(date_string, date_format)

# Format date as a string
formatted_date = date_object.strftime('%m/%d/%y')

print(formatted_date)
