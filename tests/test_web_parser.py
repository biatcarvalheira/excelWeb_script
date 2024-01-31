from datetime import datetime

# Assuming you have a datetime object
current_date = datetime.now()

# Format the date as "1/31/24" with single-digit month
formatted_date = current_date.strftime("%-m/%d/%y")

print(formatted_date)
