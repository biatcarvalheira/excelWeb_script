from datetime import datetime

# Define date strings in the format mm/dd/yy
date_str1 = "12/13/23"
date_str2 = "01/01/24"

# Parse the date strings into datetime objects
date1 = datetime.strptime(date_str1, "%m/%d/%y")
date2 = datetime.strptime(date_str2, "%m/%d/%y")

# Subtract the dates to get a timedelta object
date_difference = date2 - date1

# Access the days component of the timedelta
days_difference = date_difference.days

print(f"Days difference: {days_difference} days")
