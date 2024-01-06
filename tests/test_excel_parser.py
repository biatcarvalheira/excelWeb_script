from datetime import datetime, timedelta



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
