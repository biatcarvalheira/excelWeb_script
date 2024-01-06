from datetime import datetime, timedelta

def last_five_months():
    today = datetime.now()
    last_five_months_dates = [today - timedelta(days=30*i) for i in range(5)]

    # Formatting the month names and printing in reverse order
    formatted_months = [date.strftime('%B') for date in last_five_months_dates][::-1]

    return formatted_months

# Example usage
result = last_five_months()
print(result)
