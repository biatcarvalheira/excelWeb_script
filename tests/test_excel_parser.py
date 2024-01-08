from datetime import datetime, timedelta

trade_date = ['01/08/22', '08/23/23', '09/12/21', '03/02/19']
month_1 = []
month_2 = []
month_3 = []
month_4 = []
month_5 = []




# Get today's date
today = datetime.now()

# Calculate the last five months
last_five_months = [today - timedelta(days=30 * i) for i in range(5)]

# Extract and print only the month numbers in 'mm' format
month_numbers = [date.strftime('%m') for date in last_five_months]
print(month_numbers)
for t in trade_date:

    # Parse the date string
    date_object = datetime.strptime(t, '%m/%d/%y')

    # Extract the month and format it as 'mm'
    month = date_object.strftime('%m')
    print(month_numbers[0])
    print(month)

    if month == month_numbers[0]:
        month_5.append('x')
    else:
        month_5.append('')
    if month == month_numbers[1]:
        month_1.append('x')
    else:
        month_1.append('')
    if month == month_numbers[2]:
        month_2.append('x')
    else:
        month_2.append('')
    if month == month_numbers[3]:
        month_3.append('x')
    else:
        month_3.append('')
    if month == month_numbers[4]:
        month_4.append('x')
    else:
        month_4.append('')

print('current month:', month_5)
print('month 1:', month_1)
print('month 2:', month_2)
print('month 3:', month_3)
print('month 4:', month_4)
