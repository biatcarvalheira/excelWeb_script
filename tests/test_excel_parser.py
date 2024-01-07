import openpyxl

# Create a new Excel workbook and select the active sheet
wb = openpyxl.Workbook()
sheet = wb.active

# Sample data
data = [
    ['A', 'B', 'C'],
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
]

# Populate the sheet with data
for row in data:
    sheet.append(row)

# List of formulas
formulas = [
    '=SUM(A2:C2)/C2',  # Example formula 1
    '=SUM(A3:C3)/C3',  # Example formula 2
    '=SUM(A4:C4)/C4'   # Example formula 3
]

# Apply formulas to entire column D
for row_num, formula in enumerate(formulas, start=2):  # Start from the second row
    sheet[f'D{row_num}'].value = formula

# Apply percentage formatting to entire column D
for cell in sheet['D'][1:]:
    cell.number_format = '0.00%'

# Save the workbook
wb.save('output.xlsx')
