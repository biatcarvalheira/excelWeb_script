import pandas as pd

def save_data(data_frame, saving_directory, sheet_name='Sheet2'):
    # Use ExcelWriter to save to a specific sheet
    with pd.ExcelWriter(saving_directory, mode='a', engine='openpyxl') as writer:
        data_frame.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f'Data saved to {saving_directory} in sheet "{sheet_name}"')
    print('************ Processes completed successfully! ************')

# Example usage
# Suppose df is your DataFrame and saving_directory is the path where you want to save the Excel file
df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
save_data(df, 'output.xlsx', sheet_name='Sheet2')
