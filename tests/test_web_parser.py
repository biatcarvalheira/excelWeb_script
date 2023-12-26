import re

def remove_stock_symbol(text):
    # Define a regular expression pattern to match stock symbols
    stock_pattern = re.compile(r'\b[A-Z]+\b')
    # Find all stock symbols in the text
    if stock_pattern is not None:
        stock_symbols = stock_pattern.findall(text)
    else:
        stock_symbols = 'not found'
        print('not found')
    # return stock value
    return stock_symbols

# Example usage:
input_text = "I have some shares of TSLA and AAPL."
symbol_to_remove = "TSLA"
result_text = remove_stock_symbol(input_text)
print(result_text)
