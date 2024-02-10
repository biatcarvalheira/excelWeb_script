string_index = '1'
formula_otm_current = '=(S' + string_index + '= "Call", (I' + string_index + '-M' + string_index + ')/M' + string_index + ', (M' + string_index + '-I' + string_index + ')/M' + string_index + ')'
print(formula_otm_current)