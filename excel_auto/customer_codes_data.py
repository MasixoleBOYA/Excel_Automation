import pandas as pd

customer_codes_file = pd.read_excel("C:/Users/J1121857/Downloads/Reseller_ship-to_list.xlsx", sheet_name= 'Cust Loc (3)')

customer_code_list = [i for i in customer_codes_file['Customer No']]
customer_names_list = [ j for j in customer_codes_file['Customer Name']]
customer_cashOrTerms = [x for x in customer_codes_file['Cash/Term']]

customer_codesNames_dictionary = dict(zip(customer_code_list, customer_names_list))
customer_codesTerms_dictionary =dict(zip(customer_names_list, customer_cashOrTerms))


print("\nCUSTOMER CODES DICTIONARY\n")
for i in customer_codesNames_dictionary:
    print(f"{i}: {customer_codesNames_dictionary[i]}")
