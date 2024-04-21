import pandas as pd
from openpyxl import Workbook, load_workbook

from customer_codes_data import customer_codesNames_dictionary
from product_codes_data import product_codes_dictionary



work_book = load_workbook('C:/Users/J1121857/Downloads/GANTRY_RAW_data.xlsx')
product_codes_workbook = load_workbook('C:/Users/J1121857/Downloads/Copy of Reseller customer list 29 Mar 22.XLSX')  # Replace with the actual path
# customer_codes_workbook = load_workbook('C:/Users/J1121857/Downloads/Reseller ship-to list.xlsx')  # Replace with the actual path

depots = ["Alrode", "Bethlehem", "Cape Town", "East London", "Island View", "Klerksdorp", "Ladysmith", "Mossel Bay",
          "Nelspruit", "Port Elizabeth", "Sasolburg", "Tarlton", "Waltloo", "Witbank"]

# PART 1: renames the "Gantry AP" column for all sheets
def sheet_rename() -> None:
    for i in work_book.sheetnames:
        work_sheet = work_book[i]
        column_index = 1

        for x in work_sheet.iter_rows(min_row=2, max_row=work_sheet.max_row, min_col=column_index, max_col=column_index):
            for cell in x:
                cell.value = i

# PART 2: Append data from other sheets to a specified sheet
def appending_to_onesheet(sheet_name: str) -> None:
    appending_sheet = work_book[sheet_name]

    for depot_name in work_book.sheetnames[1:]:
        current_worksheet = work_book[depot_name]
        print(f"\n Working on sheet : {current_worksheet}\n")
        for row in current_worksheet.iter_rows(min_row=2, max_row=current_worksheet.max_row, values_only=True):
            appending_sheet.append(row)

# PART 3: Add "Customer Names" column to Alrode sheet using pandas
def add_customer_names_column():
    alrode_sheet = work_book["Alrode"]

    alrode_df = pd.DataFrame(alrode_sheet.values, columns=[col[0].value for col in alrode_sheet.iter_cols()])

    # Convert 'Customer No.' column to numeric, coercing non-numeric values to NaN
    alrode_df['Customer No.'] = pd.to_numeric(alrode_df["Customer No."], errors='coerce')

    print(f"\nXXXX TYPES XXXXX: {type(alrode_df['Customer No.'])}\n")
    
    # Drop rows with NaN values in 'Customer No.' column
    alrode_df.dropna(subset=['Customer No.'], inplace=True)

    # Convert 'Customer No.' column to integers
    alrode_df['Customer No.'] = alrode_df['Customer No.'].astype(int)

    # Initialize an empty list to store customer names
    customer_names = []

    # Iterate over the rows of the DataFrame
    for index, row in alrode_df.iterrows():
        # Get the customer number from the current row
        customer_no = row['Customer No.']
        
        # Initialize a variable to store the customer name
        customer_name = None  # Change empty string to None
        
        # Iterate over the items in the dictionary
        for key, value in customer_codesNames_dictionary.items():
            # Check if the current key matches the customer number
            if key == customer_no:
                # Assign the corresponding customer name
                customer_name = value
                # Print the customer name for verification
                print(f"Customer No. {customer_no} - Customer Name: {customer_name}")
                # Exit the loop since we found the matching customer name
                break
        
        # Append the customer name to the list
        customer_names.append(customer_name)

    print(f"CUSTOMERS : {customer_names}")
    # Add the list of customer names as a new column in the DataFrame
    alrode_df['Customer Names'] = customer_names
    
    print(f"NEW CUSTOMER NAMES COLUMN: \n {alrode_df['Customer Names']}")



# PART 4: Add "Product Names" column to Alrode sheet using pandas
def add_product_names_column():
    alrode_sheet = work_book["Alrode"]
    
    alrode_df = pd.DataFrame(alrode_sheet.values, columns=[col[0].value for col in alrode_sheet.iter_cols()])

    # Convert 'Product' column to numeric, coercing non-numeric values to NaN
    alrode_df['Product'] = pd.to_numeric(alrode_df["Product"], errors='coerce')

    print(f"\nXXXX TYPES XXXXX: {type(alrode_df['Product'])}\n")

    # Drop rows with NaN values in 'Product' column
    alrode_df.dropna(subset=['Product'], inplace=True)

    # Convert 'Product' column to integers
    alrode_df['Product'] = alrode_df['Product'].astype(int)

    # Initialize an empty list to store product names
    product_names = []

    # Iterate over the rows of the DataFrame
    for index, row in alrode_df.iterrows():
        # Get the product code from the current row
        product_code = row['Product']
        
        # Initialize a variable to store the product name
        product_name = None
        
        # Look up the product name in the dictionary
        product_name = product_codes_dictionary.get(product_code)
        
        # Append the product name to the list
        product_names.append(product_name)

    print(f"PRODUCTS : {product_names}")
    # Add the list of product names as a new column in the DataFrame
    alrode_df['Product Names'] = product_names
    
    print(f"NEW PRODUCT NAMES COLUMN: \n {alrode_df['Product Names']}")

    # Append the updated DataFrame to the Alrode sheet 
    for row in alrode_df.values:
        alrode_sheet.append(list(row))


print(f"\nWORKBOOK sheetnames:\n{work_book.sheetnames}")
sheet_rename()
appending_to_onesheet("Alrode")

add_customer_names_column()
add_product_names_column()

print("cccccccccc DONE cccccccccccc")

work_book.save('C:/Users/J1121857/Downloads/AAAAAAAAAAAA.xlsx')
