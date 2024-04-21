import pandas as pd
from openpyxl import Workbook, load_workbook


work_book = load_workbook('C:/Users/J1121857/Downloads/GANTRY_RAW_data.xlsx')
product_codes_workbook = load_workbook('C:/Users/J1121857/Downloads/Copy of Reseller customer list 29 Mar 22.XLSX')  # Replace with the actual path
customer_codes_workbook = load_workbook('C:/Users/J1121857/Downloads/Reseller ship-to list.xlsx')  # Replace with the actual path

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

    customer_codes_workbook_sheet = customer_codes_workbook["Cust Loc (3)"]
    customer_codes_workbook_df = pd.DataFrame(customer_codes_workbook_sheet.values, columns=[col[0].value for col in customer_codes_workbook_sheet.iter_cols()])

    # Set 'Customer No' column as index and identify duplicates
    customer_codes_workbook_df.set_index('Customer No', inplace=True)
    duplicates = customer_codes_workbook_df[customer_codes_workbook_df.index.duplicated(keep=False)]
    if not duplicates.empty:
        print("Duplicate Customer No:")
        print(duplicates)

    # Drop duplicate indices while keeping the first occurrence
    customer_codes_workbook_df = customer_codes_workbook_df[~customer_codes_workbook_df.index.duplicated(keep='first')]

    # Look for customer names based on the "Customer Codes" column
    alrode_df["Customer Names"] = alrode_df["Customer No."].map(customer_codes_workbook_df["Customer Name"])

    # Fill NaN values in "Customer Names" column with empty string
    alrode_df["Customer Names"].fillna("", inplace=True)

    # Clear the contents of the worksheet
    for row in alrode_sheet.iter_rows():
        for cell in row:
            cell.value = None

    # Update the Alrode sheet with the new data
    for index, row in alrode_df.iterrows():
        alrode_sheet.append(row)

# PART 4: Add "Product Names" column to Alrode sheet using pandas
def add_product_names_column():
    alrode_sheet = work_book["Alrode"]
    product_codes_workbook = product_codes_workbook['Product Code']

    alrode_df = pd.DataFrame(alrode_sheet.values, columns=[col[0].value for col in alrode_sheet.iter_cols()])

    # Looking up for product names based on the "Product" column
    alrode_df["Product Names"] = alrode_df["Product"].map(product_codes_workbook.set_index("Product code")["Product name"])

    alrode_df["Product Names"].fillna("", inplace=True)

    # Update the Alrode sheet 
    alrode_sheet.clear()
    alrode_sheet.append(list(alrode_df.columns))
    for row in alrode_df.values:
        alrode_sheet.append(list(row))

print(f"\nWORKBOOK sheetnames:\n{work_book.sheetnames}")
sheet_rename()
appending_to_onesheet("Alrode")

add_customer_names_column()
add_product_names_column()



work_book.save('C:/Users/J1121857/Downloads/AAAAAAAAAAAA.xlsx')
