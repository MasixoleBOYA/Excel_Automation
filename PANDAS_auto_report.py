import pandas as pd
from openpyxl import Workbook, load_workbook

work_book = load_workbook('C:/Users/J1121857/Downloads/GANTRY_RAW_data.xlsx')
product_codes_workbook = pd.read_excel('path_to_product_codes_workbook.xlsx')  # Replace with the actual path
customer_codes_workbook = pd.read_excel('path_to_customer_codes_workbook.xlsx')  # Replace with the actual path

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

    for depot_name in depots[1:]:
        current_worksheet = work_book[depot_name]
        for row in current_worksheet.iter_rows(min_row=2, max_row=current_worksheet.max_row, values_only=True):
            appending_sheet.append(row)

# PART 3: Add "Customer Names" column to Alrode sheet using pandas
def add_customer_names_column():
    alrode_sheet = work_book["Alrode"]

    # Convert the Alrode sheet to a pandas DataFrame
    alrode_df = pd.DataFrame(alrode_sheet.values, columns=[col[0].value for col in alrode_sheet.iter_cols()])

    # Lookup customer names based on the "Customer Codes" column
    alrode_df["Customer Names"] = alrode_df["Customer Codes"].map(customer_codes_workbook.set_index("Code")["Customer Name"])

    # Replace NaN values with an empty string (or any other desired value)
    alrode_df["Customer Names"].fillna("", inplace=True)

    # Update the Alrode sheet with the new column
    alrode_sheet.clear()
    alrode_sheet.append(list(alrode_df.columns))
    for row in alrode_df.values:
        alrode_sheet.append(list(row))

# PART 4: Add "Product Names" column to Alrode sheet using pandas
def add_product_names_column():
    alrode_sheet = work_book["Alrode"]

    # Convert the Alrode sheet to a pandas DataFrame
    alrode_df = pd.DataFrame(alrode_sheet.values, columns=[col[0].value for col in alrode_sheet.iter_cols()])

    # Lookup product names based on the "Product" column
    alrode_df["Product Names"] = alrode_df["Product"].map(product_codes_workbook.set_index("Code")["Product Name"])

    # Replace NaN values with an empty string (or any other desired value)
    alrode_df["Product Names"].fillna("", inplace=True)

    # Update the Alrode sheet with the new column
    alrode_sheet.clear()
    alrode_sheet.append(list(alrode_df.columns))
    for row in alrode_df.values:
        alrode_sheet.append(list(row))

# Print sheet names, rename columns, and append data
print(work_book.sheetnames)
sheet_rename()
appending_to_onesheet("Alrode")

# Add "Customer Names" column and "Product Names" column to Alrode sheet
add_customer_names_column()
add_product_names_column()

# Save the modified workbook to a new file
work_book.save('C:/Users/J1121857/Downloads/TEST_NEW_NEW.xlsx')
