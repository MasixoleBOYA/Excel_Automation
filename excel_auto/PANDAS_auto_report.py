import pandas as pd
from openpyxl import Workbook, load_workbook

from customer_codes_data import customer_codesNames_dictionary
'''
CUSTOMER CODES DICTIONARY

242797: EAGLE DISTRIBUTORS
244245: EAGLE DISTRIBUTORS
244288: EAGLE DISTRIBUTORS
244613: QUEST PETROLEUM
244614: QUEST PETROLEUM
244618: QUEST PETROLEUM
244622: QUEST PETROLEUM
244624: QUEST PETROLEUM
244639: GULFSTREAM
244641: GULFSTREAM
244642: GULFSTREAM
244643: GULFSTREAM
244645: GULFSTREAM
244648: GULFSTREAM
244649: GULFSTREAM
244650: GULFSTREAM
244652: GULFSTREAM
244653: GULFSTREAM
244654: GULFSTREAM
245065: QUANTUM ENERGY
245066: QUANTUM ENERGY
245127: QUEST PETROLEUM
245366: DHODA DIESELS
245371: DHODA DIESELS
245380: MAKWANDE ENERGY
245480: QUANTUM ENERGY
245910: MAKWANDE ENERGY
245922: NKOMAZI FUEL
251654: VRYPET
254017: VALSAR PETROLEUM
271547: NAMERC CONSULTING
281326: VRYPET
284368: FORCE FUEL
287369: ELEGANT FUEL
288430: NAMERC CONSULTING
290625: GULFSTREAM
290626: GULFSTREAM
290627: GULFSTREAM
290628: GULFSTREAM
290629: GULFSTREAM
291319: BLACK KNIGHT OIL
291320: BLACK KNIGHT OIL
291379: GLOBAL OIL
291382: BLACK KNIGHT OIL
291508: GLOBAL OIL
291833: BLACK KNIGHT OIL
292167: BLACK KNIGHT OIL
292622: BF FUELS
292808: ELEGANT FUEL
292809: ELEGANT FUEL
293513: NAMERC CONSULTING
294872: GLOBAL OIL
295335: AFRICOIL
297043: ROYALE ENERGY
297828: FORCE FUEL
299150: GT OIL & FUEL
299152: GT OIL & FUEL
299260: WBG
299394: AFRICA FUEL
299437: AFRICA FUEL
299486: SELATI PETROLEUM
299653: SELATI PETROLEUM
500526: ELEGANT FUEL
500918: WBG
500966: N1 PETROLEUM
500968: N1 PETROLEUM
500969: N1 PETROLEUM
501677: ECHO PETROLEUM
501810: GULFSTREAM
501811: GULFSTREAM
501812: GULFSTREAM
501928: ROYALE ENERGY
502564: GULFSTREAM
502569: GULFSTREAM
502572: GULFSTREAM
502687: GULFSTREAM
503463: NICSHA PETROLEUM
503464: NICSHA PETROLEUM
503468: NICSHA PETROLEUM
503469: NICSHA PETROLEUM
503470: NICSHA PETROLEUM
503475: NICSHA PETROLEUM
504091: GLOBAL OIL
504188: ELEGANT FUEL
504193: VOSFUELS
504200: VOSFUELS
504704: VRYHEID PETROLEUM
505570: BOVUA ENERGY
'''

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
        customer_name = ""
        
        # Iterate over the items in the dictionary
        for key, value in customer_codesNames_dictionary.items():
            # Check if the current key matches the customer number
            if key == customer_no:
                # Assign the corresponding customer name
                customer_name = value
                # Exit the loop since we found the matching customer name
                break
        
        # Append the customer name to the list
        customer_names.append(customer_name)
    
    # Add the list of customer names as a new column in the DataFrame
    alrode_df['Customer Names'] = customer_names
    
    print(f"NEW CUSTOMER NAMES COLUMN: \n {alrode_df['Customer Names']}")


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
