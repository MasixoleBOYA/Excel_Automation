from openpyxl import Workbook, load_workbook


work_book = load_workbook('C:/Users/J1121857/Downloads/GANTRY_RAW_data.xlsx') # Give path, unless in same dir 
# ws = wb.active

# print(wb.sheetnames)

depots = ["Alrode",
          "Bethlehem",
          "Cape Town",
          "East London",
          #"Island View",
          "Klerksdorp",
          "Ladysmith",
          "Mossel Bay",
          "Nelspruit",
          "Port Elizabeth",
          "Sasolburg",
          #"Tarlton",
          "Waltloo",
          "Witbank"]


# PART 1: renames the "Gantry AP" column for all sheets
def sheet_rename()->None:
    for i in work_book.sheetnames: 
        work_sheet = work_book[i] #for each depot, consider the sheet that matches the depot
        column_index = 1

        for x in work_sheet.iter_rows(min_row=2, max_row=work_sheet.max_row, min_col=column_index, max_col=column_index):
            for cell in x:
                cell.value = i



        # for col in work_sheet.iter_cols(min_row=1,max_row=1,min_col=1):
            # print(col[0].value)
            # print(col[0].column)
            # if col[0].value == "Gantry AP":    
            #     column_index = col[0].column
            #     working_column = col[0].value
            #     # break
                     
            # print("x1")
            # column = work_sheet[column_index]
            # print("x2")



            


            # for j in column:
            #     # print(f"VALUES OF COLUMN ARE: \n{j.value}")
            #     j.value = i
            # print("x3")

          

# SECOND PART
def appending_to_onesheet(sheetName: str) -> None:
    appending_sheet = work_book[sheetName]
          
    for depot_name in depots[1:]:
        current_worksheet = work_book[depot_name]
        for row in current_worksheet.iter_rows(min_row =2,
                                                     max_row = current_worksheet.max_row,
                                                     values_only = True):
            appending_sheet.append(row)

print(work_book.sheetnames)
sheet_rename()
appending_to_onesheet("Alrode")
work_book.save('C:/Users/J1121857/Downloads/TEST_NEW_NEW.xlsx')
          
