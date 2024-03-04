# from openpyxl import Workbook, load_workbook


# work_book = load_workbook('') # Give path, unless in same dir 
# ws = wb.active

# print(wb.sheetnames)

depots = ["Alrode",
          "Bethlehem",
          "Cape Town",
          "East London",
          "Island View",
          "Klerksdorp",
          "Ladysmith",
          "Mossel Bay",
          "Nelspruit",
          "Port Elizabeth",
          "Sasolburg",
          "Tarlton",
          "Waltloo",
          "Witbank"]


# PART 1: renames the "Gantry AP" column for all sheets
def sheet_rename()->None:
          for i in depots:
              work_sheet = work_book[i]
              column = work_sheet['Gantry AP']
          
              for j in column:
                  j.value = i


# SECOND PART
def appending_to_onesheet(sheetName: str) -> None:
          appending_sheet = work_book[sheetName]
          
          for depot_name in depots[1:]:
              current_worksheet = work_book[depot_name]
              for row in current_worksheet.iter_rows(min_row =2,
                                                     max_row = current_worksheet.max_rows,
                                                     values_only = True):
                  appending_sheet.append(row)

if __name__ == '__main__':
          sheet_rename()
          appending_to_onesheet("Alrode")
          
          # wb.save('xxx.xlsx')
