# from openpyxl import Workbook, load_workbook


# work_book = load_workbook('') # Give path, unless in same dir 
# ws = wb.active

# print(wb.sheetnames)

depots = ["Alrode",
          "Island View Terminal", 
          "Sasolburg",
          "Waltloo",
          "Port Elizabeth"]


for i in depots:
    work_sheet = work_book[i]
    column = work_sheet['Gantry AP']

    for j in column:
        j.value = i


# SECOND PART
appending_sheet = work_book["Alrode"]

for depot_name in depots[1:]:
    current_worksheet = work_book[depot_name]
    for row in current_worksheet.iter_rows(min_row =2,
                                           max_row = current_worksheet.max_rows,
                                           values_only = True):
        appending_sheet.append(row)


# wb.save('xxx.xlsx')