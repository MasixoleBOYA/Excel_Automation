# from openpyxl import Workbook, load_workbook


# wb = load_workbook('') # Give path, unless in same dir 
# ws = wb.active

# print(wb.sheetnames)

depots = ["Alrode",
          "Island View Terminal", 
          "Sasolburg",
          "Waltloo",
          "Port Elizabeth"]

for i in depots:
    print(i)

# wb.save('xxx.xlsx')