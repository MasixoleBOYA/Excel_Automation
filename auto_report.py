'''
C:\Users\J1121857\OneDrive - TotalEnergies\Desktop>pip install openpyxl
WARNING: Retrying (Retry(total=4, connect=None, read=None, redirect=None, status=None)) after connection broken by 'ConnectTimeoutError(<pip._vendor.urllib3.connection.HTTPSConnection object at 0x000001DAD9F5CF20>, 'Connection to pypi.org timed out. (connect timeout=15)')': /simple/openpyxl/
WARNING: Retrying (Retry(total=3, connect=None, read=None, redirect=None, status=None)) after connection broken by 'ConnectTimeoutError(<pip._vendor.urllib3.connection.HTTPSConnection object at 0x000001DADD6E2750>, 'Connection to pypi.org timed out. (connect timeout=15)')': /simple/openpyxl/
WARNING: Retrying (Retry(total=2, connect=None, read=None, redirect=None, status=None)) after connection broken by 'ConnectTimeoutError(<pip._vendor.urllib3.connection.HTTPSConnection object at 0x000001DADD7AF8F0>, 'Connection to pypi.org timed out. (connect timeout=15)')': /simple/openpyxl/
WARNING: Retrying (Retry(total=1, connect=None, read=None, redirect=None, status=None)) after connection broken by 'ConnectTimeoutError(<pip._vendor.urllib3.connection.HTTPSConnection object at 0x000001DADD7AFB00>, 'Connection to pypi.org timed out. (connect timeout=15)')': /simple/openpyxl/
WARNING: Retrying (Retry(total=0, connect=None, read=None, redirect=None, status=None)) after connection broken by 'ConnectTimeoutError(<pip._vendor.urllib3.connection.HTTPSConnection object at 0x000001DADD7AFD10>, 'Connection to pypi.org timed out. (connect timeout=15)')': /simple/openpyxl/
ERROR: Could not find a version that satisfies the requirement openpyxl (from versions: none)
ERROR: No matching distribution found for openpyxl
'''

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
