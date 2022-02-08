import openpyxl

#link absolute file path
wb = openpyxl.load_workbook("balance.xlsx")

# ws = wb['Score']
# print(ws)

# ws1 = wb['Sheet1']
# print(ws1)

#creating a new sheet and putting it at the beginning of the workbook
wb.create_sheet("New_sheet1",0)
wb.save("balance.xlsx")