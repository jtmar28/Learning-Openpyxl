import openpyxl
from openpyxl.styles import Border, Side

#link absolute file path
wb = openpyxl.load_workbook("balance.xlsx")
ws = wb['Sheet1']

top = Side(border_style='dashed',color='FF0707')
bottom = Side(border_style='double',color='10AF2A')

border = Border(top=top, bottom=bottom)

ws['B6'].border = border

wb.save("balance.xlsx")