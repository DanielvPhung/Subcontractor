from openpyxl import Workbook, load_workbook
import win32com.client

#Workbook
info_wb = load_workbook('Sinkat Contact Info.xlsx')
#Load Sheet
info_sh = info_wb.active
print(info_sh)
print(info_sh['A1'].value)





info_wb.close()
