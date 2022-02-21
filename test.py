from openpyxl import Workbook, load_workbook
import win32com.client

#Workbook
wb = load_workbook('new_template.xlsx')
#Load Sheet
ws = wb.active
print(ws)
print(ws['A1'].value)
ws['A2'].value = "Text"

print(wb.sheetnames)

wb.create_sheet("2022-02-21")

wb.save('new_template.xlsx')
wb.close()

wb_path = r'C:\Users\Daniel\Desktop\Sinkat\new_template.xlsx'
pdf_path = r'C:\Users\Daniel\Desktop\Sinkat\new.pdf'

excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.open(wb_path)
ws_index_list = [1]
wb.WorkSheets(ws_index_list).Select()
wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)\

wb.Close()
excel.Quit()