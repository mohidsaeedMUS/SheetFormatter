import openpyxl
from openpyxl.styles import PatternFill

wb = openpyxl.load_workbook("another_test.xlsx")
ws=wb.active

redFill = PatternFill(start_color='FFFF0000',
                   end_color='FFFF0000',
                   fill_type='solid')

column_string=input('Enter Column Letter with Company Name:')
iter_column=openpyxl.utils.cell.column_index_from_string(column_string) 

i = 1
while i <= ws.max_row:
    if "LLC" in ws.cell(row=i, column=iter_column).value:
        print(ws.cell(row=i,column=iter_column))
        ws.cell(row=i, column=iter_column).value = ws.cell(row=i, column=iter_column).value.replace("LLC", "")
        i+=1
    elif len(ws.cell(row=i, column=iter_column).value) > 17:
        ws.cell(row=i, column=iter_column).fill = redFill
        i+=1
    else:
        i+=1

wb.save("another_test.xlsx")