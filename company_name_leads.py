import openpyxl
wb = openpyxl.load_workbook("another_test.xlsx")
ws=wb.active

column_string=input('Enter Column Letter with Company Name:')
iter_column=openpyxl.utils.cell.column_index_from_string(column_string) 

i = 1
while i <= ws.max_row:
    if "LLC" in ws.cell(row=i, column=iter_column).value:
        print(ws.cell(row=i,column=iter_column))
        ws.cell(row=i, column=iter_column).value = ws.cell(row=i, column=iter_column).value.replace("LLC", "")
        i+=1
    else:
        i+=1

wb.save("another_test.xlsx")