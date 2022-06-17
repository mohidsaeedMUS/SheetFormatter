import openpyxl
wb = openpyxl.load_workbook("another_test.xlsx")
ws=wb.active

column_string=input('Enter Column Letter with First Name:')
iter_column=openpyxl.utils.cell.column_index_from_string(column_string) 

# TODO: let user pick column with first name
i = 1
while i <= ws.max_row:
    if ws.cell(row=i, column=iter_column).value is None:
        print(ws.cell(row=i,column=iter_column))
        ws.delete_rows(i, 1)
    else:
        i+=1

wb.save("another_test.xlsx")