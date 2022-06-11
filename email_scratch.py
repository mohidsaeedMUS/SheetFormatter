import openpyxl
wb = openpyxl.load_workbook("Test.xlsx")
ws=wb.active
sh_new = wb.create_sheet(title = "Linkedin Only")
sh_new.append(["First Name","Last Name","Company Name","Email","Email Status","First Phone","Employees","Industry","Person", "Linkedin","Company City","Company State","Company Count"])
mr=ws.max_row
mc=ws.max_column
rownum=1
column_string=input('Enter Column Letter with Email:')
#need to be able to move on to next if left blank
# column=openpyxl.utils.cell.column_index_from_string(column_string) 
for i in range (1, mr +1):
    for j in range (1, mc + 1):
        c = ws.cell(row = i, column = j)
        sh_new.cell(row = i, column = j).value = c.value
# for col in sh_new.iter_cols(min_row = 2, max_row = ws.max_row, min_col = column , max_col = column):
#     for cell in col:
#         rownum+=1
#         if cell.internal_value == "Verified":
# for row in sh_new.rows:
#     for cell in row:
#         if cell.value =="Verified":
#             sh_new.delete_rows(cell.row, 1)
for cell in sh_new[column_string][1:]:
    if cell.value == "Verified":
        sh_new.delete_rows(cell.row)
for cell in ws[column_string][1:]:
    if cell.value == "Unavailable":
        ws.delete_rows(cell.row)
wb.save("Test.xlsx")
#my worry- if doesnt have words verified or unavailable it wont work / if do not capitalize column_string it wont work/ if do not enter anything in column it has to skip
#do we arrange each of these capabilites in functions?


    

