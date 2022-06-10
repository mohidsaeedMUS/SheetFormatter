import openpyxl
#gets first active sheet and creates a workbook
wb = openpyxl.load_workbook("Test_in_vscode.xlsx")
ws=wb.active
column_string=input('Enter Column Letter with Email:')
column=openpyxl.utils.cell.column_index_from_string(column_string)
count=0
start=0
#grab a row and input it into other sheet
for row in ws.iter_rows(min_col=column,max_col=column,min_row=2,max_row=ws.max_row):
    for cell in row:
        start+=1
        if cell.internal_value=='Unavailable':
            count+=1
            if count==1:
                ws1=wb.create_sheet('Linkedin Only')
                ws1.append(["First Name","Last Name","Company Name","Email","Email Status","First Phone","Employees","Industry","Person", "Linkedin","Company City","Company State","Company Count"])
                ws.delete_rows(start, 1)
                
            else:

                ws.delete_rows(start, 1)
#print(rownum)
# print(columns)
# rownum=1
# for column in columns:
#     for cell in column:
#         rownum+=1
#         if cell.internal_value=='Unavailable':
#             print(rownum)
#             # ws1=wb.create_sheet('Linkedin Only')
#             # ws1.append(["First Name","Last Name","Company Name","Email","Email Status","First Phone","Employees","Industry","Person", "Linkedin","Company City","Company State","Company Count"])
#             # #transfer entire row where unavailable pops up to linkedin only 
#             # break
# wb.save('Test_in_vscode.xlsx')


 


