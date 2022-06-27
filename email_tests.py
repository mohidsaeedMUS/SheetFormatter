import openpyxl
wb = openpyxl.load_workbook("Test_template.xlsx")
for sheet in wb:
    mr,mc=sheet.max_row,sheet.max_column
    first_time_blank=False
    column_string=input("Enter Column Letter with Email (A or B or C or leave blank to skip editing):").upper()
    if len(column_string)>0:
        for cell in sheet[column_string][1:]:
            if cell.value is None:
                first_time_blank=True
    if first_time_blank==True:
        new_sh=wb.create_sheet('Linkedin Only')
        for i in range (1, mr +1):
            for j in range (1, mc + 1):
                c = sheet.cell(row = i, column = j)
                new_sh.cell(row = i, column = j).value = c.value
        for cell in new_sh[column_string][1:]:
            if cell.value is not None:
                new_sh.delete_rows(cell.row)
        for cell in sheet[column_string][1:]:
            if cell.value is None:
                new_sh.delete_rows(cell.row)
wb.save("Test_1.xlsx")




#add condition to delete empty sheets at end?