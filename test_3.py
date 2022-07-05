import pandas as pd,openpyxl
wb = openpyxl.load_workbook("/Users/amark/Documents/GitHub/SheetFormatter/Different_Headers.xlsx")
print(wb.sheetnames)
# lst_of_df=[pd.read_excel('/Users/amark/Documents/GitHub/SheetFormatter/Different_Headers.xlsx', sheet_name=sheet) for sheet in wb.sheetnames]


    