#do this before u do email or first?
#breaks down if not same headers for each sheet/user should be able to declare what duplicate is?
import pandas as pd
import openpyxl
wb = openpyxl.load_workbook("Test_template.xlsx")
count=-1
for sheet in wb.sheetnames:
    count+=1
    data = pd.read_excel('/Users/amark/Documents/GitHub/SheetFormatter/Test_template.xlsx',count)
    print(data.head())
    # data.drop_duplicates(keep='first',inplace=True)
    # data.to_excel('Test_Results.xlsx',sheet_name='MAIN SHEET')

