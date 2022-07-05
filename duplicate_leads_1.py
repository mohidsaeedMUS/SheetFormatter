import pandas as pd
import openpyxl
wb = openpyxl.load_workbook("/Users/amark/Documents/GitHub/SheetFormatter/Merge/Test_template.xlsx")
count=-1
dict_of_df={}
for sheet in wb.sheetnames:
    count+=1
    dict_of_df[sheet] = pd.read_excel('/Users/amark/Documents/GitHub/SheetFormatter/Merge/Test_template.xlsx',count)




