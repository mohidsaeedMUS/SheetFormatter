import pandas as pd
import openpyxl
#connect sheets by a first name/ last name/ company name/ email and automatically use that as subset for duplicate
df = pd.concat(pd.read_excel('/Users/amark/Documents/GitHub/SheetFormatter/Final Test/Test_complete copy.xlsx', sheet_name=None),ignore_index=True)
column_string=input('Enter Column Letter with First Name (or leave blank to skip): ').upper()
column_string_last=input('Enter Column Letter with Last Name (or leave blank to skip): ').upper()
i=1
column_names={}
for col in df.columns:
    column_names[i]=col
    i+=1
iter_column=openpyxl.utils.cell.column_index_from_string(column_string) 
iter_column_last=openpyxl.utils.cell.column_index_from_string(column_string_last) 
df.drop_duplicates(subset=[column_names[iter_column],column_names[iter_column_last]],keep='first',inplace=True)
df.to_excel('/Users/amark/Documents/GitHub/SheetFormatter/Final Test/Test_complete copy_result.xlsx',index=False)





