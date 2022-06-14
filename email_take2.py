import pandas as pd
df = pd.read_excel('Test.xlsx')
# column_string=input("Enter Column Letter with Email (A or B or C or leave blank to skip editing):").upper()
#i need to create a sheet and add specific values to sheet or add all values to sheet and edit both sheets but the first if condition has to be if value in column or potentially if value in row
for column in df:
    print(df[column].values)