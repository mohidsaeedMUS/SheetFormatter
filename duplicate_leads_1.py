#input - enter column with name-do i do it by first name and last name?
#do this before u do email or first
import pandas as pd
import openpyxl
data = pd.read_excel('Test_in_vscode.xlsx',sheet_name=None)
combined_df=pd.concat(data.values(), ignore_index=True)
new_df=combined_df.drop_duplicates(subset=['First Name','Last Name'],keep='first')
new_df.to_excel('Test_without_dups.xlsx', index = False, header=True)


    

