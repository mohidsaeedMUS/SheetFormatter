#do this before u do email or first?
#breaks down if not same headers for each sheet/user should be able to declare what duplicate is?
import pandas as pd
data = pd.read_excel('Test_in_vscode.xlsx',sheet_name=None)
combined_df=pd.concat(data.values(), ignore_index=True)
new_df=combined_df.drop_duplicates(subset=['First Name','Last Name'],keep='first')
new_df.to_excel('Test_without_dups.xlsx', index = False, header=True)   

