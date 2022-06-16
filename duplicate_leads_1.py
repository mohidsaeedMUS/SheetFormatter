#input - enter column with name-do i do it by first name and last name?
#do this before u do email or first
import pandas as pd
df = pd.read_excel('Test.xlsx', sheet_name=None)
new_df={}
for key,value in df.items():
    new_value=value.drop_duplicates(subset=["First Name", "Last Name"], keep="first")
    new_df[key]=new_value
print(new_df)

