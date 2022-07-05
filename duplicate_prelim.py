import pandas as pd
user_inp=input('Which headers should be considered as duplicate values? Example Answer: First Name/Last Name/Email (separate with /) ')
subset=user_inp.split("/")
for words in subset:
    words.strip()
df = pd.concat(pd.read_excel('Test_template.xlsx', sheet_name=None),ignore_index=True)
df.drop_duplicates(subset=subset,keep='first',inplace=True)
df.to_excel('Test_Results.xlsx',index=False)
