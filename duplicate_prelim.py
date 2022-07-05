import pandas as pd
user_inp=input('Type which header names should be considered as duplicate values? Example Answer: First Name/Last Name/Email (separate with /) ')
subset=user_inp.split("/")
subset=[word.strip() for word in subset]
df = pd.concat(pd.read_excel('/Users/amark/Documents/GitHub/SheetFormatter/Merge/Test_template_4.xlsx', sheet_name=None),ignore_index=True)
df.to_excel('/Users/amark/Documents/GitHub/SheetFormatter/Merge/4.xlsx',index=False)
