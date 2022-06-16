# import openpyxl
  
# load excel with its path
# wrkbk = openpyxl.load_workbook("Test.xlsx")  
# sh = wrkbk.active

# TODO: Prompt user for email status columns
# TODO: Translate user-input letters into numbers

# for col in sh.iter_cols(min_row = 2, max_row = 6, min_col = 5, max_col = 5):
#     for cell in col:
#         if cell.value.casefold() == "unavailable":
#             sh_new = wrkbk.create_sheet(title = "Linkedin Only")
#             sh_new.append(["First Name","Last Name","Company Name","Email","Email Status","First Phone","Employees","Industry","Person", "Linkedin","Company City","Company State","Company Count"])
#             print(True)
#             break
#     print()
# wrkbk.save("Test.xlsx")
#tests
