import openpyxl,pandas
wb = openpyxl.load_workbook("Test_template.xlsx")
max_rows={}
max_columns={}
sheet_num=0
for sheet in wb.sheetnames:
    data = pandas.read_excel("Test_template.xlsx",sheet_name=sheet)
    max_rows[sheet_num]=len(data)+1
    max_columns[sheet_num]=len(data.columns)
    sheet_num+=1
def access_val(num,dict_1,dict_2):
    return dict_1[num],dict_2[num]
# print(access_val(0,max_rows,max_columns))
print(max_columns)



