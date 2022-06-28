import openpyxl
wb = openpyxl.load_workbook("Test_template.xlsx")
sheet=wb['Outsourcing']
def get_maximum_rows(*, sheet_object):
    rows = 0
    for max_row, row in enumerate(sheet_object, 1):
        if not all(col.value is None for col in row):
            rows += 1
    return rows
def get_maximum_cols(*, sheet_object):
    cols = 0
    for max_col, col in enumerate(sheet_object, 1):
        if not all(row.value is None for row in col):
            cols += 1
    return cols
max_rows = get_maximum_rows(sheet_object=sheet)
max_cols=get_maximum_cols(sheet_object=sheet)
print(max_rows,max_cols)


        