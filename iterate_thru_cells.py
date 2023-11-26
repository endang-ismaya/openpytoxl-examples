import openpyxl

wb_name = "first_wb.xlsx"
wb = openpyxl.load_workbook(wb_name, read_only=True)
sh = wb["LTE_CIQ"]

# loop thru row
for row in sh.iter_rows(values_only=True):
    print(row)  # this return a tuple


# save workbook
# wb.save(wb_name)
