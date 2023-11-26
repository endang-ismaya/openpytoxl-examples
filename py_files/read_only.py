import openpyxl

# defining workbook
wb_name = "first_wb.xlsx"
wb = openpyxl.load_workbook(wb_name, read_only=True)
sh = wb["LTE_CIQ"]

# access value from a row 3
row_values = [cell.value for cell in sh[3]]
print(row_values)

# access values from column D
# col_values = [cell.value for cell in sh["D"]]
# print(col_values)

# access a range
cell_range = sh["A1:D3"]

for rows in cell_range:
    for cell in rows:
        print(cell.value)


# save file
# wb.save(wb_name)
