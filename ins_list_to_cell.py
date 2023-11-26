import openpyxl

wb_name = "first_wb.xlsx"
wb = openpyxl.load_workbook(wb_name)
sh = wb.active

numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]

# iterate thru numbers
for idx, num in enumerate(numbers):
    cell = sh.cell(row=idx + 1, column=1, value=num)

# save workbook
wb.save(wb_name)
