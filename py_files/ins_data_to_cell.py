import openpyxl

wb_name = "first_wb.xlsx"
wb = openpyxl.load_workbook(wb_name)
sh = wb.active

sh["D1"] = 55
sh["E1"] = 10

# formula
sh["F1"] = "=D1 + E1"
wb.save(wb_name)
