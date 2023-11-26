import openpyxl

wb = openpyxl.Workbook(write_only=True)
sh = wb.create_sheet()

# append sheet
sh.append(["name", "age", "address"])
sh.append(["endang", 38, "Tangerang"])


# save wb
wb.save("write-only.xlsx")
