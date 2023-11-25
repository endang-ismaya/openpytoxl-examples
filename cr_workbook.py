import openpyxl

# initiate a workbook variable
wb = openpyxl.Workbook()

# set an active sheet
sheet = wb.active

# change the name of the sheet
sheet.title = "First"

# write some values
sheet["A1"] = "Name"
sheet["B1"] = "Age"
sheet["C1"] = "Address"

# save the workbook
wb.save("first_wb.xlsx")
