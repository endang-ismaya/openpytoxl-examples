import openpyxl
from openpyxl.styles import Font

wb_name = "first_wb.xlsx"
wb = openpyxl.load_workbook(wb_name)
sh = wb["Fonts"]

# create a font style
font1 = Font(
    name="Calibri",
    bold=True,
    italic=True,
    u="double",
    color="BB9CC0",
    size=24,
    strike=True,
)

# assign the font
sh["A1"] = "Font Testing"
sh["A1"].font = font1

# save file
wb.save(wb_name)
