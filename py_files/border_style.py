import openpyxl
from openpyxl.styles import Border, Side


# defining workbook
wb_name = "first_wb.xlsx"
wb = openpyxl.load_workbook(wb_name)
sh = wb["Text_Alignment"]

# cr border obj
mix_border = Border(
    left=Side(style="thin"),
    right=Side(style="dashDot", color="2B2A4C"),
    top=Side(style="dashDotDot", color="265073"),
    bottom=Side(style="thick", color="6B240C"),
)

sh["B4"] = "Visual Studio Code"
sh["B4"].border = mix_border

# save file
wb.save(wb_name)
