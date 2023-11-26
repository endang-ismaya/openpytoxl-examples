import openpyxl
from openpyxl.styles import Alignment

"""
horizontal
vertical
wrap_text
shrink_to_fit
indent
realtive_indent
justity_last_line
reading_order
"""

wb_name = "first_wb.xlsx"
wb = openpyxl.load_workbook(wb_name)
sh = wb["Text_Alignment"]

# create alignment
align1 = Alignment(
    horizontal="left",
    vertical="top",
    wrap_text=False,
    shrink_to_fit=False,
    indent=2,
)

ipsum = (
    "It is a long established fact that a reader will be distracted by the"
    + "readable content of a page when looking at its layout. The point of "
    + "using Lorem Ipsum is that it has a more-or-less normal distribution "
    + "of letters, as opposed to using 'Content here, content here', making "
    + "it look like readable English. Many desktop publishing packages and web"
    + " page editors now use Lorem Ipsum as their default model text, "
    + "and a search for 'lorem ipsum' will uncover many web sites "
    + "still in their infancy. Various versions have evolved over the years, "
    + "sometimes by accident, sometimes on purpose "
    + "(injected humour and the like)."
)
sh["A1"] = ipsum
sh["A1"].alignment = align1

# save file
wb.save(wb_name)
