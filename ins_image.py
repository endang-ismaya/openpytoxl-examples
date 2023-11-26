import openpyxl
from openpyxl.drawing.image import Image

wb_name = "first_wb.xlsx"
wb = openpyxl.load_workbook(wb_name)
sh = wb["First"]

# create img and resize
# img = Image("pexels-pixabay-45246.jpg")
# img.width = 200
# img.height = 200

# position the image in the xls
# sh.add_image(img=img, anchor="A2")

# multiple images
image_files = ["pexels-pixabay-45246.jpg", "pexels-photo-3280908.jpeg"]
row_data = ["A", "E"]

for idx, image_file in enumerate(image_files, start=1):
    img = Image(image_file)
    row_d = row_data[idx - 1]
    img.width = 200
    img.height = 200
    sh.add_image(img=img, anchor=f"{row_d}1")

# save file
wb.save(wb_name)
