import openpyxl
import numpy as np

wb_name = "first_wb.xlsx"
ws_name = "family_data"
wb = openpyxl.load_workbook(wb_name)
ws = wb[ws_name]

ages = [cell.value for cell in ws["B"] if str(cell.value).isdigit()]
np_ages = np.array(ages)

# sum and average
sum_ = np.sum(np_ages)  # 113
avg_ = np.mean(np_ages)  # 16.142857142857142

print(sum_)
print(avg_)

# wb.save(wb_name)
