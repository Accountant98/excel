# import pandas as pd
# form_syo=pd.read_excel("data_clone.xlsx",sheet_name="Sheet1",header=None)
# print(form_syo)

# import pandas as pd
# import numpy as np


# dict_check={}
# dict_sub1={"A":"xxx","B":"ccc","C":"zzz"}
# dict_sub2={"A":"xxxyyy","B":"ccc","C":"zzz"}

# for key, value in dict_sub2.items():
#     if "xy" in value:
#         dict_sub2[key]="w"

# print(dict_sub2)
# ls=[dict_sub1,dict_sub2]
# for dic in ls:
#     for key, value in dic.items():
#         if key not in dict_check.keys():
#             dict_check[key]=value
#         else:
#             dict_check[key]=dict_check[key]+" + "+value


# print(dict_check)

import openpyxl
workbook = openpyxl.load_workbook('xxx.xlsx')
sheet = workbook.active
def get_color_code(cell):
    fill = cell.fill
    if fill and fill.start_color.index:
        color_code = fill.start_color.index
        if color_code.startswith('FF'):
            return color_code
        else:
            return f"Color Index: {color_code}"
    return None

cell_address = 'A1'  # Thay thế bằng địa chỉ ô cụ thể mà bạn muốn lấy mã màu
cell = sheet[cell_address]

# Lấy mã màu từ ô cụ thể
color_code = get_color_code(cell)
if color_code:
    print(f"Cell {cell_address} has color code {color_code}")
else:
    print(f"Cell {cell_address} has no color or is default color.")