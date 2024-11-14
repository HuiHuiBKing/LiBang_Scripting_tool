import os
import zipfile
import openpyxl

"""
    实现需求：解压当前路径下所有压缩文件，解压后遍历当前文件夹及子文件夹路径所有excel文件；
    读取指定单元格，因部分文件格式不同，如果读取为空则读取另一单元格，并组合逐行写入新excel。
"""


# 当前目录路径
current_directory = os.getcwd()

# 遍历当前目录下的所有文件
for file_name in os.listdir(current_directory):
    # 判断是否为压缩文件，简单判断
    if '.zip' in file_name:
        # 构造完整路径
        zip_path = os.path.join(current_directory, file_name)
        # 解压路径，去掉.zip 以原文件名命名
        extract_path = os.path.join(current_directory, file_name[:-4])

        # 解压 zip 文件
        with zipfile.ZipFile(zip_path, 'r') as f:
            f.extractall(extract_path)

# 初始化结果表格
output_file = "审核记录_test.xlsx"
if os.path.exists(output_file):
    os.remove(output_file)

# 实例化工作表工作簿
wb = openpyxl.Workbook()
ws = wb.active

for root, dirs, files in os.walk(current_directory):
    for file_name in files:
        if '.xlsx' in file_name:
            file_path = os.path.join(root, file_name)
            # 打开excel
            wb_input = openpyxl.load_workbook(file_path, data_only=True)
            sheet = wb_input.active  # 读取第一个工作表

            # 获取ID
            ID = root.split('\\')[-1]

            # 获取指定单元格的值
            d5 = sheet["D5"].value
            if not d5:
                d5 = sheet["C4"].value
            d8 = sheet["D8"].value
            if not d8:
                d8 = sheet["C7"].value
            d7 = sheet["D7"].value
            if not d7:
                d7 = sheet["C6"].value
            d9 = sheet["D9"].value
            if not d9:
                d9 = sheet["C8"].value
            f19 = sheet["F19"].value
            if not f19:
                f19 = sheet["E18"].value
            g13 = sheet["G13"].value
            if g13 == 0:
                g13 = sheet["f12"].value
            h18 = sheet["H18"].value
            if not h18:
                h18 = sheet["g17"].value

            # 将数据写入结果表
            ws.append([ID, d5, d8, d7, d9, f19, g13, h18])

# 保存结果表格
wb.save(output_file)
