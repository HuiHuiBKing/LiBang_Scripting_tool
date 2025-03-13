import os
import shutil

"""
脚本功能：
1. 在当前目录及其子目录中查找包含 "单证备案电子档" 关键字的文件夹。  √
2. 将匹配到的文件夹及结构复制到 new_files 目录中，文件夹为 **序号-报关单号**。  √
3. 在 new_files 目录内保留原始目录结构，将原路径下的子文件夹复制至new_files 目录的相同结构路径下。  √
4. 访问 new_files 路径下的报关单号文件夹，遍历处理后的报关单号，访问公共盘某一路径，以当前报关单号做索引搜索。
5. 将搜索的结果文件复制至当前的报关单号文件夹内。
"""

# 目标文件夹路径
new_files_path = "./new_files"

# 删除后重新创建文件夹
if os.path.exists(new_files_path):
    shutil.rmtree(new_files_path)
os.makedirs(new_files_path)


def search_folders_containing_keyword(keyword="单证备案电子档"):
    """ 查找当前目录下包含关键字的文件夹 """
    folders_list = []
    for dirpath, dirnames, _ in os.walk('.'):
        for dirname in dirnames:
            if keyword in dirname:
                full_path = os.path.join(dirpath, dirname)
                folders_list.append(full_path)
    return folders_list


# 查找符合条件的文件夹
folders = search_folders_containing_keyword()
print(folders)

# 复制文件夹内容并保留原路径结构
for folder in folders:
    target_path = os.path.join(new_files_path, os.path.relpath(folder, start="."))  # 组装目标文件夹路径
    os.makedirs(target_path, exist_ok=True)  # 创建目标文件夹
    # print(target_path)
    for item in os.listdir(folder):  # 遍历原文件夹的内容
        old_item = os.path.join(folder, item)  # 组装原文件夹的完整路径
        new_item = os.path.join(target_path, item)  # 组装目标路径
        # print(old_item, new_item)
        if os.path.isdir(old_item):
            shutil.copytree(old_item, new_item, dirs_exist_ok=True)  # 复制子文件夹
