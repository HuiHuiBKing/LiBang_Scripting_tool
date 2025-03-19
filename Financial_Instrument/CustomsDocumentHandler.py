import os
import shutil

"""
脚本功能：
1. 在当前目录及其子目录中查找包含 "单证备案电子档" 关键字的文件夹。  √
2. 将匹配到的文件夹及结构复制到 new_files 目录中，文件夹为 **序号-报关单号**。  √
3. 在 new_files 目录内保留原始目录结构，将原路径下的子文件夹复制至new_files 目录的相同结构路径下。  √
4. 访问 new_files 路径下的报关单号文件夹，遍历处理后的报关单号，访问公共盘某一路径，以当前报关单号做索引搜索。  √
5. 将搜索的结果文件复制至当前的报关单号文件夹内。  √
6. 将公共文件内无搜索结果的文件名写入txt中。  √
"""

# 目标文件夹路径
new_files_path = "./new_files"

# 公共盘路径
public_disk_path = r"\\npcdnas01\public\Public\臻辅材出口退税单证备案资料"

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


def search_in_public_disk(keyword, path):
    """在公共盘路径及其子路径中搜索包含关键字的文件或文件夹"""
    search_list = []

    for dirpath, dirnames, filenames in os.walk(path):
        # 搜索文件夹名
        for dirname in dirnames:
            if keyword in dirname:
                search_list.append(os.path.join(dirpath, dirname))
        # 搜索文件名
        for filename in filenames:
            if keyword in filename:
                search_list.append(os.path.join(dirpath, filename))
    return search_list


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

with open(r'.\未处理文件集合.txt', 'w') as f:
    # 遍历 new_files 目录，检查并从公共盘复制对应的文件和文件夹
    for folder in folders:
        new_folder_path = os.path.join(new_files_path, os.path.relpath(folder, start="."))
        # print(new_folder_path)

        for files_name in os.listdir(new_folder_path):  # 遍历 new_files 里的子文件夹
            integrity_path = os.path.join(new_folder_path, files_name)  # 完整路径
            # print(integrity_path)

            # 假设文件名格式为 "订单-编号"，我们提取 "编号" 作为搜索关键字
            format_files_name = files_name.split('-')[1]  # 搜索关键字格式化处理

            # 在公共盘搜索匹配的文件或文件夹
            matching_items = search_in_public_disk(format_files_name, public_disk_path)

            if matching_items:
                for pub_item_path in matching_items:
                    target_item_path = os.path.join(integrity_path, os.path.basename(pub_item_path))  # 计算目标路径
                    # print(target_item_path)

                    if os.path.isdir(pub_item_path):
                        # 复制整个文件夹
                        shutil.copytree(pub_item_path, target_item_path, dirs_exist_ok=True)
                        # print(f"复制文件夹: {pub_item_path} -> {target_item_path}")
                    else:
                        # 复制单个文件
                        shutil.copy2(pub_item_path, target_item_path)
                        # print(f"复制文件: {pub_item_path} -> {target_item_path}")
            else:
                f.write(format_files_name)
                f.write('\n')
    f.close()
