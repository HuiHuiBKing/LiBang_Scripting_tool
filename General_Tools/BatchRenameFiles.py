import os
import re


def batch_rename_files(directory, keyword, rules):
    """
    实现需求：定义规则进行批量重命名文件，根据参数，保留关键字前或后的内容，并据此重命名文件。
    如有重名文件，删除后续出现的同名文件，保留第一个被重命名的文件。
    """

    existing_files = set()  # 重命名过的文件集合set

    for root, _, files in os.walk(directory):
        for filename in files:
            if keyword in filename:
                if rules:
                    # rules is 1，保留关键字后的内容
                    match = re.search(rf'{keyword}(.*)', filename)
                else:
                    #  rules is 0，保留关键字前的内容
                    match = re.search(rf'(.*?){keyword}', filename)

                if match:
                    new_filename = match.group(1).strip()

                    # get文件扩展名拼接
                    file_ext = os.path.splitext(filename)[1]

                    if rules:
                        # 保留关键字后的内容，不包括关键字
                        new_filename_with_ext = new_filename.split(keyword)[-1] + file_ext
                    else:
                        # 保留关键字前的内容，不包括关键字
                        new_filename_with_ext = new_filename.split(keyword)[0] + file_ext

                    # 原文件路径和新文件路径
                    old_file_path = os.path.join(root, filename)
                    new_file_path = os.path.join(root, new_filename_with_ext)

                    # 重命名并添加当前名字进入set后续判断
                    if new_filename_with_ext not in existing_files:
                        os.rename(old_file_path, new_file_path)
                        existing_files.add(new_filename_with_ext)
                        print(f'Renamed: {old_file_path} -> {new_file_path}')
                    else:
                        # 如果文件已被重命名，则删除当前文件
                        os.remove(old_file_path)
                        print(f'Removed duplicate: {old_file_path}')


# 示例调用
directory = input("请粘贴目标路径：")
keyword = input("请输入关键字：")
rules = int(input("输入 '1' 表示保留关键字后的内容，输入 '0' 表示保留关键字前的内容："))


batch_rename_files(directory, keyword, rules)

