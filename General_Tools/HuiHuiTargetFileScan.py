import os


def convert_size(size_bytes, unit):
    size_to_bytes = {'K': 1024, 'M': 1024 * 1024, 'G': 1024 * 1024 * 1024}
    return str(int(size_bytes / size_to_bytes[unit]))


def files_scanning(disk, num, unit):
    size_to_bytes = {'K': 1024, 'M': 1024 * 1024, 'G': 1024 * 1024 * 1024}
    target_files_dict = {}
    size_bytes = num * size_to_bytes[unit]
    for root, _, files in os.walk(disk):
        for f in files:
            target_file = os.path.join(root, f)
            try:
                file_size = os.path.getsize(target_file)
                if file_size > size_bytes:
                    target_files_dict[target_file] = convert_size(file_size, unit)
            except FileNotFoundError:
                continue
    return dict(sorted(target_files_dict.items(), key=lambda k: k[1], reverse=True))


def run():
    disk = input("请输入磁盘盘符: ") + ":\\"
    num = int(input("请输入文件大小: "))
    unit = input("请输入文件大小单位 (K/M/G): ").strip().upper()  # 将单位去除空格并转换为大写

    if unit not in ['K', 'M', 'G']:
        print("无效单位，请输入 K, M, G")
        return

    results = files_scanning(disk, num, unit)

    with open(".\满足条件文件路径集合.txt", 'w+') as f:
        for path, size in results.items():
            f.write(f'{path} : {size} {unit} \n')
        f.close()


run()
