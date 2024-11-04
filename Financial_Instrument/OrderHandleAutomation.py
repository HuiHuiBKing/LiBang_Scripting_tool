import time
import pyautogui
import shutil
import getpass
import os
import pandas as pd
import openpyxl

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

"""

# 前置执行条件：1、公司内网  2、chromedriver.exe位置写死，需协助用户使用对应版本chrome及固定driver文件存放位置后运行
实现需求：根据用户输入订单号打开对应工作网站查询对应订单号详情并导出到本地
并隐藏指定列及开启工作表保护设置密码及启用自动筛选功能
"""

# 获取当前用户名
current_user = getpass.getuser()

# 浏览器下载文件路径
source_folder = f'C:\\Users\\{current_user}\\Downloads'

# 目标文件夹路径
destination_folder = f'C:\\Users\\{current_user}\\Desktop\\Script destination folder'

# 确保文件存放文件夹为空，删除后创建
if os.path.exists(destination_folder):
    shutil.rmtree(destination_folder)

os.makedirs(destination_folder)

#
order_number = input("请输入订单号: ")
#
# 驱动路径
chromedriver_path = fr'C:\Users\{current_user}\Desktop\chromedriver_win32\chromedriver.exe'

# selenium 兼容高版本
service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service)

try:
    driver.implicitly_wait(10)

    # 账号密码写入url访问 省略登录步骤
    driver.get(
        'http://caoxingqiong:Cao19912-@mscrmprdapp.nipponpaint.com.cn/_LAYOUTS/NPCRM/TU/StrategyTest/'
        'StrategyTestAllSearch.aspx')

    # 输入订单号
    application_number = driver.find_element(By.ID, 'ctl00_PlaceHolderMain_DB__BillId')
    application_number.send_keys(order_number)

    # 搜索
    submit_btn = driver.find_element(By.ID, 'ctl00_PlaceHolderMain_btnSearch')
    submit_btn.click()

    # 点击搜索结果对应订单号进入详情页
    link = driver.find_element(By.XPATH, f'//a[text()="{order_number}"]')
    link.click()

    # 切换句柄
    handles = driver.window_handles
    driver.switch_to.window(handles[-1])

    # 导出（导出后文件名为集采测算申请明细表_xxxx
    element = driver.find_element(By.ID, 'ctl00_PlaceHolderMain_btnExport_tuliao')

    element.click()

    # 系统确认按钮
    pyautogui.press('enter')
    time.sleep(7)

finally:
    driver.quit()

# 模糊查找确定文件名
matching_files = []
for filename in os.listdir(source_folder):
    if '集采测算申请明细表_' in filename:
        matching_files.append(filename)

# 确保只存在一个符合条件的文件，并将文件重名名为primitive_订单号.xls
if matching_files:
    # 如有多个满足筛选条件文件，取第一个
    first_file = matching_files[0]
    source_file = os.path.join(source_folder, first_file)
    destination_file = os.path.join(destination_folder, f"primitive_{order_number}.xls")

    # 移动和重命名文件
    shutil.move(source_file, destination_file)

# 修改文件名后的原始文件路径
primitive_file_path = destination_folder + "\\" + f"primitive_{order_number}.xls"

# 读取原始文件中的表格写入新的Excel文件
tables = pd.read_html(primitive_file_path)
df = tables[0]

new_file_path = fr'C:\\Users\\{current_user}\Desktop\Script destination folder\{order_number}.xlsx'
df.to_excel(new_file_path, sheet_name="Sheet1", index=False)

# 处理 Excel 文件
wb = openpyxl.load_workbook(new_file_path)
ws = wb.active

# 隐藏指定列
columns_hide = ['D', 'E', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC',
                'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT',
                'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BC', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BM', 'BO', 'BQ',
                'BR', 'BS', 'BT']
for col in columns_hide:
    ws.column_dimensions[col].hidden = True

# 设置保护密码
ws.protection.password = '112233'
# 启用工作表保护
ws.protection.enable()
# 启用自动筛选功能
ws.auto_filter.ref = ws.dimensions

# 统一保存文件
wb.save(new_file_path)
