import os
import re
import time
import getpass
import pyautogui
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta

"""
#前置执行条件：1、chromedriver.exe位置写死，需协助用户使用对应版本chrome及固定driver文件存放位置后运行
            2、执行过程中因使用pyautogui进行保存网页内容操作，在执行未结束时中尽量避免手动操作
实现需求：运行后打开税务网站，登录后自动选择开始时间-结束时间，获取时间范围内的多个发票文件名，组装多个url访问发票详情页后进行下载；
        下载完成后对比本地文件数量是否符合预期，并将结果写入桌面result.txt中
"""

# 获取运行当前用户名
current_user = getpass.getuser()

# 指定 chromedriver.exe 的路径
chromedriver_path = fr'C:\Users\{current_user}\Desktop\chromedriver_win32\chromedriver.exe'

# 发票号集合
Invoice_numbers_values = set()


def get_start_end_dates():
    """根据当前月份获取上个月的第一天和当前月的第一天"""
    today = datetime.today()  # 获取今天的日期

    # 当前月份的第一天
    first_day_current_month = today.replace(day=1)

    # 上个月的第一天  当前月第一天减去一天后，获得上月最后一天，将天数replace为 1， 获得上月第一天
    first_day_last_month = (first_day_current_month - timedelta(days=1)).replace(day=1)

    # 截取日期
    start_date = str(first_day_last_month).split()[0]
    end_date = str(first_day_current_month).split()[0]

    return start_date, end_date


start_date_str, end_date_str = get_start_end_dates()

print(f"start date: {start_date_str}")
print(f"end date: {end_date_str}")

# 初始化 WebDriver
service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service)

driver.implicitly_wait(20)
# 使用 显性等待 验证，翻页和提取数据
wait = WebDriverWait(driver, 20)

driver.get('https://i.chinaport.gov.cn/deskserver/sw/deskIndex?menu_id=rtx')
driver.maximize_window()

# 登录
login_type = driver.find_element(By.ID, "cardTabBtn")
login_type.click()
login_password = driver.find_element(By.ID, "password")
login_password.send_keys('tf172168')
login_agreement = driver.find_element(By.ID, "checkboxIntel")
login_agreement.click()
login_ptn = driver.find_element(By.ID, "loginbutton")
login_ptn.click()

# 点击对应发票选项
odd_number_inquire_title = driver.find_element(By.XPATH, '//a[@menuid="rtx001"]')
odd_number_inquire_title.click()
odd_number_inquire_details = driver.find_element(By.XPATH, '//a[@menuid="rtx001leaf1"]')
odd_number_inquire_details.click()

driver.switch_to.frame('iframe01')  # 切换到 iframe

# 选择时间及发票时间范围
date_of_export = driver.find_element(By.XPATH, '//input[@id="dateRadio"]')
date_of_export.click()
# 开始时间
start_date_elem = driver.find_element(By.ID, "startDate")
start_date_elem.click()
time.sleep(2)
date_to_select = driver.find_element(By.XPATH, fr'//td[@lay-ymd="{start_date_str}"]')
date_to_select.click()
# 结束时间
end_date_elem = driver.find_element(By.ID, "endDate")
end_date_elem.click()
time.sleep(2)
date_to_select = driver.find_element(By.XPATH, fr'//td[@lay-ymd="{end_date_str}"]')
date_to_select.click()
# 查找
search_btn = driver.find_element(By.ID, "searchBtn")
search_btn.click()
# 页面滚动至最底部
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(1)


def extract_odd_numbers():
    """提取当前页面中的所有发票号"""
    odd_numbers = driver.find_elements(By.XPATH, '//body//a[@style="color: red"]')
    for element in odd_numbers:
        Invoice_numbers_values.add(element.text)


while True:
    try:
        extract_odd_numbers()

        next_button = wait.until(
            EC.presence_of_element_located((By.XPATH, '//li[@class="page-next"]/a'))
        )
        next_button.click()

        # 等待新页面内容完全加载
        time.sleep(1)

    except TimeoutException:
        break

print("invoice number values:", Invoice_numbers_values, len(Invoice_numbers_values))


def save_pdf(pdf_url, file_name):
    """拼接url访问对应发票详情页，页面无保存按钮，使用 pyautogui 模拟按键操作来保存页面为 PDF"""
    # 打开发票详情页
    driver.execute_script(f"window.open('{pdf_url}', '_blank');")
    driver.switch_to.window(driver.window_handles[-1])

    # 等待页面完全加载
    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    print(f"url: {driver.current_url}")

    # 使用 pyautogui 模拟按键操作来保存文件
    pyautogui.hotkey('ctrl', 's')  # 打开保存对话框
    time.sleep(1)  # 等待对话框打开

    # 文件名命名为当前访问发票名
    pyautogui.write(file_name)  # 输入文件名
    time.sleep(1)
    pyautogui.press('enter')  # 确认保存
    time.sleep(2)  # 等待文件下载完成

    # 关闭当前标签页
    driver.close()
    time.sleep(1)
    # 切换回第一个标签页
    driver.switch_to.window(driver.window_handles[0])


# 访问指定的页面并保存为 PDF
for value in Invoice_numbers_values:
    pdf_url = f'https://i.chinaport.gov.cn/rtxwebserver/sw/rtx/decDetail/preview/{value},1'
    save_pdf(pdf_url, value)

# 关闭 WebDriver
driver.quit()


def extract_pdf_filenames():
    """将下载文件夹中满足 18位数字.pdf 的文件名添加至集合"""
    pdf_files = set()  # 符合条件的pdf文件名集合
    download_folder_path = rf'C:\Users\{current_user}\Downloads'  # pdf下载文件夹路径 默认为downloads
    for filename in os.listdir(download_folder_path):
        # 发票号文件名字为 18位数字.pdf
        match = re.match(r'^\d{18}\.pdf$', filename)
        if match:
            pdf_files.add(filename)
    return pdf_files


output_file = rf'C:\Users\{current_user}\Desktop\result.txt'


def check_files(values, result_file):
    """将获取的发票集合与本地下载文件夹中的 pdf集合 对比，检查本地与预期是否一致"""

    # 获取下载文件夹中所有符合条件的pdf文件名
    downloaded_files = extract_pdf_filenames()

    # 将获取的发票集合内元素组装添加 .pdf ，后续与下载文件中的pdf集合作对比
    expected_files = set()
    for value in values:
        expected_files.add(f"{value}.pdf")

    # 与下载文件中的pdf集合作对比，找出缺失的文件
    missing_files = expected_files - downloaded_files

    # 写入到result.txt
    with open(result_file, 'w') as f:
        total_files = len(values)
        f.write(f'需要下载的文件数量为：{total_files}\n')
        f.write("未正确下载的文件:\n")
        for file in missing_files:
            f.write(f"{file}\n")
            invoice_id = file.replace('.pdf', '')  # 提取文件名中的发票id
            # 组装未被下载的发票号写入txt，方便用户自行访问下载
            url = f'https://i.chinaport.gov.cn/rtxwebserver/sw/rtx/decDetail/preview/{invoice_id},1'
            f.write(f"URL: {url}\n")


# 执行对比
check_files(Invoice_numbers_values, output_file)
