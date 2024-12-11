import requests
import subprocess
import platform
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pandas as pd
import socket
import time  # 导入 time 模块，用于延时

# 定义 ping 函数，并返回目标 IP 地址
def ping(host):
    param = "-n" if platform.system().lower() == "windows" else "-c"
    command = ["tcping", param, "1", host]  # 发送一次 ping 请求

    try:
        response = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if response.returncode == 0:
            ip_address = socket.gethostbyname(host)  # 获取 IP 地址
            return "可达", ip_address
        else:
            return "不可达", None
    except Exception as e:
        return f"错误: {e}", None

# 定义 HTTP 请求函数，跟随跳转并判断响应状态
def check_http_status(url):
    try:
        # 设置 30 秒的超时时间，自动跟随 301 和 302 跳转
        response = requests.get(url, timeout=30, allow_redirects=True)

        # 判断最终响应状态码
        if response.status_code == 200:
            return "可达"
        elif response.status_code in [301, 302]:
            return "重定向"
        else:
            return f"HTTP 错误: {response.status_code}"
    except requests.exceptions.Timeout:
        return "超时"
    except requests.exceptions.RequestException as e:
        return f"错误: {e}"

# 设置 Selenium Chrome 配置
def setup_driver():
    options = Options()
    options.headless = True  # 无头模式，不打开浏览器窗口
    options.add_argument('--ignore-certificate-errors')  # 忽略 SSL 证书错误
    options.add_argument('--disable-extensions')  # 禁用扩展程序
    options.add_argument('--disable-gpu')  # 禁用 GPU 加速
    options.add_argument('--no-sandbox')  # 避免权限问题

    # 指定 chromedriver 的路径
    chrome_driver_path = "chromedriver.exe"  # 请替换为实际路径
    
    # 使用 Service 来指定驱动程序路径
    service = Service(executable_path=chrome_driver_path)
    
    # 初始化 Chrome 浏览器
    driver = webdriver.Chrome(service=service, options=options)
    return driver

# 等待页面加载完成的函数
def wait_for_element(driver, element_xpath):
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, element_xpath))
        )
    except Exception as e:
        print(f"元素加载失败: {e}")

# 截图并保存到文件夹
def capture_screenshot(driver, url, status):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"screenshots/{status}_{timestamp}.png"
    
    # 创建截图文件夹
    os.makedirs("screenshots", exist_ok=True)
    
    driver.get(url)  # 访问网页

    # 等待页面上的某个特定元素加载完成，确保页面完全渲染
    wait_for_element(driver, '//*[contains(text(), "特定页面元素")]')  # 使用适当的元素定位符
    
    driver.save_screenshot(filename)  # 保存截图
    
    print(f"已保存 {url} 的截图，文件名为 {filename}")
    return filename

# 准备存储数据的列表
data = []

# 读取 txt 文件中的 URL 或 IP 地址并进行 ping 测试
with open('urls.txt', 'r') as file:
    urls = file.readlines()

# 设置 Selenium 浏览器驱动
driver = setup_driver()

for url in urls:
    url = url.strip()  # 去掉两端的空白字符
    
    # 确保 URL 包含 https:// 前缀
    if not url.lower().startswith('https://'):
        url = 'https://' + url
    
    # 执行 ping 测试并获取目标 IP 地址
    status, ip_address = ping(url.split("//")[-1].split("/")[0])  # 只传入主机名部分
    
    # 执行 HTTP 状态检查
    http_status = check_http_status(url)
    
    # 确认可达并且 HTTP 状态正常后，等待 3 秒钟
    if status == "可达" and http_status == "可达":
        time.sleep(3)  # 等待 3 秒
    
    # 设置 Cookie
    driver.get(url)  # 访问一个页面，开始浏览器会话
    driver.add_cookie({
        'name': 'YTH-IDP-ACCESS-TOKEN',
        'value': 'c33fac2164d638d3c1b3ab3adfde7a27_20220621215445466-DC6A-A2D2E087D',
        'domain': url.split("//")[1].split("/")[0]  # 从 URL 中提取域名
    })
    driver.add_cookie({
        'name': 'SESSION',
        'value': 'YTEzY2RmZGYtYmQ3My00MTBjLWE1YmQtYTVhYjQ1MjE2Yzlm',
        'domain': url.split("//")[1].split("/")[0]  # 从 URL 中提取域名
    })
    
    # 截图并保存
    screenshot = capture_screenshot(driver, url, http_status)
    
    # 将 URL, 状态, IP 地址和截图路径添加到列表
    data.append([url, status, ip_address, http_status, screenshot])

# 关闭浏览器驱动
driver.quit()

# 创建一个 DataFrame 并保存到 Excel
df = pd.DataFrame(data, columns=["URL", "连接状态", "IP 地址", "HTTP 状态", "截图路径"])

# 导出到 Excel
df.to_excel("url_connectivity_with_screenshots.xlsx", index=False, engine="openpyxl")

print("结果已保存到 'url_connectivity_with_screenshots.xlsx' 文件中。")
