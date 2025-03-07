import os
import time

from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

# 指向驱动位置
# 下载地址：https://chromedriver.storage.googleapis.com/index.html path = Service( '../venv/chromedriver.exe')
path = Service('')
# 配置选项
options = webdriver.ChromeOptions()
# ---------------------------优化选项---------------------------------- #
# 忽略证书错误
options.add_argument('--ignore-certificate-errors')
# 忽略 Bluetooth: bluetooth_adapter_winrt.cc:1075 Getting Default Adapter failed. 错误
options.add_experimental_option('excludeSwitches', ['enable-automation'])
# 忽略 DevTools listening on ws://127.0.0.1... 提示
options.add_experimental_option('excludeSwitches', ['enable-logging'])
# 防止登录后自动关闭浏览器
options.add_experimental_option('detach', True)
# 关闭Chrome浏览器受自动控制的提示
options.add_argument("--disable-blink-features=AutomationControlled")  # 检查是否为软件自动控制
options.add_argument("--ignore-certificate-errors")  # 忽略证书错误
options.add_argument("--ignore-ssl-errors")  # 忽略ssl错误
options.add_argument('--disable-gpu')  # 关闭GPU加速
options.add_argument('--no-sandbox')  # 以最高权限运行
options.add_argument('--log_level=3')  # 日志级别（在命令行里不会报一堆干扰信息）
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_experimental_option('excludeSwitches', ['enable-automation'])  # 也是检查是否为软件自动控制
options.add_experimental_option("useAutomationExtension", False)  # 还是检查是否为软件自动控制
options.add_argument("--force-device-scale-factor=0.97")
# 禁用密码保存弹窗
options.add_argument("--disable-infobars")
options.add_argument("--disable-extensions")
options.add_argument("--disable-popup-blocking")
options.add_experimental_option('prefs', {
    'credentials_enable_service': False,
    'profile.default_content_setting_values': {
        'passwords': 2
    }
})
date_dir = os.path.join(os.getcwd(), 'S2虚拟化平台')
options.add_argument(f"--user-data-dir={date_dir}")
options.add_argument('--start-maximized')  # 窗口最大化

driver = webdriver.Chrome(service=path, options=options)


class S2XJPT_run:
    def __init__(self):
        self.login_S2xnj()

    def login_S2xnj(self):
        url = 'https://10.140.10.1/ui/'
        username = 'administrator@vsphere.local'
        password = 'P@ssw0rd01!'
        # 打开链接
        driver.get(url)
        time.sleep(2)
        user_input = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//input[@id="username"]')))
        pw_input = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//input[@id="password"]')))
        # 输入用户名和密码，点击登录
        user_input.send_keys(username)
        pw_input.send_keys(password)
        login_button = driver.find_element(By.XPATH, '//*[@id="submit"]')
        driver.execute_script("arguments[0].click();", login_button)

        login_button2 = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, '//button[text()="虚拟机"]')))
        driver.execute_script("arguments[0].click();", login_button2)
        element_xj = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//vsc-dg-cell-container[text()="已打开电源"]')))
        if element_xj:
            driver.save_screenshot('S2-虚拟机.png')
            # 加载图像
            image = Image.open('S2-虚拟机.png')
            # 获取图像的宽度和高度
            left = 390
            top = 20
            # 设定你想要裁剪的区域，确保不超出原图的边界
            crop_box = (left, top, 1300, 730)
            # 裁剪图像
            cropped_image = image.crop(crop_box)
            # 保存或处理裁剪后的图像
            cropped_image.save('S2-虚拟机.png')
            time.sleep(5)
            login_button3 = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//button[text()="数据存储"]')))
            driver.execute_script("arguments[0].click();", login_button3)

        element_store = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//div[@class="datagrid-row-scrollable"]')))
        if element_store:
            driver.save_screenshot('S2-虚机存储.png')
            # 加载图像
            image = Image.open('S2-虚机存储.png')
            # 获取图像的宽度和高度
            left = 390
            top = 25
            # 设定你想要裁剪的区域，确保不超出原图的边界
            crop_box = (left, top, 1300, 400)
            # 裁剪图像
            cropped_image = image.crop(crop_box)
            # 保存或处理裁剪后的图像
            cropped_image.save('S2-虚机存储.png')
        time.sleep(3)
        driver.close()


if __name__ == "__main__":
    S2XJPT_run()
