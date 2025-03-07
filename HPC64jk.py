import os
import time

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
options.add_argument("--disable-javascript")
options.add_argument("--force-device-scale-factor=0.85")
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
date_dir = os.path.join(os.getcwd(), './HPC64')
options.add_argument(f"--user-data-dir={date_dir}")
options.add_argument('--start-maximized')  # 窗口最大化

driver = webdriver.Chrome(service=path, options=options)


class HPC63JK_run:
    def __init__(self):
        self.login_HPC64()
        self.picture_cat()

    def login_HPC64(self):
        url = 'http://10.190.103.11:18080/'
        username = 'admin123'
        password = 'Mnbv%67890'
        # 打开链接
        driver.get(url)
        # 浏览器全屏，可有可无
        user_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//input[@id="username"]')))
        pw_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="password"]')))
        # 输入用户名和密码，点击登录
        time.sleep(1)
        user_input.clear()
        user_input.send_keys(username)
        time.sleep(1)
        pw_input.clear()
        pw_input.send_keys(password)
        time.sleep(1)
        driver.find_element(by=By.XPATH, value='//button[@type="button"]').click()
        time.sleep(2)

    def picture_cat(self):
        driver.find_element(By.XPATH, '//p[text()="微模块监控"]').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//li[text()="液冷机柜"]').click()
        time.sleep(10)
        driver.save_screenshot('HPC64-液冷机柜.png')
        wd_pick = ['//p[text()="液冷监控"]', '//li[text()="水冷空调监控"]', '//li[text()="CDU(35KW)监控"]',
                   '//li[text()="通信电源监控"]', '//input[@type="text"]', '//span[text()="通信电源2"]']
        wd_cat = ['', "HPC64-水冷空调监控.png", "HPC64-CDU(35KW)监控.png", "HPC64-通信电源监控1.png", '',
                  "HPC64-通信电源监控2.png"]
        picture = zip(wd_pick, wd_cat)

        for wd_hj, wd_picture in picture:
            pick_picture = driver.find_element(By.XPATH, wd_hj)
            pick_picture.click()
            time.sleep(2)
            driver.save_screenshot(wd_picture)

        af_pick = ['//p[text()="告警管理"]', '//li[text()="实时告警"]']
        for af_picture in af_pick:
            af_cat = driver.find_element(By.XPATH, af_picture)
            if af_picture == '//p[text()="告警管理"]':
                af_cat.click()
                time.sleep(2)
            else:
                af_cat.click()
                time.sleep(2)
                driver.save_screenshot('HPC64-实时告警.png')
        driver.close()


if __name__ == "__main__":
    HPC63JK_run()
