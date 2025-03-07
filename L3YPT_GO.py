import time

from selenium import webdriver
from selenium.common import TimeoutException, InvalidSessionIdException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

import login

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
options.add_experimental_option('excludeSwitches', ['enable-automation'])  # 也是检查是否为软件自动控制
options.add_experimental_option("useAutomationExtension", False)  # 还是检查是否为软件自动控制
options.add_experimental_option("detach", True)  # 程序结束时浏览器不会自动关闭
options.add_argument("--force-device-scale-factor=0.95")
# 启用隐身模式
options.add_argument("--incognito")

driver = webdriver.Chrome(service=path, options=options)


# driver.implicitly_wait(3)


class L3YPT_run:
    def __init__(self):
        self.url = login.L3YPT_url
        self.username = login.L3YPT_username
        self.password = login.L3YPT_password
        self.login_L3YPT()
        self.judge_L3YPT()

    def login_L3YPT(self):
        # 打开链接
        driver.get(self.url)
        time.sleep(1)
        # 浏览器全屏，可有可无
        driver.maximize_window()
        # 找到输入框，这里需要自行在F12的Elements中找输入框的位置，然后在这里写入
        user_input = WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.XPATH, '//input[@placeholder="请填写用户邮箱"]')))
        pw_input = driver.find_element(by=By.XPATH, value='//input[@placeholder="请填写密码"]')
        # 输入用户名和密码，点击登录
        time.sleep(1)
        user_input.send_keys(self.username)
        time.sleep(1)
        pw_input.send_keys(self.password)
        time.sleep(1)
        driver.find_element(by=By.XPATH, value='//button[@type="submit"]').click()
        time.sleep(1)
        try:
            error_message = WebDriverWait(driver, 10).until(EC.presence_of_element_located
                                                            ((By.CSS_SELECTOR, 'p[class="es-con-tips"]')))
            if error_message.text == '此邮箱不是平台用户，请再次确认。' or '用户名或密码错误':
                time.sleep(2)
                return driver.close()
        except (TimeoutException, InvalidSessionIdException):
            pass

    def judge_L3YPT(self):
        try:
            driver.find_element(By.XPATH, '//span[text()="云监控服务"]').click()
            time.sleep(5)
            # 判断是否进入截图界面并截图保存
            logged_in = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'span[class="name-wrapper"]')))

            if logged_in:
                time.sleep(10)
                driver.save_screenshot('word_L3YPT.png')
                time.sleep(1)
                driver.close()
            else:
                print('截图失败')
        except TimeoutException:
            pass


if __name__ == "__main__":
    L3YPT_run()
