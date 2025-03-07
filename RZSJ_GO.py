import time

import ddddocr
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

import login

ocr = ddddocr.DdddOcr(show_ad=False, beta=True)
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
options.add_experimental_option('useAutomationExtension', False)
options.add_experimental_option('excludeSwitches', ['enable-automation'])
options.add_argument("--disable-blink-features=AutomationControlled")  # 检查是否为软件自动控制
options.add_argument("--ignore-certificate-errors")  # 忽略证书错误
options.add_argument("--ignore-ssl-errors")  # 忽略ssl错误
options.add_argument('--disable-gpu')  # 关闭GPU加速
options.add_argument('--no-sandbox')  # 以最高权限运行
options.add_argument('--log_level=3')  # 日志级别（在命令行里不会报一堆干扰信息）
options.add_experimental_option('excludeSwitches', ['enable-automation'])  # 也是检查是否为软件自动控制
options.add_experimental_option("useAutomationExtension", False)  # 还是检查是否为软件自动控制
options.add_experimental_option("detach", True)  # 程序结束时浏览器不会自动关闭
# 启用隐身模式
options.add_argument("--incognito")

driver = webdriver.Chrome(service=path, options=options)


class RZSJ_run:
    def __init__(self):
        self.url = login.RZSJ_url
        self.username = login.RZSJ_username
        self.password = login.RZSJ_password
        self.login_rzsj()
        self.captcha_rzsj()
        self.judge_rzsj()

    def login_rzsj(self):
        # 打开链接
        driver.get(self.url)
        time.sleep(1)
        # 浏览器全屏，可有可无
        driver.maximize_window()

        # 找到输入框，这里需要自行在F12的Elements中找输入框的位置，然后在这里写入
        user_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="text"]')))
        pw_input = driver.find_element(by=By.XPATH, value='//input[@id="password"]')
        # 输入用户名和密码，点击登录
        time.sleep(1)
        user_input.send_keys(self.username)
        time.sleep(1)
        pw_input.send_keys(self.password)
        time.sleep(1)

    def captcha_rzsj(self):
        # 验证码截图保存
        captcha_element = driver.find_element(by=By.XPATH, value='//img[@id="validCodeImg"]')
        captcha_element.screenshot('img_RZSJ.png')
        driver.find_element(by=By.XPATH, value='//input[@id="captcha"]').clear()
        captcha_input = driver.find_element(by=By.XPATH, value='//input[@id="captcha"]')
        # 验证码识别输入
        with open('img_RZSJ.png', 'rb') as f:  # 打开图片
            img_bytes = f.read()  # 读取图片
        res = ocr.classification(img_bytes)
        captcha_input.send_keys(res)
        # 点击登录
        driver.find_element(by=By.XPATH, value='//button[@type="submit"]').click()
        time.sleep(2)
        counter = 0
        while True:
            try:
                # 判断验证码输入状态
                error_message = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class=msg-error]')))
                if error_message.text == '验证码错误！':
                    counter += 1
                    self.captcha_rzsj()  # 验证码输入错误重新输入
                elif counter == 10:
                    break
                else:
                    break
            except TimeoutException:
                pass
            break

    def judge_rzsj(self):
        # 点击进入需要截图界面
        action = ActionChains(driver)
        action.move_to_element(driver.find_element(By.XPATH, "//a[text()='告警']")).perform()
        driver.find_element(By.XPATH, "//a[text()='告警摘要']").click()
        time.sleep(2)
        # 判断是否进入截图界面并截图保存
        logged_in = driver.find_element(by=By.XPATH, value='//div[@class="panel-title"]')
        if logged_in.text == '查询条件':
            time.sleep(2)
            driver.save_screenshot('word_RZSJ.png')
            time.sleep(1)
            driver.close()
        else:
            print('截图失败')


if __name__ == "__main__":
    RZSJ_run()
