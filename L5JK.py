import os
import time

from selenium import webdriver
from selenium.common import InvalidSelectorException, TimeoutException
from selenium.webdriver import ActionChains
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
date_dir = os.path.join(os.getcwd(), 'L5')
options.add_argument(f"--user-data-dir={date_dir}")
options.add_argument('--start-maximized')  # 窗口最大化

driver = webdriver.Chrome(service=path, options=options)


class L5JK_run:
    def __init__(self):
        self.login_L5jk()
        self.picture_cat()

    def login_L5jk(self):
        url = 'http://10.200.10.200/dcom/#'
        username = 'admin'
        password = 'admin'
        # 打开链接
        driver.get(url)
        time.sleep(2)
        try:
            user_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//input[@id="user"]')))
            pw_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//input[@id="password"]')))
            # 输入用户名和密码，点击登录
            user_input.send_keys(username)
            pw_input.send_keys(password)
            driver.find_element(by=By.XPATH, value='//div[@class="login-button"]').click()
        except TimeoutException:
            pass
            login_jk = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//span[@title="进入动环监控系统"]')))
            if login_jk.text == '进入动环监控系统':
                login_jk.click()
            else:
                ActionChains(driver).click(on_element=None).perform()
                time.sleep(1)
                login_jk.click()

        click_button = ['//span[text()="5"]', '//span[text()="1"]', '//span[text()="2"]', '//span[text()="1"]',
                        '//span[text()="0"]', '//span[text()="1"]']
        for button in click_button:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, button))).click()
        try:
            # 判断是否进入截图界面并截图保存
            logged_in = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="content"]')))
            if logged_in:
                time.sleep(20)
                driver.save_screenshot('word_L5JKZY.png')
                driver.find_element(by=By.XPATH, value='//div[@class="pieCharts"]').click()
                time.sleep(5)
                driver.save_screenshot('word_L5KGJ.png')
            else:
                print('截图失败')
        except InvalidSelectorException:
            pass

    def picture_cat(self):
        driver.find_element(by=By.XPATH, value='//i[text()="menu"]').click()
        time.sleep(1)
        driver.find_element(by=By.XPATH, value='//span[text()="设备监控"]').click()
        ise = [4, 13, 14, 17, 22, 27, 31]
        time.sleep(3)
        for i in ise:
            time.sleep(1)
            driver.find_elements(By.CLASS_NAME, "fancytree-expander")[i].click()
        time.sleep(1)

        pick_button = ["//span[text()='智能电表（512整机及IT系统电量）']"]
        cat_dy = ["word_L5-512-智能电表.png"]
        cat_gl = ["word_L5-512-总有功功率.png"]
        zipped_dygl = zip(pick_button, cat_dy, cat_gl)

        for button, dy, gl in zipped_dygl:
            driver.find_element(By.XPATH, button).click()
            time.sleep(3)
            element1 = driver.find_element(By.CLASS_NAME, 'cabinet-box')
            element1.screenshot(dy)
            driver.find_element(By.XPATH, "//span[text()='总有功功率(kW)']").click()
            time.sleep(3)
            element1kw = driver.find_element(By.CLASS_NAME, 'historyData')
            time.sleep(3)
            element1kw.screenshot(gl)
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'span[translate="DOOR.CLOSE"]'))).click()
            time.sleep(2)

        wd_pick = ['//span[text()="温湿度-512-RACK1"]', '//span[text()="温湿度-512-RACK2"]', '//span[text()="温湿度-512-RACK3"]',
                   '//span[text()="温湿度-512-RACK4"]', '//span[text()="温湿度-动力配套室东"]', '//span[text()="温湿度-水泵房"]'
            , '//span[text()="CDU-1"]']
        wd_cat = ["L5-温湿度-512-RACK1.png", "L5-温湿度-512-RACK2.png", "L5-温湿度-512-RACK3.png", "L5-温湿度-512-RACK4.png",
                  "L5-温湿度-动力配套室东.png", "L5-温湿度-水泵房.png", "L5-CDU-1.png"]
        hjwd = zip(wd_pick, wd_cat)
        for wd_hj, wd_picture in hjwd:
            driver.find_element(by=By.XPATH, value=wd_hj).click()
            time.sleep(5)
            wd1 = driver.find_element(By.CLASS_NAME, 'cabinet-box')
            wd1.screenshot(wd_picture)
        driver.find_element(by=By.XPATH, value='//span[text()="门禁管理"]').click()
        time.sleep(2)
        driver.find_element(By.XPATH, '//span[text()="门禁监控"]').click()
        time.sleep(2)
        driver.save_screenshot('L5-门禁监控.png')
        driver.close()


if __name__ == "__main__":
    L5JK_run()
