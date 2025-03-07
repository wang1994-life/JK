import os
import time

from PIL import Image
from selenium import webdriver
from selenium.common import TimeoutException, StaleElementReferenceException
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
# 防止登录后自动关闭浏览器
options.add_experimental_option('detach', True)
# 关闭Chrome浏览器受自动控制的提示
options.add_argument('--no-sandbox')  # 允许无沙盒模式
options.add_argument('--disable-dev-shm-usage')  # 解决资源限制问题
options.add_argument("--disable-blink-features=AutomationControlled")  # 检查是否为软件自动控制
options.add_argument('--disable-gpu')  # 关闭GPU加速
options.add_argument('--no-sandbox')  # 以最高权限运行
options.add_argument('--log_level=3')  # 日志级别（在命令行里不会报一堆干扰信息）
options.add_experimental_option('excludeSwitches', ['enable-automation'])  # 也是检查是否为软件自动控制
options.add_experimental_option("useAutomationExtension", False)  # 还是检查是否为软件自动控制
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
date_dir = os.path.join(os.getcwd(), 'L2FD')
options.add_argument(f"--user-data-dir={date_dir}")
options.add_argument('--start-maximized')  # 窗口最大化

driver = webdriver.Chrome(service=path, options=options)


class L2FDPT_run:
    def __init__(self):
        self.login_L2fd()

    def login_L2fd(self):
        url = 'http://192.168.30.80/dcom/#'
        username = 'admin'
        password = 'admin'
        # 打开链接
        driver.get(url)

        try:
            user_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//input[@id="user"]')))
            pw_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//input[@id="password"]')))
            # 输入用户名和密码，点击登录
            user_input.clear()
            user_input.send_keys(username)
            pw_input.clear()
            pw_input.send_keys(password)
            driver.find_element(by=By.XPATH, value='//span[@class="ng-binding"]').click()
        except (TimeoutException, StaleElementReferenceException):
            pass

        login_fd = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//span[@title="进入超算系统"]')))
        if login_fd.text == '进入超算系统':
            login_fd.click()
        else:
            ActionChains(driver).click(on_element=None).perform()
            time.sleep(1)
            login_fd.click()
        time.sleep(2)
        click_button = ['//span[text()="5"]', '//span[text()="1"]', '//span[text()="2"]', '//span[text()="1"]',
                        '//span[text()="0"]', '//span[text()="1"]']
        for button in click_button:
            driver.find_element(By.XPATH, button).click()

        read_picture = ['第一套立方体监控视图', '第二套立方体监控视图', '平台机柜监控图']
        save_picture = ["L2第一套立方体监控视图.png", "L2第二套立方体监控视图.png", "L2平台机柜监控图.png"]

        zipped = zip(read_picture, save_picture)

        for pick in read_picture:
            pick_name = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.LINK_TEXT, pick)))
            pick_name.click()
            if pick_name.text == '第二套立方体监控视图':
                pick_name.click()
                time.sleep(60)

        for cat_picture, down_picture in zipped:
            tag_name = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.LINK_TEXT, cat_picture)))
            tag_name.click()
            element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.TAG_NAME, 'html')))
            if element:
                driver.save_screenshot(down_picture)
            # 加载图像
            image = Image.open(down_picture)

            # 获取图像的宽度和高度
            width, height = image.size
            left = 200
            top = 50
            # 设定你想要裁剪的区域，确保不超出原图的边界
            crop_box = (left, top, width, height)
            # 裁剪图像
            cropped_image = image.crop(crop_box)
            # 保存或处理裁剪后的图像
            cropped_image.save(down_picture)
        driver.close()


if __name__ == "__main__":
    L2FDPT_run()
