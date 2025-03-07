import os
import shutil
import time

import win32con
import win32gui
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
options.add_argument("--force-device-scale-factor=0.98")
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
options.add_argument('--start-maximized')  # 窗口最大化

driver = webdriver.Chrome(service=path, options=options)


class S2JKSP_run:
    def __init__(self):
        self.login_S2jksp()



    def login_S2jksp(self, ):
        target_title = "10.190.101.13"
        url = 'http://10.190.101.13/'
        username = 'admin'
        password = 'IPc5123568'
        # 打开链接
        driver.get(url)
        time.sleep(2)
        user_input = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//input[@id="szUserName"]')))
        pw_input = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//input[@id="szUserPasswdSrc"]')))
        # 输入用户名和密码，点击登录
        user_input.clear()
        user_input.send_keys(username)
        pw_input.clear()
        pw_input.send_keys(password)
        time.sleep(1)
        driver.find_element(by=By.XPATH, value='//span[@title="登录"]').click()
        time.sleep(5)
        iframe_element = driver.find_element(By.ID, 'container')
        driver.switch_to.frame(iframe_element)
        driver.find_element(by=By.XPATH, value='//button[@id="setWinNum"]').click()
        time.sleep(3)
        button = driver.find_element(by=By.XPATH, value='//*[@id="chsWinDlg"]/button[1]')
        driver.execute_script("arguments[0].click();", button)

        video_button = ['//span[@title="D1 (摄像机 01)"]','//span[@title="D2 (摄像机 02)"]']
        picture_name = ['S2-监控后视图.png','S2-监控前视图.png']

        zip_files = zip(video_button,picture_name)
        for cat_picture,save_picture in zip_files:
            driver.find_element(by=By.XPATH, value=cat_picture).click()
            time.sleep(3)
            driver.find_element(by=By.XPATH, value='//button[@id="snap"]').click()
            time.sleep(2)
            driver.find_element(by=By.XPATH, value='//span[@class="msg-detail ellipsis"]').click()
            source_path = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, '//a[@class="a-exec-func"]')))
            source_path.click()
            content = source_path.text
            # 提取文件路径
            colon_index = content.find(' ')
            time.sleep(3)
            self.close_specific_explorer_window(target_title)
            if colon_index != -1:
                # 截取冒号后面的部分并去除首尾空格
                file_path = content[colon_index + 1:].strip()

                # 检查文件是否存在
                if os.path.exists(file_path):
                    try:
                        # 获取当前目录
                        current_dir = os.getcwd()
                        # 定义新的文件名
                        new_file_name = f"{save_picture}"
                        # 构建目标文件路径
                        dest_path = os.path.join(current_dir, new_file_name)
                        # 移动文件
                        shutil.move(file_path, dest_path)
                        image_path = save_picture
                        image = Image.open(image_path)
                        converted_image_path = save_picture
                        image.save(converted_image_path, 'png')
                        print(f"文件 {new_file_name} 已成功移动到当前目录。")
                    except PermissionError:
                        print("权限不足，无法移动文件。")
                    except Exception as e:
                        print(f"移动文件时出现错误: {e}")
                else:
                    print(f"文件 {file_path} 不存在，无法移动。")
            else:
                print("输入字符串中未找到冒号，无法提取文件路径。")

        time.sleep(2)
        driver.close()

    def close_specific_explorer_window(self,window_title):
        def callback(hwnd, _):
            class_name = win32gui.GetClassName(hwnd)
            # 检查窗口类名是否为文件资源管理器的类名
            if class_name == 'CabinetWClass':
                # 获取窗口标题
                title = win32gui.GetWindowText(hwnd)
                if window_title in title:
                    # 发送关闭窗口的消息
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

        # 枚举所有顶级窗口，并对每个窗口调用 callback 函数
        win32gui.EnumWindows(callback, None)

if __name__ == "__main__":
    S2JKSP_run()

