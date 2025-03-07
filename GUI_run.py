import logging
import sys
import tkinter as tk
from tkinter import scrolledtext

import psutil

import GUI_module

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class StdoutRedirector:
    def __init__(self, text_widget):
        self.text_space = text_widget

    def write(self, string):
        self.text_space.insert(tk.END, string)
        self.text_space.see(tk.END)

    def flush(self):
        pass


# 用于关闭指定名称进程并确保所有后台进程结束的函数
def close_all():
    process_names = ['chrome', 'firefox']
    for proc in psutil.process_iter(['pid', 'name']):
        process_name = proc.info['name'].lower()
        for target_name in process_names:
            if target_name in process_name:
                try:
                    logging.info(f"尝试终止进程: {proc.info['name']} (PID: {proc.info['pid']})")
                    proc.terminate()  # 尝试优雅地终止进程
                    proc.wait(timeout=3)  # 等待进程结束，最多等待3秒
                except psutil.TimeoutExpired:
                    try:
                        logging.warning(f"进程 {proc.info['name']} (PID: {proc.info['pid']}) 未在3秒内终止，尝试强制杀死")
                        proc.kill()  # 强制杀死进程
                    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
                        logging.error(f"无法杀死进程 {proc.info['name']} (PID: {proc.info['pid']}): {e}")
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
                    logging.error(f"无法终止进程 {proc.info['name']} (PID: {proc.info['pid']}): {e}")
    # 如果你有其他后台线程或者进程需要退出，进行处理
    # sys.exit(0)  # 强制退出程序


def main():
    def close_all_and_quit():
        root.destroy()

    root = tk.Tk()
    root.title("巡检日报半自动生成器")
    root.geometry("1000x600")
    text_area = scrolledtext.ScrolledText(root, width=139, height=19)
    text_area.place(x=5, y=340)

    # 重定向标准输出和标准错误输出
    sys.stdout = StdoutRedirector(text_area)
    sys.stderr = StdoutRedirector(text_area)

    open_L3button = tk.Button(root, text="L3巡检文档生成点击", command=GUI_module.open_L3file, width=20, height=2)
    open_L3button.place(x=350, y=10)
    open_L2button = tk.Button(root, text="L2巡检文档生成点击", command=GUI_module.open_L2file, width=20, height=2)
    open_L2button.place(x=350, y=70)
    open_L5button = tk.Button(root, text="L5巡检文档生成点击", command=GUI_module.open_L5file, width=20, height=2)
    open_L5button.place(x=350, y=130)
    open_S2button = tk.Button(root, text="S2巡检文档生成点击", command=GUI_module.open_S2file, width=20, height=2)
    open_S2button.place(x=350, y=200)
    open_HPC64button = tk.Button(root, text="HPC64巡检文档生成点击", command=GUI_module.open_HPC64file, width=20, height=2)
    open_HPC64button.place(x=350, y=270)
    open_L6ZZJDbutton = tk.Button(root, text="L6总装基地巡检文档生成点击", command=GUI_module.open_L6ZZJD_file, width=23,height=2)
    open_L6ZZJDbutton.place(x=770, y=270)
    L3py_button = tk.Button(root, text="L3截图程序运行点击", command=GUI_module.run_L3py, width=20, height=2)
    L3py_button.place(x=150, y=10)
    L3GLQ_button = tk.Button(root, text="L3隔离区FBD截图程序运行", command=GUI_module.run_L3GLQFBD, width=25, height=2)
    L3GLQ_button.place(x=550, y=10)
    L2py_button = tk.Button(root, text="L2截图程序运行点击", command=GUI_module.run_L2py, width=20, height=2)
    L2py_button.place(x=150, y=70)
    L5py_button = tk.Button(root, text="L5截图程序运行点击", command=GUI_module.run_L5py, width=20, height=2)
    L5py_button.place(x=150, y=130)
    S2py_button = tk.Button(root, text="S2截图程序运行点击", command=GUI_module.run_S2py, width=20, height=2)
    S2py_button.place(x=150, y=200)
    HPC64_button = tk.Button(root, text="HPC64截图程序运行点击", command=GUI_module.run_HPC64py, width=20, height=2)
    HPC64_button.place(x=150, y=270)
    L6ZZJD_button = tk.Button(root, text="L6总装基地截图程序运行点击", command=GUI_module.run_L6ZZJDpy, width=23, height=2)
    L6ZZJD_button.place(x=550, y=270)
    exit_button = tk.Button(root, text="关闭所有谷歌浏览器", command=close_all)
    exit_button.place(x=880, y=0)
    # 创建一个按钮，并绑定文件选择功能
    execl_button = tk.Button(root, text="选择巡检信息表", command=GUI_module.load_excel)
    execl_button.place(x=20, y=20)

    open_L3button.config(font=("Arial", 12))
    open_L2button.config(font=("Arial", 12))
    open_L5button.config(font=("Arial", 12))
    open_S2button.config(font=("Arial", 12))
    open_HPC64button.config(font=("Arial", 12))
    open_L6ZZJDbutton.config(font=("Arial", 12))
    L3py_button.config(font=("Arial", 12))
    L3GLQ_button.config(font=("Arial", 12))
    L2py_button.config(font=("Arial", 12))
    L5py_button.config(font=("Arial", 12))
    S2py_button.config(font=("Arial", 12))
    HPC64_button.config(font=("Arial", 12))
    L6ZZJD_button.config(font=("Arial", 12))

    root.protocol("WM_DELETE_WINDOW", close_all_and_quit)
    root.mainloop()


if __name__ == "__main__":
    main()
