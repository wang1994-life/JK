import threading
import time
from tkinter import messagebox


def L3py_files():
    import WLSJ_GO, WAF_GO, RZSJ_GO, LDSM_GO, IPS_GO, L3YPT_GO, FBD_GO, L3JK, L3FDPT
    # 指定要运行的Python文件
    scripts = [WLSJ_GO.WLSJ_run, WAF_GO.WAF_run, RZSJ_GO.RZSJ_run, LDSM_GO.LDSM_run, IPS_GO.IPS_run, L3YPT_GO.L3YPT_run,
               FBD_GO.FBD_run, L3JK.L3jk_run, L3FDPT.L3FDPT_run]
    # 间隔时间（秒）
    interval = 3
    processes = []
    # 异步启动所有脚本
    for script in scripts:
        process = threading.Thread(target=script, daemon=True)
        processes.append(process)
        process.start()
        # 在启动下一个脚本之前等待一段时间
        time.sleep(interval)

    # 等待所有进程结束
    for process in processes:
        process.join()
    py_completed()

def L3GLQoy_run():
    import GLQFBD
    scripts = [GLQFBD.GLQFBD_run]
    # 间隔时间（秒）
    interval = 3
    processes = []
    # 异步启动所有脚本
    for script in scripts:
        process = threading.Thread(target=script, daemon=True)
        processes.append(process)
        process.start()
        # 在启动下一个脚本之前等待一段时间
        time.sleep(interval)

    # 等待所有进程结束
    for process in processes:
        process.join()
    py_completed()



def L2py_files():
    import L2JK, L2FDPT
    # 指定要运行的Python文件
    scripts = [L2JK.L2JK_run, L2FDPT.L2FDPT_run]
    # 间隔时间（秒）
    interval = 5
    processes = []
    # 异步启动所有脚本
    for script in scripts:
        process = threading.Thread(target=script, daemon=True)
        processes.append(process)
        process.start()
        # 在启动下一个脚本之前等待一段时间
        time.sleep(interval)

        # 等待所有进程结束
    for process in processes:
        process.join()
    py_completed()


def L5py_files():
    import L5JK, L5XJPT
    # 指定要运行的Python文件
    scripts = [L5JK.L5JK_run, L5XJPT.L5XJPT_run]
    # 间隔时间（秒）
    interval = 5
    processes = []
    # 异步启动所有脚本
    for script in scripts:
        process = threading.Thread(target=script, daemon=True)
        processes.append(process)
        process.start()
        # 在启动下一个脚本之前等待一段时间
        time.sleep(interval)

        # 等待所有进程结束
    for process in processes:
        process.join()
    py_completed()


def S2py_files():
    import S2JK, S2XJPT, S2JKSP
    # 指定要运行的Python文件
    scripts = [S2JK.S2JK_run, S2XJPT.S2XJPT_run, S2JKSP.S2JKSP_run]
    # 间隔时间（秒）
    interval = 5
    processes = []
    # 异步启动所有脚本
    for script in scripts:
        process = threading.Thread(target=script, daemon=True)
        processes.append(process)
        process.start()
        # 在启动下一个脚本之前等待一段时间
        time.sleep(interval)

        # 等待所有进程结束
    for process in processes:
        process.join()
    py_completed()


def HPC64py_files():
    import HPC64jk
    # 指定要运行的Python文件
    scripts = [HPC64jk.HPC63JK_run]
    # 间隔时间（秒）
    interval = 5
    processes = []
    # 异步启动所有脚本
    for script in scripts:
        process = threading.Thread(target=script, daemon=True)
        processes.append(process)
        process.start()
        # 在启动下一个脚本之前等待一段时间
        time.sleep(interval)

        # 等待所有进程结束
    for process in processes:
        process.join()
    py_completed()


def L6ZZJD_files():
    import L6ZZJD, L6ZZJD_XJPT
    # 指定要运行的Python文件
    scripts = [L6ZZJD.L6JK_run, L6ZZJD_XJPT.L6XJPT_run]
    # 间隔时间（秒）
    interval = 5
    processes = []
    # 异步启动所有脚本
    for script in scripts:
        process = threading.Thread(target=script, daemon=True)
        processes.append(process)
        process.start()
        # 在启动下一个脚本之前等待一段时间
        time.sleep(interval)

        # 等待所有进程结束
    for process in processes:
        process.join()
    py_completed()


def py_completed():
    messagebox.showinfo("任务完成", "任务已经成功完成！")


if __name__ == "__main__":
    print('')
