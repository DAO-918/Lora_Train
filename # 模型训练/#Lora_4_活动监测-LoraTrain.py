import time
import pynvml
import psutil
import pygetwindow as gw
import pyautogui


def get_gpu_utilization():
    try:
        handle = pynvml.nvmlDeviceGetHandleByIndex(0)
        utilization = pynvml.nvmlDeviceGetUtilizationRates(handle).gpu
        return utilization
    except pynvml.NVMLError as e:
        print(f"NVML error: {e}")
        return None


def get_gpu_memory_usage():
    try:
        handle = pynvml.nvmlDeviceGetHandleByIndex(0)
        memory_info = pynvml.nvmlDeviceGetMemoryInfo(handle)
        used_memory = memory_info.used / (1024 ** 3)
        return used_memory
    except pynvml.NVMLError as e:
        print(f"NVML error: {e}")
        return None


def find_bat_window():
    """尝试查找批处理窗口，支持多种窗口标题匹配方式
    返回找到的窗口对象，如果未找到则返回None
    """
    # 可能的窗口标题列表
    possible_titles = [
        "A启动脚本.bat",  # 直接使用批处理文件名作为标题
        "C:\\WINDOWS\\system32\\cmd.exe"  # 使用cmd.exe作为标题（批处理脚本运行时的实际窗口标题）
    ]
    
    # 尝试每一种可能的标题
    for title in possible_titles:
        windows = gw.getWindowsWithTitle(title)
        if windows:
            print(f"找到窗口: {title}，共 {len(windows)} 个")
            return windows[0]  # 返回第一个匹配的窗口
    
    # 如果所有标题都未找到匹配的窗口
    return None

def activate_bat_window_and_press_enter():
    """激活批处理窗口并按回车，支持多种窗口标题匹配方式"""
    try:
        # 尝试查找批处理窗口
        window = find_bat_window()
        if window:
            window.activate()
            pyautogui.press('enter')
            print(f"已激活批处理窗口 '{window.title}' 并按回车")
        else:
            print("未找到批处理窗口，请先手动启动批处理文件")
    except Exception as e:
        print(f"激活窗口时出错: {e}")


def shutdown_computer():
    import os
    os.system("shutdown /s /t 0")


if __name__ == "__main__":
    pynvml.nvmlInit()
    try:
        while True:
            utilization_records = []
            memory_records = []
            for _ in range(12):
                utilization = get_gpu_utilization()
                memory_usage = get_gpu_memory_usage()
                if utilization is not None:
                    utilization_records.append(utilization)
                if memory_usage is not None:
                    memory_records.append(memory_usage)
                time.sleep(5)

            if all(util < 50 for util in utilization_records) and any(mem > 15 for mem in memory_records):
                activate_bat_window_and_press_enter()
            if all(util < 10 for util in utilization_records) and all(mem < 3 for mem in memory_records):
                shutdown_computer()

    except KeyboardInterrupt:
        print("监测已停止。")
    finally:
        pynvml.nvmlShutdown()
    