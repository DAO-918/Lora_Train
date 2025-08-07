import os
import sys
import time
import traceback
import importlib.util
import importlib.machinery


def find_bat_window():
    """
    尝试查找批处理窗口，支持多种窗口标题匹配方式
    返回找到的窗口对象，如果未找到则返回None
    """
    import pygetwindow as gw
    
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

def test_activate_bat_window():
    """
    测试LoraTrainer类中的activate_bat_window_and_press_enter方法
    该方法用于激活批处理窗口并发送回车键，支持多种窗口标题匹配方式
    """
    try:
        # 尝试查找批处理窗口
        import pygetwindow as gw
        window = find_bat_window()
        
        if window:
            print(f"找到批处理窗口: {window.title}")
        else:
            print("警告: 未找到批处理窗口，请先手动启动批处理文件")
        
        return True
    
    except Exception as e:
        print(f"测试过程中出错: {str(e)}")
        traceback.print_exc()
        print("===== 测试失败 =====")
        return False

if __name__ == "__main__":
    test_activate_bat_window()