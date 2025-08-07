import subprocess
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
import random

class ChromeManager:
    def __init__(self, config):
        self.chrome_path = config['chrome_path']
        self.chromedriver_path = config['chromedriver_path']
        self.user_data_dir = config['user_data_dir']
        self.port = config.get('remote_debugging_port', 9222 + random.randint(0, 1000))  # 随机端口避免冲突
        self.headless = config.get('enable_headless', False)
        self.extensions = config.get('enable_extensions', False)
        self.proxy = config.get('using_proxy', False)
        self.images = config.get('enable_images', False)
        self.process = None
        self.driver = None
        self.wait = None
    
    def get_chrome_command(self):
        command = [
            self.chrome_path,
            f"--remote-debugging-port={self.port}",
            f"--user-data-dir={self.user_data_dir}",
            "--excludeSwitches=enable-automation",
            "--window-size=1920,1080",
            "--disable-infobars",
        ]
        if self.headless:
            command.append("--headless=new") # 启用即无头模式
        if self.extensions:  # 注意逻辑改为not，因为默认是禁用
            command.append("--disable-extensions") # 启用即无扩展
        if self.proxy:
            command.append("--proxy-server=http://127.0.0.1:7890") # 使用代理
        if self.images:
            command.append("--blink-settings=imagesEnabled=false") # 禁用图像
        return command
    
    def start_chrome(self):
        if self.process and self.process.poll() is None:
            return True  # Chrome已在运行
        
        try:
            self.process = subprocess.Popen(self.get_chrome_command())
            time.sleep(3)  # 减少等待时间
            print(f"Chrome 已启动于端口 {self.port}")
            return True
        except Exception as e:
            print(f"启动 Chrome 失败: {e}")
            return False
    
    def get_driver(self):
        if self.driver and self.driver.service.is_connectable():
            return self.driver, self.wait

        try:
            service = Service(executable_path=self.chromedriver_path)
            options = Options()
            options.binary_location = self.chrome_path
            options.debugger_address = f'127.0.0.1:{self.port}'
            
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 10)
            return self.driver, self.wait
        except Exception as e:
            print(f"连接 Chrome 失败: {e}")
            return None, None
    
    def open_url(self, url, max_retries=3):
        for attempt in range(max_retries):
            if not self.start_chrome():
                continue
                
            driver, wait = self.get_driver()
            if not driver:
                continue
            
            try:
                driver.get(url)
                print(f"成功打开网页: {url}")
                return driver, wait
            except Exception as e:
                print(f"尝试 {attempt + 1}/{max_retries} 打开 {url} 失败: {e}")
                self.cleanup()
                time.sleep(2)
        
        print(f"经过 {max_retries} 次尝试仍无法打开 {url}")
        return None, None
    
    def cleanup(self):
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None
        if self.process and self.process.poll() is None:
            self.process.terminate()
            self.process = None
    
    def __del__(self):
        self.cleanup()

# 使用示例
if __name__ == "__main__":
    config = {
        'chrome_path': r'D:\Code\.module\chrome\chrome-win-115\chrome.exe',
        'chromedriver_path': r'D:\Code\.module\chrome-driver\115\chromedriver.exe',
        'user_data_dir': r"D:\Code\.module\chrome-cache\chrome-cache-115\AutomationProfile",
        'remote_debugging_port': 9222,  # 可选，默认为随机端口
        'headless': False,
        'enable_extensions': False,
        'using_proxy': False,
        'enable_images': False
    }

    # 创建实例
    chrome_manager = ChromeManager(config)
    
    # 打开网页
    driver, wait = chrome_manager.open_url("https://www.baidu.com")
    print("=========")
    # 使用完成后清理（可选，因为析构函数会自动清理）
    # chrome_manager.cleanup()