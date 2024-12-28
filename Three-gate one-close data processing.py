import os
import pandas as pd
import requests
import json
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import sys
import time
import glob
import logging
import subprocess
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.common.exceptions import SessionNotCreatedException

CONFIG_FILE = 'config1.json'
LOG_FILE = '三关一闭.log'


# 配置log
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def get_driver_path(browser):
    # 如果程序被打包成可执行文件，则从临时文件夹中查找驱动
    base_path = os.path.join(sys._MEIPASS, 'drivers') if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'drivers')
    if browser == 'chrome':
        return os.path.join(base_path, 'chromedriver.exe')
    elif browser == 'firefox':
        return os.path.join(base_path, 'geckodriver.exe')
    elif browser == 'edge':
        return os.path.join(base_path, 'msedgedriver.exe')
    return None

class ConfigWindow:
    def __init__(self, parent=None):
        self.window = tk.Toplevel(parent) if parent else tk.Tk()
        self.window.title("配置信息")
        self.window.geometry("500x350")  # 调整宽度以容纳按钮
        self.window.resizable(False, False)

        if parent:
            parent.iconify() # 最小化父窗口
            self.window.protocol("WM_DELETE_WINDOW", lambda: self.on_closing(parent))
        else:
            self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.config = self.load_config()
        self.create_widgets()

    def create_widgets(self):
        # PIN码
        pin_frame = tk.Frame(self.window)
        pin_frame.pack(pady=5, fill=tk.X, padx=10)
        tk.Label(pin_frame, text="PIN码:", font=('Arial', 10), width=15, anchor='w').pack(side=tk.LEFT)
        self.pin_entry = tk.Entry(pin_frame)
        self.pin_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Webhook URL
        webhook_frame = tk.Frame(self.window)
        webhook_frame.pack(pady=5, fill=tk.X, padx=10)
        tk.Label(webhook_frame, text="Webhook URL:", font=('Arial', 10), width=15, anchor='w').pack(side=tk.LEFT)
        self.webhook_entry = tk.Entry(webhook_frame)
        self.webhook_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 下载文件夹路径
        download_frame = tk.Frame(self.window)
        download_frame.pack(pady=5, fill=tk.X, padx=10)
        tk.Label(download_frame, text="下载文件夹路径:", font=('Arial', 10), width=15, anchor='w').pack(side=tk.LEFT)
        self.download_path_var = tk.StringVar(self.window)
        self.download_path_entry = tk.Entry(download_frame, textvariable=self.download_path_var)
        self.download_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(download_frame, text="选择", command=self.browse_directory).pack(side=tk.LEFT, padx=5)

        # 默认浏览器
        browser_frame = tk.Frame(self.window)
        browser_frame.pack(pady=5, fill=tk.X, padx=10)
        tk.Label(browser_frame, text="选择浏览器:", font=('Arial', 10), width=15, anchor='w').pack(side=tk.LEFT)
        self.browser_var = tk.StringVar(self.window)
        self.browser_combobox = ttk.Combobox(
            browser_frame,
            textvariable=self.browser_var,
            values=["chrome", "firefox", "edge"]
        )
        self.browser_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 加载配置信息到界面
        if self.config:
            self.pin_entry.insert(0, self.config.get('pin_code', ''))
            self.webhook_entry.insert(0, self.config.get('webhook_url', ''))
            self.browser_var.set(self.config.get('browser', 'chrome'))
            self.download_path_var.set(self.config.get('download_path', ''))

        self.status_label = tk.Label(self.window, text="", fg="green")
        self.status_label.pack(pady=5)

        # 保存配置按钮居中
        save_button_frame = tk.Frame(self.window)
        save_button_frame.pack(side=tk.BOTTOM, pady=15)
        ttk.Button(
            save_button_frame,
            text="保存配置",
            command=self.save
        ).pack()

    def browse_directory(self):
        # 弹出文件夹选择对话框
        directory = filedialog.askdirectory(initialdir=self.download_path_var.get() or os.path.expanduser("~"))
        if directory:
            self.download_path_var.set(directory)

    def load_config(self):
        logging.info("Loading configuration...")
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    return json.load(f)
            except json.JSONDecodeError as e:
                logging.error(f"Error decoding config file: {e}")
                messagebox.showerror("Error", f"Error decoding config file: {e}")
                return None
        return None

    def save(self):
        pin_code = self.pin_entry.get().strip()
        webhook_url = self.webhook_entry.get().strip()
        browser = self.browser_var.get()
        download_path = self.download_path_var.get().strip()

        if not pin_code or not webhook_url or not browser or not download_path:
            messagebox.showerror("错误", "请填写所有字段！")
            return

        config = {
            'pin_code': pin_code,
            'webhook_url': webhook_url,
            'browser': browser,
            'download_path': download_path
        }

        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=4)
            self.status_label.config(text="配置已保存！", fg="green")
            logging.info("Configuration saved successfully.")
            self.window.after(2000, lambda: self.status_label.config(text="")) # 2秒后清除状态信息
        except Exception as e:
            logging.error(f"Error saving config file: {e}")
            messagebox.showerror("Error", f"Error saving config file: {e}")

    def on_closing(self, parent=None):
        # 关闭配置窗口时的处理
        if not os.path.exists(CONFIG_FILE):
            if messagebox.askokcancel("退出", "配置尚未保存，确定要退出吗？"):
                self.window.destroy()
                if parent:
                    parent.deiconify() # 恢复父窗口
        else:
            self.window.destroy()
            if parent:
                parent.deiconify() # 恢复父窗口

class MainWindow:
    def __init__(self):
        # 确保 drivers 文件夹存在
        script_dir = os.path.dirname(os.path.abspath(__file__))
        drivers_dir = os.path.join(script_dir, 'drivers')
        if not os.path.exists(drivers_dir):
            os.makedirs(drivers_dir)

        self.root = tk.Tk()
        self.root.title("三关一闭未完成项提醒")
        self.root.geometry("350x200") # 调整高度
        self.create_widgets()
        self.bind_shortcuts()
        self.root.protocol("WM_DELETE_WINDOW", self.close_application)

        # 检查是否以自动模式运行
        if len(sys.argv) > 1 and sys.argv[1] == '--auto':
            logging.info("Running in auto mode.")
            self.run_automation()
        else:
            logging.info("Running in GUI mode.")

    def create_widgets(self):
        ttk.Button(
            self.root,
            text="设置",
            command=self.open_config
        ).pack(pady=20)

        ttk.Button(
            self.root,
            text="运行",
            command=self.run_automation
        ).pack(pady=20)

        self.status_label = tk.Label(self.root, text="", fg="black")
        self.status_label.pack(pady=10)

    def open_config(self):
        ConfigWindow(self.root)

    def run_automation(self):
        self.root.iconify() # 最小化主窗口
        self.status_label.config(text="正在运行...", fg="blue")
        self.root.update() # 立即更新状态标签

        logging.info(f"Starting automation")

        try:
            if not os.path.exists(CONFIG_FILE):
                raise Exception("配置文件不存在，请先进行配置")
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
            selected_browser = config.get('browser', 'chrome')  # 如果未设置，默认为 chrome
            self.perform_automation(selected_browser)
            self.status_label.config(text="运行完成！", fg="green")
            logging.info("Automation completed successfully.")
        except Exception as e:
            error_message = f"运行出错: {str(e)}"
            self.status_label.config(text=error_message, fg="red")
            logging.error(f"Automation failed: {e}")
            messagebox.showerror("错误", error_message)
        finally:
            self.root.deiconify() # 恢复主窗口

    def bind_shortcuts(self):
        # 绑定快捷键
        self.root.bind('<Control-Alt-R>', lambda event: self.run_automation())
        self.root.bind('<Control-Alt-S>', lambda event: self.open_config())

    def setup_driver(self, browser_name, download_path):
        logging.info(f"Setting up driver for {browser_name}")
        driver_path = get_driver_path(browser_name)

        try:
            if browser_name == "chrome":
                chrome_options = webdriver.ChromeOptions()
                prefs = {"download.default_directory": download_path,
                         "safebrowsing.enabled": False} # 关闭安全浏览警告
                chrome_options.add_experimental_option("prefs", prefs)
                chrome_options.add_argument('--no-sandbox') # 禁用沙箱模式，可能解决某些环境下的问题
                chrome_options.add_argument('--disable-gpu') # 禁用 GPU 加速，可能解决某些环境下的问题
                if driver_path and os.path.exists(driver_path):
                    service = ChromeService(driver_path)
                    driver = webdriver.Chrome(service=service, options=chrome_options)
                    logging.info("Using packaged ChromeDriver.")
                else:
                    driver_path = ChromeDriverManager().install()
                    service = ChromeService(driver_path)
                    driver = webdriver.Chrome(service=service, options=chrome_options)
                    logging.info("Downloaded ChromeDriver using webdriver-manager.")
            elif browser_name == "firefox":
                firefox_options = webdriver.FirefoxOptions()
                firefox_options.set_preference("browser.download.folderList", 2) # 使用自定义下载目录
                firefox_options.set_preference("browser.download.manager.showWhenStarting", False) # 下载时不显示下载管理器
                firefox_options.set_preference("browser.download.dir", download_path)
                firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel") # 自动保存 Excel 文件
                if driver_path and os.path.exists(driver_path):
                    service = FirefoxService(driver_path)
                    driver = webdriver.Firefox(service=service, options=firefox_options)
                    logging.info("Using packaged GeckoDriver.")
                else:
                    driver_path = GeckoDriverManager().install()
                    service = FirefoxService(driver_path)
                    driver = webdriver.Firefox(service=service, options=firefox_options)
                    logging.info("Downloaded GeckoDriver using webdriver-manager.")
            elif browser_name == "edge":
                edge_options = webdriver.EdgeOptions()
                prefs = {"download.default_directory": download_path,
                         "safebrowsing.enabled": False} # 关闭安全浏览警告
                edge_options.add_experimental_option("prefs", prefs)
                if driver_path and os.path.exists(driver_path):
                    service = EdgeService(driver_path)
                    driver = webdriver.Edge(service=service, options=edge_options)
                    logging.info("Using packaged EdgeDriver.")
                else:
                    driver_path = EdgeChromiumDriverManager().install()
                    service = EdgeService(driver_path)
                    driver = webdriver.Edge(service=service, options=edge_options)
                    logging.info("Downloaded EdgeDriver using webdriver-manager.")
            else:
                raise ValueError(f"Unsupported browser: {browser_name}")
            return driver
        except SessionNotCreatedException as e:
            logging.warning(f"Session creation failed with local driver: {e}. Trying webdriver-manager.")
            # 如果使用本地驱动创建会话失败，则尝试使用 webdriver-manager 下载驱动
            if browser_name == "chrome":
                chrome_options = webdriver.ChromeOptions()
                prefs = {"download.default_directory": download_path,
                         "safebrowsing.enabled": False}
                chrome_options.add_experimental_option("prefs", prefs)
                chrome_options.add_argument('--no-sandbox')
                chrome_options.add_argument('--disable-gpu')
                driver_path = ChromeDriverManager().install()
                service = ChromeService(driver_path)
                driver = webdriver.Chrome(service=service, options=chrome_options)
            elif browser_name == "firefox":
                firefox_options = webdriver.FirefoxOptions()
                firefox_options.set_preference("browser.download.folderList", 2)
                firefox_options.set_preference("browser.download.manager.showWhenStarting", False)
                firefox_options.set_preference("browser.download.dir", download_path)
                firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel")
                driver_path = GeckoDriverManager().install()
                service = FirefoxService(driver_path)
                driver = webdriver.Firefox(service=service, options=firefox_options)
            elif browser_name == "edge":
                edge_options = webdriver.EdgeOptions()
                prefs = {"download.default_directory": download_path,
                         "safebrowsing.enabled": False}
                edge_options.add_experimental_option("prefs", prefs)
                driver_path = EdgeChromiumDriverManager().install()
                service = EdgeService(driver_path)
                driver = webdriver.Edge(service=service, options=edge_options)
            logging.info(f"Successfully created session using webdriver-manager for {browser_name}")
            return driver
        except Exception as e:
            logging.error(f"Error setting up {browser_name} driver: {e}")
            raise

    def perform_automation(self, browser_name):
        logging.info("Starting browser automation...")
        if not os.path.exists(CONFIG_FILE):
            raise Exception("配置文件不存在，请先进行配置")

        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)

        pin_code = config.get('pin_code')
        webhook_url = config.get('webhook_url')
        download_path = config.get('download_path')

        if not pin_code or not webhook_url or not download_path:
            raise Exception("配置信息不完整")

        driver = self.setup_driver(browser_name, download_path)
        if not driver:
            logging.error(f"Failed to initialize {browser_name} driver.")
            return

        try:
            driver.get("https://ehs.crland.com.cn/home")
            wait = WebDriverWait(driver, 20) # 设置最长等待时间为 20 秒

            # 定位 PIN 码输入框并输入
            pin_input = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//input[@type="password" and @placeholder="pin码" and contains(@class, "el-input__inner")]')))
            pin_input.send_keys(pin_code)
            logging.info("PIN code entered.")
            time.sleep(1) # 等待 1 秒

            # 定位登录按钮并点击
            login_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div/div/div[2]/div[1]/div[2]/div/div[1]/div/div[2]/div[2]/form/div[3]/div/button/span')))
            login_button.click()
            logging.info("Login button clicked.")

            driver.maximize_window() # 最大化浏览器窗口
            time.sleep(5) # 等待页面加载

            # 导航到合作伙伴安全管理菜单
            partner_safety_management_menu = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div/div[1]/div/ul/li[4]/a/span')))
            partner_safety_management_menu.click()
            time.sleep(1)

            # 导航到三关一闭菜单
            three_closes_menu = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div/div[1]/div/ul/li[4]/ul/li[3]/a/span')))
            three_closes_menu.click()
            time.sleep(1)

            # 导航到三关一闭管理菜单
            three_closes_management_menu = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div/div[1]/div/ul/li[4]/ul/li[3]/ul/li[2]/a/span')))
            three_closes_management_menu.click()
            time.sleep(2)

            # 点击日期选择器
            date_picker = wait.until(EC.element_to_be_clickable(
                (By.XPATH,
                 '/html/body/div[1]/div/div[5]/div/div[2]/div/div[1]/div[1]/div[1]/div/form/div[2]/div[3]/div/div/div/div[1]/div/span')))
            date_picker.click()
            time.sleep(1)

            # 选择起始日期的昨天
            yesterday_button_start = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[contains(@class, 'ivu-date-picker-cells-cell-today')]/preceding-sibling::span[1]")))
            yesterday_button_start.click()
            logging.info("Start date selected as yesterday.")

            # 再次点击日期选择器以选择结束日期
            date_picker.click()
            time.sleep(1)

            # 选择结束日期的昨天
            yesterday_button_end = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[contains(@class, 'ivu-date-picker-cells-cell-today')]/preceding-sibling::span[1]")))
            yesterday_button_end.click()
            logging.info("End date selected as yesterday.")

            # 点击查询按钮
            query_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div/div[5]/div/div[2]/div/div[1]/div[1]/div[2]/div/div[2]/button[2]')))
            query_button.click()
            logging.info("Query button clicked.")
            time.sleep(2) # 等待查询结果加载

            # 点击导出按钮
            export_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div/div[5]/div/div[2]/div/div[1]/div[3]/div/div[2]/div/button')))
            export_button.click()
            logging.info("Export button clicked.")

            time.sleep(5)  # 等待下载完成

            # 获取最新下载的 Excel 文件
            excel_file = self.get_latest_excel_file(download_path)
            self.process_excel_and_send_message(excel_file, webhook_url)

        except Exception as e:
            logging.error(f"Automation error: {e}")
            raise
        finally:
            if driver:
                driver.quit()
                logging.info(f"{browser_name} driver closed.")

    def get_latest_excel_file(self, download_path):
        # 获取最新下载的 Excel 文件
        pattern = os.path.join(download_path, "*三关一闭记录*.xls*")
        files = glob.glob(pattern)
        if not files:
            raise Exception("未找到三关一闭记录的Excel文件")
        latest_file = max(files, key=os.path.getmtime)
        return latest_file

    def process_excel_and_send_message(self, excel_file, webhook_url):
        try:
            # 读取 Excel 文件
            df = pd.read_excel(excel_file)
            # 筛选未闭店和未申请的店铺
            df = df[df['状态'].isin(['未闭店', '未申请'])]
            # 选择需要的列
            df = df[['店铺名称', '状态']]

            if df.empty:
                message = "昨日店铺已全部闭店"
            else:
                message_lines = ["昨日以下店铺未完成闭店："]
                for index, (shop_name, status) in enumerate(zip(df['店铺名称'], df['状态']), 1):
                    message_lines.append(f"{index}、{shop_name}：{status}")
                message = "\n".join(message_lines)

            # 添加 @全体成员
            message = f"<at user_id=\"all\"></at>\n{message}"

            # 构建 Webhook Payload
            payload = {
                "msg_type": "text",
                "content": {"text": message}
            }

            # 发送 Webhook 消息
            response = requests.post(webhook_url, json=payload)
            response.raise_for_status()

            logging.info("Message sent successfully.")
            logging.info(f"Sent message: {message}")

        except Exception as e:
            logging.error(f"Error processing Excel and sending message: {e}")
            raise

    def close_application(self):
        logging.info("Closing application.")
        self.root.destroy()

    def run(self):
        # 检查配置文件是否存在，如果不存在则先打开配置窗口
        if not os.path.exists(CONFIG_FILE):
            ConfigWindow(self.root)
        self.root.mainloop()

if __name__ == "__main__":
    app = MainWindow()
    app.run()
