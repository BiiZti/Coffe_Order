import os
import shutil
import pandas as pd
import tkinter as tk
from datetime import datetime
import zipfile
import logging
import time
import errno
import openpyxl
import win32com.client
import winsound
import threading
import json

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ==== 日志设置 ====
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
log_path = os.path.join(desktop, "coffee_log.txt")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_path, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
print = logging.info

# ==== 其他配置 ====
download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
# 创建项目数据文件夹
project_data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(project_data_dir, exist_ok=True)
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:80")
# 添加基础后台运行选项（兼容性更好）
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--disable-popup-blocking")
chrome_options.add_argument("--disable-default-apps")
chrome_options.add_argument("--no-first-run")
chrome_options.add_argument("--no-default-browser-check")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-plugins")
chrome_options.add_argument("--disable-images")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
# 尝试禁用下载栏（使用更兼容的选项）
chrome_options.add_argument("--disable-features=DownloadBubble")
chrome_options.add_argument("--disable-features=DownloadShelf")
# 添加连接池配置，解决connection pool问题
chrome_options.add_argument("--disable-background-timer-throttling")
chrome_options.add_argument("--disable-backgrounding-occluded-windows")
chrome_options.add_argument("--disable-renderer-backgrounding")
chrome_options.add_argument("--disable-features=TranslateUI")
chrome_options.add_argument("--disable-ipc-flooding-protection")
# 移除不兼容的Chrome选项
# chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
# chrome_options.add_experimental_option('useAutomationExtension', False)
# 连接重试机制
max_retries = 3
retry_count = 0
driver = None

while retry_count < max_retries and driver is None:
    try:
        print(f"🔄 尝试连接Chrome浏览器 (第{retry_count + 1}次)...")
        # 尝试使用本地ChromeDriver
        service = Service("chromedriver.exe")
        # 配置ChromeDriver连接池参数
        service.creation_flags = 0x08000000  # 禁用控制台窗口
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("✅ 使用本地ChromeDriver连接到Chrome浏览器")
        break
    except Exception as e:
        retry_count += 1
        print(f"❌ 本地ChromeDriver连接失败 (第{retry_count}次): {e}")
        if retry_count < max_retries:
            print(f"⏳ 等待3秒后重试...")
            time.sleep(3)
        else:
            print("💡 请确保已运行 start_chrome.py 启动Chrome浏览器")
            raise Exception("无法连接到Chrome浏览器，请先运行 start_chrome.py")

root = tk.Tk()
root.withdraw()
coffee_path = None

# 文件锁管理
class FileLockManager:
    def __init__(self, lock_file_path):
        self.lock_file_path = lock_file_path
        self.lock_file = None
        self.lock_acquired = False
    
    def acquire_lock(self, timeout=30):
        """获取文件锁"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # 创建锁文件
                self.lock_file = open(self.lock_file_path, 'w')
                # 写入锁信息
                lock_info = {
                    'pid': os.getpid(),
                    'timestamp': time.time(),
                    'process': 'coffee_monitor'
                }
                json.dump(lock_info, self.lock_file)
                self.lock_file.flush()
                self.lock_acquired = True
                print(f"🔒 已获取文件锁: {self.lock_file_path}")
                return True
            except Exception as e:
                print(f"⚠️ 获取锁失败，等待重试: {e}")
                time.sleep(1)
                continue
        print(f"❌ 获取文件锁超时: {self.lock_file_path}")
        return False
    
    def release_lock(self):
        """释放文件锁"""
        if self.lock_acquired and self.lock_file:
            try:
                self.lock_file.close()
                if os.path.exists(self.lock_file_path):
                    os.remove(self.lock_file_path)
                self.lock_acquired = False
                print(f"🔓 已释放文件锁: {self.lock_file_path}")
            except Exception as e:
                print(f"⚠️ 释放锁时出错: {e}")

# 创建全局锁管理器
lock_manager = FileLockManager(os.path.join(project_data_dir, "data.lock"))

# 全局最小化监控标志
minimize_monitor_active = False

def cleanup_connections():
    """清理连接池"""
    try:
        # 清理Selenium连接
        if 'driver' in globals() and driver:
            try:
                driver.quit()
            except:
                pass
        # 清理其他可能的连接
        import gc
        gc.collect()
        print("🧹 已清理连接池")
    except Exception as e:
        print(f"⚠️ 清理连接池时出错: {e}")

# 注册程序退出时的清理函数
import atexit
atexit.register(cleanup_connections)


def start_minimize_monitor():
    """启动持续最小化监控"""
    global minimize_monitor_active
    if minimize_monitor_active:
        return
    
    minimize_monitor_active = True
    print("🔒 启动持续最小化监控...")
    
    def monitor_minimize():
        global minimize_monitor_active
        while minimize_monitor_active:
            try:
                # 强制最小化窗口
                driver.minimize_window()
                # 设置窗口为后台
                driver.execute_script("window.focus = function() {};")
                driver.execute_script("window.blur();")
                # 移动窗口到屏幕外
                driver.execute_script("window.moveTo(-1000, -1000);")
                driver.execute_script("window.resizeTo(1, 1);")
            except Exception as e:
                pass
            time.sleep(1)  # 每秒检查一次
    
    # 在后台线程中运行监控
    import threading
    monitor_thread = threading.Thread(target=monitor_minimize, daemon=True)
    monitor_thread.start()


def stop_minimize_monitor():
    """停止持续最小化监控"""
    global minimize_monitor_active
    minimize_monitor_active = False
    print("🔓 停止持续最小化监控")


def is_valid_xlsx(path):
    """验证Excel文件是否有效"""
    if not os.path.exists(path):
        print(f"❌ 文件不存在: {path}")
        return False
    
    if os.path.getsize(path) == 0:
        print(f"❌ 文件为空: {path}")
        return False
    
    try:
        # 尝试用pandas读取文件
        df = pd.read_excel(path, nrows=1)  # 只读取第一行来验证
        print(f"✅ Excel文件有效: {os.path.basename(path)}")
        return True
    except Exception as e:
        print(f"❌ Excel文件无效 {os.path.basename(path)}: {e}")
        return False


def cleanup_old_data():
    """清理旧的数据文件（带文件锁保护）"""
    # 尝试获取文件锁
    if not lock_manager.acquire_lock(timeout=10):
        print("⚠️ 无法获取文件锁，跳过数据清理（可能有其他程序正在访问数据）")
        return
    
    try:
        # 清理项目数据文件夹中的旧文件
        for file in os.listdir(project_data_dir):
            file_path = os.path.join(project_data_dir, file)
            if os.path.isfile(file_path) and file.endswith('.xlsx') and not file.endswith('.lock'):
                # 跳过前端咖啡订单文件
                if '前端咖啡订单' in file:
                    print(f"🛡️ 保护前端文件: {file}")
                    continue
                
                try:
                    os.remove(file_path)
                    print(f"🗑️ 删除旧数据文件: {file}")
                except Exception as e:
                    print(f"⚠️ 删除文件失败 {file}: {e}")
        
        # 清理下载目录中的临时文件
        today_str = datetime.now().strftime("%Y%m%d")
        for file in os.listdir(download_dir):
            if file.endswith('.xlsx') and not file.startswith(today_str):
                file_path = os.path.join(download_dir, file)
                try:
                    os.remove(file_path)
                    print(f"🗑️ 删除临时文件: {file}")
                except Exception as e:
                    print(f"⚠️ 删除临时文件失败 {file}: {e}")
                    
        print("✅ 旧数据清理完成")
    except Exception as e:
        print(f"⚠️ 清理旧数据时出错: {e}")
    finally:
        # 释放文件锁
        lock_manager.release_lock()


def move_file_to_project_data(source_path, target_filename):
    """将文件移动到项目数据文件夹（带文件锁保护）"""
    # 尝试获取文件锁
    if not lock_manager.acquire_lock(timeout=10):
        print("⚠️ 无法获取文件锁，跳过文件移动（可能有其他程序正在访问数据）")
        return None
    
    try:
        target_path = os.path.join(project_data_dir, target_filename)
        shutil.move(source_path, target_path)
        print(f"📁 文件已移动到项目数据文件夹: {target_filename}")
        return target_path
    except Exception as e:
        print(f"❌ 移动文件失败: {e}")
        return None
    finally:
        # 释放文件锁
        lock_manager.release_lock()


def switch_to_target_tab():
    """切换到目标标签页或创建新标签页（云电脑兼容版）"""
    target_url_prefix = "https://zhst.cmft.com.cn/mgmt/index.html#"
    
    # 首先尝试找到现有的目标页面
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if driver.current_url.startswith(target_url_prefix):
            print(f"✅ 已切换到现有Dashboard标签页: {driver.current_url}")
            return
    
    # 如果没有找到目标页面，创建新标签页并导航到登录页面
    print("🔍 未找到目标页面，创建新标签页并导航到登录页面...")
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[-1])
    
    # 导航到登录页面
    login_url = "https://zhst.cmft.com.cn/mgmt/index.html#/login?redirect=%2Fmerchant-service%2Fmerchant-mgmt%2FMerchantInfoMgmtNew2"
    driver.get(login_url)
    print(f"✅ 已导航到登录页面: {login_url}")
    
    # 等待用户登录
    print("⏳ 等待用户登录...")
    print("💡 请在浏览器中完成登录，然后输入 '已登录' 继续程序")
    
    while True:
        user_input = input("请输入 '已登录' 继续程序: ").strip()
        if user_input == "已登录":
            print("✅ 用户确认已登录，继续程序...")
            break
        else:
            print("❌ 输入错误，请输入 '已登录'")
    
    # 用户登录确认后，启动持续最小化监控
    print("🔒 启动持续最小化监控...")
    start_minimize_monitor()
    
    # 强制最小化窗口并设置为后台运行（云电脑环境）
    try:
        # 多次尝试最小化，确保在云电脑环境中生效
        for i in range(3):
            driver.minimize_window()
            time.sleep(0.5)
        
        # 设置窗口为后台运行
        driver.execute_script("window.focus = function() {};")
        driver.execute_script("window.blur();")
        
        # 禁用窗口激活事件（云电脑环境）
        driver.execute_script("""
            // 禁用所有可能的焦点事件
            window.addEventListener('focus', function(e) {
                e.preventDefault();
                e.stopPropagation();
                window.blur();
                return false;
            }, true);
            
            window.addEventListener('activate', function(e) {
                e.preventDefault();
                e.stopPropagation();
                return false;
            }, true);
            
            // 设置窗口属性
            window.name = 'background_window';
            window.opener = null;
            
            // 尝试禁用下载栏
            if (window.chrome && window.chrome.downloads) {
                window.chrome.downloads.onChanged.addListener(function(downloadDelta) {
                    // 阻止下载栏显示
                    return false;
                });
            }
        """)
        print("✅ 已最小化后台标签页窗口并设置为后台运行")
    except Exception as e:
        print(f"⚠️ 窗口设置警告: {e}")
        print("✅ 已创建新标签页用于后台操作")


def check_login_status():
    """检查登录状态"""
    try:
        # 检查是否在登录页面
        if "login" in driver.current_url.lower():
            print("⚠️ 检测到仍在登录页面，请先完成登录")
            return False
        
        # 检查是否有登录相关的元素
        login_elements = driver.find_elements(By.XPATH, "//input[@type='password'] | //button[contains(text(),'登录')]")
        if login_elements:
            print("⚠️ 检测到登录表单，请先完成登录")
            return False
        
        print("✅ 登录状态检查通过")
        return True
    except Exception as e:
        print(f"⚠️ 登录状态检查失败: {e}")
        return False


def click_waimai_menu():
    """静默跳转到外卖订单管理页面，确保在后台运行（云电脑兼容版）"""
    current_url = driver.current_url
    target_url = "https://zhst.cmft.com.cn/mgmt/index.html#/report-form/take-out-order-mgmt/OlOrderMgmt"
    
    # 检查登录状态
    if not check_login_status():
        print("❌ 登录状态检查失败，无法继续")
        return False
    
    # 如果当前不在目标页面，则静默跳转
    if not current_url.endswith("OlOrderMgmt"):
        print("🔄 准备跳转到外卖订单管理页面...")
        
        # 强制窗口最小化（云电脑环境）
        try:
            # 多次尝试最小化，确保在云电脑环境中生效
            for i in range(3):
                driver.minimize_window()
                time.sleep(0.5)
            
            # 设置窗口为后台运行
            driver.execute_script("window.focus = function() {};")
            driver.execute_script("window.blur();")
            
            # 禁用窗口激活事件（更强制的方式）
            driver.execute_script("""
                // 禁用所有可能的焦点事件
                window.addEventListener('focus', function(e) {
                    e.preventDefault();
                    e.stopPropagation();
                    window.blur();
                    return false;
                }, true);
                
                window.addEventListener('activate', function(e) {
                    e.preventDefault();
                    e.stopPropagation();
                    return false;
                }, true);
                
                // 设置窗口属性
                window.name = 'background_window';
                window.opener = null;
                
                // 持续监控并最小化
                setInterval(function() {
                    if (window.innerHeight > 100) {
                        window.resizeTo(1, 1);
                        window.moveTo(-1000, -1000);
                    }
                }, 100);
            """)
            
            print("✅ 已设置窗口为后台运行模式")
        except Exception as e:
            print(f"⚠️ 窗口设置警告: {e}")
        
        # 使用JavaScript进行静默跳转
        try:
            driver.execute_script(f"window.location.href = '{target_url}';")
            print("✅ 已静默跳转到外卖订单管理页面")
        except Exception as e:
            print(f"⚠️ JavaScript跳转失败，使用直接跳转: {e}")
            driver.get(target_url)
        
        # 等待页面加载完成
        time.sleep(3)
        
        # 再次强制最小化（云电脑环境需要多次尝试）
        try:
            for i in range(2):
                driver.minimize_window()
                time.sleep(0.5)
            
            # 再次设置后台运行
            driver.execute_script("window.focus = function() {};")
            driver.execute_script("window.blur();")
            print("✅ 已确保窗口保持最小化状态")
        except Exception as e:
            print(f"⚠️ 最终窗口设置警告: {e}")
    else:
        print("✅ 已在外卖订单管理页面")
    
    return True


def click_query_button():
    query_btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class,'el-button') and .//span[contains(text(),'查询')]]"))
    )
    driver.execute_script("arguments[0].click();", query_btn)
    print("✅ 点击查询按钮完成，等待数据刷新...")


def click_export_button():
    export_btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class,'el-button') and .//span[contains(text(),'导出明细')]]"))
    )
    driver.execute_script("arguments[0].click();", export_btn)
    print("✅ 点击导出明细按钮完成，开始等待文件下载...")
    
    # 处理可能的下载权限弹窗
    try:
        # 等待弹窗出现
        time.sleep(2)
        
        # 尝试查找并点击"允许"按钮
        allow_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'允许') or contains(text(),'Allow')]"))
        )
        allow_button.click()
        print("✅ 已自动允许下载权限")
        time.sleep(1)
    except:
        # 如果没有弹窗，继续正常流程
        print("✅ 无需处理下载权限弹窗")


def handle_popups():
    """处理可能出现的各种弹窗"""
    try:
        # 处理下载权限弹窗
        allow_button = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'允许') or contains(text(),'Allow')]"))
        )
        allow_button.click()
        print("✅ 已处理下载权限弹窗")
        time.sleep(1)
        return True
    except:
        pass
    
    try:
        # 处理确认弹窗
        confirm_button = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'确定') or contains(text(),'确认') or contains(text(),'OK')]"))
        )
        confirm_button.click()
        print("✅ 已处理确认弹窗")
        time.sleep(1)
        return True
    except:
        pass
    
    return False


def start_download_monitor():
    """开始监控下载目录，返回监控状态"""
    print(f"📁 开始监控下载目录: {download_dir}")
    
    # 获取当前目标文件列表
    try:
        all_files = os.listdir(download_dir)
        current_files = {f for f in all_files if f.startswith("外卖订单商品明细_")}
        print(f"📋 当前目标文件: {list(current_files)}")
        return {
            'start_time': time.time(),
            'before_files': current_files,
            'download_dir': download_dir
        }
    except Exception as e:
        print(f"⚠️ 无法读取下载目录: {e}")
        return None


def wait_for_download_complete(monitor_info, timeout=60):
    """等待下载完成，基于监控状态"""
    if not monitor_info:
        print("❌ 监控信息无效，使用默认检测")
        return wait_for_new_download(timeout)
    
    print(f"⏳ 等待文件下载，超时时间: {timeout}秒")
    before_files = monitor_info['before_files']
    
    for i in range(timeout):
        # 每10秒显示一次进度
        if i % 10 == 0 and i > 0:
            print(f"⏳ 已等待下载 {i} 秒...")
        
        # 检查并处理弹窗
        handle_popups()
        
        # 检查新文件
        try:
            all_files = os.listdir(monitor_info['download_dir'])
            current_files = {f for f in all_files if f.startswith("外卖订单商品明细_")}
            new_files = current_files - before_files
            
            # 显示所有新文件（用于调试）
            if new_files:
                print(f"🔍 发现新目标文件: {list(new_files)}")
            
            for f in new_files:
                path = os.path.join(monitor_info['download_dir'], f)
                print(f"✅ 检测到新下载文件: {f}")
                print(f"📂 文件路径: {path}")
                
                # 等待文件完全下载（检查文件大小是否稳定）
                time.sleep(2)
                return path
                
        except Exception as e:
            print(f"⚠️ 检查下载目录时出错: {e}")
        
        # 等待1秒
        time.sleep(1)
    
    # 超时后显示当前下载目录的所有目标文件
    try:
        all_files = os.listdir(monitor_info['download_dir'])
        current_files = {f for f in all_files if f.startswith("外卖订单商品明细_")}
        print(f"📋 下载目录当前目标文件: {list(current_files)}")
    except Exception as e:
        print(f"⚠️ 无法读取下载目录: {e}")
    
    print(f"❌ 下载超时，{timeout}秒内未检测到新目标文件")
    return None


def wait_for_new_download(timeout=60):
    """等待新文件下载完成（兼容旧版本）"""
    return wait_for_download_complete(None, timeout)


def wait_for_frontend_processing(timeout=300):
    """等待前端处理完数据"""
    print("⏳ 等待前端处理数据...")
    
    for i in range(timeout):
        # 每10秒检查一次前端状态
        if i % 10 == 0 and i > 0:
            print(f"⏳ 已等待前端处理 {i} 秒...")
        
        # 检查前端是否还在访问数据文件
        lock_file_path = os.path.join(project_data_dir, "data.lock")
        if os.path.exists(lock_file_path):
            try:
                with open(lock_file_path, 'r') as f:
                    lock_info = json.load(f)
                    lock_time = lock_info.get('timestamp', 0)
                    # 如果锁文件超过60秒，可能是僵尸锁
                    if time.time() - lock_time > 60:
                        print("⚠️ 检测到僵尸锁，继续处理")
                        break
                    else:
                        print("🔒 前端正在处理数据，继续等待...")
                        time.sleep(1)
                        continue
            except:
                pass
        
        # 如果没有锁文件，说明前端处理完成
        if not os.path.exists(lock_file_path):
            print("✅ 前端数据处理完成")
            break
            
        time.sleep(1)
    
    print("✅ 前端处理等待完成")


def create_processing_flag():
    """创建处理标志，通知前端有新数据"""
    flag_file = os.path.join(project_data_dir, "new_data_ready.flag")
    try:
        with open(flag_file, 'w') as f:
            json.dump({
                'timestamp': time.time(),
                'message': '新数据已准备就绪'
            }, f)
        print("🚩 已创建新数据标志，通知前端")
    except Exception as e:
        print(f"⚠️ 创建标志文件失败: {e}")


def remove_processing_flag():
    """移除处理标志"""
    flag_file = os.path.join(project_data_dir, "new_data_ready.flag")
    try:
        if os.path.exists(flag_file):
            os.remove(flag_file)
            print("🚩 已移除处理标志")
    except Exception as e:
        print(f"⚠️ 移除标志文件失败: {e}")


def wait_until_unlocked(filepath, timeout=60):
    start = time.time()
    attempt = 0
    while True:
        attempt += 1
        try:
            with open(filepath, "a"):
                print(f"✅ 文件已释放: {filepath} (尝试次数: {attempt})")
                return
        except OSError as e:
            if e.errno != errno.EACCES:
                raise
        if time.time() - start > timeout:
            raise TimeoutError(f"等待文件释放超时: {filepath}")
        time.sleep(0.5)


# def process_excel(file_path):
#     df = pd.read_excel(file_path)
#     df.columns = [col.strip() for col in df.columns]
#
#     if '订单分类' not in df.columns or '订单编号' not in df.columns:
#         raise Exception("Excel中未找到‘订单分类’或‘订单编号’列，请检查表格格式")
#
#     df['订单编号'] = df['订单编号'].astype(str).str.strip()
#     df['订单分类'] = df['订单分类'].fillna("").astype(str).str.strip()
#
#     coffee_df = df[
#         (df['订单编号'].str.match(r"^[A-Za-z0-9]+$")) &
#         (df['订单分类'].str.contains("咖啡", case=False, na=False))
#     ]
#
#     print(f"📊 提取出 {len(coffee_df)} 条咖啡订单")
#     return coffee_df

def get_status_name(status_code):
    """根据状态代码获取状态名称"""
    status_map = {
        '2': '备货中',
        '5': '已完成'
    }
    return status_map.get(str(status_code), f'未知状态({status_code})')


def process_excel(file_path):
    df = pd.read_excel(file_path)
    df.columns = [col.strip() for col in df.columns]

    if '订单分类' not in df.columns or '订单编号' not in df.columns:
        raise Exception("Excel中未找到‘订单分类’或‘订单编号’列，请检查表格格式")

    df['订单编号'] = df['订单编号'].astype(str).str.strip()
    df['订单分类'] = df['订单分类'].fillna("").astype(str).str.strip()

    # 首先过滤出咖啡订单
    coffee_df = df[
        (df['订单编号'].str.match(r"^[A-Za-z0-9]+$")) &
        (df['订单分类'].str.contains("咖啡", case=False, na=False))
    ]

    print(f"🔍 原始数据: {len(df)} 条")
    print(f"🔍 过滤后咖啡订单: {len(coffee_df)} 条")
    
    # 显示订单分类统计
    if '订单分类' in df.columns:
        category_counts = df['订单分类'].value_counts()
        print("📊 订单分类统计:")
        for category, count in category_counts.head(10).items():
            print(f"   {category}: {count} 条")

    # 只保留你关心的列
    required_columns = [
        "订单编号", "手机号码", "姓名", "部门", "支付时间", "订单分类", "订单备注"
    ]
    
    # 过滤出存在的列，但保持咖啡订单的过滤结果
    available_columns = [col for col in required_columns if col in coffee_df.columns]
    coffee_df = coffee_df[available_columns]
    
    # 确保所有必需的列都存在
    for col in ["订单编号", "手机号码", "姓名", "部门", "支付时间", "订单分类"]:
        if col not in coffee_df.columns:
            coffee_df[col] = ""
    
    # 如果订单备注列不存在，添加空列
    if "订单备注" not in coffee_df.columns:
        coffee_df["订单备注"] = ""

    print(f"📊 提取出 {len(coffee_df)} 条咖啡订单")
    return coffee_df


def update_frontend_excel(new_coffee_df):
    """更新前端专用Excel文件，通过订单号同步"""
    today_str = datetime.now().strftime("%Y%m%d")
    frontend_excel_path = os.path.join(project_data_dir, f"{today_str}_前端咖啡订单.xlsx")
    
    print(f"🔄 开始更新前端Excel文件: {frontend_excel_path}")
    print(f"📊 新数据包含 {len(new_coffee_df)} 条记录")
    
    try:
        # 检查新数据是否为空
        if new_coffee_df.empty:
            print("⚠️ 新数据为空，跳过前端Excel更新")
            return False
        
        # 检查订单编号列是否存在
        if '订单编号' not in new_coffee_df.columns:
            print("❌ 新数据中未找到'订单编号'列")
            print(f"   可用列: {list(new_coffee_df.columns)}")
            return False
        
        # 获取文件锁
        if not lock_manager.acquire_lock(timeout=15):
            print("⚠️ 无法获取文件锁，跳过前端Excel更新")
            return False
        
        try:
            # 读取现有的前端Excel文件
            existing_df = None
            if os.path.exists(frontend_excel_path) and is_valid_xlsx(frontend_excel_path):
                try:
                    existing_df = pd.read_excel(frontend_excel_path)
                    print(f"📄 读取现有前端Excel文件，包含 {len(existing_df)} 条记录")
                except Exception as e:
                    print(f"⚠️ 读取现有前端Excel失败: {e}")
                    existing_df = None
            
            # 创建新的DataFrame用于前端
            frontend_data = []
            
            # 处理新数据
            for index, new_row in new_coffee_df.iterrows():
                try:
                    order_id = str(new_row.get('订单编号', '')).strip()
                    
                    # 检查订单编号是否有效
                    if not order_id or order_id == 'nan':
                        print(f"⚠️ 跳过无效订单编号 (行 {index + 1}): {order_id}")
                        continue
                    
                    # 检查是否已存在
                    existing_row = None
                    if existing_df is not None and '订单编号' in existing_df.columns:
                        existing_matches = existing_df[existing_df['订单编号'].astype(str).str.strip() == order_id]
                        if not existing_matches.empty:
                            existing_row = existing_matches.iloc[0]
                    
                    if existing_row is not None:
                        # 使用现有状态，更新其他信息
                        existing_status = str(existing_row.get('状态', '2'))
                        status_name = get_status_name(existing_status)
                        frontend_row = {
                            '订单编号': order_id,
                            '手机号码': str(new_row.get('手机号码', '')).strip(),
                            '姓名': str(new_row.get('姓名', '')).strip(),
                            '部门': str(new_row.get('部门', '')).strip(),
                            '支付时间': str(new_row.get('支付时间', '')).strip(),
                            '订单分类': str(new_row.get('订单分类', '')).strip(),
                            '订单备注': str(new_row.get('订单备注', '')).strip(),
                            '状态': existing_status,  # 保持现有状态
                            '状态名称': status_name,  # 添加状态名称
                            '处理时间': str(existing_row.get('处理时间', '')).strip()
                        }
                        print(f"🔄 更新订单: {order_id} (保持状态: {status_name})")
                    else:
                        # 新订单
                        frontend_row = {
                            '订单编号': order_id,
                            '手机号码': str(new_row.get('手机号码', '')).strip(),
                            '姓名': str(new_row.get('姓名', '')).strip(),
                            '部门': str(new_row.get('部门', '')).strip(),
                            '支付时间': str(new_row.get('支付时间', '')).strip(),
                            '订单分类': str(new_row.get('订单分类', '')).strip(),
                            '订单备注': str(new_row.get('订单备注', '')).strip(),
                            '状态': '2',  # 默认备货中状态
                            '状态名称': '备货中',  # 添加状态名称
                            '处理时间': ''
                        }
                        print(f"🆕 新增订单: {order_id} (状态: 备货中)")
                    
                    frontend_data.append(frontend_row)
                    
                except Exception as row_error:
                    print(f"❌ 处理第 {index + 1} 行数据时出错: {row_error}")
                    continue
            
            # 检查是否有有效数据
            if not frontend_data:
                print("⚠️ 没有有效数据，跳过保存")
                return False
            
            # 保存到前端Excel文件
            frontend_df = pd.DataFrame(frontend_data)
            
            # 确保目录存在
            os.makedirs(os.path.dirname(frontend_excel_path), exist_ok=True)
            
            # 保存文件
            frontend_df.to_excel(frontend_excel_path, index=False, engine='openpyxl')
            print(f"✅ 前端Excel文件已更新: {os.path.basename(frontend_excel_path)}")
            print(f"📊 总订单数: {len(frontend_df)}")
            
            # 验证文件是否成功保存
            if os.path.exists(frontend_excel_path):
                file_size = os.path.getsize(frontend_excel_path)
                print(f"📁 文件大小: {file_size} 字节")
                return True
            else:
                print("❌ 文件保存失败，文件不存在")
                return False
            
        finally:
            # 释放文件锁
            lock_manager.release_lock()
            
    except Exception as e:
        print(f"❌ 更新前端Excel失败: {e}")
        import traceback
        traceback.print_exc()
        return False



def open_excel(path):
    """打开Excel文件（已禁用，避免干扰前端用户体验）"""
    # 注释掉自动打开功能，避免干扰前端用户
    # print(f"📄 打开Excel：{path}")
    # os.startfile(path)
    pass


def save_and_close_excel():
    print(f"🔒 尝试保存并关闭所有打开的Excel文件")
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        print(f"⚠️ 没有运行中的Excel实例，跳过保存关闭")
        return
    try:
        for wb in excel.Workbooks:
            print(f"💾 保存工作簿: {wb.Name}")
            wb.Save()
            wb.Close(False)
        excel.Quit()
        time.sleep(3)
        print("✅ 所有Excel文件已保存并关闭")
    except Exception as e:
        logging.exception("❌ 保存关闭Excel出错")


def play_no_new_sound():
    winsound.MessageBeep(0x00000040)


def play_new_sound():
    winsound.MessageBeep(0x00000030)


def show_message(title, message, timeout=5000):
    win = tk.Toplevel()
    win.title(title)
    win.geometry("350x120+600+300")
    win.attributes("-topmost", True)
    win.resizable(False, False)
    label = tk.Label(win, text=message, font=("Arial", 12), wraplength=320)
    label.pack(padx=20, pady=20)
    win.after(timeout, win.destroy)


def do_check():
    global coffee_path
    today_str = datetime.now().strftime("%Y%m%d")
    
    # 清理旧数据
    cleanup_old_data()
    
    # 设置文件路径
    full_order_filename = f"{today_str}_所有外卖订单.xlsx"
    coffee_filename = f"{today_str}_咖啡订单.xlsx"
    full_order_path = os.path.join(project_data_dir, full_order_filename)
    coffee_path = os.path.join(project_data_dir, coffee_filename)

    # 确保启动持续最小化监控（无论是否已登录）
    if not minimize_monitor_active:
        print("🔒 启动持续最小化监控...")
        start_minimize_monitor()

    # 在点击导出按钮之前开始监控下载目录
    print("🔍 开始监控下载目录...")
    download_monitor = start_download_monitor()

    try:
        switch_to_target_tab()
        if not click_waimai_menu():
            print("❌ 登录状态检查失败，等待用户登录...")
            root.after(60000, do_check)  # 1分钟后重试
            return
        click_query_button()
        click_export_button()
    except Exception as e:
        logging.exception("❌ 点击页面元素失败")
        # 如果是connection pool错误，尝试清理连接
        if "connection pool is full" in str(e).lower():
            print("🔄 检测到连接池问题，尝试清理连接...")
            cleanup_connections()
            time.sleep(2)
        root.after(30000, do_check)
        return

    # 等待下载完成
    file_path = wait_for_download_complete(download_monitor)
    if not file_path or not os.path.exists(file_path):
        print("❌ 未找到最新下载文件，重试中...")
        root.after(30000, do_check)
        return

    # 移动文件到项目数据文件夹
    moved_path = move_file_to_project_data(file_path, full_order_filename)
    if not moved_path:
        print("❌ 文件移动失败，重试中...")
        root.after(30000, do_check)
        return
    
    print(f"✅ 最新外卖总表已保存到项目数据文件夹")

    try:
        coffee_df = process_excel(full_order_path)
    except Exception:
        logging.exception("❌ 处理Excel出错")
        root.after(30000, do_check)
        return

    # 更新前端专用Excel文件
    print("🔄 开始更新前端专用Excel文件...")
    if update_frontend_excel(coffee_df):
        print("✅ 前端Excel文件更新成功")
        play_new_sound()  # 播放提示音
    else:
        print("⚠️ 前端Excel文件更新失败")
        play_no_new_sound()

    # 等待前端处理完数据后再进行下一次检查
    print("⏳ 等待前端处理数据...")
    wait_for_frontend_processing(timeout=30)  # 最多等待30秒
    
    # 移除处理标志
    remove_processing_flag()
    
    # 缩短下次检查间隔，提高响应速度
    print("🔄 准备进行下一次数据检查...")
    root.after(30000, do_check)  # 改为30秒间隔


def start_program():
    global coffee_path
    today_str = datetime.now().strftime("%Y%m%d")
    coffee_path = os.path.join(project_data_dir, f"{today_str}_咖啡订单.xlsx")
    
    print("🚀 咖啡订单监控程序启动中...")
    print("=" * 50)
    print("📋 程序功能：")
    print("   • 自动监控外卖系统中的咖啡订单")
    print("   • 每30秒检测一次新订单（快速响应）")
    print("   • 自动筛选咖啡类订单")
    print("   • 声音提示（无弹窗干扰）")
    print("   • 自动管理Excel文件")
    print("   • 静默运行，不干扰前端用户")
    print("   • 持续最小化监控，防止窗口弹出")
    print("=" * 50)
    
    if is_valid_xlsx(coffee_path):
        print(f"📄 发现现有咖啡订单表: {os.path.basename(coffee_path)}")
    else:
        print(f"📄 将创建新的咖啡订单表: {os.path.basename(coffee_path)}")

    print("🔍 正在连接浏览器并检查登录状态...")
    
    do_check()


if __name__ == "__main__":
    try:
        root.after(0, start_program)
        root.mainloop()
    except KeyboardInterrupt:
        print("\n🛑 程序被用户中断")
    except Exception as e:
        logging.exception("❌ 程序运行出错")
    finally:
        # 停止最小化监控
        stop_minimize_monitor()
        print("👋 程序结束")
