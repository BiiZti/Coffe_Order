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
try:
    # 尝试使用本地ChromeDriver
    service = Service("chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=chrome_options)
    print("✅ 使用本地ChromeDriver连接到Chrome浏览器")
except Exception as e:
    print(f"❌ 本地ChromeDriver连接失败: {e}")
    print("💡 请确保已运行 start_chrome.py 启动Chrome浏览器")
    raise Exception("无法连接到Chrome浏览器，请先运行 start_chrome.py")

root = tk.Tk()
root.withdraw()
coffee_path = None


def is_valid_xlsx(path):
    if not os.path.exists(path):
        return False
    try:
        with zipfile.ZipFile(path, 'r') as z:
            if z.testzip() is not None:
                return False
    except zipfile.BadZipFile:
        return False
    return True


def cleanup_old_data():
    """清理旧的数据文件"""
    try:
        # 清理项目数据文件夹中的旧文件
        for file in os.listdir(project_data_dir):
            file_path = os.path.join(project_data_dir, file)
            if os.path.isfile(file_path) and file.endswith('.xlsx'):
                os.remove(file_path)
                print(f"🗑️ 删除旧数据文件: {file}")
        
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


def move_file_to_project_data(source_path, target_filename):
    """将文件移动到项目数据文件夹"""
    try:
        target_path = os.path.join(project_data_dir, target_filename)
        shutil.move(source_path, target_path)
        print(f"📁 文件已移动到项目数据文件夹: {target_filename}")
        return target_path
    except Exception as e:
        print(f"❌ 移动文件失败: {e}")
        return None


def switch_to_target_tab():
    target_url_prefix = "https://zhst.cmft.com.cn/mgmt/index.html#"
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if driver.current_url.startswith(target_url_prefix):
            print(f"✅ 已切换到Dashboard标签页: {driver.current_url}")
            return
    raise Exception("❌ 未找到目标页面标签页，请先打开并登录")


def click_waimai_menu():
    driver.get("https://zhst.cmft.com.cn/mgmt/index.html#/report-form/take-out-order-mgmt/OlOrderMgmt")
    print("✅ 已直接跳转到外卖订单管理页面")


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


def wait_for_new_download(timeout=30):
    before = set(os.listdir(download_dir))
    for _ in range(timeout):
        root.update()
        after = set(os.listdir(download_dir))
        new_files = after - before
        for f in new_files:
            if f.endswith(".xlsx"):
                path = os.path.join(download_dir, f)
                root.after(1000)
                return path
        root.after(1000)
    return None


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

def process_excel(file_path):
    df = pd.read_excel(file_path)
    df.columns = [col.strip() for col in df.columns]

    if '订单分类' not in df.columns or '订单编号' not in df.columns:
        raise Exception("Excel中未找到‘订单分类’或‘订单编号’列，请检查表格格式")

    df['订单编号'] = df['订单编号'].astype(str).str.strip()
    df['订单分类'] = df['订单分类'].fillna("").astype(str).str.strip()

    coffee_df = df[
        (df['订单编号'].str.match(r"^[A-Za-z0-9]+$")) &
        (df['订单分类'].str.contains("咖啡", case=False, na=False))
    ]

    # 只保留你关心的列
    required_columns = [
        "订单编号", "手机号码", "姓名", "部门", "支付时间", "订单分类", "Unnamed: 39", "订单备注"
    ]
    coffee_df = coffee_df[[col for col in required_columns if col in coffee_df.columns]]

    print(f"📊 提取出 {len(coffee_df)} 条咖啡订单")
    return coffee_df



def open_excel(path):
    print(f"📄 打开Excel：{path}")
    os.startfile(path)


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

    try:
        switch_to_target_tab()
        click_waimai_menu()
        click_query_button()
        click_export_button()
    except Exception as e:
        logging.exception("❌ 点击页面元素失败")
        root.after(30000, do_check)
        return

    # 等待下载完成
    file_path = wait_for_new_download()
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

    existing_order_ids = set()
    if is_valid_xlsx(coffee_path):
        try:
            existing_df = pd.read_excel(coffee_path)
            if '订单编号' in existing_df.columns:
                existing_order_ids = set(existing_df['订单编号'].astype(str))
        except Exception:
            print(f"⚠️ 无法读取现有咖啡表，将重新创建")
    else:
        print(f"⚠️ 现有咖啡表无效或不存在，将重新创建")

    new_rows = coffee_df[~coffee_df['订单编号'].isin(existing_order_ids)]

    if new_rows.empty:
        print("✅ 当前没有新增咖啡")
        play_no_new_sound()
    else:
        print(f"🚨 检测到新增咖啡：{len(new_rows)} 条")
        play_new_sound()
        # 移除弹窗提示，只保留声音提示
        try:
            save_and_close_excel()
            wait_until_unlocked(coffee_path, timeout=60)

            if is_valid_xlsx(coffee_path):
                wb = openpyxl.load_workbook(coffee_path)
                ws = wb.active
                ws.append([])
                for _, row in new_rows.iterrows():
                    ws.append(row.tolist())
                wb.save(coffee_path)
                wb.close()
            else:
                coffee_df.to_excel(coffee_path, index=False)

            # 移除自动打开Excel，避免干扰用户工作
            print(f"✅ 咖啡订单数据已保存到: {os.path.basename(coffee_path)}")
        except Exception:
            logging.exception("❌ 写入咖啡Excel时出错")

    root.after(30000, do_check)


def start_program():
    global coffee_path
    today_str = datetime.now().strftime("%Y%m%d")
    coffee_path = os.path.join(project_data_dir, f"{today_str}_咖啡订单.xlsx")
    
    print("🚀 咖啡订单监控程序启动中...")
    print("=" * 50)
    print("📋 程序功能：")
    print("   • 自动监控外卖系统中的咖啡订单")
    print("   • 每30秒检测一次新订单")
    print("   • 自动筛选咖啡类订单")
    print("   • 声音和弹窗通知")
    print("   • 自动管理Excel文件")
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
    except Exception:
        logging.exception("❌ 主程序运行失败")
