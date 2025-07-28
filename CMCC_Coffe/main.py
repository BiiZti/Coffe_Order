#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CMCC Coffee 订单管理系统 - 主入口文件
负责系统启动和初始化
"""

import webbrowser
import threading
import time
from app import init_app, run_app

# 全局标志，防止重复打开浏览器
_browser_opened = False
_browser_lock = threading.Lock()

def open_browser():
    """延迟打开浏览器"""
    global _browser_opened
    
    # 使用锁确保线程安全
    with _browser_lock:
        if _browser_opened:
            print("⚠️  浏览器已经打开，跳过重复操作")
            return
        _browser_opened = True
    
    time.sleep(2)  # 等待Flask启动
    try:
        # 检查是否已经有浏览器窗口打开
        import psutil
        browser_processes = []
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if any(browser in proc.info['name'].lower() for browser in ['chrome', 'firefox', 'edge', 'iexplore']):
                    browser_processes.append(proc.info['name'])
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        if browser_processes:
            print(f"🔍 检测到已运行的浏览器: {', '.join(browser_processes)}")
            print("🌐 尝试在新标签页中打开应用...")
        
        webbrowser.open('http://localhost:5000')
        print("🌐 已自动打开浏览器")
    except Exception as e:
        print(f"⚠️  自动打开浏览器失败: {e}")
        print("请手动访问: http://localhost:5000")

def main():
    """主函数"""
    print("=" * 60)
    print("🚀 CMCC Coffee 订单管理系统启动中...")
    print("=" * 60)
    
    # 初始化应用
    init_app()
    
    print("✅ 系统启动成功！")
    print("🌐 访问地址: http://localhost:5000")
    print("📁 Excel文件目录: data/")
    print("=" * 60)
    
    # 自动打开浏览器
    print("✅ 将在2秒后自动打开浏览器")
    browser_thread = threading.Thread(target=open_browser, daemon=True)
    browser_thread.start()
    
    # 启动Flask应用
    run_app()

if __name__ == '__main__':
    main() 