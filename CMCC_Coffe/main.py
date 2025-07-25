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

def open_browser():
    """延迟打开浏览器"""
    time.sleep(2)  # 等待Flask启动
    try:
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
    print("📁 Excel文件目录: orders/")
    print("=" * 60)
    
    # 启动自动打开浏览器的线程
    browser_thread = threading.Thread(target=open_browser, daemon=True)
    browser_thread.start()
    
    # 启动Flask应用
    run_app()

if __name__ == '__main__':
    main() 