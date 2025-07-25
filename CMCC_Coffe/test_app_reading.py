#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试app.py中的文件读取逻辑
"""

import os
import glob
import pandas as pd
from datetime import datetime

def test_app_reading_logic():
    """测试app.py中的文件读取逻辑"""
    print("🔍 测试app.py中的文件读取逻辑")
    print("=" * 60)
    
    # 模拟app.py中的配置
    EXCEL_FOLDER = os.path.join(os.path.expanduser("~"), "Desktop")
    EXCEL_PATTERN = "*咖啡订单*.xlsx"
    
    print(f"桌面路径: {EXCEL_FOLDER}")
    print(f"文件匹配模式: {EXCEL_PATTERN}")
    
    # 检查桌面路径
    if not os.path.exists(EXCEL_FOLDER):
        print("❌ 桌面路径不存在")
        return
    
    print("✅ 桌面路径正常")
    
    # 查找咖啡订单Excel文件
    excel_files = glob.glob(os.path.join(EXCEL_FOLDER, EXCEL_PATTERN))
    
    if not excel_files:
        print("❌ 未找到咖啡订单Excel文件")
        return
    
    print(f"✅ 找到 {len(excel_files)} 个咖啡订单Excel文件:")
    for file in excel_files:
        filename = os.path.basename(file)
        print(f"  📄 {filename}")
    
    # 选择最新的文件
    latest_file = max(excel_files, key=os.path.getctime)
    print(f"\n📊 选择最新文件: {os.path.basename(latest_file)}")
    
    # 尝试读取Excel文件
    try:
        print(f"\n🔍 正在读取Excel文件...")
        df = pd.read_excel(latest_file, engine='openpyxl')
        print(f"✅ 成功读取Excel文件")
        print(f"📊 数据行数: {len(df)}")
        print(f"📋 列名: {list(df.columns)}")
        
        # 显示前几行数据
        if len(df) > 0:
            print(f"\n📝 前3行数据预览:")
            print(df.head(3).to_string())
        
    except Exception as e:
        print(f"❌ 读取Excel文件失败: {e}")
    
    print("\n" + "=" * 60)
    print("测试完成")

if __name__ == '__main__':
    test_app_reading_logic() 