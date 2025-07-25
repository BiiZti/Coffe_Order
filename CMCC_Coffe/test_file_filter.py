#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试文件过滤功能
"""

import os
import glob

def test_file_filter():
    """测试文件过滤功能"""
    print("🔍 测试文件过滤功能")
    print("=" * 50)
    
    # 桌面路径
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    print(f"桌面路径: {desktop_path}")
    
    # 测试不同的文件匹配模式
    patterns = [
        "*.xlsx",  # 所有Excel文件
        "*咖啡订单*.xlsx",  # 只包含"咖啡订单"的文件
        "*所有外卖订单*.xlsx"  # 只包含"所有外卖订单"的文件
    ]
    
    for pattern in patterns:
        print(f"\n📁 匹配模式: {pattern}")
        files = glob.glob(os.path.join(desktop_path, pattern))
        
        if files:
            print(f"找到 {len(files)} 个文件:")
            for file in files:
                filename = os.path.basename(file)
                print(f"  ✅ {filename}")
        else:
            print("❌ 未找到匹配的文件")
    
    print("\n" + "=" * 50)
    print("测试完成")

if __name__ == '__main__':
    test_file_filter() 