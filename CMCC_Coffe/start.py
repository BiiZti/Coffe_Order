#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CMCC Coffee 订单管理系统启动脚本
"""

import subprocess
import sys
import os

def main():
    """启动主程序"""
    print("=" * 50)
    print("🚀 CMCC Coffee 订单管理系统")
    print("=" * 50)
    
    # 检查Python版本
    if sys.version_info < (3, 7):
        print("❌ 错误: 需要Python 3.7或更高版本")
        sys.exit(1)
    
    # 检查依赖
    try:
        import flask
        import pandas
        import openpyxl
        print("✅ 依赖检查通过")
    except ImportError as e:
        print(f"❌ 缺少依赖: {e}")
        print("请运行: pip install -r requirements.txt")
        sys.exit(1)
    
    # 检查项目data文件夹中的咖啡订单Excel文件
    project_root = os.path.dirname(os.path.abspath(__file__))
    data_path = os.path.join(project_root, "..", "data")
    
    if not os.path.exists(data_path):
        print(f"❌ 无法访问data文件夹: {data_path}")
        print("请确保已运行咖啡订单监控程序生成数据文件")
    else:
        print(f"✅ data文件夹路径正常: {data_path}")
    
    coffee_excel_files = []
    if os.path.exists(data_path):
        coffee_excel_files = [f for f in os.listdir(data_path) if '咖啡订单' in f and f.endswith('.xlsx')]
    
    if not coffee_excel_files:
        print("⚠️  警告: data文件夹中没有咖啡订单Excel文件")
        print("请先运行咖啡订单监控程序生成数据文件")
    else:
        print(f"📊 data文件夹找到 {len(coffee_excel_files)} 个咖啡订单Excel文件")
        for file in coffee_excel_files:
            print(f"   - {file}")
    
    print("\n🎯 启动系统...")
    print("=" * 50)
    
    # 启动Flask应用
    try:
        # 获取当前脚本所在目录
        script_dir = os.path.dirname(os.path.abspath(__file__))
        main_path = os.path.join(script_dir, 'main.py')
        subprocess.run([sys.executable, main_path], check=True)
    except KeyboardInterrupt:
        print("\n👋 系统已停止")
    except Exception as e:
        print(f"❌ 启动失败: {e}")

if __name__ == '__main__':
    main() 