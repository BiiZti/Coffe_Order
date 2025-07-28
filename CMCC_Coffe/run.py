#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CMCC Coffee 订单管理系统 - 启动脚本
"""

import os
import sys
import signal
import threading
import time
from app import init_app, run_app

def signal_handler(signum, frame):
    """处理退出信号"""
    print("\n🛑 收到退出信号，正在关闭应用...")
    sys.exit(0)

def main():
    """主函数"""
    print("=" * 60)
    print("🚀 CMCC Coffee 订单管理系统启动中...")
    print("=" * 60)
    
    # 设置信号处理器
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)
    
    # 设置环境变量
    os.environ.setdefault('FLASK_ENV', 'development')
    
    try:
        # 初始化应用
        print("📊 初始化数据读取...")
        init_app()
        
        print("✅ 系统启动成功！")
        print("🌐 访问地址: http://localhost:5000")
        print("📁 Excel文件目录: data/")
        print("🔧 环境: " + os.environ.get('FLASK_ENV', 'development'))
        print("=" * 60)
        
        # 启动Flask应用
        run_app()
        
    except KeyboardInterrupt:
        print("\n👋 用户中断，系统已停止")
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == '__main__':
    main() 