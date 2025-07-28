#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Connection Pool 问题修复脚本
用于解决ChromeDriver和Selenium的connection pool问题
"""

import os
import sys
import time
import subprocess
import psutil
import signal

def kill_chrome_processes():
    """强制结束所有Chrome相关进程"""
    print("🔍 正在查找Chrome相关进程...")
    
    chrome_processes = []
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            proc_info = proc.info
            proc_name = proc_info['name'].lower() if proc_info['name'] else ''
            cmdline = ' '.join(proc_info['cmdline']) if proc_info['cmdline'] else ''
            
            # 查找Chrome相关进程
            if any(keyword in proc_name for keyword in ['chrome', 'chromedriver']):
                chrome_processes.append(proc)
                print(f"   找到Chrome进程: PID={proc_info['pid']}, 名称={proc_info['name']}")
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    
    if not chrome_processes:
        print("✅ 没有找到Chrome相关进程")
        return
    
    print(f"🔄 正在结束 {len(chrome_processes)} 个Chrome进程...")
    
    for proc in chrome_processes:
        try:
            proc.terminate()
            print(f"   ✅ 已终止进程 PID={proc.pid}")
        except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
            print(f"   ❌ 无法终止进程 PID={proc.pid}: {e}")
    
    # 等待进程结束
    time.sleep(3)
    
    # 强制结束仍在运行的进程
    for proc in chrome_processes:
        try:
            if proc.is_running():
                proc.kill()
                print(f"   🔥 强制结束进程 PID={proc.pid}")
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    
    print("✅ Chrome进程清理完成")

def clear_chrome_data():
    """清理Chrome缓存和数据"""
    print("🧹 正在清理Chrome缓存和数据...")
    
    # Chrome用户数据目录
    chrome_data_dirs = [
        os.path.expanduser("~\\AppData\\Local\\Google\\Chrome\\User Data"),
        os.path.expanduser("~\\AppData\\Roaming\\Google\\Chrome\\User Data")
    ]
    
    for data_dir in chrome_data_dirs:
        if os.path.exists(data_dir):
            try:
                # 清理缓存目录
                cache_dir = os.path.join(data_dir, "Default", "Cache")
                if os.path.exists(cache_dir):
                    import shutil
                    shutil.rmtree(cache_dir)
                    print(f"   ✅ 已清理缓存: {cache_dir}")
                
                # 清理会话存储
                session_dir = os.path.join(data_dir, "Default", "Session Storage")
                if os.path.exists(session_dir):
                    import shutil
                    shutil.rmtree(session_dir)
                    print(f"   ✅ 已清理会话存储: {session_dir}")
                    
            except Exception as e:
                print(f"   ⚠️ 清理失败 {data_dir}: {e}")
    
    print("✅ Chrome数据清理完成")

def restart_chrome():
    """重启Chrome浏览器"""
    print("🔄 正在重启Chrome浏览器...")
    
    try:
        # 启动Chrome
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Users\43887\AppData\Local\Google\Chrome\Application\chrome.exe"
        ]
        
        chrome_started = False
        for chrome_path in chrome_paths:
            if os.path.exists(chrome_path):
                try:
                    subprocess.Popen([chrome_path, "--remote-debugging-port=80"])
                    print(f"✅ Chrome已启动: {chrome_path}")
                    chrome_started = True
                    break
                except Exception as e:
                    print(f"❌ 启动Chrome失败 {chrome_path}: {e}")
                    continue
        
        if not chrome_started:
            print("❌ 无法启动Chrome浏览器")
            return False
            
        # 等待Chrome启动
        time.sleep(5)
        print("✅ Chrome重启完成")
        return True
        
    except Exception as e:
        print(f"❌ 重启Chrome失败: {e}")
        return False

def check_connection_pool():
    """检查connection pool状态"""
    print("🔍 检查connection pool状态...")
    
    try:
        # 检查端口80是否被占用
        import socket
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        result = sock.connect_ex(('127.0.0.1', 80))
        sock.close()
        
        if result == 0:
            print("✅ Chrome调试端口80正常")
            return True
        else:
            print("❌ Chrome调试端口80未响应")
            return False
            
    except Exception as e:
        print(f"❌ 检查connection pool失败: {e}")
        return False

def main():
    """主函数"""
    print("=" * 60)
    print("🔧 Connection Pool 问题修复工具")
    print("=" * 60)
    
    print("📋 修复步骤:")
    print("1. 结束所有Chrome相关进程")
    print("2. 清理Chrome缓存和数据")
    print("3. 重启Chrome浏览器")
    print("4. 检查connection pool状态")
    print("=" * 60)
    
    # 步骤1: 结束Chrome进程
    print("\n1️⃣ 结束Chrome进程...")
    kill_chrome_processes()
    
    # 步骤2: 清理Chrome数据
    print("\n2️⃣ 清理Chrome数据...")
    clear_chrome_data()
    
    # 步骤3: 重启Chrome
    print("\n3️⃣ 重启Chrome浏览器...")
    if not restart_chrome():
        print("❌ Chrome重启失败，请手动启动Chrome")
        return
    
    # 步骤4: 检查状态
    print("\n4️⃣ 检查connection pool状态...")
    if check_connection_pool():
        print("\n✅ 修复完成！现在可以运行咖啡订单监控程序了")
        print("💡 建议:")
        print("   - 运行 start_chrome.py 启动Chrome")
        print("   - 然后运行 咖啡订单监控.py")
    else:
        print("\n❌ 修复失败，请手动检查Chrome状态")
        print("💡 建议:")
        print("   - 手动启动Chrome浏览器")
        print("   - 确保Chrome正常运行")
        print("   - 重新运行此修复脚本")

if __name__ == "__main__":
    main() 