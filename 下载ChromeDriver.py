#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ChromeDriver下载脚本
用于下载与本地Chrome版本匹配的ChromeDriver
"""

import os
import sys
import subprocess
import requests
import zipfile
import re
from pathlib import Path

def get_chrome_version():
    """获取本地Chrome版本"""
    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Users\43887\AppData\Local\Google\Chrome\Application\chrome.exe",
        r"C:\Users\43887\AppData\Local\Programs\Google\Chrome\Application\chrome.exe"
    ]
    
    print("🔍 正在查找Chrome浏览器...")
    for path in chrome_paths:
        print(f"   检查路径: {path}")
        if os.path.exists(path):
            print(f"   ✅ 找到Chrome: {path}")
        else:
            print(f"   ❌ 路径不存在: {path}")
    
    for path in chrome_paths:
        if os.path.exists(path):
            try:
                # 尝试从Chrome的安装目录读取版本信息
                version_file = os.path.join(os.path.dirname(path), "chrome.exe")
                if os.path.exists(version_file):
                    # 使用PowerShell获取文件版本信息
                    ps_command = f'[System.Diagnostics.FileVersionInfo]::GetVersionInfo("{version_file}").FileVersion'
                    result = subprocess.run(['powershell', '-Command', ps_command], 
                                          capture_output=True, text=True, timeout=10)
                    
                    if result.returncode == 0 and result.stdout.strip():
                        version = result.stdout.strip()
                        print(f"✅ 找到Chrome版本: {version}")
                        # 提取主版本号
                        major_version = version.split('.')[0]
                        return major_version
                    else:
                        print(f"⚠️ 无法获取版本信息: {result.stderr}")
                
                # 备用方法：直接使用 --version 参数
                result = subprocess.run([path, "--version"], 
                                      capture_output=True, text=True, timeout=10)
                if result.returncode == 0 and result.stdout:
                    version_match = re.search(r'Chrome\s+(\d+)\.(\d+)\.(\d+)', result.stdout)
                    if version_match:
                        major_version = version_match.group(1)
                        print(f"✅ 找到Chrome版本: {version_match.group(0)}")
                        return major_version
            except Exception as e:
                print(f"⚠️ 检查Chrome版本失败 {path}: {e}")
                continue
    
    print("❌ 无法找到Chrome浏览器或获取版本信息")
    return None

def download_chromedriver(version):
    """下载ChromeDriver"""
    try:
        # 获取ChromeDriver版本信息
        print(f"🔍 正在获取ChromeDriver {version} 版本信息...")
        
        # 这里使用一个简单的下载方式，实际使用时可能需要根据网络情况调整
        download_url = f"https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{version}"
        
        try:
            response = requests.get(download_url, timeout=10)
            if response.status_code == 200:
                driver_version = response.text.strip()
                print(f"✅ 找到ChromeDriver版本: {driver_version}")
            else:
                print(f"❌ 无法获取ChromeDriver版本信息，状态码: {response.status_code}")
                return False
        except Exception as e:
            print(f"❌ 网络连接失败: {e}")
            print("💡 请检查网络连接或使用VPN")
            return False
        
        # 下载ChromeDriver
        download_url = f"https://chromedriver.storage.googleapis.com/{driver_version}/chromedriver_win32.zip"
        print(f"📥 正在下载ChromeDriver: {download_url}")
        
        response = requests.get(download_url, timeout=30)
        if response.status_code == 200:
            # 保存到当前目录
            zip_path = "chromedriver_win32.zip"
            with open(zip_path, 'wb') as f:
                f.write(response.content)
            
            # 解压文件
            print("📦 正在解压ChromeDriver...")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(".")
            
            # 删除zip文件
            os.remove(zip_path)
            
            print("✅ ChromeDriver下载完成！")
            print(f"📁 文件位置: {os.path.abspath('chromedriver.exe')}")
            return True
        else:
            print(f"❌ 下载失败，状态码: {response.status_code}")
            return False
            
    except Exception as e:
        print(f"❌ 下载过程中出错: {e}")
        return False

def main():
    print("🚀 ChromeDriver下载工具")
    print("=" * 50)
    
    # 获取Chrome版本
    chrome_version = get_chrome_version()
    if not chrome_version:
        print("💡 请确保已安装Chrome浏览器")
        return
    
    # 下载ChromeDriver
    if download_chromedriver(chrome_version):
        print("\n🎉 ChromeDriver安装成功！")
        print("💡 现在可以运行面点订单监控程序了")
    else:
        print("\n❌ ChromeDriver下载失败")
        print("💡 请手动下载ChromeDriver并放置到项目目录中")

if __name__ == "__main__":
    main() 