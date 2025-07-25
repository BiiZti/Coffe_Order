# 咖啡订单监控系统

## 📋 项目简介
这是一个自动化咖啡订单监控系统，用于监控外卖系统中的咖啡类订单，自动下载、筛选、更新咖啡订单数据。

## 🚀 主要功能
- **自动监控**：每30秒检测一次新订单
- **智能筛选**：自动筛选咖啡类订单
- **增量更新**：只添加新订单，避免重复
- **声音通知**：有新订单时播放提示音
- **弹窗提醒**：显示新增订单数量
- **Excel管理**：自动保存和管理Excel文件

## 📁 项目结构

### 核心程序
- **`咖啡订单监控.py`** - 主要的咖啡订单监控程序

### 辅助工具
- **`start_chrome.py`** - Chrome浏览器启动器
- **`kill_auto_cafe.py`** - 进程终止工具
- **`下载ChromeDriver.py`** - ChromeDriver下载工具

### 测试工具
- **`test_app_reading.py`** - Excel文件读取测试
- **`test_file_filter.py`** - 文件过滤测试

### 配置文件
- **`*.spec`** - PyInstaller打包配置文件
- **`chromedriver.exe`** - ChromeDriver驱动程序

## 🛠️ 使用方法

### 1. 启动Chrome浏览器
```powershell
python start_chrome.py
```

### 2. 运行咖啡订单监控
```powershell
python "咖啡订单监控.py"
```

### 3. 手动登录系统
程序会自动打开登录页面，请手动完成登录操作。

### 4. 自动监控
登录成功后，程序会自动开始监控咖啡订单。

## 📊 输出文件
- **`YYYYMMDD_所有外卖订单.xlsx`** - 原始订单数据
- **`YYYYMMDD_咖啡订单.xlsx`** - 筛选后的咖啡订单数据

## 🔧 技术栈
- **Python** - 主要编程语言
- **Selenium** - 网页自动化
- **Pandas** - 数据处理
- **Tkinter** - GUI界面
- **openpyxl** - Excel文件操作
- **win32com** - Windows COM接口

## 📝 注意事项
1. 需要预先安装Chrome浏览器
2. 需要ChromeDriver驱动程序
3. 需要网络连接访问目标系统
4. 需要手动完成登录操作

## 🆘 故障排除
- 如果ChromeDriver下载失败，请手动下载并放置到项目目录
- 如果连接失败，请检查Chrome浏览器是否已启动
- 如果登录失败，请检查网络连接和登录凭据 