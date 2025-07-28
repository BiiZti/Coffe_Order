# Connection Pool 问题修复说明

## 问题描述

咖啡订单监控程序出现 "Connection pool is full, discarding connection: localhost. Connection pool size: 1" 错误。

## 问题原因

1. **Selenium WebDriver连接池限制**：ChromeDriver默认连接池大小只有1，无法处理并发连接
2. **Chrome浏览器进程残留**：之前的Chrome进程没有正确关闭，占用连接
3. **缓存和会话数据**：Chrome的缓存和会话数据可能导致连接问题

## 解决方案

### 方案1：使用修复脚本（推荐）

运行专门的修复脚本：

```bash
python fix_connection_pool.py
```

这个脚本会：
1. 结束所有Chrome相关进程
2. 清理Chrome缓存和数据
3. 重启Chrome浏览器
4. 检查connection pool状态

### 方案2：手动修复

#### 步骤1：结束Chrome进程
1. 打开任务管理器（Ctrl+Shift+Esc）
2. 找到所有Chrome相关进程
3. 结束这些进程

#### 步骤2：清理Chrome数据
1. 关闭Chrome浏览器
2. 删除Chrome缓存目录：
   - `%LOCALAPPDATA%\Google\Chrome\User Data\Default\Cache`
   - `%LOCALAPPDATA%\Google\Chrome\User Data\Default\Session Storage`

#### 步骤3：重启Chrome
1. 运行 `start_chrome.py` 启动Chrome
2. 或者手动启动Chrome并添加参数：`--remote-debugging-port=80`

### 方案3：代码修复

已经对代码进行了以下修复：

1. **添加Chrome选项**：
   ```python
   chrome_options.add_argument("--disable-background-timer-throttling")
   chrome_options.add_argument("--disable-backgrounding-occluded-windows")
   chrome_options.add_argument("--disable-renderer-backgrounding")
   chrome_options.add_argument("--disable-features=TranslateUI")
   chrome_options.add_argument("--disable-ipc-flooding-protection")
   ```

2. **连接重试机制**：
   ```python
   max_retries = 3
   retry_count = 0
   while retry_count < max_retries and driver is None:
       # 尝试连接Chrome
   ```

3. **连接池清理**：
   ```python
   def cleanup_connections():
       # 清理Selenium连接
       if 'driver' in globals() and driver:
           driver.quit()
   ```

4. **错误处理**：
   ```python
   if "connection pool is full" in str(e).lower():
       cleanup_connections()
       time.sleep(2)
   ```

## 预防措施

### 1. 正确关闭程序
- 使用Ctrl+C正确退出程序
- 不要强制关闭命令行窗口

### 2. 定期清理
- 定期运行修复脚本
- 定期清理Chrome缓存

### 3. 监控连接状态
- 注意观察日志中的connection pool警告
- 及时处理连接问题

## 故障排除

### 如果修复脚本失败：

1. **手动结束进程**：
   ```bash
   taskkill /f /im chrome.exe
   taskkill /f /im chromedriver.exe
   ```

2. **检查端口占用**：
   ```bash
   netstat -ano | findstr :80
   ```

3. **重启系统**：
   如果问题持续，考虑重启系统

### 如果仍然有问题：

1. **检查Chrome版本**：
   确保Chrome和ChromeDriver版本匹配

2. **更新依赖**：
   ```bash
   pip install --upgrade selenium webdriver-manager
   ```

3. **使用不同的端口**：
   修改Chrome启动参数使用不同端口

## 相关文件

- `fix_connection_pool.py` - 修复脚本
- `咖啡订单监控.py` - 已修复的主程序
- `start_chrome.py` - Chrome启动脚本
- `requirements.txt` - 依赖列表

## 技术支持

如果问题仍然存在，请：
1. 查看详细的错误日志
2. 检查Chrome和ChromeDriver版本
3. 确认网络连接正常
4. 尝试在不同的环境中运行 