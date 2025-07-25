import subprocess

# 创建目录
import os
os.makedirs(r"C:\selenium", exist_ok=True)

# 用 cmd 的 start 命令启动 chrome（依赖系统PATH能找到chrome.exe）
subprocess.Popen(
    'cmd /c start chrome.exe --remote-debugging-port=80 --user-data-dir="C:\\selenium" --disable-features=DownloadBubble,DownloadBubbleV2,DownloadShelf,DownloadShelfV2,DownloadBubbleV3,DownloadShelfV3 --disable-default-apps --disable-notifications --disable-download-notification --disable-download-bubble --disable-download-shelf',
    shell=True
)
