import os

def main():
    # 这里执行taskkill命令
    os.system('taskkill /f /im auto_cafe_final.exe')

if __name__ == "__main__":
    main()
