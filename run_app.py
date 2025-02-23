import streamlit.web.cli as stcli
import os
import sys
import threading
import pystray
from PIL import Image
import webbrowser
 
# 系统托盘图标回调函数：点击退出时停止图标并结束进程
def on_exit(icon, item):
    icon.stop()
    os._exit(0)
 
# 点击“打开网页”菜单时的回调函数
def open_website(icon, item):
    webbrowser.open("http://localhost:8501/")
 
def resolve_path(path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.getcwd()
    resolved_path = os.path.abspath(os.path.join(base_path, path))
    return resolved_path
 
def main():
 
    # 初始化系统托盘图标并设置菜单项
    icon = pystray.Icon("Streamlit App")
    icon.icon = Image.open(resolve_path("icon.ico"))  # 替换为你的图标路径
    icon.menu = pystray.Menu(
        pystray.MenuItem("打开网页", open_website),
        pystray.MenuItem("退出", on_exit)
    )
 
    # 启动 Streamlit 应用
    sys.argv = [
        "streamlit",
        "run",
        resolve_path("app.py"),
        "--global.developmentMode=false",
    ]
 
    # 启动系统托盘图标线程
    def run_icon():
        icon.run()
 
    threading.Thread(target=run_icon, daemon=True).start()
 
    # 启动 Streamlit 应用
    sys.exit(stcli.main())
 
if __name__ == "__main__":
    main()