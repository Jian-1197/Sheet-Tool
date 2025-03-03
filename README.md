# 上课啦考勤表制作工具

## 简介
这是一个用于制作上课啦平时考勤文件以及学期末汇总文件的自动化办公小项目，分别运用Streamlit和Tkinter构建了两种用户界面，利用Pandas,python-docx和OpenPyXL等库进行数据处理文件生成。

项目已部署到streamlit cloud community([demo](https://sheet-tool-zsjsj.streamlit.app/))

同时项目用pyinstaller打包成exe，可在release中查看，也可自行根据spec文件构建。

ui构建借助ai工具辅助完成。

## 文件结构
```
main
├── app.py                       # Streamlit应用程序代码
├── run_app.py                   # streamlit执行app.py
├── ctk_ui.py                    # 基于Tkinter,customtkinter的GUI程序
├── process_attendance_files.py  # 处理周/月考勤数据的逻辑
├── process_confirm_sheets.py    # 处理学期考勤数据的逻辑
├── tools.py                     # 处理数据时相关函数工具
├── run_app.spec                 # pyinstaller打包规范文件
├── ctk_ui.spec                  # pyinstaller打包规范文件
├── requirements.txt             # 项目依赖
└── README.md                    # 项目文档
```

## 快速开始

### 🚀 python部署

clone项目到本地
```
git clone https://github.com/Jian-1197/Sheet-Tool.git

```

推荐使用conda虚拟环境
```
conda create -n sheet-tool python=3.12

```

为项目配置好刚刚创建的python解释器后安装项目依赖
```
pip install -r requirements.txt

```

最后终端streamlit启动app.py
```
streamlit run app.py

```
或者启动run_app.py
```
python run_app.py

```

或者启动customtkinter_ui.py
```
python ctk_ui.py

```

### 🛸 直接运行EXE文件


在release中下载相关文件直接运行即可！🎉


## 交流学习


欢迎后续同学依据需求提问或拉取请求！😊
