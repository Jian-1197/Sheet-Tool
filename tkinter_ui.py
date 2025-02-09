import os
import re
import sys
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# 获取脚本所在路径，并添加到sys.path中（便于导入自定义模块）
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

# 导入自定义模块
from tools import zip_files, delete_files_and_folders
from process_attendance_files import process_attendance_files
from process_confirm_sheets import process_confirm_sheets


# 全局变量：记录生成的压缩包路径及对应的生成文件夹（用于删除）
attendance_zip_path = None
attendance_folder = None
confirm_zip_path = None
confirm_folders = None

# 辅助函数：选择文件，返回文件路径
def select_file(filetypes):
    filename = filedialog.askopenfilename(filetypes=filetypes)
    return filename if filename != "" else None

# 考勤表生成函数
def generate_attendance_report():
    global attendance_zip_path, attendance_folder
    file1 = attendance_tab.file1_path.get()
    file2 = attendance_tab.file2_path.get()
    date_period = attendance_tab.date_var.get()
    date_str = attendance_tab.date_entry.get()

    if not (file1 and file2):
        messagebox.showwarning("警告", "请先选择所有必需文件！")
        return

    try:
        calendar = datetime.strptime(date_str, "%Y-%m-%d")
    except Exception as e:
        messagebox.showerror("错误", f"日期格式错误：请按 YYYY-MM-DD 格式输入\n{e}")
        return

    try:
        year = calendar.strftime('%Y')
        month = calendar.strftime('%m')
        day = calendar.strftime('%d')

        data = pd.read_excel(file1)
        attendance_folder = f"第{date_period}"
        os.makedirs(attendance_folder, exist_ok=True)

        # 换名另存原始数据文件
        shutil.copy2(file1, os.path.join(attendance_folder, "原始数据.xlsx"))
        shutil.copy2(file2, os.path.join(attendance_folder, f"计算机科学与技术学院第{date_period}上课啦考勤明细.xlsx"))

        process_attendance_files(data, date_period, year, month, day, attendance_folder)
        zip_files([attendance_folder], attendance_folder)
        attendance_zip_path = f"{attendance_folder}.zip"

        # 自动触发另存压缩包保存
        dest = filedialog.asksaveasfilename(
            initialfile=os.path.basename(attendance_zip_path),
            defaultextension=".zip",
            filetypes=[("Zip Files", "*.zip")],
            title="保存压缩包"
        )
        if dest:
            shutil.copy2(attendance_zip_path, dest)
            delete_files_and_folders([attendance_folder, attendance_zip_path])
            messagebox.showinfo("成功", "保存并清理生成文件成功！")
        else:
            messagebox.showwarning("提示", "未选择保存路径，生成的临时文件未清理！")
    except Exception as e:
        messagebox.showerror("错误", f"处理过程中出现错误：{str(e)}")

# 确认签字表生成函数
def generate_confirm_sheet():
    global confirm_zip_path, confirm_folders
    file_att = confirm_tab.file_attendance_path.get()
    if not file_att:
        messagebox.showwarning("警告", "请先选择考勤数据文件！")
        return

    study_year = confirm_tab.study_year_combo.get().strip()
    custom_study_year = confirm_tab.custom_year_entry.get().strip()
    if custom_study_year:
        study_year = custom_study_year

    semester = confirm_tab.semester_combo.get().strip()
    grade = confirm_tab.grade_entry.get().strip()
    max_classes = confirm_tab.max_classes_spin.get()

    start_date_str = confirm_tab.start_date_entry.get().strip()
    end_date_str = confirm_tab.end_date_entry.get().strip()

    if not grade:
        messagebox.showwarning("警告", "年级范围必填！")
        return

    pattern = re.compile(r"^\d{2}\s*-\s*\d{2}$")
    if not pattern.match(grade):
        messagebox.showerror("错误", "格式错误：请按示例格式输入（如22-24）")
        return

    try:
        start_grade, end_grade = [int(x.strip()) for x in grade.split('-')]
        if start_grade > end_grade:
            messagebox.showerror("错误", "年级范围输入错误：请按从小到大的顺序输入")
            return
        grade_list = [str(g) for g in range(start_grade, end_grade + 1)]
        class_list = [f"{g}{str(i).zfill(2)}" for g in grade_list for i in range(1, int(max_classes) + 1)]
    except Exception as e:
        messagebox.showerror("错误", f"年级转换错误：{str(e)}")
        return

    try:
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
    except Exception as e:
        messagebox.showerror("错误", f"日期格式错误：请按 YYYY-MM-DD 格式输入\n{e}")
        return

    try:
        data_attendance = pd.read_excel(file_att)
        start_year = start_date.strftime('%Y')
        start_month = start_date.strftime('%m')
        start_day = start_date.strftime('%d')
        end_year = end_date.strftime('%Y')
        end_month = end_date.strftime('%m')
        end_day = end_date.strftime('%d')

        output_folder_1 = "上课啦确认签字表"
        output_folder_2 = "学期汇总表+违规违纪名单"
        os.makedirs(output_folder_1, exist_ok=True)
        os.makedirs(output_folder_2, exist_ok=True)
        confirm_folders = [output_folder_1, output_folder_2]

        process_confirm_sheets(
            data_attendance, study_year, semester,
            start_year, start_month, start_day,
            end_year, end_month, end_day,
            output_folder_1, output_folder_2, class_list
        )
        zip_files(confirm_folders, "确认签字表+汇总表")
        confirm_zip_path = "确认签字表+汇总表.zip"

        # 自动触发另存压缩包保存
        dest = filedialog.asksaveasfilename(
            initialfile=os.path.basename(confirm_zip_path),
            defaultextension=".zip",
            filetypes=[("Zip Files", "*.zip")],
            title="保存压缩包"
        )
        if dest:
            shutil.copy2(confirm_zip_path, dest)
            delete_files_and_folders(confirm_folders + [confirm_zip_path])
            messagebox.showinfo("成功", "保存并清理生成文件成功！")
        else:
            messagebox.showwarning("提示", "未选择保存路径，生成的临时文件未清理！")
    except Exception as e:
        messagebox.showerror("错误", f"生成过程中出现错误：{str(e)}")

# 考勤表Tab
class AttendanceTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding=10)
        # 定义变量
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.date_var = tk.StringVar(value="一周")
        self.date_entry = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))

        # 设置网格权重，实现自适应布局
        self.columnconfigure(0, weight=1)

        # ── 数据上传区域 (文件上传按钮并排) ──
        file_frame = ttk.Labelframe(self, text="数据上传")
        file_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        file_frame.columnconfigure(0, weight=1)
        file_frame.columnconfigure(1, weight=1)

        # 第一个文件上传
        file1_frame = ttk.Frame(file_frame)
        file1_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        file1_frame.columnconfigure(0, weight=1)
        ttk.Button(file1_frame, text="选择周/月考勤数据", command=self.select_file1).grid(row=0, column=0, sticky="ew", padx=5)
        ttk.Label(file1_frame, textvariable=self.file1_path, relief="sunken").grid(row=1, column=0, sticky="ew", padx=5, pady=2)

        # 第二个文件上传
        file2_frame = ttk.Frame(file_frame)
        file2_frame.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        file2_frame.columnconfigure(0, weight=1)
        ttk.Button(file2_frame, text="选择考勤明细表", command=self.select_file2).grid(row=0, column=0, sticky="ew", padx=5)
        ttk.Label(file2_frame, textvariable=self.file2_path, relief="sunken").grid(row=1, column=0, sticky="ew", padx=5, pady=2)

        # ── 参数设置区域 (两列布局) ──
        param_frame = ttk.Labelframe(self, text="参数设置")
        param_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        param_frame.columnconfigure(0, weight=1)
        param_frame.columnconfigure(1, weight=1)

        # 左侧参数：选择周/月
        ttk.Label(param_frame, text="选择周/月").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        combo = ttk.Combobox(param_frame, textvariable=self.date_var, state="readonly",
                             values=["一周", "二周", "三周", "四周", "五周", "六周",
                                     "七周", "八周", "九周", "十周", "十一周", "十二周",
                                     "十三周", "十四周", "十五周", "十六周", "十七周", "十八周",
                                     "一月", "二月", "三月", "四月", "五月", "六月",
                                     "七月", "八月", "九月", "十月", "十一月", "十二月"])
        combo.grid(row=1, column=0, sticky="ew", padx=5, pady=2)

        # 右侧参数：制表日期
        ttk.Label(param_frame, text="制表日期 (YYYY-MM-DD)").grid(row=0, column=1, sticky="w", padx=5, pady=2)
        date_entry_widget = ttk.Entry(param_frame, textvariable=self.date_entry)
        date_entry_widget.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        # ── 生成按钮 (填充整行) ──
        ttk.Button(self, text="生成考勤通报", command=generate_attendance_report)\
            .grid(row=2, column=0, sticky="ew", padx=5, pady=10)

    def select_file1(self):
        path = select_file([("Excel Files", "*.xls *.xlsx")])
        if path:
            self.file1_path.set(path)

    def select_file2(self):
        path = select_file([("Excel Files", "*.xls *.xlsx")])
        if path:
            self.file2_path.set(path)

# 确认签字表Tab
class ConfirmTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding=10)
        # 定义变量
        self.file_attendance_path = tk.StringVar()
        self.study_year_combo = tk.StringVar(value="2024-2025")
        self.custom_year_entry = tk.StringVar()
        self.semester_combo = tk.StringVar(value="第一学期")
        self.grade_entry = tk.StringVar()
        self.max_classes_spin = tk.StringVar(value="5")
        self.start_date_entry = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.end_date_entry = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))

        self.columnconfigure(0, weight=1)

        # ── 数据上传区域 ──
        file_frame = ttk.Labelframe(self, text="数据上传")
        file_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        file_frame.columnconfigure(0, weight=1)
        file_frame.columnconfigure(1, weight=1)
        # 此处只有一个文件上传按钮，跨越两列居中显示
        ttk.Button(file_frame, text="选择考勤数据文件", command=self.select_file)\
            .grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        ttk.Label(file_frame, textvariable=self.file_attendance_path, relief="sunken")\
            .grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=2)

        # ── 参数设置区域（重新排列） ──
        param_frame = ttk.Labelframe(self, text="参数设置")
        param_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        param_frame.columnconfigure(0, weight=1)
        param_frame.columnconfigure(1, weight=1)

        # 第一行：选择学年 和 或手动输入学年
        left_frame = ttk.Frame(param_frame)
        left_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=2)
        left_frame.columnconfigure(0, weight=1)
        ttk.Label(left_frame, text="选择学年").grid(row=0, column=0, sticky="w")
        study_year_combo_widget = ttk.Combobox(left_frame, textvariable=self.study_year_combo, 
                                            state="readonly", values=["2024-2025", "2025-2026", "2026-2027", "2027-2028"])
        study_year_combo_widget.grid(row=1, column=0, sticky="ew", pady=2)

        right_frame = ttk.Frame(param_frame)
        right_frame.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        right_frame.columnconfigure(0, weight=1)
        ttk.Label(right_frame, text="或手动输入学年 (YYYY-YYYY)").grid(row=0, column=0, sticky="w")
        ttk.Entry(right_frame, textvariable=self.custom_year_entry).grid(row=1, column=0, sticky="ew", pady=2)

        # 第二行：选择学期 (跨两列)
        semester_frame = ttk.Frame(param_frame)
        semester_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=2)
        semester_frame.columnconfigure(0, weight=1)
        ttk.Label(semester_frame, text="选择学期").grid(row=0, column=0, sticky="w")
        semester_combo_widget = ttk.Combobox(semester_frame, textvariable=self.semester_combo, 
                                            state="readonly", values=["第一学期", "第二学期"])
        semester_combo_widget.grid(row=1, column=0, sticky="ew", pady=2)

        # 第三行：输入年级范围 和 最大班级数量
        left_frame2 = ttk.Frame(param_frame)
        left_frame2.grid(row=2, column=0, sticky="ew", padx=5, pady=2)
        left_frame2.columnconfigure(0, weight=1)
        ttk.Label(left_frame2, text="输入年级范围 (YY-YY)").grid(row=0, column=0, sticky="w")
        ttk.Entry(left_frame2, textvariable=self.grade_entry).grid(row=1, column=0, sticky="ew", pady=2)

        right_frame2 = ttk.Frame(param_frame)
        right_frame2.grid(row=2, column=1, sticky="ew", padx=5, pady=2)
        right_frame2.columnconfigure(0, weight=1)
        ttk.Label(right_frame2, text="最大班级数量").grid(row=0, column=0, sticky="w")
        # 添加验证命令：确保输入的值在1到10之间
        vcmd = (self.register(self.validate_max_classes_spin), '%P')
        spinbox = ttk.Spinbox(
            right_frame2, from_=1, to=10, textvariable=self.max_classes_spin, width=5,
            validate='key', validatecommand=vcmd
        )
        spinbox.grid(row=1, column=0, sticky="ew", pady=2)
        
        # 第四行：学期开始日期 和 学期结束日期
        left_frame3 = ttk.Frame(param_frame)
        left_frame3.grid(row=3, column=0, sticky="ew", padx=5, pady=2)
        left_frame3.columnconfigure(0, weight=1)
        ttk.Label(left_frame3, text="学期开始日期 (YYYY-MM-DD)").grid(row=0, column=0, sticky="w")
        ttk.Entry(left_frame3, textvariable=self.start_date_entry).grid(row=1, column=0, sticky="ew", pady=2)

        right_frame3 = ttk.Frame(param_frame)
        right_frame3.grid(row=3, column=1, sticky="ew", padx=5, pady=2)
        right_frame3.columnconfigure(0, weight=1)
        ttk.Label(right_frame3, text="学期结束日期 (YYYY-MM-DD)").grid(row=0, column=0, sticky="w")
        ttk.Entry(right_frame3, textvariable=self.end_date_entry).grid(row=1, column=0, sticky="ew", pady=2)

        # ── 生成按钮 ──
        ttk.Button(self, text="生成确认签字表", command=generate_confirm_sheet)\
            .grid(row=2, column=0, sticky="ew", padx=5, pady=10)

    def select_file(self):
        path = select_file([("Excel Files", "*.xls *.xlsx")])
        if path:
            self.file_attendance_path.set(path)
            
    def validate_max_classes_spin(self, new_value):
        if new_value == "":
            return True
        try:
            val = int(new_value)
            return 1 <= val <= 10
        except ValueError:
            return False

def main():
    # 主窗口
    root = ttk.Window(title="上课啦数据处理工具",themename="minty")
    root.geometry("800x600")

    # 调整主窗口网格布局
    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    # ── 主题切换区域 ──
    theme_frame = ttk.Frame(root, padding=5)
    theme_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
    theme_frame.columnconfigure(1, weight=1)
    
    ttk.Label(theme_frame, text="Theme✨：").grid(row=0, column=0, sticky="w")
    theme_names = list(root.style.theme_names())
    theme_combobox = ttk.Combobox(theme_frame, values=theme_names, state="readonly")
    theme_combobox.set("minty")
    theme_combobox.grid(row=0, column=1, sticky="ew", padx=5)

    def change_theme(event=None):
        new_theme = theme_combobox.get()
        root.style.theme_use(new_theme)

    theme_combobox.bind("<<ComboboxSelected>>", change_theme)

    # ── Notebook 区域 ──
    notebook = ttk.Notebook(root)
    notebook.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

    # 添加各个 Tab
    global attendance_tab, confirm_tab
    attendance_tab = AttendanceTab(notebook)
    confirm_tab = ConfirmTab(notebook)
    notebook.add(attendance_tab, text="考勤表")
    notebook.add(confirm_tab, text="确认签字表")

    root.mainloop()

if __name__ == "__main__":
    main()