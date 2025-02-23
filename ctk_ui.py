import os
import re
import sys
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from datetime import datetime
import warnings
import threading

warnings.filterwarnings('ignore')

# 添加脚本所在文件夹到sys.path中
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

# 导入自定义模块
from tools import zip_files, delete_files_and_folders
from process_attendance_files import process_attendance_files
from process_confirm_sheets import process_confirm_sheets

def select_file(filetypes):
    filename = filedialog.askopenfilename(filetypes=filetypes)
    return filename if filename != "" else None

class SheetToolApp:
    def __init__(self):
        ctk.set_appearance_mode("system")
        self.root = ctk.CTk()
        self.root.title("上课啦数据处理工具")
        self.root.geometry("800x600")
        
        # 配置根布局的行列权重
        self.root.grid_columnconfigure(0, weight=0)  # 侧边栏列固定宽度
        self.root.grid_columnconfigure(1, weight=1)  # 主内容区自适应
        self.root.grid_rowconfigure(1, weight=1)
        
        # 使用变量保存展开时的侧边栏宽度
        self.sidebar_width = 80
        
        # 侧边栏框架（可折叠）
        self.sidebar = ctk.CTkFrame(self.root, width=self.sidebar_width, corner_radius=0)
        self.sidebar.grid(row=0, column=0, rowspan=3, sticky="nsew")
        self.sidebar.grid_propagate()  
        self.sidebar.grid_rowconfigure(4, weight=1)
        
        # 折叠按钮
        self.collapse_btn = ctk.CTkButton(
            self.sidebar, 
            text="折叠 ◀", 
            width=self.sidebar_width-20,  
            command=self.toggle_sidebar
        )
        self.collapse_btn.grid(row=0, column=0, padx=10, pady=40, sticky="ew")
        
        # 侧边栏选项
        self.attendance_btn = ctk.CTkButton(
            self.sidebar,
            text="考勤通报表",
            width=self.sidebar_width-20,
            command=lambda: self.show_tab("attendance")
        )
        self.attendance_btn.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        
        self.confirm_btn = ctk.CTkButton(
            self.sidebar,
            text="确认签字表",
            width=self.sidebar_width-20,
            command=lambda: self.show_tab("confirm")
        )
        self.confirm_btn.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        
        # 内容区域框架
        self.content_frame = ctk.CTkFrame(self.root)
        self.content_frame.grid(row=0, column=1, rowspan=3, sticky="nsew", padx=20, pady=20)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)
        
        # 创建内容页
        self.tabs = {
            "attendance": AttendanceTab(self.content_frame, self),
            "confirm": ConfirmTab(self.content_frame, self)
        }
        self.current_tab = None
        
        # 进度条组件
        self.progress_canvas = None
        self.progress_active = False
        self.after_id = None
        
        # 初始化显示第一个标签页
        self.show_tab("attendance")
        self.is_sidebar_collapsed = False

    def toggle_sidebar(self):
        """切换侧边栏展开/折叠状态"""
        if self.is_sidebar_collapsed:
            new_width = self.sidebar_width
            self.collapse_btn.configure(text="折叠 ◀", width=new_width-20)
            self.is_sidebar_collapsed = False
        else:
            new_width = 30
            self.collapse_btn.configure(text="▶", width=new_width-10)
            self.is_sidebar_collapsed = True
            
        # 更新侧边栏组件尺寸
        self.sidebar.configure(width=new_width)
        self.attendance_btn.configure(
            width=new_width-20 if new_width > 40 else new_width-10,
            text="📅" if new_width < 40 else "考勤通报表"
        )
        self.confirm_btn.configure(
            width=new_width-20 if new_width > 40 else new_width-10,
            text="📑" if new_width < 40 else "确认签字表"
        )

    def show_tab(self, tab_name):
        """显示指定的标签页"""
        if self.current_tab:
            self.current_tab.pack_forget()
            
        self.current_tab = self.tabs[tab_name]
        self.current_tab.pack(expand=True, fill="both", padx=10, pady=10)

    def start_progress(self):
        """启动小块移动的进度条动画"""
        # 创建一个Canvas作为进度条背景
        self.progress_canvas = tk.Canvas(self.root, height=10, bg="#cccccc", highlightthickness=0)
        self.progress_canvas.grid(row=2, column=1, sticky="ew", padx=10, pady=5)
        # 确保Canvas尺寸更新
        self.progress_canvas.update()
        self.canvas_width = self.progress_canvas.winfo_width()
        self.block_width = 80  # 小块的宽度（像素）
        # 在Canvas中创建小块（蓝色矩形）
        self.block_id = self.progress_canvas.create_rectangle(0, 0, self.block_width, 10, fill="blue", width=0)
        self.progress_active = True
        self.update_progress()

    def update_progress(self):
        if not self.progress_active:
            return
        # 获取当前小块的x坐标
        coords = self.progress_canvas.coords(self.block_id)
        current_x = coords[0]
        # 每次移动10像素
        new_x = current_x + 10
        if new_x > self.canvas_width - self.block_width:
            new_x = 0
        # 更新小块坐标
        self.progress_canvas.coords(self.block_id, new_x, 0, new_x + self.block_width, 10)
        self.after_id = self.root.after(50, self.update_progress)

    def stop_progress(self):
        self.progress_active = False
        if self.after_id:
            self.root.after_cancel(self.after_id)
            self.after_id = None
        if self.progress_canvas:
            self.progress_canvas.destroy()
            self.progress_canvas = None

    def run(self):
        self.root.mainloop()

class AttendanceTab(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        # 定义变量
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.date_var = tk.StringVar(value="一周")
        self.date_entry = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))

        # 文件上传区域
        file_frame = ctk.CTkFrame(self)
        file_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)
        file_frame.grid_columnconfigure((0, 1), weight=1)

        self.generate_button = ctk.CTkButton(self, text="生成考勤通报表", command=self.generate_attendance_report)
        self.generate_button.grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        # 第一个文件上传
        ctk.CTkButton(file_frame, text="选择周/月考勤数据", 
                      command=self.select_file1).grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        ctk.CTkLabel(file_frame, textvariable=self.file1_path).grid(row=1, column=0, sticky="ew", padx=5)

        # 第二个文件上传
        ctk.CTkButton(file_frame, text="选择考勤明细表",
                      command=self.select_file2).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ctk.CTkLabel(file_frame, textvariable=self.file2_path).grid(row=1, column=1, sticky="ew", padx=5)

        # 参数设置区域
        param_frame = ctk.CTkFrame(self)
        param_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        param_frame.grid_columnconfigure((0, 1), weight=1)

        # 选择周/月
        ctk.CTkLabel(param_frame, text="选择周/月").grid(row=0, column=0, sticky="w", padx=5)
        weeks = ["一周", "二周", "三周", "四周", "五周", "六周",
                 "七周", "八周", "九周", "十周", "十一周", "十二周",
                 "十三周", "十四周", "十五周", "十六周", "十七周", "十八周",
                 "一月", "二月", "三月", "四月", "五月", "六月",
                 "七月", "八月", "九月", "十月", "十一月", "十二月"]
        combo = ctk.CTkOptionMenu(param_frame, values=weeks,
                                  variable=self.date_var)
        combo.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

        # 制表日期
        ctk.CTkLabel(param_frame, text="制表日期 (YYYY-MM-DD)").grid(row=0, column=1, sticky="w", padx=5)
        ctk.CTkEntry(param_frame, textvariable=self.date_entry).grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        # # 生成按钮
        # ctk.CTkButton(self, text="生成考勤通报", 
        #               command=self.generate_attendance_report).grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        # 配置网格权重
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

    def select_file1(self):
        path = select_file([("Excel Files", "*.xls *.xlsx")])
        if path:
            self.file1_path.set(path)

    def select_file2(self):
        path = select_file([("Excel Files", "*.xls *.xlsx")])
        if path:
            self.file2_path.set(path)

    def generate_attendance_report(self):
        self.generate_button.configure(state="disabled")
        file1 = self.file1_path.get()
        file2 = self.file2_path.get()
        date_period = self.date_var.get()
        date_str = self.date_entry.get()

        if not (file1 and file2):
            messagebox.showwarning("警告", "请先选择所有必需文件！")
            self.generate_button.configure(state="normal")
            return

        try:
            calendar = datetime.strptime(date_str, "%Y-%m-%d")
        except Exception as e:
            messagebox.showerror("错误", f"日期格式错误：请按 YYYY-MM-DD 格式输入\n{e}")
            self.generate_button.configure(state="normal")
            return

        self.app.start_progress()

        def thread_task():
            try:
                year = calendar.strftime('%Y')
                month = calendar.strftime('%m')
                day = calendar.strftime('%d')

                data = pd.read_excel(file1)
                attendance_folder = f"第{date_period}"
                os.makedirs(attendance_folder, exist_ok=True)

                shutil.copy2(file1, os.path.join(attendance_folder, "原始数据.xlsx"))
                shutil.copy2(file2, os.path.join(attendance_folder, f"计算机科学与技术学院第{date_period}上课啦考勤明细.xlsx"))

                process_attendance_files(data, date_period, year, month, day, attendance_folder)
                zip_files([attendance_folder], attendance_folder)
                zip_path = f"{attendance_folder}.zip"

                self.app.root.after(0, lambda: self.save_attendance_file(zip_path, attendance_folder))
            except Exception as e:
                self.app.root.after(0, lambda: messagebox.showerror("错误", f"处理过程中出现错误：{str(e)}"))
            finally:
                self.app.root.after(0, self.app.stop_progress)
                self.app.root.after(0, lambda: self.generate_button.configure(state="normal"))

        threading.Thread(target=thread_task, daemon=True).start()

    def save_attendance_file(self, zip_path, folder):
        dest = filedialog.asksaveasfilename(
            initialfile=os.path.basename(zip_path),
            defaultextension=".zip",
            filetypes=[("Zip Files", "*.zip")],
            title="保存压缩包"
        )
        if dest:
            shutil.copy2(zip_path, dest)
            delete_files_and_folders([folder, zip_path])
            messagebox.showinfo("成功", "保存并清理生成文件成功！")
        else:
            messagebox.showwarning("提示", "未选择保存路径，生成的临时文件未清理！")

class ConfirmTab(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        # 定义变量
        self.file_attendance_path = tk.StringVar()
        self.study_year_combo = tk.StringVar(value="2024-2025")
        self.custom_year_entry = tk.StringVar()
        self.semester_combo = tk.StringVar(value="第一学期")
        self.grade_entry = tk.StringVar()
        self.max_classes_spin = tk.StringVar(value="5")
        self.start_date_entry = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.end_date_entry = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))

        # 文件上传区域
        file_frame = ctk.CTkFrame(self)
        file_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)
        file_frame.grid_columnconfigure(0, weight=1)

        self.generate_button = ctk.CTkButton(self, text="生成确认签字表", command=self.generate_confirm_sheet)
        self.generate_button.grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        ctk.CTkButton(file_frame, text="选择考勤数据文件",
                      command=self.select_file).grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        ctk.CTkLabel(file_frame, textvariable=self.file_attendance_path).grid(row=1, column=0, sticky="ew", padx=5)

        # 参数设置区域
        param_frame = ctk.CTkFrame(self)
        param_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        param_frame.grid_columnconfigure((0, 1), weight=1)

        # 学年设置
        ctk.CTkLabel(param_frame, text="选择学年").grid(row=0, column=0, sticky="w", padx=5)
        years = ["2024-2025", "2025-2026", "2026-2027", "2027-2028"]
        ctk.CTkOptionMenu(param_frame, values=years,
                          variable=self.study_year_combo).grid(row=1, column=0, sticky="ew", padx=5, pady=5)

        ctk.CTkLabel(param_frame, text="或手动输入学年 (YYYY-YYYY)").grid(row=0, column=1, sticky="w", padx=5)
        ctk.CTkEntry(param_frame, textvariable=self.custom_year_entry).grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        # 学期选择
        ctk.CTkLabel(param_frame, text="选择学期").grid(row=2, column=0, columnspan=2, sticky="w", padx=5)
        semesters = ["第一学期", "第二学期"]
        ctk.CTkOptionMenu(param_frame, values=semesters,
                          variable=self.semester_combo).grid(row=3, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

        # 年级范围和班级数量
        ctk.CTkLabel(param_frame, text="输入年级范围 (YY-YY)").grid(row=4, column=0, sticky="w", padx=5)
        ctk.CTkEntry(param_frame, textvariable=self.grade_entry).grid(row=5, column=0, sticky="ew", padx=5, pady=5)
        
        ctk.CTkLabel(param_frame, text="最大班级数量").grid(row=4, column=1, sticky="w", padx=5)
        spin_frame = ctk.CTkFrame(param_frame)
        spin_frame.grid(row=5, column=1, sticky="ew", padx=5, pady=5)
        def validate_spin(value):
            if value == "":  # 允许空输入（用户删除默认值时）
                return True
            try:
                val = int(value)
                return 1 <= val <= 10
            except ValueError:
                return False  # 非数字输入直接拒绝
        ctk.CTkEntry(spin_frame, textvariable=self.max_classes_spin, 
                validate="key", validatecommand=(self.register(validate_spin), '%P'),
                ).pack(side="left", fill="x", expand=True)

        # 日期选择
        ctk.CTkLabel(param_frame, text="学期开始日期 (YYYY-MM-DD)").grid(row=6, column=0, sticky="w", padx=5)
        ctk.CTkEntry(param_frame, textvariable=self.start_date_entry).grid(row=7, column=0, sticky="ew", padx=5, pady=5)

        ctk.CTkLabel(param_frame, text="学期结束日期 (YYYY-MM-DD)").grid(row=6, column=1, sticky="w", padx=5)
        ctk.CTkEntry(param_frame, textvariable=self.end_date_entry).grid(row=7, column=1, sticky="ew", padx=5, pady=5)

        # 生成按钮
        # ctk.CTkButton(self, text="生成确认签字表",
        #               command=self.generate_confirm_sheet).grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        # 配置网格权重
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

    def select_file(self):
        path = select_file([("Excel Files", "*.xls *.xlsx")])
        if path:
            self.file_attendance_path.set(path)

    def generate_confirm_sheet(self):
        self.generate_button.configure(state="disabled")
        file_att = self.file_attendance_path.get()
        if not file_att:
            messagebox.showwarning("警告", "请先选择考勤数据文件！")
            self.generate_button.configure(state="normal")
            return

        study_year = self.study_year_combo.get().strip()
        custom_study_year = self.custom_year_entry.get().strip()
        if custom_study_year:
            study_year = custom_study_year

        semester = self.semester_combo.get().strip()
        grade = self.grade_entry.get().strip()
        max_classes = self.max_classes_spin.get()

        start_date_str = self.start_date_entry.get().strip()
        end_date_str = self.end_date_entry.get().strip()

        if not grade:
            messagebox.showwarning("警告", "年级范围必填！")
            self.generate_button.configure(state="normal")
            return

        pattern = re.compile(r"^\d{2}\s*-\s*\d{2}$")
        if not pattern.match(grade):
            messagebox.showerror("错误", "格式错误：请按示例格式输入（如22-24）")
            self.generate_button.configure(state="normal")
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
            self.generate_button.configure(state="normal")
            return

        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
        except Exception as e:
            messagebox.showerror("错误", f"日期格式错误：请按 YYYY-MM-DD 格式输入\n{e}")
            self.generate_button.configure(state="normal")
            return

        self.app.start_progress()

        def thread_task():
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
                zip_path = "确认签字表+汇总表.zip"

                self.app.root.after(0, lambda: self.save_confirm_file(zip_path, confirm_folders))
            except Exception as e:
                self.app.root.after(0, lambda: messagebox.showerror("错误", f"生成过程中出现错误：{str(e)}"))
            finally:
                self.app.root.after(0, self.app.stop_progress)
                self.app.root.after(0, lambda: self.generate_button.configure(state="normal"))

        threading.Thread(target=thread_task, daemon=True).start()

    def save_confirm_file(self, zip_path, folders):
        dest = filedialog.asksaveasfilename(
            initialfile=os.path.basename(zip_path),
            defaultextension=".zip",
            filetypes=[("Zip Files", "*.zip")],
            title="保存压缩包"
        )
        if dest:
            shutil.copy2(zip_path, dest)
            delete_files_and_folders(folders + [zip_path])
            messagebox.showinfo("成功", "保存并清理生成文件成功！")
        else:
            messagebox.showwarning("提示", "未选择保存路径，生成的临时文件未清理！")

if __name__ == "__main__":
    app = SheetToolApp()
    app.run()