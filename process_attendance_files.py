import os
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

from tools import write_attendance_sheet, is_chinese

import pandas as pd
from openpyxl import Workbook
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx2pdf import convert     # win系统下需有word，利用docx2pdf库将docx文件转换为pdf文件
import subprocess                # 其他平台需安装libreOffice用命令行实现转换，保真度差一些

def process_attendance_files(data, date, year, month, day, output_folder):

    data["姓名是否中文"] = data["姓名"].apply(is_chinese)

    # 考勤表excel
    def create_excel():
        wb = Workbook()
        
        # 本科生 sheet
        stu_cn = wb.active
        stu_cn.title = "本科生"
        cols = ["学号", "姓名", "学院", "班级", "旷课次数", "迟到次数", "早退次数", "旷课课时"]
        filtered_data = data[(data["旷课课时"] >= 2) & data["姓名是否中文"]]
        filtered_data = filtered_data.sort_values(by=["旷课课时", "旷课次数"], ascending=[False, False])
        write_attendance_sheet(stu_cn, filtered_data, cols, [15, 10, 10, 10, 10, 10, 10, 10], ["学院", "班级"])
        
        # 留学生 sheet
        international = wb.create_sheet(title="留学生")
        filtered_data = data[(data["旷课课时"] >= 2) & ~data["姓名是否中文"]]
        filtered_data = filtered_data.sort_values(by=["旷课课时", "旷课次数"], ascending=[False, False])
        write_attendance_sheet(international, filtered_data, cols, [15, 35, 10, 10, 10, 10, 10, 10], ["学院", "班级"])

        # 保存文件
        excel_output = f"{output_folder}/计算机科学与技术学院学生第{date}上课啦系统缺勤情况.xlsx"
        wb.save(excel_output)

    # 考勤通报word文档
    def create_docx():
        # 创建 Word 文档
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'  # 必须先设置font.name
        style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

        title = f"第{date}“上课啦”考勤通报"
        para_1 = doc.add_paragraph()
        run_1 = para_1.add_run(title)
        run_1.font.size = Pt(22)
        run_1.font.bold = True
        para_1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 设置标题居中

        # 段落内容
        content = f"以下同学在第{date}“上课啦”考勤中未正常出勤, 旷课学时计入个人档案, 并纳入日常考评。"
        para_2 = doc.add_paragraph()
        run_2 = para_2.add_run(content)
        run_2.font.size = Pt(16)
        para_2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY  # 设置内容两端对齐

        # 添加表格
        filtered_data = data[(data["旷课课时"] >= 2) & data["姓名是否中文"]]
        filtered_data = filtered_data.sort_values(by=["旷课课时", "旷课次数"], ascending=[False, False])

        table = doc.add_table(rows=1, cols=6, style="Table Grid")
        col_width_dict = {0: 1.6, 1: 1.12, 2: 0.7638, 3: 0.7638, 4: 0.7638, 5: 0.7638}
        row_height = Pt(25)
        
        # 设置列宽
        for col_num in range(6):
            table.cell(0, col_num).width = Inches(col_width_dict[col_num])

        # 设置表头
        headers = ["学号", "姓名", "旷课次数", "迟到次数", "早退次数", "旷课课时"]
        header_cells = table.rows[0].cells
        for idx, header in enumerate(headers):
            cell = header_cells[idx]
            cell.text = header                      # 设置单元格文本
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True            # 设置文字加粗

        # 添加数据行
        for _, row in filtered_data.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row["学号"])
            row_cells[1].text = row["姓名"]
            row_cells[2].text = str(row["旷课次数"])
            row_cells[3].text = str(row["迟到次数"])
            row_cells[4].text = str(row["早退次数"])
            row_cells[5].text = str(row["旷课课时"])

        # 全局设置行高和垂直对齐方式
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = 1         # 单元格垂直居中
                row.height = row_height             # 设置行高
                for paragraph in cell.paragraphs:
                    paragraph.alignment = 1         # 设置单元格文本居中
                    for run in paragraph.runs:
                        run.font.size = Pt(11)

        # 添加说明
        note_1 = "\n注:\n一、根据学生手册中“浙江师范大学学生违纪处分规定”中第三章第二十七条规定,学生一学期内旷课累计满10学时的, 给予警告处分; 满20学时的, 给予严重警告处分;满30学时的, 给与记过处分; 满40学时的, 给与留校察看处分; 因旷课屡次受到纪律处分并经教育不改的, 可以给予开除学籍处分;\n二、学时计算方法如下: \n(1) 旷课1小节为1学时, 未经请假缺勤1天, 不足5学时按5学时计; 超过5学时的, 按实际学时数计; \n(2) 学生无故迟到或早退达3次, 作旷课1学时计; "

        para_3 = doc.add_paragraph()
        run_3 = para_3.add_run(note_1)
        run_3.font.size = Pt(12)
        run_3.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY  # 设置内容两端对齐

        note_3 = "特此通报！"
        para_4 = doc.add_paragraph()
        run_4 = para_4.add_run(note_3)
        run_4.font.size = Pt(16)
        run_4.font.bold = True
        run_4.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # 设置内容左对齐

        para_5 = doc.add_paragraph()
        para_5.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT  # 设置段落右对齐
        run_5 = para_5.add_run(f"计算机科学与技术学院学工办\n{year}年{month}月{day}日")
        run_5.font.size = Pt(16)

        # 保存文件
        docx_output = f"{output_folder}/计算机科学与技术学院学生第{date}上课啦系统缺勤通报.docx"
        doc.save(docx_output)
        return docx_output

    # 考勤通报pdf
    def create_pdf(docx_path):
        def convert_docx_to_pdf(input_file, output_dir='.'):
            # 构建命令字符串，包含输出目录参数
            command = [
                'soffice',  # LibreOffice/OpenOffice 的命令行工具
                '--headless',  # 不显示图形用户界面
                '--invisible',  # 运行时不可见
                '--convert-to', 'pdf:writer_pdf_Export:AutoTableColumnWidths=false',  # 转换格式为目标格式
                '--outdir', output_dir,  # 指定输出目录
                input_file  # 输入文件路径
            ]

            try:
                # 使用 subprocess.run 来执行命令
                result = subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                print("转换成功", result.stdout.decode('utf-8', errors='ignore'))
            except subprocess.CalledProcessError as e:
                print("转换失败", e.stderr.decode('utf-8', errors='ignore'))

        pdf_output = f"{output_folder}/计算机科学与技术学院学生第{date}上课啦系统缺勤通报.pdf"

        # windows下可采用docx2pdf实现转换
        if os.name == "nt":
            convert(docx_path, pdf_output)
        else:
            convert_docx_to_pdf(docx_path, output_folder)

    # 执行
    create_excel()
    docx_output = create_docx()
    create_pdf(docx_output)
