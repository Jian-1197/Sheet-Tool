import os
import sys
import re
import pandas as pd
import streamlit as st
import warnings
warnings.filterwarnings("ignore")

script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

from tools import zip_files, delete_files_and_folders
from process_attendance_files import process_attendance_files
from process_confirm_sheets import process_confirm_sheets

# 页面配置
st.set_page_config(
    page_title="上课啦表格处理中心",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 主标题
st.markdown('<h1 class="title">📊 上课啦表格处理中心</h1>', unsafe_allow_html=True)

# 标签页
tab1, tab2 = st.tabs(["📅 考勤表", "📑 确认签字表"])
                
with tab1:
    st.subheader("📤 上传周/月考勤数据")
    with st.expander("数据上传（点击展开/收起）",expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            uploaded_file_1 = st.file_uploader("上传周/月考勤数据", type=["xls", "xlsx"])
        with col2:
            uploaded_file_2 = st.file_uploader("上传考勤明细表", type=["xls", "xlsx"])
        if uploaded_file_1 and uploaded_file_2:
            cols = st.columns(2)
            with cols[0]:
                st.write("周/月考勤数据预览")
                st.dataframe(pd.read_excel(uploaded_file_1,nrows=3))
            with cols[1]:
                st.write("考勤明细数据预览")
                st.dataframe(pd.read_excel(uploaded_file_2,nrows=3))

    with st.container():
        st.subheader("⚙️ 参数设置")
        with st.expander("🔧 设置参数",expanded=True):
            cols = st.columns(2)
            with cols[0]:
                selected_date = ['一周','二周','三周','四周','五周','六周','七周','八周','九周','十周',
                                '十一周','十二周','十三周','十四周','十五周','十六周','十七周','十八周',
                                '一月','二月','三月','四月','五月','六月','七月','八月','九月','十月',
                                '十一月','十二月']
                date = st.selectbox("选择周/月", selected_date)
            with cols[1]:
                calendar = st.date_input("选择制表日期")
        
        if st.button("生成考勤通报",use_container_width=True):
            if uploaded_file_1 and uploaded_file_2:
                with st.spinner('正在处理数据，请稍候...'):
                    try:
                        # 为了避免多线程冲突，需要初始化COM组件
                        import pythoncom
                        pythoncom.CoInitialize()

                        year = calendar.strftime('%Y')
                        month = calendar.strftime('%m')
                        day = calendar.strftime('%d')
                        data = pd.read_excel(uploaded_file_1)


                        output_folder_3 = f"第{date}"
                        os.makedirs(output_folder_3, exist_ok=True)

                        # 保存原始数据
                        file1 = os.path.join(output_folder_3, "原始数据.xlsx")
                        file2 = os.path.join(output_folder_3, f"计算机科学与技术学院第{date}上课啦考勤明细.xlsx")
                        with open(file1, 'wb') as f:
                            f.write(uploaded_file_1.getvalue())
                        with open(file2, 'wb') as f:
                            f.write(uploaded_file_2.getvalue())
                            
                        process_attendance_files(data, date, year, month, day, output_folder_3)
                        
                        zip_files([output_folder_3], f"{output_folder_3}")
                        st.success("✅ 考勤通报生成成功！")
                        
                        download_zip_path = f'{output_folder_3}.zip'
                        with open(download_zip_path, 'rb') as f:
                            st.download_button(
                                label='📥 下载考勤通报',
                                data=f,
                                file_name=download_zip_path,
                                mime='application/zip',
                                use_container_width=True,
                                on_click=lambda: delete_files_and_folders([output_folder_3, download_zip_path])
                            )
                    except Exception as e:
                        st.warning(f"⚠️ 处理过程中出现错误：{str(e)}")
            else:
                st.warning("⚠️ 请先上传所有必需文件")

with tab2:
    st.subheader("📤 上传学期考勤数据")
    with st.expander("数据上传（点击展开/收起）",expanded=True):
        # 文件上传组件
        uploaded_file_attendance = st.file_uploader(
            "请选择考勤数据文件（支持.xls/.xlsx）",
            type=["xls", "xlsx"],
            key="attendance_upload"
        )
        # 数据预览
        if uploaded_file_attendance is not None:
            st.dataframe(pd.read_excel(uploaded_file_attendance,nrows=3))

    with st.container():
        st.subheader("⚙️ 参数设置")
        with st.form("confirm_sheet_settings"):
            cols = st.columns(2)
            with cols[0]:
                study_year_options = ["2024-2025", "2025-2026", "2026-2027", "2027-2028"]
                study_year = st.selectbox("选择学年", study_year_options)
            with cols[1]:
                custom_study_year = st.text_input("或手动输入学年（格式示例：2024-2025）", "",help="注意：手动输入后将覆盖右边选择")
                if custom_study_year:
                    study_year = custom_study_year

            semester_options = ["第一学期", "第二学期"]
            semester = st.selectbox("选择学期", semester_options)

            grade_class_cols = st.columns(2)
            with grade_class_cols[0]:
                grade = st.text_input("输入年级范围，如22-24（必填）", "")
            with grade_class_cols[1]:
                max_classes = st.number_input("最大班级数量（必填）", 1, 10, 5, help="如最大班级后缀为05，则输入5")
            
            date_cols = st.columns(2)
            with date_cols[0]:
                start_date = st.date_input("学期开始日期")
            with date_cols[1]:
                end_date = st.date_input("学期结束日期")
            
            if grade:
                pattern = re.compile(r"^\d{2}\s*-\s*\d{2}$")
                if pattern.match(grade):
                    try:
                        start_grade, end_grade = [int(x.strip()) for x in grade.split('-')]
                        if start_grade > end_grade:
                            st.error("⚠️ 年级范围输入错误：请按从小到大的顺序输入")
                        else:
                            grade_list = [str(g) for g in range(start_grade, end_grade+1)]
                            class_list = [f"{g}{str(i).zfill(2)}" for g in grade_list for i in range(1, max_classes+1)]
                    except Exception as e:
                        st.error(f"❌ 年级转换错误：{str(e)}")
                else:
                    st.error("⚠️ 格式错误：请按示例格式输入（如22-24）")

            submit_button_clicked = st.form_submit_button("🎈 提交表单",use_container_width=True)

        process_button_clicked = st.button("生成确认签字表",use_container_width=True)
        if process_button_clicked:
            if uploaded_file_attendance is not None:
                with st.spinner('正在生成文件，请稍候...'):
                    try:
                        data_attendance = pd.read_excel(uploaded_file_attendance)
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

                        process_confirm_sheets(data_attendance, study_year, semester, 
                                            start_year, start_month, start_day, end_year, end_month, end_day,
                                            output_folder_1, output_folder_2, class_list)

                        zip_files([output_folder_1, output_folder_2], "确认签字表+汇总表")
                        st.success("✅ 文件生成成功！")

                        download_zip_path = '确认签字表+汇总表.zip'
                        with open(download_zip_path, 'rb') as f:
                            st.download_button(
                                label='📥 下载生成文件包',
                                data=f,
                                file_name=download_zip_path,
                                mime='application/zip',
                                use_container_width=True,
                                on_click=lambda: delete_files_and_folders([output_folder_1, output_folder_2, download_zip_path])
                            )

                    except Exception as e:
                        st.warning(f"⚠️ 生成过程中出现错误，请检查表单是否填写完整并提交：{str(e)}")
            else:
                st.warning("⚠️ 请先上传考勤数据文件")
