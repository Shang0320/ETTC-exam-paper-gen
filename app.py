import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ORIENT
import random
import io

# Web UI 標題
st.title("考試試卷生成工具")

# 使用者輸入基本信息
class_name = st.text_input("請輸入班級名稱（如113-1）", value="113-X")
exam_type = st.selectbox("請選擇考試類型", ["期中", "期末"])
subject = st.selectbox("請選擇科目", ["法律", "專業"])

# 上傳題庫檔案
st.markdown("### 上傳題庫檔案（請選擇6個Excel檔案）")
uploaded_files = st.file_uploader("請上傳6個Excel檔案（題庫）", accept_multiple_files=True, type=["xlsx"])

if uploaded_files and len(uploaded_files) == 6:
    st.success(f"已成功上傳 {len(uploaded_files)} 個檔案！")

    # 開始生成試卷
    if st.button("生成試卷"):
        for paper_type in ["A卷", "B卷"]:
            doc = Document()

            # 設置頁面大小與邊距
            section = doc.sections[-1]
            section.page_height, section.page_width = Cm(42.0), Cm(29.7)
            section.orientation = WD_ORIENT.LANDSCAPE
            section.top_margin = section.bottom_margin = Cm(1.5 / 2.54)
            section.left_margin = section.right_margin = Cm(2 / 2.54)

            # 添加標題
            header_para = doc.add_paragraph()
            header_run = header_para.add_run(f"海巡署教育訓練測考中心{class_name}梯志願士兵司法警察專長班{exam_type}測驗階段考試（{subject}{paper_type}）")
            header_run.font.name, header_run.font.size = '標楷體', Pt(20)
            header_run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            header_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            # 添加考試信息
            exam_info_para = doc.add_paragraph("選擇題：100％（共50題，每題2分）")
            exam_info_para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            for run in exam_info_para.runs:
                run.font.name = '標楷體'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
                run.font.size = Pt(16)

            difficulty_counts = {'難': 0, '中': 0, '易': 0}
            question_number = 1

            # 處理題庫檔案
            for i, file in enumerate(uploaded_files):
                df = pd.read_excel(file)
                random_seed = 1 if paper_type == "A卷" else 2
                questions_to_select = 10 if i == len(uploaded_files) - 1 else 8
                selected_rows = df.sample(n=questions_to_select, random_state=random_seed)

                for _, row in selected_rows.iterrows():
                    difficulty_counts['難' if '（難）' in row.iloc[1] else '中' if '（中）' in row.iloc[1] else '易'] += 1
                    question_para = doc.add_paragraph(f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}")
                    question_para.paragraph_format.left_indent = Cm(1)
                    question_para.paragraph_format.first_line_indent = Cm(-1)
                    for run in question_para.runs:
                        run.font.name, run.font.size = '標楷體', Pt(16)
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

                    question_number += 1

            # 添加難度統計
            summary_text = f"難：{difficulty_counts['難']}，中：{difficulty_counts['中']}，易：{difficulty_counts['易']}"
            doc.add_paragraph(summary_text)

            # 保存試卷
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            # 提供下載連結
            filename = f"{class_name}_{exam_type}_{subject}_{paper_type}.docx"
            st.download_button(label=f"下載 {paper_type}", data=buffer, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
