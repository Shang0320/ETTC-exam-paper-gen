import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
import random
import io
import time

# 主題設定
st.set_page_config(page_title="試卷生成器", page_icon="📄", layout="wide")

# 頁面標題與簡介
st.markdown("""
# 📄 志兵班試卷生成器WEB UI
**輕鬆生成專業格式的試卷！**
""")

# 主體內容佈局
col1, col2 = st.columns([1, 2])

with col1:
    st.markdown("## 📋 基本設定")
    class_name = st.text_input("班級名稱", value="113-X")
    exam_type = st.selectbox("考試類型", ["期中", "期末"])
    subject = st.selectbox("科目", ["法律", "專業"])
    num_hard_questions = st.number_input("選擇難題數量", min_value=0, max_value=50, value=10, step=1)

with col2:
    st.markdown("## 📤 上傳題庫")
    uploaded_files = st.file_uploader("上傳題庫檔案（最多 6 個）", accept_multiple_files=True, type=["xlsx"])

# 生成試卷函數
def generate_paper(paper_type, question_banks, num_hard_questions):
    doc = Document()
    # ... (頁面設置和標題保持不變)

    random.seed(int(time.time()) if paper_type == "A卷" else int(time.time() + 1))
    difficulty_counts = {'難': 0, '中': 0, '易': 0}
    question_number = 1
    questions_per_file = [8, 8, 8, 8, 8, 10]

    total_hard = sum(len(bank[bank.iloc[:, 1].str.contains('（難）', na=False) & ~bank['selected']]) for bank in question_banks)
    hard_for_this_paper = min(num_hard_questions, total_hard // 2 if paper_type == "A卷" else total_hard)

    base_hard_pattern = [2, 3, 3, 1, 3, 3]
    base_total = sum(base_hard_pattern)
    
    hard_per_file = []
    for i in range(6):
        ratio = base_hard_pattern[i] / base_total
        calculated_hard = int(hard_for_this_paper * ratio)
        available_hard = len(question_banks[i][question_banks[i].iloc[:, 1].str.contains('（難）', na=False) & ~question_banks[i]['selected']])
        hard_per_file.append(min(calculated_hard, questions_per_file[i], available_hard))
    
    current_total = sum(hard_per_file)
    if current_total < hard_for_this_paper:
        remaining = hard_for_this_paper - current_total
        for i in range(6):
            if remaining == 0:
                break
            available_hard = len(question_banks[i][question_banks[i].iloc[:, 1].str.contains('（難）', na=False) & ~question_banks[i]['selected']])
            max_additional = min(questions_per_file[i], available_hard) - hard_per_file[i]
            additional = min(remaining, max_additional)
            hard_per_file[i] += additional
            remaining -= additional

    for i, bank in enumerate(question_banks):
        hard_questions = bank[bank.iloc[:, 1].str.contains('（難）', na=False) & ~bank['selected']]
        if hard_per_file[i] > 0 and not hard_questions.empty:
            selected_hard = hard_questions.sample(n=min(hard_per_file[i], len(hard_questions)))
            for _, row in selected_hard.iterrows():
                bank.loc[row.name, 'selected'] = True
                difficulty_counts['難'] += 1
                doc.add_paragraph(f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}")
                question_number += 1

    for i, bank in enumerate(question_banks):
        remaining_to_draw = questions_per_file[i] - hard_per_file[i]
        available = bank[~bank['selected']]
        if len(available) < remaining_to_draw:
            st.error(f"{paper_type} 生成失敗：檔案 {i+1} 剩餘題目不足！")
            return None
        selected = available.sample(n=remaining_to_draw)
        for _, row in selected.iterrows():
            bank.loc[row.name, 'selected'] = True
            difficulty = '難' if '（難）' in row.iloc[1] else ('中' if '（中）' in row.iloc[1] else '易')
            difficulty_counts[difficulty] += 1
            doc.add_paragraph(f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}")
            question_number += 1

    doc.add_paragraph(f"難：{difficulty_counts['難']}，中：{difficulty_counts['中']}，易：{difficulty_counts['易']}")
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# 主程式
if 'exam_papers' not in st.session_state:
    st.session_state.exam_papers = {}

if uploaded_files and len(uploaded_files) == 6:
    question_banks = [pd.read_excel(file) for file in uploaded_files]
    for i, bank in enumerate(question_banks):
        bank['selected'] = False
        min_required = 16 if i < 5 else 20
        if len(bank) < min_required:
            st.error(f"檔案 {i+1} 題目數 ({len(bank)}) 不足，至少需要 {min_required} 題！")
            break
    else:
        total_hard = sum(len(bank[bank.iloc[:, 1].str.contains('（難）', na=False)]) for bank in question_banks)
        if total_hard < num_hard_questions:
            st.warning(f"總難題數 ({total_hard}) 小於需求 ({num_hard_questions})，將按比例分配至 A、B 卷。")
        
        if st.button("✨ 開始生成試卷"):
            with st.spinner("正在生成試卷，請稍候..."):
                st.session_state.exam_papers["A卷"] = generate_paper("A卷", question_banks, num_hard_questions)
                st.session_state.exam_papers["B卷"] = generate_paper("B卷", question_banks, num_hard_questions)
            st.success("🎉 試卷生成完成！")

# 下載按鈕
if "exam_papers" in st.session_state and st.session_state.exam_papers:
    st.markdown("## 📥 下載試卷")
    for paper_type, file_data in st.session_state.exam_papers.items():
        st.download_button(
            label=f"下載 {paper_type}",
            data=file_data,
            file_name=f"{class_name}_{exam_type}_{subject}_{paper_type}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
