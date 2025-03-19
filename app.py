import streamlit as st
import pandas as pd
from docx import Document
import random
import time
import io

# ... (其他導入和頁面設置保持不變)

# 生成試卷的函數
def generate_exam_paper(class_name, exam_type, subject, num_hard_questions, uploaded_files, paper_type):
    doc = Document()
    # ... (頁面設置和標題部分保持不變)

    # 合併所有題庫
    all_questions = pd.DataFrame()
    for file in uploaded_files:
        df = pd.read_excel(file)
        all_questions = pd.concat([all_questions, df], ignore_index=True)

    # 設置動態隨機種子
    random.seed(int(time.time()))  # 使用當前時間戳作為種子，每次不同

    difficulty_counts = {'難': 0， '中': 0， '易': 0}
    question_number = 1
    total_questions = 0

    # 優先抽取難題
    hard_questions = all_questions[all_questions.iloc[:, 1]。str。contains('（難）', na=False)]
    remaining_hard_questions = min(num_hard_questions, len(hard_questions))
    if remaining_hard_questions > 0:
        selected_hard = hard_questions.sample(n=remaining_hard_questions)
        for _, row in selected_hard.iterrows():
            difficulty_counts['難'] += 1
            question_text = f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}"
            question_para = doc.add_paragraph(question_text)
            # ... (段落格式設置保持不變)
            question_number += 1
            total_questions += 1

    # 從剩餘題目中抽取其他題目
    remaining_questions = 50 - total_questions
    other_questions = all_questions[~all_questions.index。isin(hard_questions.index)]
    if remaining_questions > 0 和 not other_questions.empty:
        selected_other = other_questions.sample(n=min(remaining_questions, len(other_questions)))
        for _, row in selected_other.iterrows():
            difficulty = '中' if '（中）' in row.iloc[1] else '易'
            difficulty_counts[difficulty] += 1
            question_text = f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}"
            question_para = doc.add_paragraph(question_text)
            # ... (段落格式設置保持不變)
            question_number += 1
            total_questions += 1

    # 添加難度統計
    summary_text = f"難：{difficulty_counts['難']}，中：{difficulty_counts['中']}，易：{difficulty_counts['易']}"
    doc.add_paragraph(summary_text)

    # 保存到內存
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# 主程式部分
if uploaded_files 和 len(uploaded_files) == 6:
    if st.button("✨ 開始生成試卷"):
        with st.spinner("正在生成試卷，請稍候..."):
            for paper_type in ["A卷"， "B卷"]:
                file_data = generate_exam_paper(class_name, exam_type, subject, num_hard_questions, uploaded_files, paper_type)
                st.session_state。exam_papers[paper_type] = file_data
        st.success("🎉 試卷生成完成！")

# ... (下載按鈕部分保持不變)
