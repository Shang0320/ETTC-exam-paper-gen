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

# ... (頁面設置保持不變)

# 生成試卷函數
def generate_paper(paper_type, question_banks, num_hard_questions):
    doc = Document()
    # ... (頁面設置和標題保持不變)

    random.seed(int(time.time()) if paper_type == "A卷" else int(time.time() + 1))
    difficulty_counts = {'難': 0, '中': 0, '易': 0}  # 使用半形逗號
    question_number = 1
    questions_per_file = [8, 8, 8, 8, 8, 10]  # 每個檔案的總抽題數

    # 計算此卷的難題數量
    total_hard = sum(len(bank[bank.iloc[:, 1].str.contains('(難)', na=False) & ~bank['selected']]) for bank in question_banks)
    hard_for_this_paper = min(num_hard_questions, total_hard // 2 if paper_type == "A卷" else total_hard)

    # 基準難題分配比例 [2, 3, 3, 1, 3, 3]，總和 = 15
    base_hard_pattern = [2, 3, 3, 1, 3, 3]
    base_total = sum(base_hard_pattern)
    
    # 動態計算每個檔案的難題數
    hard_per_file = []
    for i in range(6):
        # 按比例調整
        ratio = base_hard_pattern[i] / base_total
        calculated_hard = int(hard_for_this_paper * ratio)
        # 限制不超過該檔案總抽題數和可用難題數
        available_hard = len(question_banks[i][question_banks[i].iloc[:, 1].str.contains('(難)', na=False) & ~question_banks[i]['selected']])
        hard_per_file.append(min(calculated_hard, questions_per_file[i], available_hard))
    
    # 調整總和至 hard_for_this_paper
    current_total = sum(hard_per_file)
    if current_total < hard_for_this_paper:
        remaining = hard_for_this_paper - current_total
        for i in range(6):
            if remaining == 0:
                break
            available_hard = len(question_banks[i][question_banks[i]。iloc[:, 1]。str。contains('(難)', na=False) & ~question_banks[i]['selected']])
            max_additional = min(questions_per_file[i], available_hard) - hard_per_file[i]
            additional = min(remaining, max_additional)
            hard_per_file[i] += additional
            remaining -= additional

    # 抽取難題
    for i, bank in enumerate(question_banks):
        hard_questions = bank[bank.iloc[:, 1]。str。contains('(難)', na=False) & ~bank['selected']]
        if hard_per_file[i] > 0 和 not hard_questions.empty:
            selected_hard = hard_questions.sample(n=min(hard_per_file[i]， len(hard_questions)))
            for _, row in selected_hard.iterrows():
                bank.loc[row.name， 'selected'] = True
                difficulty_counts['難'] += 1
                doc.add_paragraph(f"({row.iloc[0]}){question_number}，{row.iloc[1]}")
                question_number += 1

    # 補充中、易題至指定數量
    for i, bank in enumerate(question_banks):
        remaining_to_draw = questions_per_file[i] - hard_per_file[i]
        available = bank[~bank['selected']]
        if len(available) < remaining_to_draw:
            st.error(f"{paper_type} 生成失敗: 檔案 {i+1} 剩餘題目不足!")
            return None
        selected = available.sample(n=remaining_to_draw)
        for _, row in selected.iterrows():
            bank.loc[row.name， 'selected'] = True
            difficulty = '難' if '(難)' in row.iloc[1] else ('中' if '(中)' in row.iloc[1] else '易')
            difficulty_counts[difficulty] += 1
            doc.add_paragraph(f"({row.iloc[0]}){question_number}，{row.iloc[1]}")
            question_number += 1

    doc.add_paragraph(f"難:{difficulty_counts['難']},中:{difficulty_counts['中']},易:{difficulty_counts['易']}")
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# 主程式
if uploaded_files 和 len(uploaded_files) == 6:
    question_banks = [pd.read_excel(file) for file in uploaded_files]
    for i, bank in enumerate(question_banks):
        bank['selected'] = False
        min_required = 16 if i < 5 else 20
        if len(bank) < min_required:
            st.error(f"檔案 {i+1} 題目數 ({len(bank)}) 不足, 至少需要 {min_required} 題!")
            break
    else:
        total_hard = sum(len(bank[bank.iloc[:, 1]。str。contains('(難)', na=False)]) for bank in question_banks)
        if total_hard < num_hard_questions:
            st.warning(f"總難題數 ({total_hard}) 小於需求 ({num_hard_questions}), 將按比例分配至 A,B 卷.")
        
        if st.button("✨ 開始生成試卷"):
            with st.spinner("正在生成試卷, 請稍候..."):
                st.session_state。exam_papers["A卷"] = generate_paper("A卷", question_banks, num_hard_questions)
                st.session_state。exam_papers["B卷"] = generate_paper("B卷", question_banks, num_hard_questions)
            st.success("🎉 試卷生成完成!")

# ... (下載按鈕保持不變)
