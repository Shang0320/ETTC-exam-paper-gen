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

# 主體內容
if uploaded_files and len(uploaded_files) == 6:
    # 將 6 個題庫放入容器並分區
    question_banks = [pd.read_excel(file) for file in uploaded_files]
    
    # 檢查每個題庫的題目數量
    for i, bank in enumerate(question_banks):
        if i < 5 and len(bank) < 16:
            st.error(f"檔案 {i+1} 題目數 ({len(bank)}) 不足，至少需要 16 題！")
            break
        elif i == 5 and len(bank) < 20:
            st.error(f"檔案 6 題目數 ({len(bank)}) 不足，至少需要 20 題！")
            break
    else:  # 如果題庫數量足夠
        # 檢查總難題數量
        total_hard = sum(len(bank[bank.iloc[:, 1].str.contains('（難）', na=False)]) for bank in question_banks)
        if total_hard < num_hard_questions:
            st.warning(f"題庫中困難題目總數 ({total_hard}) 小於需求 ({num_hard_questions})，已調整為 {total_hard} 題。")
            num_hard_questions = total_hard
        
        if st.button("✨ 開始生成試卷"):
            with st.spinner("正在生成試卷，請稍候..."):
                # 初始化已選取標記
                for bank in question_banks:
                    bank['selected'] = False
                
                # 生成 A 卷
                a_paper = Document()
                # ... (頁面設置和標題保持不變)
                
                random.seed(int(time.time()))  # 動態種子
                difficulty_counts_a = {'難': 0, '中': 0, '易': 0}
                question_number = 1
                
                # 抽取困難題目
                hard_questions = pd.concat([bank[bank.iloc[:, 1].str.contains('（難）', na=False) & ~bank['selected']] for bank in question_banks])
                selected_hard_a = hard_questions.sample(n=min(num_hard_questions, len(hard_questions)))
                for _, row in selected_hard_a.iterrows():
                    question_banks[row.name // len(question_banks[0])]['selected'].iloc[row.name % len(question_banks[0])] = True
                    difficulty_counts_a['難'] += 1
                    a_paper.add_paragraph(f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}")
                    question_number += 1
                
                # 從每個檔案抽取指定數量的題目
                questions_per_file = [8， 8， 8， 8， 8， 10]
                for i, bank in enumerate(question_banks):
                    available = bank[~bank['selected']]
                    if len(available) < questions_per_file[i]:
                        st.error(f"檔案 {i+1} 剩餘題目不足！")
                        break
                    selected = available.sample(n=questions_per_file[i])
                    for _, row in selected.iterrows():
                        bank.loc[row.name， 'selected'] = True
                        difficulty = '難' if '（難）' in row.iloc[1] else ('中' if '（中）' in row.iloc[1] else '易')
                        difficulty_counts_a[difficulty] += 1
                        a_paper.add_paragraph(f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}")
                        question_number += 1
                
                a_paper.add_paragraph(f"難：{difficulty_counts_a['難']}，中：{difficulty_counts_a['中']}，易：{difficulty_counts_a['易']}")
                buffer_a = io.BytesIO()
                a_paper.save(buffer_a)
                buffer_a.seek(0)
                st.session_state。exam_papers["A卷"] = buffer_a.getvalue()

                # 生成 B 卷
                b_paper = Document()
                # ... (頁面設置和標題保持不變)
                
                random.seed(int(time.time() + 1))  # 不同種子
                difficulty_counts_b = {'難': 0， '中': 0， '易': 0}
                question_number = 1
                
                # 抽取困難題目
                hard_questions = pd.concat([bank[bank.iloc[:, 1]。str。contains('（難）', na=False) & ~bank['selected']] for bank in question_banks])
                selected_hard_b = hard_questions.sample(n=min(num_hard_questions, len(hard_questions)))
                for _, row in selected_hard_b.iterrows():
                    question_banks[row.name // len(question_banks[0])]['selected']。iloc[row.name % len(question_banks[0])] = True
                    difficulty_counts_b['難'] += 1
                    b_paper.add_paragraph(f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}")
                    question_number += 1
                
                # 從每個檔案抽取指定數量的題目
                for i, bank in enumerate(question_banks):
                    available = bank[~bank['selected']]
                    if len(available) < questions_per_file[i]:
                        st.error(f"檔案 {i+1} 剩餘題目不足！")
                        break
                    selected = available.sample(n=questions_per_file[i])
                    for _, row in selected.iterrows():
                        bank.loc[row.name， 'selected'] = True
                        difficulty = '難' if '（難）' in row.iloc[1] else ('中' if '（中）' in row.iloc[1] else '易')
                        difficulty_counts_b[difficulty] += 1
                        b_paper.add_paragraph(f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}")
                        question_number += 1
                
                b_paper.add_paragraph(f"難：{difficulty_counts_b['難']}，中：{difficulty_counts_b['中']}，易：{difficulty_counts_b['易']}")
                buffer_b = io.BytesIO()
                b_paper.save(buffer_b)
                buffer_b.seek(0)
                st.session_state。exam_papers["B卷"] = buffer_b.getvalue()

            st.success("🎉 試卷生成完成！")

# ... (下載按鈕保持不變)
