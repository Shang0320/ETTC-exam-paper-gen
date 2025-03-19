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
按照以下步驟完成試卷生成：
1. 填寫基本資訊。
2. 上傳題庫檔案（6 個 Excel 文件）。
3. 點擊生成按鈕，下載標準化的 A 卷與 B 卷。
4. 題庫下載點－ https://drive.google.com/drive/folders/17Bcgo8ZeHz0yVhfIxBk7L2wzoiZcyoXt?usp=sharing
""")

# 分隔線
st.divider()

# 主體內容佈局
col1, col2 = st.columns([1, 2])

with col1:
    st.markdown("## 📋 基本設定")
    class_name = st.text_input("班級名稱", value="113-X", help="請輸入班級名稱，例如：113-1")
    exam_type = st.selectbox("考試類型", ["期中", "期末"], help="選擇期中或期末考試")
    subject = st.selectbox("科目", ["法律", "專業"], help="選擇科目類型")
    num_hard_questions = st.number_input("選擇難題數量", min_value=0, max_value=50, value=10, step=1, help="設定生成試卷中難題的數量")  # 修正語法

with col2:
    st.markdown("## 📤 上傳題庫")
    st.markdown("請上傳 **6 個 Excel 文件**，每個文件代表一個題庫")
    uploaded_files = st.file_uploader("上傳題庫檔案（最多 6 個）", accept_multiple_files=True, type=["xlsx"])

if uploaded_files:
    st.success(f"✅ 已成功上傳 {len(uploaded_files)} 個檔案！")
    if len(uploaded_files) != 6:
        st.warning("⚠️ 請上傳 6 個文件，否則無法生成完整試卷。")

# 初始化 Session State 中的緩存
if "exam_papers" not in st.session_state:
    st.session_state.exam_papers = {}

# 分隔線
st.divider()

# 生成試卷函數
def generate_paper(paper_type, question_banks, num_hard_questions):
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
    header_run.font.name = '標楷體'
    header_run.font.size = Pt(20)
    header_run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    header_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # 添加考試信息
    exam_info_para = doc.add_paragraph("選擇題：100％（共50題，每題2分）")
    exam_info_para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in exam_info_para.runs:
        run.font.name = '標楷體'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
        run.font.size = Pt(16)

    random.seed(int(time.time()) if paper_type == "A卷" else int(time.time() + 1))
    difficulty_counts = {'難': 0, '中': 0, '易': 0}
    question_number = 1
    questions_per_file = [8, 8, 8, 8, 8, 10]  # 每個檔案的總抽題數

    # 計算此卷的難題數量
    total_hard = sum(len(bank[bank.iloc[:, 1].str.contains('（難）', na=False) & ~bank['selected']]) for bank in question_banks)
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
        available_hard = len(question_banks[i][question_banks[i].iloc[:, 1].str.contains('（難）', na=False) & ~question_banks[i]['selected']])
        hard_per_file.append(min(calculated_hard, questions_per_file[i], available_hard))
    
    # 調整總和至 hard_for_this_paper
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

    # 抽取難題
    for i, bank in enumerate(question_banks):
        hard_questions = bank[bank.iloc[:, 1].str.contains('（難）', na=False) & ~bank['selected']]
        if hard_per_file[i] > 0 and not hard_questions.empty:
            selected_hard = hard_questions.sample(n=min(hard_per_file[i], len(hard_questions)))
            for _, row in selected_hard.iterrows():
                bank.loc[row.name, 'selected'] = True
                difficulty_counts['難'] += 1
                question_para = doc.add_paragraph(f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}")
                paragraph_format = question_para.paragraph_format
                paragraph_format.left_indent = Cm(0)
                paragraph_format.right_indent = Cm(0)
                paragraph_format.hanging_indent = Pt(8 * 0.35)
                paragraph_format.space_after = Pt(0)
                paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                for run in question_para.runs:
                    run.font.name = '標楷體'
                    run.font.size = Pt(16)
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
                question_number += 1

    # 補充中、易題至指定數量
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
            question_para = doc.add_paragraph(f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}")
            paragraph_format = question_para.paragraph_format
            paragraph_format.left_indent = Cm(0)
            paragraph_format.right_indent = Cm(0)
            paragraph_format.hanging_indent = Pt(8 * 0.35)
            paragraph_format.space_after = Pt(0)
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            for run in question_para.runs:
                run.font.name = '標楷體'
                run.font.size = Pt(16)
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            question_number += 1

    # 添加難度統計
    summary_para = doc.add_paragraph(f"難：{difficulty_counts['難']}，中：{difficulty_counts['中']}，易：{difficulty_counts['易']}")
    summary_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # 保存到內存
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# 主程式
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
            start_time = time.time()
            with st.spinner("正在生成試卷，請稍候..."):
                st.session_state.exam_papers["A卷"] = generate_paper("A卷", question_banks, num_hard_questions)
                st.session_state.exam_papers["B卷"] = generate_paper("B卷", question_banks, num_hard_questions)
            end_time = time.time()
            elapsed_time = end_time - start_time
            st.success(f"🎉 試卷生成完成！耗時：{elapsed_time:.2f} 秒")

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
