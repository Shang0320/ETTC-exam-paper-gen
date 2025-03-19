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
    num_hard_questions = st.number_input("選擇難題數量", min_value=0, max_value=50, value=10, step=1, help="設定生成試卷中難題的數量")

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

if uploaded_files and len(uploaded_files) == 6:
    if st.button("✨ 開始生成試卷"):
        with st.spinner("正在生成試卷，請稍候..."):
            for paper_type in ["A卷", "B卷"]:
                doc = Document()
                # ... (頁面設置和標題保持不變)

                random.seed(int(time.time()))  # 動態種子
                difficulty_counts = {'難': 0, '中': 0, '易': 0}
                question_number = 1
                total_questions = 0

                # 平均分配：每個檔案抽 8-9 題
                base_questions_per_file = 50 // 6  # 每個檔案基本抽 8 題
                extra_questions = 50 % 6  # 剩餘 2 題分配給前 2 個檔案

                selected_questions = []
                for i, file in enumerate(uploaded_files):
                    df = pd.read_excel(file)
                    questions_to_draw = base_questions_per_file + (1 if i < extra_questions else 0)
                    
                    # 優先抽難題
                    hard_questions = df[df.iloc[:, 1].str.contains('（難）', na=False)]
                    hard_to_draw = min(num_hard_questions - difficulty_counts['難'], len(hard_questions))
                    if hard_to_draw > 0:
                        selected_hard = hard_questions.sample(n=hard_to_draw)
                        selected_questions.extend(selected_hard.iterrows())
                        difficulty_counts['難'] += hard_to_draw

                    # 補充其他題目
                    other_questions = df[~df.index.isin(hard_questions.index)]
                    remaining_to_draw = questions_to_draw - hard_to_draw
                    if remaining_to_draw > 0 and not other_questions.empty:
                        selected_other = other_questions.sample(n=min(remaining_to_draw, len(other_questions)))
                        selected_questions.extend(selected_other.iterrows())

                # 寫入試卷
                for _, row in selected_questions[:50]:  # 確保不超過 50 題
                    difficulty = '難' if '（難）' in row.iloc[1] else ('中' if '（中）' in row.iloc[1] else '易')
                    difficulty_counts[difficulty] += 1
                    question_text = f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}"
                    # ... (段落格式設置)
                    question_number += 1
                    total_questions += 1


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
