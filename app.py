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
        start_time = time.time()  # 記錄開始時間

        # 定義各題庫抽取題數分配（總共 50 題）
        question_distribution = [8, 8, 8, 8, 9, 9]  # 可根據需求自行調整

        # 使用亂數產生各題庫的難題數量，範圍 1~5，且 1 與 5 出現的機率較低
        def generate_random_hard_distribution():
            possibilities = [1, 2, 3, 4, 5]
            weights = [1, 2, 2, 2, 1]  # 1與5的權重較低
            return random.choices(possibilities, weights=weights, k=6)
        
        hard_distribution = generate_random_hard_distribution()
        st.write("本次亂數生成的難題分配：", hard_distribution)

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
            header_run.font.name = '標楷體'
            header_run.font.size = Pt(20)
            header_run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            header_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            # 添加考試信息
            exam_info_para = doc.add_paragraph("選擇題：100％（共50題，每題2分）")
            exam_info_para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            for run in exam_info_para.runs:
                run.font.name = '標楷體'
                run.font.size = Pt(16)
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

            question_number = 1  # 全卷題號起始值
            # 初始化難度統計，這部分一定要保留
            difficulty_counts = {'難': 0, '中': 0, '易': 0}

            # 依題庫抽題
            for i, file in enumerate(uploaded_files):
                df = pd.read_excel(file)
                total_needed = question_distribution[i]
                desired_hard = hard_distribution[i]
                # 設定隨機種子，以區分 A 卷與 B 卷，並結合題庫編號
                random_seed = (1 if paper_type == "A卷" else 2) + i

                # 先從該題庫中抽取難題
                df_hard = df[df.iloc[:, 1].str.contains('（難）', na=False)]
                n_hard_available = len(df_hard)
                n_hard_to_select = min(desired_hard, total_needed, n_hard_available)
                if n_hard_to_select > 0:
                    selected_hard = df_hard.sample(n=n_hard_to_select, random_state=random_seed)
                else:
                    selected_hard = pd.DataFrame(columns=df.columns)

                # 剩餘題數從非難題中抽取
                remaining = total_needed - n_hard_to_select
                df_nonhard = df[~df.index.isin(df_hard.index)]
                n_nonhard_available = len(df_nonhard)
                n_nonhard_to_select = min(remaining, n_nonhard_available)
                if n_nonhard_to_select > 0:
                    selected_nonhard = df_nonhard.sample(n=n_nonhard_to_select, random_state=random_seed)
                else:
                    selected_nonhard = pd.DataFrame(columns=df.columns)

                # 合併難題與非難題並打亂順序
                selected_questions = pd.concat([selected_hard, selected_nonhard])
                selected_questions = selected_questions.sample(frac=1, random_state=random_seed).reset_index(drop=True)

                # 將該題庫的抽取題目依序加入文件
                for _, row in selected_questions.iterrows():
                    question_text = f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}"
                    question_para = doc.add_paragraph(question_text)
                    
                    # 段落格式設定
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
                    
                    # 更新難度統計：若題目文字中含有「（難）」，則計入難題；否則依據是否含有「（中）」計算中或易題
                    if '（難）' in row.iloc[1]:
                        difficulty_counts['難'] += 1
                    elif '（中）' in row.iloc[1]:
                        difficulty_counts['中'] += 1
                    else:
                        difficulty_counts['易'] += 1

                    question_number += 1

            # 添加難度統計（此段不可刪除）
            summary_text = f"難：{difficulty_counts['難']}，中：{difficulty_counts['中']}，易：{difficulty_counts['易']}"
            summary_para = doc.add_paragraph(summary_text)
            summary_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

            # 保存文件到內存
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            st.session_state.exam_papers[paper_type] = buffer.getvalue()

        end_time = time.time()  # 記錄結束時間
        elapsed_time = end_time - start_time
        st.success(f"🎉 試卷生成完成！耗時：{elapsed_time:.2f} 秒")

# 顯示下載按鈕
if "exam_papers" in st.session_state and st.session_state.exam_papers:
    st.markdown("## 📥 下載試卷")
    for paper_type, file_data in st.session_state.exam_papers.items():
        st.download_button(
            label=f"下載 {paper_type}",
            data=file_data,
            file_name=f"{class_name}_{exam_type}_{subject}_{paper_type}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
