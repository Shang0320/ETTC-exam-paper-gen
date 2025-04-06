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
請依下列步驟完成試卷生成：
1. 填寫基本資訊。
2. 上傳題庫檔案（6 個 Excel 文件）。
3. 點擊生成按鈕，下載 A 卷與 B 卷試卷。
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
    
    # 修改後的顯示答案選項，添加註記
    show_answers = st.checkbox("在試卷上顯示答案[※上簽出題]", value=False)
    
    # 新增列印學生考卷版本功能
    print_student_version = st.checkbox("列印學生考卷版本[※刪除答案與難度]", value=False)

with col2:
    st.markdown("## 📤 上傳題庫")
    st.markdown("請上傳 **6 個 Excel 文件**，每個文件代表一個題庫")
    uploaded_files = st.file_uploader("上傳題庫檔案（最多 6 個）", accept_multiple_files=True, type=["xlsx"])

if uploaded_files:
    st.success(f"✅ 已成功上傳 {len(uploaded_files)} 個檔案！")
    if len(uploaded_files) != 6:
        st.warning("⚠️ 請上傳 6 個文件，否則無法生成完整試卷。")
    else:
        st.subheader("檢查上傳文件頭部（前5行）")
        for i, file in enumerate(uploaded_files):
            df = pd.read_excel(file)
            st.write(f"檔案 {i+1} 的頭部：")
            st.write(df.head())

# 初始化 Session State 中的緩存
if "exam_papers" not in st.session_state:
    st.session_state.exam_papers = {}

# 建立一個全域列表，用來記錄各題庫中 A 卷和 B 卷已抽取題目的原始索引
used_indices = [set() for _ in range(len(uploaded_files))]

# 分隔線
st.divider()

if uploaded_files and len(uploaded_files) == 6:
    if st.button("✨ 開始生成試卷"):
        start_time = time.time()  # 記錄開始時間

        # 各題庫總抽題分配（總題數 50 題）
        total_distribution = [9, 9, 8, 8, 8, 8]

        # A 卷較偏難，設定較高難題數分配
        A_hard_distribution = [4, 3, 3, 3, 3, 3]
        # B 卷較偏易，設定較低難題數分配
        B_hard_distribution = [2, 2, 2, 2, 2, 2]

        # 用於保存選取的題目索引，確保學生版本與標準版本一致
        selected_indices = {"A卷": [], "B卷": []}

        # 定義生成試卷的函式
        def generate_exam(paper_type, total_distribution, hard_distribution, student_version=False):
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
            difficulty_counts = {'難': 0, '中': 0, '易': 0}
            answer_key = []

            # 逐一處理每個題庫
            for i, file in enumerate(uploaded_files):
                file.seek(0)
                
                try:
                    df = pd.read_excel(file)
                    
                    if len(df.columns) < 5:
                        st.error(f"檔案 {i+1} 的列數不足，請確保題庫格式正確！")
                        return None
                    
                    expected_columns = ['序號', '難度', '答案', '題目', '選項1', '選項2', '選項3', '選項4']
                    current_columns = df.columns.tolist()
                    mapping = {}
                    for expected in expected_columns:
                        for current in current_columns:
                            if expected.lower().strip() in current.lower().strip():
                                mapping[current] = expected
                    
                    df = df.rename(columns=mapping)
                    missing = [col for col in expected_columns if col not in df.columns]
                    if missing:
                        st.error(f"檔案 {i+1} 缺少必要欄位：{missing}")
                        return None
                    
                    df = df.dropna(subset=['題目', '答案'])
                    
                    if paper_type == "B卷":
                        df = df[~df.index.isin(used_indices[i])]
                    
                    df['難度'] = df['難度'].astype(str).str.strip()
                    df.loc[~df['難度'].isin(['難', '中', '易']), '難度'] = '中'
                    
                    df['答案'] = df['答案'].astype(str).str.strip()
                    df.loc[~df['答案'].isin(['1', '2', '3', '4']), '答案'] = '1'
                    
                    seed_shuffle = i + (100 if paper_type == "A卷" else 200)
                    df = df.sample(frac=1, random_state=seed_shuffle)

                    total_needed = total_distribution[i]
                    desired_hard = hard_distribution[i]
                    random_seed = (1 if paper_type == "A卷" else 2) + i

                    df_hard = df[df['難度'] == '難']
                    df_medium = df[df['難度'] == '中']
                    df_easy = df[df['難度'] == '易']

                    if paper_type == "A卷":
                        n_hard_to_select = min(desired_hard, len(df_hard))
                        additional_hard = df_hard.sample(n=n_hard_to_select, random_state=random_seed) if n_hard_to_select > 0 else pd.DataFrame()
                        
                        remaining = total_needed - len(additional_hard)
                        df_remaining = df[~df.index.isin(additional_hard.index)]
                        n_remaining_to_select = min(remaining, len(df_remaining))
                        additional_questions = df_remaining.sample(n=n_remaining_to_select, random_state=random_seed) if n_remaining_to_select > 0 else pd.DataFrame()
                        
                        selected_questions = pd.concat([additional_hard, additional_questions]).sample(frac=1, random_state=random_seed)
                        selected_indices["A卷"].extend(selected_questions.index.tolist())
                        used_indices[i].update(selected_questions.index.tolist())
                    else:
                        n_hard_to_select = min(B_hard_distribution[i], len(df_hard))
                        additional_hard = df_hard.sample(n=n_hard_to_select, random_state=random_seed) if n_hard_to_select > 0 else pd.DataFrame()
                        
                        remaining = total_needed - len(additional_hard)
                        df_remaining = df[~df.index.isin(additional_hard.index)]
                        n_easy_to_select = min(remaining, len(df_easy))
                        additional_easy = df_easy.sample(n=n_easy_to_select, random_state=random_seed) if n_easy_to_select > 0 else pd.DataFrame()
                        
                        remaining_after_easy = remaining - len(additional_easy)
                        df_final_remaining = df_remaining[~df_remaining.index.isin(additional_easy.index)]
                        n_final_to_select = min(remaining_after_easy, len(df_final_remaining))
                        additional_questions = df_final_remaining.sample(n=n_final_to_select, random_state=random_seed) if n_final_to_select > 0 else pd.DataFrame()
                        
                        selected_questions = pd.concat([additional_hard, additional_easy, additional_questions]).sample(frac=1, random_state=random_seed)
                        selected_indices["B卷"].extend(selected_questions.index.tolist())

                    for _, row in selected_questions.iterrows():
                        answer = row['答案']
                        question_text = row['題目']
                        options = [row['選項1'], row['選項2'], row['選項3'], row['選項4']]
                        difficulty = row['難度']
                        
                        cleaned_options = [str(opt).strip() for opt in options]
                        options_text = "".join([f"({i+1}){opt}" for i, opt in enumerate(cleaned_options)])

                        answer_key.append((question_number, answer))
                        
                        if print_student_version:
                            # 學生考卷版本：移除答案和難易度，只保留題目和選項，題號前括號為空
                            question_para = doc.add_paragraph(f"(){question_number}、{question_text} {options_text}")
                        else:
                            if show_answers:
                                question_para = doc.add_paragraph(f"（{answer}）{question_number}、{question_text} {options_text}（{difficulty}）")
                            else:
                                question_para = doc.add_paragraph(f"{question_number}、{question_text} {options_text}（{difficulty}）")
                        
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
                        
                        if not print_student_version:
                            difficulty_counts[difficulty] += 1
                        question_number += 1

                except Exception as e:
                    st.error(f"處理檔案 {i+1} 時發生錯誤: {str(e)}")
                    return None

            if not print_student_version:
                summary_text = f"難：{difficulty_counts['難']}，中：{difficulty_counts['中']}，易：{difficulty_counts['易']}"
                summary_para = doc.add_paragraph(summary_text)
                summary_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            
            if not print_student_version and not show_answers:
                doc.add_page_break()
                answer_title = doc.add_paragraph()
                answer_title.add_run(f"{subject}{paper_type} 答案卷").bold = True
                answer_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                
                answers_per_row = 5
                num_rows = (len(answer_key) + answers_per_row - 1) // answers_per_row
                
                for row in range(num_rows):
                    answer_row = doc.add_paragraph()
                    for col in range(answers_per_row):
                        idx = row * answers_per_row + col
                        if idx < len(answer_key):
                            q_num, ans = answer_key[idx]
                            answer_row.add_run(f"{q_num}. {ans}     ")
                    answer_row.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            return buffer.getvalue(), answer_key

        # 生成標準版本和學生版本
        standard_A, answer_key_A = generate_exam("A卷", total_distribution, A_hard_distribution, student_version=False)
        standard_B, answer_key_B = generate_exam("B卷", total_distribution, B_hard_distribution, student_version=False)

        if print_student_version:
            # 生成學生版本，使用相同的題目索引
            student_A, _ = generate_exam("A卷", total_distribution, A_hard_distribution, student_version=True)
            student_B, _ = generate_exam("B卷", total_distribution, B_hard_distribution, standard_version=False)

            st.session_state.exam_papers["A卷"] = standard_A
            st.session_state.exam_papers["B卷"] = standard_B
            st.session_state.exam_papers["學生A卷"] = student_A
            st.session_state.exam_papers["學生B卷"] = student_B
        else:
            st.session_state.exam_papers["A卷"] = standard_A
            st.session_state.exam_papers["B卷"] = standard_B

        end_time = time.time()
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
