import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
import random
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account

# 主題設定
st.set_page_config(page_title="試卷生成器", page_icon="📄", layout="wide")

# Google Drive 資料夾 ID
ROOT_FOLDER_ID = '17Bcgo8ZeHz0yVhfIxBk7L2wzoiZcyoXt'
SUBJECT_MAPPING = {
    "法律": "法律",
    "專業": "專業"
}

# 建立 Google Drive API 服務
def create_drive_service():
    service_account_info = st.secrets["service_account_json"]
    credentials = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    return build('drive', 'v3', credentials=credentials)

# 遞迴列出指定資料夾及其子資料夾內的所有檔案
def list_files_recursively(service, folder_id):
    all_files = []
    folders_to_process = [folder_id]

    while folders_to_process:
        current_folder_id = folders_to_process.pop()
        query = f"'{current_folder_id}' in parents and trashed=false"
        result = service.files().list(q=query, fields='files(id, name, mimeType)').execute()
        files = result.get('files', [])

        for file in files:
            if file['mimeType'] == 'application/vnd.google-apps.folder':
                folders_to_process.append(file['id'])
            else:
                all_files.append(file)

    return all_files

# 下載檔案為二進位格式
def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return fh

# 列出所有題庫
def display_topics_selection(service, subject_folder_id):
    files = list_files_recursively(service, subject_folder_id)
    topics = {file['name']: file['id'] for file in files if file['mimeType'] == 'application/vnd.google-apps.folder'}
    selected_topics = st.multiselect("選擇題庫", list(topics.keys()))

    if len(selected_topics) != 6:
        st.warning("請選擇 6 個題庫來生成試卷！")
        return None

    if st.button("生成考卷"):
        return {topic: topics[topic] for topic in selected_topics}

# 生成試卷
def generate_exam(selected_topics, service, class_name, exam_type, subject):
    exam_papers = {}

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
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(16)

        question_number = 1
        difficulty_counts = {'難': 0, '中': 0, '易': 0}

        for topic, topic_id in selected_topics.items():
            files = list_files_recursively(service, topic_id)
            excel_files = [file for file in files if file['mimeType'] in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel']]

            for file in excel_files:
                file_content = download_file(service, file['id'])
                df = pd.read_excel(file_content, engine='openpyxl')
                random.seed(1 if paper_type == "A卷" else 2)
                selected_rows = df.sample(n=min(10, len(df)))

                for _, row in selected_rows.iterrows():
                    difficulty_counts['難' if '（難）' in row.iloc[1] else '中' if '（中）' in row.iloc[1] else '易'] += 1
                    question_text = f"（{row.iloc[0]}）{question_number}、{row.iloc[1]}"
                    question_para = doc.add_paragraph(question_text)

                    # 段落格式設置
                    paragraph_format = question_para.paragraph_format
                    paragraph_format.left_indent = Cm(0)
                    paragraph_format.right_indent = Cm(0)
                    paragraph_format.hanging_indent = Pt(4 * 0.35)
                    paragraph_format.space_after = Pt(0)
                    paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                    for run in question_para.runs:
                        run.font.name = '標楷體'
                        run.font.size = Pt(16)
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

                    question_number += 1

        # 添加難度統計
        summary_text = f"難：{difficulty_counts['難']}，中：{difficulty_counts['中']}，易：{difficulty_counts['易']}"
        summary_para = doc.add_paragraph(summary_text)
        summary_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        exam_papers[paper_type] = buffer.getvalue()

    return exam_papers

# 主程式
service = create_drive_service()
st.markdown("## 📋 基本設定")
class_name = st.text_input("班級名稱", value="113-X", help="請輸入班級名稱，例如：113-1")
exam_type = st.selectbox("考試類型", ["期中", "期末"], help="選擇期中或期末考試")
subject = st.selectbox("科目", ["", "法律", "專業"], help="選擇科目類型")

if subject:
    subject_folder_name = SUBJECT_MAPPING[subject]
    files = list_files_recursively(service, ROOT_FOLDER_ID)
    subject_folder_id = next((file['id'] for file in files if file['name'] == subject_folder_name), None)

    if subject_folder_id:
        selected_topics = display_topics_selection(service, subject_folder_id)

        if selected_topics:
            st.info("正在生成試卷，請稍候...")
            exam_papers = generate_exam(selected_topics, service, class_name, exam_type, subject)
            st.success("試卷生成完成！")

            for paper_type, file_data in exam_papers.items():
                st.download_button(
                    label=f"下載 {paper_type}",
                    data=file_data,
                    file_name=f"{class_name}_{exam_type}_{subject}_{paper_type}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
