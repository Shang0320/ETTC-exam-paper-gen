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

# 顯示題庫選項
def display_subject_selection(service, root_folder_id):
    files = list_files_recursively(service, root_folder_id)
    subjects = {file['name']: file['id'] for file in files if file['mimeType'] == 'application/vnd.google-apps.folder'}
    selected_subject = st.radio("選擇大項目", list(subjects.keys()))

    if st.button("下一步"):
        return subjects[selected_subject]

# 列出子資料夾中的科目
def display_topics_selection(service, subject_folder_id):
    files = list_files_recursively(service, subject_folder_id)
    topics = {file['name']: file['id'] for file in files if file['mimeType'] == 'application/vnd.google-apps.folder'}
    selected_topics = st.multiselect("選擇科目", list(topics.keys()))

    if len(selected_topics) != 6:
        st.warning("請選擇 6 個科目來生成試卷！")
        return None

    if st.button("生成考卷"):
        return {topic: topics[topic] for topic in selected_topics}

# 生成試卷
def generate_exam(selected_topics, service):
    exam_papers = {}

    for paper_type in ["A卷", "B卷"]:
        doc = Document()

        for topic, topic_id in selected_topics.items():
            files = list_files_recursively(service, topic_id)
            excel_files = [file for file in files if file['mimeType'] in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel']]

            for file in excel_files:
                file_content = download_file(service, file['id'])
                df = pd.read_excel(file_content, engine='openpyxl')
                random.seed(1 if paper_type == "A卷" else 2)
                selected_rows = df.sample(n=min(10, len(df)))

                for _, row in selected_rows.iterrows():
                    question_text = f"{row.iloc[0]}、{row.iloc[1]}"
                    doc.add_paragraph(question_text)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        exam_papers[paper_type] = buffer.getvalue()

    return exam_papers

# 主程式
service = create_drive_service()
subject_folder_id = display_subject_selection(service, ROOT_FOLDER_ID)

if subject_folder_id:
    selected_topics = display_topics_selection(service, subject_folder_id)

    if selected_topics:
        st.info("正在生成試卷，請稍候...")
        exam_papers = generate_exam(selected_topics, service)
        st.success("試卷生成完成！")

        for paper_type, file_data in exam_papers.items():
            st.download_button(
                label=f"下載 {paper_type}",
                data=file_data,
                file_name=f"{paper_type}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
