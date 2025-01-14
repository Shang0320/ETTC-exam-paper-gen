import streamlit as st
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account
import pandas as pd
import io

# Google Drive 資料夾 ID
FOLDER_ID = '17Bcgo8ZeHz0yVhfIxBk7L2wzoiZcyoXt'

def create_drive_service():
    """以 Service Account 建立 Google Drive API 服務，從 Streamlit Secrets 中讀取憑證。"""
    service_account_info = st.secrets["service_account_json"]
    credentials = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    return build('drive', 'v3', credentials=credentials)

def list_files_recursively(service, folder_id):
    """遞迴列出指定資料夾及其所有子資料夾內的所有檔案。"""
    all_files = []
    folders_to_process = [folder_id]  # 初始化待處理資料夾清單

    while folders_to_process:
        current_folder_id = folders_to_process.pop()
        query = f"'{current_folder_id}' in parents and trashed=false"
        result = service.files().list(q=query, fields='files(id, name, mimeType)').execute()
        files = result.get('files', [])
        
        for file in files:
            # 如果是子資料夾，加入待處理清單
            if file['mimeType'] == 'application/vnd.google-apps.folder':
                folders_to_process.append(file['id'])
            else:
                all_files.append(file)  # 只加入非資料夾檔案
    
    return all_files

def download_file(service, file_id):
    """從 Google Drive 下載檔案為二進位格式。"""
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return fh

def main():
    st.title("Google Drive 檔案選擇器")

    # 建立 Drive 服務
    service = create_drive_service()

    # 遞迴列出檔案
    files = list_files_recursively(service, FOLDER_ID)
    if not files:
        st.error("該資料夾及其所有子資料夾中沒有任何檔案，或 Service Account 無法讀取。")
        return

    # 過濾 Excel 檔案
    st.write("檔案資訊：")
    for f in files:
        st.write(f"檔案名稱: {f['name']}, MIME 類型: {f['mimeType']} ")

    excel_files = [f for f in files if f['mimeType'] in [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ]]

    file_options = {f['name']: f['id'] for f in excel_files}
    if not file_options:
        st.warning("該資料夾及其所有子資料夾中沒有任何 Excel 檔案。")
        return

    selected_files = st.multiselect("選擇要處理的檔案", options=list(file_options.keys()))

    if st.button("下載並讀取選擇的檔案"):
        if not selected_files:
            st.warning("請至少選擇一個檔案！")
        else:
            for filename in selected_files:
                file_id = file_options[filename]
                file_content = download_file(service, file_id)

                # 使用 Pandas 讀取 Excel 檔案
                df = pd.read_excel(file_content, engine='openpyxl')
                st.write(f"檔案: {filename}")
                st.dataframe(df.head())

if __name__ == "__main__":
    main()