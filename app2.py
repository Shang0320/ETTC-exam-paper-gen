import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
import streamlit as st

# 從 Streamlit Secrets 讀取 Service Account 憑證
service_account_info = st.secrets["service_account_json"]

# 初始化 Google Drive API 憑證
credentials = service_account.Credentials.from_service_account_info(
    service_account_info,
    scopes=["https://www.googleapis.com/auth/drive.readonly"]
)

# 建立 Google Drive API 服務
service = build('drive', 'v3', credentials=credentials)

# 測試 API 功能
def list_drive_files():
    results = service.files().list(pageSize=10, fields="files(id, name)").execute()
    items = results.get('files', [])
    if not items:
        st.write("No files found.")
    else:
        st.write("Files:")
        for item in items:
            st.write(f"{item['name']} ({item['id']})")

# Streamlit 頁面顯示
st.title("Google Drive API 測試")
if st.button("列出檔案"):
    list_drive_files()
