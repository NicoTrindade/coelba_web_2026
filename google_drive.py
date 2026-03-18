import pandas as pd
from io import BytesIO
import streamlit as st
from googleapiclient.discovery import build
from googleapiclient.http import MediaInMemoryUpload
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/drive']


def get_service():
    credentials = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPES
    )
    return build('drive', 'v3', credentials=credentials)


def upload_file(service, file_bytes, file_name, folder_id, mime_type):

    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }

    media = MediaInMemoryUpload(file_bytes, mimetype=mime_type)

    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    file_id = file.get('id')

    service.permissions().create(
        fileId=file_id,
        body={'role': 'reader', 'type': 'anyone'}
    ).execute()

    return f"https://drive.google.com/file/d/{file_id}/view"


def process_csv_and_excel(csv_bytes, file_name, csv_folder, excel_folder):

    service = get_service()

    # 🔹 Upload CSV
    csv_link = upload_file(
        service,
        csv_bytes,
        file_name,
        csv_folder,
        "text/csv"
    )

    # 🔹 Converter para DataFrame
    df = pd.read_csv(BytesIO(csv_bytes))

    # 🔹 Criar Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    excel_bytes = output.getvalue()
    excel_name = file_name.replace(".csv", ".xlsx")

    # 🔹 Upload Excel
    excel_link = upload_file(
        service,
        excel_bytes,
        excel_name,
        excel_folder,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    return csv_link, excel_link