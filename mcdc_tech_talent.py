import os
from pathlib import Path

import gspread
import pandas as pd
from src.google_api import create_service
from src.google_api import get_file_id, download_file
from src.process_bls import process_bls


DATA_DIR = 'data'
DB_DIR = 'db'

DATA_PATH = Path() / DATA_DIR
DB_PATH = Path() / DB_DIR

API_KEY = 'key.json'


def download_project_data(secret):
    """ Downloads all remote data files to local project folder"""

    scope = ['https://www.googleapis.com/auth/drive']
    service = create_service(secret, 'drive', 'v3', scope)

    data_folder_id = get_file_id(service,
                                 file_name=DATA_DIR,
                                 mime_type='application/vnd.google-apps.folder')

    sub_folders = ['bls', 'bg', 'ipeds', 'lookups']

    for folder in sub_folders:
        dl_folder_id = get_file_id(service, folder, parent_id=data_folder_id)
        results = service.files().list(
            q="parents in '{}' and trashed=False".format(dl_folder_id),
            fields='files(name, id)').execute()

        download_path = DATA_PATH / folder

        print("Downloading remote files to local project...")
        for file in results['files']:
            download_file(service, file['id'], file['name'], download_path)


def upload_db_data(secret, db_path):
    """ Loads processed data files to remote Google Sheets """

    scope = ['https://www.googleapis.com/auth/drive']
    drive_service = create_service(secret, 'drive', 'v3', scope)
    sheet_service = gspread.service_account('key.json')

    db_folder_id = get_file_id(drive_service,
                               file_name=DB_DIR,
                               mime_type='application/vnd.google-apps.folder')

    files = [f.name for f in os.scandir(DB_PATH) if f.name.endswith('.xlsx')]

    for file in files:
        print("Loading {} to Google Sheets".format(file))
        df = pd.read_excel(db_path / file)

        file_id = get_file_id(drive_service,
                              file[:-5],
                              parent_id=db_folder_id)

        sh = sheet_service.open_by_key(file_id)
        sh.values_clear("'Sheet1'!A:AAA")
        sh.sheet1.update([df.columns.values.tolist()] + df.fillna('').values.tolist())


def main():
    download_project_data(API_KEY)
    process_bls(DATA_PATH, DB_PATH)
    upload_db_data(API_KEY, DB_PATH)


if __name__ == "__main__":
    main()
