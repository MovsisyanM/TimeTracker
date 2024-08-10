import os
from time import sleep
from pathlib import Path
from PIL import ImageGrab
from datetime import datetime

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http  import MediaFileUpload

import winsound
frequency = 2500
duration = 1500

sheet_uid = "1EJdMk8qWqvtSAvMkHwHIJB3JwVLppQaC8n0kn77mw1k"

credentials = service_account.Credentials.from_service_account_file(
    Path("app") / "key.json",
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        'https://www.googleapis.com/auth/drive'
    ]
)
sheets_service = build("sheets", "v4", credentials=credentials)
drive_service = build("drive", "v3", credentials=credentials)


spreadsheets = sheets_service.spreadsheets()
values_api = spreadsheets.values()
request = values_api.get(spreadsheetId=sheet_uid, range="'Time log'!A:B")
data = request.execute()["values"]

row = len(data) + 1
# col = len(data[-1])

def start_time() -> dict[str, str]:
    values: list[str] = [datetime.now().strftime("%Y/%m/%d %H:%M"), "Currently working..."]
    values_api.update(
        spreadsheetId=sheet_uid,
        range=f"'Time log'!A{row}:B{row}",
        valueInputOption="USER_ENTERED",
        body={
            "values": [
                values
            ]
        }
    ).execute()

    task = input("Task: ")

    values_api.update(
        spreadsheetId=sheet_uid,
        range=f"'Time log'!D{row}",
        valueInputOption="USER_ENTERED",
        body={
            "values": [
                [task]
            ]
        }
    ).execute()

    return {
        "start_time": values[0],
        "task": task,
    }


def end_time(row: int, task: str):
    values_api.update(
        spreadsheetId=sheet_uid,
        range=f"'Time log'!B{row}:F{row}",
        valueInputOption="USER_ENTERED",
        body={
            "values": [[
                datetime.now().strftime("%Y/%m/%d %H:%M"),
                f'=IF(ISBLANK(A{row}), "", TEXT(IF(EQ(B{row}, "Currently working..."), NOW(), B{row})-A{row}, "h:mm"))',
                task,
                f'=TEXT(C{row}, "h")*G{row} + RIGHT(C{row}, 2)*G{row}/60', 0
            ]]
        }
    ).execute()


def get_drive_folder() -> str:
    # Define the folder name you want to check for
    folder_name = datetime.now().strftime("%Y_%m_%d__%H_00")

    # Define the search query
    query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder'"
    
    folder_permission: dict = {
        'type': 'anyone',
        'role': 'reader'
    }

    # Search for the folder
    results = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
    items = results.get('files', [])

    if items:
        # Folder exists
        folder_id = items[0]['id']
    else:
        # Folder does not exist
        folder_metadata: dict = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }

        folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
        folder_id = folder.get('id')
    drive_service.permissions().create(fileId=folder_id, body=folder_permission).execute()

    return folder_id

def screenshot() -> str:
    ss_name: str = datetime.now().strftime("%Y_%m_%d__%H_%M.jpg")
    screenshot = ImageGrab.grab(all_screens=True)
    screenshot.save(str((Path("app") / "screenshots" / ss_name).resolve()))

    return ss_name


def upload_screenshot(ss_name: str) -> str:
    folder_id = get_drive_folder()
    ss_path = Path("app") / "screenshots" / ss_name

    file_metadata = {
        'name':  ss_name,
        "parents": [folder_id]
    }
    
    media = MediaFileUpload(str(ss_path), mimetype=None)
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

    shareable_link = f"https://drive.google.com/drive/folders/{folder_id}?usp=sharing"
    return shareable_link


def share_ss_folder(sharable_link: str):
    values_api.update(
        spreadsheetId=sheet_uid,
        range=f"'Time log'!H{row}",
        valueInputOption="USER_ENTERED",
        body={
            "values": [[sharable_link]]
        }
    ).execute()

if __name__ == "__main__":
    task_info = start_time()

    try:
        while True:
            sleep(60*5 - 4)
            winsound.Beep(frequency, duration)
            sleep(3)
            ss_name = screenshot()
            sharable_link = upload_screenshot(ss_name)
            share_ss_folder(sharable_link)
    except KeyboardInterrupt as ke:
        end_time(row, task_info["task"])
