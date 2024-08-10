from PIL import ImageGrab
from datetime import datetime
from time import sleep
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build

sheet_uid = "1EJdMk8qWqvtSAvMkHwHIJB3JwVLppQaC8n0kn77mw1k"

credentials = service_account.Credentials.from_service_account_file(Path("app") / "key.json", scopes=["https://www.googleapis.com/auth/spreadsheets"])
service = build("sheets", "v4", credentials=credentials)

# screenshot = ImageGrab.grab(all_screens=True)
# screenshot.save("screenshot.png")

spreadsheets = service.spreadsheets()
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


task_info = start_time()
sleep(5)
end_time(row, task_info["task"])


# print(data)