import os
import datetime
import requests
import gspread
from google.oauth2.service_account import Credentials

from run_monthly import download_excel_for_month, run_pipeline

TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID = os.environ["TELEGRAM_CHAT_ID"]
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def send_telegram(msg: str) -> None:
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    requests.post(url, json={
        "chat_id": TELEGRAM_CHAT_ID,
        "text": msg,
    }, timeout=30)


def send_telegram_file(file_path: str, caption: str = "") -> None:
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    with open(file_path, "rb") as f:
        requests.post(
            url,
            data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption},
            files={"document": f},
            timeout=120,
        )


def get_last_month() -> str:
    today = datetime.date.today()
    first = today.replace(day=1)
    last_month = first - datetime.timedelta(days=1)
    return last_month.strftime("%Y%m")


def get_gspread_client():
    creds = Credentials.from_service_account_file(
        "service_account.json",
        scopes=SCOPES,
    )
    return gspread.authorize(creds)


def sheet_exists(month: str) -> bool:
    gc = get_gspread_client()
    ss = gc.open_by_key(SPREADSHEET_ID)
    target = f"{month}流水"
    return any(ws.title == target for ws in ss.worksheets())


def main():
    month = os.environ.get("TEST_MONTH") or get_last_month()

    try:
        send_telegram(f"🚀 开始检查 {month} 流水")

        if sheet_exists(month):
            send_telegram(f"✅ {month}流水已存在，今日跳过")
            return

        excel_path = download_excel_for_month(month)
        send_telegram_file(str(excel_path), f"📥 已抓取 {month} Excel")

        run_pipeline(month)

        if sheet_exists(month):
            send_telegram(f"🎉 {month} 流水生成成功")
        else:
            send_telegram(f"❌ {month} 流水未生成成功，请检查日志")
            raise RuntimeError(f"{month}流水未生成成功")

    except Exception as e:
        send_telegram(f"❌ {month} 流水任务失败：{e}")
        raise


if __name__ == "__main__":
    main()