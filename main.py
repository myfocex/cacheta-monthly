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
    requests.post(
        url,
        json={
            "chat_id": TELEGRAM_CHAT_ID,
            "text": msg,
        },
        timeout=30,
    )


def send_telegram_file(file_path: str, caption: str = "") -> None:
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    with open(file_path, "rb") as f:
        requests.post(
            url,
            data={
                "chat_id": TELEGRAM_CHAT_ID,
                "caption": caption,
            },
            files={"document": f},
            timeout=120,
        )


def get_last_month() -> str:
    today = datetime.date.today()
    first = today.replace(day=1)
    last_month = first - datetime.timedelta(days=1)
    return last_month.strftime("%Y%m")


def get_target_month() -> str:
    # 手动测试优先，例如 TEST_MONTH=202601
    test_month = os.environ.get("TEST_MONTH", "").strip()
    if test_month:
        return test_month
    return get_last_month()


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


def should_stop_retry() -> bool:
    """
    正式自动模式下：
    每月15号开始检查“上个月”
    如果到了本月15号之前，其实不需要跑
    但你现在是每天定时跑一次，所以这里只做“是否已过14号”判断的可扩展预留。

    当前版本先不硬性拦截，避免影响 TEST_MONTH 测试。
    后面你如果要严格控制到“次月14号停止”，我可以再补。
    """
    return False


def is_month_not_ready_error(err: Exception) -> bool:
    msg = str(err)
    keywords = [
        "没找到月份按钮",
        "未找到月份按钮",
        "没找到月份",
        "month_label",
    ]
    return any(k in msg for k in keywords)


def main():
    month = get_target_month()

    try:
        send_telegram(f"🚀 开始检查 {month} 流水")

        if should_stop_retry():
            send_telegram(f"⛔ 已超过重试截止日期，停止 {month} 抓取")
            return

        # 1. 已存在则直接跳过
        if sheet_exists(month):
            send_telegram(f"✅ {month}流水已存在，今日跳过")
            return

        # 2. 下载 Excel
        try:
            excel_path = download_excel_for_month(month)
        except RuntimeError as e:
            if is_month_not_ready_error(e):
                send_telegram(f"📭 网站还没有 {month} 的流水数据，明天继续检查")
                return
            raise

        if not excel_path:
            send_telegram(f"📭 网站还没有 {month} 的流水数据，明天继续检查")
            return

        # 3. 发 Excel 到 Telegram
        send_telegram_file(str(excel_path), f"📥 已抓取 {month} Excel")

        # 4. 生成 Google Sheet 流水
        run_pipeline(month)

        # 5. 再检查一次是否成功生成
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