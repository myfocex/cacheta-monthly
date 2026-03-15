from __future__ import annotations

import argparse
import os
import re
import sys
import time
from pathlib import Path
from typing import Tuple

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

from write_to_sheet import write_finance_excel_to_sheet
from validate_finance import validate_finance_excel

DOWNLOAD_DIR = Path("downloads")


def validate_month(month: str) -> str:
    month = str(month).strip()
    if not re.fullmatch(r"\d{6}", month):
        raise ValueError("月份格式必须是 YYYYMM，例如 202601")

    year = int(month[:4])
    mon = int(month[4:6])

    if year < 2020 or year > 2100:
        raise ValueError("年份不合理")
    if mon < 1 or mon > 12:
        raise ValueError("月份必须在 01-12 之间")

    return month


def expected_excel_path(month: str) -> Path:
    return DOWNLOAD_DIR / f"{month}.xlsx"


def download_excel_for_month(month: str) -> Path:
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    excel_path = expected_excel_path(month)

    year = int(month[:4])
    mon = int(month[4:6])

    month_map = {
        1: "Jan",
        2: "Feb",
        3: "Mar",
        4: "Apr",
        5: "May",
        6: "Jun",
        7: "Jul",
        8: "Aug",
        9: "Sep",
        10: "Oct",
        11: "Nov",
        12: "Dec",
    }
    month_label = month_map[mon]

    username = os.getenv("CACHETA_USERNAME", "").strip()
    password = os.getenv("CACHETA_PASSWORD", "").strip()

    storage_state_path = Path("storage_state.json")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
        )

        context_kwargs = {
            "accept_downloads": True,
        }

        if storage_state_path.exists():
            context_kwargs["storage_state"] = str(storage_state_path)

        context = browser.new_context(**context_kwargs)
        page = context.new_page()

        try:
            page.goto("https://cachetafinance.com/", wait_until="networkidle", timeout=60000)

            if page.locator('input[type="password"]').count() > 0:
                if not username or not password:
                    raise ValueError("检测到需要登录，但未设置 CACHETA_USERNAME / CACHETA_PASSWORD")

                print("[INFO] 检测到登录页，开始登录...")
                page.locator('input[type="text"], input[name="username"]').first.fill(username)
                page.locator('input[type="password"]').first.fill(password)

                if page.locator('button:has-text("Sign In")').count() > 0:
                    page.locator('button:has-text("Sign In")').first.click()
                else:
                    page.keyboard.press("Enter")

                page.wait_for_load_state("networkidle", timeout=60000)
                context.storage_state(path=str(storage_state_path))
                print("[INFO] 登录成功，已保存 storage_state.json")

            if page.get_by_text("Consolidated View", exact=False).count() == 0:
                if page.get_by_text("Consolidated", exact=False).count() > 0:
                    page.get_by_text("Consolidated", exact=False).first.click()
                    page.wait_for_load_state("networkidle", timeout=30000)

            page.wait_for_timeout(1500)

            export_btn = None
            export_candidates = [
                page.get_by_text("Export Excel", exact=False),
                page.locator('button:has-text("Export Excel")'),
                page.locator('[role="button"]:has-text("Export Excel")'),
                page.locator('a:has-text("Export Excel")'),
            ]

            for loc in export_candidates:
                try:
                    for i in range(loc.count()):
                        item = loc.nth(i)
                        if item.is_visible():
                            export_btn = item
                            break
                    if export_btn is not None:
                        break
                except Exception:
                    continue

            if export_btn is None:
                raise RuntimeError("没找到 Export Excel 按钮")

            print(f"[INFO] 切换年份：{year}")

            def visible_year_locator():
                return page.locator("text=/^20\\d{2}$/")

            def get_current_year_text() -> str | None:
                loc = visible_year_locator()
                try:
                    for i in range(loc.count()):
                        t = loc.nth(i)
                        if t.is_visible():
                            txt = t.inner_text().strip()
                            if re.fullmatch(r"20\d{2}", txt):
                                return txt
                except Exception:
                    pass
                return None

            def click_previous_year_arrow() -> bool:
                current_year = get_current_year_text()
                if not current_year:
                    return False

                year_loc = page.get_by_text(current_year, exact=True)
                try:
                    target = year_loc.first
                    arrow = target.locator(
                        "xpath=preceding::*[(self::button or @role='button')][1]"
                    )
                    if arrow.count() > 0 and arrow.first.is_visible():
                        arrow.first.click(timeout=3000)
                        return True
                except Exception:
                    pass

                fallback_candidates = [
                    page.locator('button[aria-label*="previous"]'),
                    page.locator('[role="button"][aria-label*="previous"]'),
                    page.locator('button:has(svg)'),
                    page.locator('[role="button"]:has(svg)'),
                ]
                for loc in fallback_candidates:
                    try:
                        for i in range(loc.count()):
                            btn = loc.nth(i)
                            if btn.is_visible():
                                btn.click(timeout=3000)
                                return True
                    except Exception:
                        continue
                return False

            current_year = get_current_year_text()
            if current_year is None:
                raise RuntimeError("没读到当前年份")

            print(f"[INFO] 当前顶部年份：{current_year}")

            for _ in range(6):
                current_year = get_current_year_text()
                if current_year == str(year):
                    break

                if not click_previous_year_arrow():
                    raise RuntimeError(f"无法点击年份切换箭头，当前年份：{current_year}")

                page.wait_for_timeout(1000)

            current_year = get_current_year_text()
            if current_year != str(year):
                raise RuntimeError(f"年份未切换成功，当前年份是：{current_year}，目标年份是：{year}")

            print(f"[INFO] 已切到年份：{current_year}")

            print(f"[INFO] 切换月份：{month_label}")
            month_clicked = False
            month_candidates = [
                page.get_by_text(month_label, exact=True),
                page.locator(f'button:has-text("{month_label}")'),
                page.locator(f'[role="button"]:has-text("{month_label}")'),
                page.locator(f'span:has-text("{month_label}")'),
                page.locator(f'div:has-text("{month_label}")'),
            ]

            for loc in month_candidates:
                try:
                    for i in range(loc.count()):
                        item = loc.nth(i)
                        if item.is_visible():
                            item.click(timeout=3000)
                            month_clicked = True
                            break
                    if month_clicked:
                        break
                except Exception:
                    continue

            if not month_clicked:
                print(f"[INFO] 网站还没有月份数据：{month_short}/{month_full}")
                return None

            page.wait_for_load_state("networkidle", timeout=30000)
            page.wait_for_timeout(1500)

            print("[INFO] 点击 Export Excel")
            with page.expect_download(timeout=120000) as download_info:
                export_btn.click()

            download = download_info.value
            if excel_path.exists():
                excel_path.unlink()

            download.save_as(str(excel_path))
            print(f"[INFO] 下载完成：{excel_path}")

            return excel_path

        except PlaywrightTimeoutError as e:
            raise RuntimeError(f"下载超时：{e}") from e

        finally:
            context.close()
            browser.close()


def run_pipeline(month: str) -> Tuple[str, str, Path]:
    month = validate_month(month)

    print(f"[INFO] 开始处理月份：{month}")

    excel_path = download_excel_for_month(month)
    print(f"[INFO] Excel 文件：{excel_path}")

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel 文件不存在：{excel_path}")

    sheet_name, url = write_finance_excel_to_sheet(str(excel_path))
    print(f"[INFO] 已生成工作表：{sheet_name}")
    print(f"[INFO] Google Sheet：{url}")

    validation = validate_finance_excel(str(excel_path))

    if validation["ok"]:
        print("[INFO] 校验通过：总收入、总支出、净额、税费、顶部汇总均一致，且没有未映射类型。")
    else:
        print("[WARNING] 校验未通过")
        for err in validation["errors"]:
            print(f"[WARNING] {err}")

    return sheet_name, url, excel_path


def main():
    parser = argparse.ArgumentParser(description="下载财务 Excel 并写入 Google Sheet")
    parser.add_argument(
        "months",
        nargs="+",
        help="月份列表，格式 YYYYMM，例如 202510 202511 202512"
    )
    args = parser.parse_args()

    success_list = []
    failed_list = []

    for idx, month in enumerate(args.months):
        print("\n" + "=" * 60)
        print(f"[RUN] 开始处理 {month}")
        print("=" * 60)

        try:
            sheet_name, url, excel_path = run_pipeline(month)
            success_list.append({
                "month": month,
                "excel": str(excel_path),
                "sheet_name": sheet_name,
                "url": url,
            })
            print(f"[DONE] {month} 完成")

        except Exception as e:
            failed_list.append({
                "month": month,
                "error": str(e),
            })
            print(f"[ERROR] {month} 失败：{e}")

        if idx < len(args.months) - 1:
            print("[INFO] 等待 20 秒，避免 Google Sheets 写入限额...")
            time.sleep(20)

    print("\n" + "=" * 60)
    print("批量执行结果")
    print("=" * 60)

    if success_list:
        print("\n成功：")
        for item in success_list:
            print(f"- {item['month']} -> {item['sheet_name']}")
            print(f"  Excel: {item['excel']}")
            print(f"  Link : {item['url']}")

    if failed_list:
        print("\n失败：")
        for item in failed_list:
            print(f"- {item['month']} -> {item['error']}")

    if failed_list:
        sys.exit(1)


if __name__ == "__main__":
    main()