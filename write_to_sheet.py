import time
import random
from typing import Any, List, Tuple, Optional

import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

from parse_finance_excel import parse_finance_excel


SERVICE_ACCOUNT_FILE = "service_account.json"
SPREADSHEET_ID = "181As6eTY5-kVzLvyIQC3ncGBT1FZuS9xhhAbqb5h9Rk"
TEMPLATE_SHEET_NAME = "流水模板"
BRAZIL_DATA_SHEET_NAME = "巴西数据"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

LAYOUT = {
    "title_cell": "B1",
    "title_range": "B1:T1",

    "summary_left_start": (2, 3),   # C2:D5
    "summary_mid_start": (2, 6),    # F2:G5

    # 右侧红框：主标签 I/J/L，主值 K/M
    "summary_right_cells": {
        "row1_label": "I2",
        "row1_value": "K2",

        "row2_label": "I3",
        "row2_value": "K3",
        "row2_extra_label": "L3",
        "row2_extra_value": "M3",

        "row3_label": "I4",
        "row3_value": "K4",
        "row3_extra_label": "L4",
        "row3_extra_value": "M4",

        "row4_label": "I5",
        "row4_value": "K5",
        "row4_extra_label": "L5",
        "row4_extra_value": "M5",
    },

    "blocks": {
        "ABEMC": {
            "start_row": 9,
            "start_col": 3,   # C
            "max_rows": 95,
            "closing_label": "期末余额(D)",
            "total_label": "总额(A)",
        },
        "FORRA": {
            "start_row": 9,
            "start_col": 9,   # I
            "max_rows": 95,
            "closing_label": "期末余额(E)",
            "total_label": "总额(B)",
        },
        "MIND SPORTS": {
            "start_row": 9,
            "start_col": 15,  # O
            "max_rows": 95,
            "closing_label": "期末余额(F)",
            "total_label": "总额(C)",
        },
    },
}


def get_gspread_client():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES,
    )
    return gspread.authorize(creds)


def open_spreadsheet():
    gc = get_gspread_client()
    return gc.open_by_key(SPREADSHEET_ID)


def gs_retry(func, *args, **kwargs):
    last_err = None

    for attempt in range(8):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            msg = str(e)
            last_err = e

            if "429" in msg or "Quota exceeded" in msg:
                sleep_s = 20 + attempt * 15
                print(f"[RETRY] Google Sheets 限流，等待 {sleep_s} 秒后重试...")
                time.sleep(sleep_s)
                continue

            if any(code in msg for code in ["502", "503", "500"]):
                sleep_s = (2 ** attempt) * 0.8 + random.random() * 0.3
                time.sleep(sleep_s)
                continue

            raise

    raise last_err


def cell_addr(row: int, col: int) -> str:
    return rowcol_to_a1(row, col)


def a1_range(start_row: int, start_col: int, num_rows: int, num_cols: int) -> str:
    end_row = start_row + num_rows - 1
    end_col = start_col + num_cols - 1
    return f"{cell_addr(start_row, start_col)}:{cell_addr(end_row, end_col)}"


def normalize_cell_value(v: Any):
    if v is None:
        return ""
    return v


def normalize_2d(values: List[List[Any]]) -> List[List[Any]]:
    return [[normalize_cell_value(v) for v in row] for row in values]


def clear_range(ws, start_row: int, start_col: int, num_rows: int, num_cols: int):
    if num_rows <= 0 or num_cols <= 0:
        return
    rng = a1_range(start_row, start_col, num_rows, num_cols)
    gs_retry(ws.batch_clear, [rng])


def write_block(ws, start_row: int, start_col: int, values: List[List[Any]]):
    if not values:
        return
    rng = a1_range(start_row, start_col, len(values), len(values[0]))
    gs_retry(
        ws.update,
        values=normalize_2d(values),
        range_name=rng,
        value_input_option="USER_ENTERED",
    )


def duplicate_template_worksheet(sh, new_sheet_name: str):
    try:
        old_ws = sh.worksheet(new_sheet_name)
        sh.del_worksheet(old_ws)
    except gspread.WorksheetNotFound:
        pass

    template_ws = sh.worksheet(TEMPLATE_SHEET_NAME)
    return template_ws.duplicate(new_sheet_name=new_sheet_name)


def to_num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return round(float(v), 2)

    s = str(v).strip().replace("R$", "").replace(",", "")
    if s == "":
        return 0.0
    return round(float(s), 2)


def get_brazil_crm_monthly_metrics(sh, yyyymm: str, worksheet_name: str = BRAZIL_DATA_SHEET_NAME) -> dict:
    ws = sh.worksheet(worksheet_name)

    values = gs_retry(
        ws.get,
        "A:K",
        value_render_option="UNFORMATTED_VALUE",
    )

    if not values or len(values) < 2:
        raise ValueError("『巴西数据』页没有数据")

    headers = values[0]
    rows = values[1:]
    idx = {str(h).strip(): i for i, h in enumerate(headers)}

    required_cols = [
        "日期",
        "总抽水(A+B+C)",
        "代理抽水(C)",
        "平台抽水(A)",
        "自家代理抽水(B)",
    ]
    for col in required_cols:
        if col not in idx:
            raise ValueError(f"『巴西数据』页缺少列：{col}")

    total_rake = 0.0
    platform_rake = 0.0
    agent_rake_c = 0.0
    own_agent_rake_b = 0.0

    for row in rows:
        if not row:
            continue

        date_val = row[idx["日期"]] if idx["日期"] < len(row) else ""
        date_str = str(date_val).strip().split(".")[0]

        if not date_str.startswith(yyyymm):
            continue

        total_rake += to_num(row[idx["总抽水(A+B+C)"]] if idx["总抽水(A+B+C)"] < len(row) else 0)
        agent_rake_c += to_num(row[idx["代理抽水(C)"]] if idx["代理抽水(C)"] < len(row) else 0)
        platform_rake += to_num(row[idx["平台抽水(A)"]] if idx["平台抽水(A)"] < len(row) else 0)
        own_agent_rake_b += to_num(row[idx["自家代理抽水(B)"]] if idx["自家代理抽水(B)"] < len(row) else 0)

    return {
        "crm_total_rake": round(total_rake, 2),
        "crm_platform_rake": round(platform_rake, 2),
        "crm_agent_rake": round(agent_rake_c + own_agent_rake_b, 2),
    }


def clear_template_input_areas(ws):
    gs_retry(ws.batch_clear, [LAYOUT["title_range"]])

    left_r, left_c = LAYOUT["summary_left_start"]
    mid_r, mid_c = LAYOUT["summary_mid_start"]

    clear_range(ws, left_r, left_c, 4, 2)
    clear_range(ws, mid_r, mid_c, 4, 2)

    right_cells = LAYOUT["summary_right_cells"]
    gs_retry(ws.batch_clear, [
        right_cells["row1_label"], right_cells["row1_value"],
        right_cells["row2_label"], right_cells["row2_value"],
        right_cells["row2_extra_label"], right_cells["row2_extra_value"],
        right_cells["row3_label"], right_cells["row3_value"],
        right_cells["row3_extra_label"], right_cells["row3_extra_value"],
        right_cells["row4_label"], right_cells["row4_value"],
        right_cells["row4_extra_label"], right_cells["row4_extra_value"],
    ])

    for _, cfg in LAYOUT["blocks"].items():
        start_row = cfg["start_row"]
        start_col = cfg["start_col"]
        max_rows = cfg["max_rows"]
        clear_range(ws, start_row, start_col, max_rows + 4, 5)


def format_range_borders(ws, rng: str):
    gs_retry(
        ws.format,
        rng,
        {
            "borders": {
                "top": {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}},
                "bottom": {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}},
                "left": {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}},
                "right": {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}},
            }
        },
    )


def write_title(ws, title: str):
    gs_retry(
        ws.update,
        values=[[title]],
        range_name=LAYOUT["title_cell"],
        value_input_option="USER_ENTERED",
    )


def write_summary(ws, summary_left, summary_mid, summary_right):
    left_r, left_c = LAYOUT["summary_left_start"]
    mid_r, mid_c = LAYOUT["summary_mid_start"]

    write_block(ws, left_r, left_c, summary_left)
    write_block(ws, mid_r, mid_c, summary_mid)

    right_cells = LAYOUT["summary_right_cells"]

    # row1
    gs_retry(
        ws.update,
        values=[[summary_right[0][0]]],
        range_name=right_cells["row1_label"],
        value_input_option="USER_ENTERED",
    )
    gs_retry(
        ws.update,
        values=[[summary_right[0][1]]],
        range_name=right_cells["row1_value"],
        value_input_option="USER_ENTERED",
    )

    # row2
    gs_retry(
        ws.update,
        values=[[summary_right[1][0]]],
        range_name=right_cells["row2_label"],
        value_input_option="USER_ENTERED",
    )
    gs_retry(
        ws.update,
        values=[[summary_right[1][1]]],
        range_name=right_cells["row2_value"],
        value_input_option="USER_ENTERED",
    )
    gs_retry(
        ws.update,
        values=[[summary_right[1][2]]],
        range_name=right_cells["row2_extra_label"],
        value_input_option="USER_ENTERED",
    )
    gs_retry(
        ws.update,
        values=[[summary_right[1][3]]],
        range_name=right_cells["row2_extra_value"],
        value_input_option="USER_ENTERED",
    )

    # row3
    gs_retry(
        ws.update,
        values=[[summary_right[2][0]]],
        range_name=right_cells["row3_label"],
        value_input_option="USER_ENTERED",
    )
    gs_retry(
        ws.update,
        values=[[summary_right[2][1]]],
        range_name=right_cells["row3_value"],
        value_input_option="USER_ENTERED",
    )
    gs_retry(
        ws.update,
        values=[[summary_right[2][2]]],
        range_name=right_cells["row3_extra_label"],
        value_input_option="USER_ENTERED",
    )
    gs_retry(
        ws.update,
        values=[[summary_right[2][3]]],
        range_name=right_cells["row3_extra_value"],
        value_input_option="USER_ENTERED",
    )

    # row4
    gs_retry(
        ws.update,
        values=[[summary_right[3][0]]],
        range_name=right_cells["row4_label"],
        value_input_option="USER_ENTERED",
    )
    gs_retry(
        ws.update,
        values=[[summary_right[3][1]]],
        range_name=right_cells["row4_value"],
        value_input_option="USER_ENTERED",
    )
    gs_retry(
        ws.update,
        values=[[summary_right[3][2]]],
        range_name=right_cells["row4_extra_label"],
        value_input_option="USER_ENTERED",
    )
    gs_retry(
        ws.update,
        values=[[summary_right[3][3]]],
        range_name=right_cells["row4_extra_value"],
        value_input_option="USER_ENTERED",
    )


def validate_block_capacity(company_name: str, rows: List[List[Any]]):
    cfg = LAYOUT["blocks"][company_name]
    max_rows = cfg["max_rows"]

    if len(rows) > max_rows:
        raise ValueError(
            f"{company_name} 数据行数 {len(rows)} 超过模板预留 {max_rows} 行，请把模板预留行数调大。"
        )


def build_running_balance_formula(
    row: int,
    income_col: int,
    expense_col: int,
    detail_start_row: int,
    opening_balance_cell: str,
) -> str:
    income_cell = cell_addr(row, income_col)
    expense_cell = cell_addr(row, expense_col)

    if row == detail_start_row:
        return f'=IF(COUNTA({income_cell}:{expense_cell})=0,"",{opening_balance_cell}+{income_cell}-{expense_cell})'

    prev_balance = cell_addr(row - 1, expense_col + 1)
    return f'=IF(COUNTA({income_cell}:{expense_cell})=0,"",{prev_balance}+{income_cell}-{expense_cell})'


def build_month_total_formulas(
    total_row: int,
    detail_start_row: int,
    detail_end_row: int,
    income_col: int,
    expense_col: int,
) -> Tuple[str, str, str]:
    income_rng = f"{cell_addr(detail_start_row, income_col)}:{cell_addr(detail_end_row, income_col)}"
    expense_rng = f"{cell_addr(detail_start_row, expense_col)}:{cell_addr(detail_end_row, expense_col)}"

    income_formula = f'=ROUND(SUM({income_rng}),2)'
    expense_formula = f'=ROUND(SUM({expense_rng}),2)'
    month_total_formula = f'=ROUND({cell_addr(total_row, income_col)}-{cell_addr(total_row, expense_col)},2)'

    return income_formula, expense_formula, month_total_formula


def prev_month_yyyymm(year: int, month_num: int) -> Optional[str]:
    if month_num == 1:
        return f"{year - 1}12"
    return f"{year}{month_num - 1:02d}"


def get_previous_sheet(sh, current_year: int, current_month: int):
    prev = prev_month_yyyymm(current_year, current_month)
    if not prev:
        return None

    prev_sheet_name = f"{prev}流水"
    try:
        return sh.worksheet(prev_sheet_name)
    except gspread.WorksheetNotFound:
        return None


def read_closing_balance_from_company_block(prev_ws, company_name: str) -> float:
    cfg = LAYOUT["blocks"][company_name]
    start_row = cfg["start_row"]
    start_col = cfg["start_col"]
    max_rows = cfg["max_rows"]
    closing_label = cfg["closing_label"]

    rng = a1_range(start_row, start_col, max_rows + 4, 5)
    values = gs_retry(prev_ws.get, rng, value_render_option="UNFORMATTED_VALUE")

    for row in values:
        if not row:
            continue
        first_cell = str(row[0]).strip() if len(row) >= 1 else ""
        if first_cell == closing_label:
            if len(row) >= 5:
                return to_num(row[4])
            return 0.0

    return 0.0


def get_previous_closing_balance(sh, current_year: int, current_month: int, company_name: str) -> float:
    if current_year == 2025 and current_month == 10:
        return 0.0

    prev_ws = get_previous_sheet(sh, current_year, current_month)
    if prev_ws is None:
        return 0.0

    return read_closing_balance_from_company_block(prev_ws, company_name)


def read_label_value_from_sheet(ws, label: str) -> float:
    values = gs_retry(ws.get_all_values)
    for row in values:
        for idx, cell in enumerate(row):
            if str(cell).strip() == label:
                for j in range(len(row) - 1, idx, -1):
                    if str(row[j]).strip() != "":
                        return to_num(row[j])
    return 0.0


def get_previous_cumulative_distribution(sh, current_year: int, current_month: int) -> float:
    if current_year == 2025 and current_month == 10:
        return 0.0

    prev_ws = get_previous_sheet(sh, current_year, current_month)
    if prev_ws is None:
        return 0.0

    # 固定读取上个月工作表 K4：
    # I4 = 截至当月累计已派发分红 (G)
    # K4 = G 的值
    v = gs_retry(
        prev_ws.acell,
        "K4",
        value_render_option="UNFORMATTED_VALUE",
    ).value
    return to_num(v)


def format_opening_row(ws, row_num: int, start_col: int):
    category_col = start_col
    name_col = start_col + 1
    balance_col = start_col + 4

    try:
        gs_retry(
            ws.merge_cells,
            f"{cell_addr(row_num, category_col)}:{cell_addr(row_num, name_col)}"
        )
    except Exception:
        pass

    rng = f"{cell_addr(row_num, category_col)}:{cell_addr(row_num, balance_col)}"
    format_range_borders(ws, rng)

    gs_retry(
        ws.format,
        f"{cell_addr(row_num, category_col)}:{cell_addr(row_num, name_col)}",
        {
            "horizontalAlignment": "RIGHT",
            "verticalAlignment": "MIDDLE",
            "textFormat": {"bold": True},
        },
    )

    gs_retry(
        ws.format,
        cell_addr(row_num, balance_col),
        {
            "horizontalAlignment": "RIGHT",
            "verticalAlignment": "MIDDLE",
            "textFormat": {"bold": True},
        },
    )


def format_closing_row(ws, row_num: int, start_col: int):
    category_col = start_col
    name_col = start_col + 1
    balance_col = start_col + 4

    try:
        gs_retry(
            ws.merge_cells,
            f"{cell_addr(row_num, category_col)}:{cell_addr(row_num, name_col)}"
        )
    except Exception:
        pass

    rng = f"{cell_addr(row_num, category_col)}:{cell_addr(row_num, balance_col)}"
    format_range_borders(ws, rng)

    gs_retry(
        ws.format,
        f"{cell_addr(row_num, category_col)}:{cell_addr(row_num, name_col)}",
        {
            "horizontalAlignment": "RIGHT",
            "verticalAlignment": "MIDDLE",
            "textFormat": {"bold": True},
        },
    )

    gs_retry(
        ws.format,
        cell_addr(row_num, balance_col),
        {
            "horizontalAlignment": "RIGHT",
            "verticalAlignment": "MIDDLE",
            "textFormat": {"bold": True},
        },
    )


def format_total_row(ws, row_num: int, start_col: int):
    category_col = start_col
    name_col = start_col + 1
    income_col = start_col + 2
    expense_col = start_col + 3
    balance_col = start_col + 4

    try:
        gs_retry(
            ws.merge_cells,
            f"{cell_addr(row_num, category_col)}:{cell_addr(row_num, name_col)}"
        )
    except Exception:
        pass

    total_rng = f"{cell_addr(row_num, category_col)}:{cell_addr(row_num, balance_col)}"
    format_range_borders(ws, total_rng)

    gs_retry(
        ws.format,
        f"{cell_addr(row_num, category_col)}:{cell_addr(row_num, name_col)}",
        {
            "horizontalAlignment": "RIGHT",
            "verticalAlignment": "MIDDLE",
            "textFormat": {"bold": True},
        },
    )

    gs_retry(
        ws.format,
        f"{cell_addr(row_num, income_col)}:{cell_addr(row_num, balance_col)}",
        {
            "horizontalAlignment": "RIGHT",
            "verticalAlignment": "MIDDLE",
            "textFormat": {"bold": True},
        },
    )


def merge_category_cells(ws, rows: List[List[Any]], detail_start_row: int, category_col: int):
    current_category = None
    group_start_row = None

    for i, row in enumerate(rows):
        category = row[0]
        current_row_num = detail_start_row + i

        if category:
            if current_category is not None and group_start_row is not None:
                prev_end_row = current_row_num - 1
                if prev_end_row > group_start_row:
                    try:
                        gs_retry(
                            ws.merge_cells,
                            f"{cell_addr(group_start_row, category_col)}:{cell_addr(prev_end_row, category_col)}"
                        )
                    except Exception:
                        pass

            current_category = category
            group_start_row = current_row_num

    if current_category is not None and group_start_row is not None:
        last_end_row = detail_start_row + len(rows) - 1
        if last_end_row > group_start_row:
            try:
                gs_retry(
                    ws.merge_cells,
                    f"{cell_addr(group_start_row, category_col)}:{cell_addr(last_end_row, category_col)}"
                )
            except Exception:
                pass


def write_company_block_dynamic(ws, sh, data: dict, company_name: str, rows: List[List[Any]]):
    validate_block_capacity(company_name, rows)

    cfg = LAYOUT["blocks"][company_name]
    start_row = cfg["start_row"]
    start_col = cfg["start_col"]
    closing_label = cfg["closing_label"]
    total_label = cfg["total_label"]

    category_col = start_col
    name_col = start_col + 1
    income_col = start_col + 2
    expense_col = start_col + 3
    balance_col = start_col + 4

    year = data["meta"]["year"]
    month_num = data["meta"]["month_num"]

    opening_balance = get_previous_closing_balance(sh, year, month_num, company_name)

    opening_row = start_row
    gs_retry(
        ws.update,
        values=[["期初余额", "", "", "", opening_balance]],
        range_name=f"{cell_addr(opening_row, category_col)}:{cell_addr(opening_row, balance_col)}",
        value_input_option="USER_ENTERED",
    )
    format_opening_row(ws, opening_row, start_col)

    detail_start_row = start_row + 1

    if rows:
        write_block(ws, detail_start_row, start_col, rows)

        opening_balance_cell = cell_addr(opening_row, balance_col)
        balance_values = []

        for i in range(len(rows)):
            row_num = detail_start_row + i
            formula = build_running_balance_formula(
                row=row_num,
                income_col=income_col,
                expense_col=expense_col,
                detail_start_row=detail_start_row,
                opening_balance_cell=opening_balance_cell,
            )
            balance_values.append([formula])

        write_block(ws, detail_start_row, balance_col, balance_values)

        data_rng = f"{cell_addr(detail_start_row, category_col)}:{cell_addr(detail_start_row + len(rows) - 1, balance_col)}"
        format_range_borders(ws, data_rng)

        merge_category_cells(ws, rows, detail_start_row, category_col)

        detail_end_row = detail_start_row + len(rows) - 1
        closing_row = detail_end_row + 1
        total_row = closing_row + 1

        closing_formula = f"={cell_addr(detail_end_row, balance_col)}"

        income_formula, expense_formula, month_total_formula = build_month_total_formulas(
            total_row=total_row,
            detail_start_row=detail_start_row,
            detail_end_row=detail_end_row,
            income_col=income_col,
            expense_col=expense_col,
        )
    else:
        closing_row = detail_start_row
        total_row = closing_row + 1
        closing_formula = f"={cell_addr(opening_row, balance_col)}"
        income_formula = "=0"
        expense_formula = "=0"
        month_total_formula = "=0"

    gs_retry(
        ws.update,
        values=[[closing_label, "", "", "", closing_formula]],
        range_name=f"{cell_addr(closing_row, category_col)}:{cell_addr(closing_row, balance_col)}",
        value_input_option="USER_ENTERED",
    )
    format_closing_row(ws, closing_row, start_col)

    gs_retry(
        ws.update,
        values=[[total_label, "", income_formula, expense_formula, month_total_formula]],
        range_name=f"{cell_addr(total_row, category_col)}:{cell_addr(total_row, balance_col)}",
        value_input_option="USER_ENTERED",
    )
    format_total_row(ws, total_row, start_col)


def write_company_blocks(ws, sh, data: dict):
    flat_rows = data["flat_rows"]

    write_company_block_dynamic(ws, sh, data, "ABEMC", flat_rows["ABEMC"])
    write_company_block_dynamic(ws, sh, data, "FORRA", flat_rows["FORRA"])
    write_company_block_dynamic(ws, sh, data, "MIND SPORTS", flat_rows["MIND SPORTS"])


def make_sheet_name(data: dict) -> str:
    year = data["meta"]["year"]
    month_num = data["meta"]["month_num"]
    return f"{year}{month_num:02d}流水"


def write_finance_excel_to_sheet(excel_path: str) -> Tuple[str, str]:
    data = parse_finance_excel(excel_path)
    sheet_name = make_sheet_name(data)

    sh = open_spreadsheet()

    year = data["meta"]["year"]
    month_num = data["meta"]["month_num"]
    yyyymm = f"{year}{month_num:02d}"

    crm = get_brazil_crm_monthly_metrics(sh, yyyymm, worksheet_name=BRAZIL_DATA_SHEET_NAME)

    a = float(data["pnl"]["ABEMC"]["final_net_profit"] or 0)
    b = float(data["pnl"]["FORRA"]["final_net_profit"] or 0)
    c = float(data["pnl"]["MIND SPORTS"]["final_net_profit"] or 0)

    prev_d = get_previous_closing_balance(sh, year, month_num, "ABEMC")
    prev_e = get_previous_closing_balance(sh, year, month_num, "FORRA")
    prev_f = get_previous_closing_balance(sh, year, month_num, "MIND SPORTS")

    d = round(prev_d + a, 2)
    e = round(prev_e + b, 2)
    f = round(prev_f + c, 2)

    dist = data.get("distribution") or {}
    current_distribution = float(dist.get("distributed", 0) or 0)
    prev_g = get_previous_cumulative_distribution(sh, year, month_num)
    g = round(prev_g + current_distribution, 2)

    gross_revenue = float(data["pnl"]["ABEMC"]["gross_revenue"] or 0)
    other_income_total = float(data["summary_left"][1][1] or 0)
    total_month_balance = round(a + b + c, 2)
    total_expense = round(gross_revenue + other_income_total - total_month_balance, 2)

    summary_left = [
        [f"{month_num}月总营收", round(gross_revenue, 2)],
        [f"{month_num}月其他收入", round(other_income_total, 2)],
        [f"{month_num}月总支出", round(total_expense, 2)],
        [f"{month_num}月总净利(A+B+C)", total_month_balance],
    ]

    summary_mid = [
        [f"{month_num}月总抽水(CRM)", crm["crm_total_rake"]],
        [f"{month_num}月平台抽水(CRM)", crm["crm_platform_rake"]],
        [f"{month_num}月所有代理抽水(CRM)", crm["crm_agent_rake"]],
        [f"三家累计总余额(D+E+F-G)", round(d + e + f - g, 2)],
    ]

    summary_right = [
        ["当月可派发分红 [(B+C)*90%]", round((b + c) * 0.9, 2)],
    
        ["当月已派发分红",
        round(current_distribution, 2),
        "30% 分红",
        round(current_distribution * 0.3, 2)],
    
        ["截至当月累计已派发分红 (G)",
        g,
        "30% 累计分红",
        round(g * 0.3, 2)],
    
        ["ABEMC累计总余额(D)",
        round(d, 2),
        "FORRA+MIND累计总余额 (E+F-G)",
        round(e + f - g, 2)],
    ]

    ws = duplicate_template_worksheet(sh, new_sheet_name=sheet_name)

    title = data["title"]
    clear_template_input_areas(ws)
    write_title(ws, title)
    write_summary(ws, summary_left, summary_mid, summary_right)
    write_company_blocks(ws, sh, data)

    spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
    return sheet_name, spreadsheet_url


if __name__ == "__main__":
    excel_path = "downloads/202601.xlsx"
    sheet_name, url = write_finance_excel_to_sheet(excel_path)
    print(f"已生成工作表：{sheet_name}")
    print(url)