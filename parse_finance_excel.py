import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd


# =========================
# 基础工具
# =========================

MONTH_NAME_TO_NUM = {
    "january": 1,
    "february": 2,
    "march": 3,
    "april": 4,
    "may": 5,
    "june": 6,
    "july": 7,
    "august": 8,
    "september": 9,
    "october": 10,
    "november": 11,
    "december": 12,
}


def round2(x: Any) -> float:
    try:
        return round(float(x), 2)
    except Exception:
        return 0.0


def safe_num(x: Any) -> Optional[float]:
    if pd.isna(x):
        return None
    try:
        return float(x)
    except Exception:
        return None


def normalize_text(x: Any) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def normalize_company_name(name: str) -> str:
    s = normalize_text(name).replace("\xa0", " ").strip().upper()
    if "ABEMC" in s:
        return "ABEMC"
    if "FORRA" in s:
        return "FORRA"
    if "MIND" in s:
        return "MIND SPORTS"
    return s


def read_raw_sheet(excel_path: str, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(excel_path, sheet_name=sheet_name, header=None)


def fmt_label(template: str, month_num: int) -> str:
    return template.format(m=month_num)


# =========================
# 月份 / 标题解析
# =========================

def parse_period_from_cover(excel_path: str) -> Tuple[int, int]:
    try:
        df = read_raw_sheet(excel_path, "📋 Cover")
    except Exception:
        return 2026, 1

    year = 2026
    month_num = 1

    for i in range(len(df)):
        row_texts = [normalize_text(v) for v in df.iloc[i].tolist()]
        joined = " ".join([x for x in row_texts if x])

        m = re.search(
            r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})",
            joined,
            re.I,
        )
        if m:
            month_name = m.group(1).lower()
            year = int(m.group(2))
            month_num = MONTH_NAME_TO_NUM[month_name]
            return year, month_num

    return year, month_num


def make_title(year: int, month_num: int) -> str:
    return f"{year}年{month_num}月流水"


# =========================
# 各公司类别顺序
# =========================

ABEMC_ORDER = [
    "总营收",
    "其他收入",
    "代理渠道成本",
    "技术与产品",
    "Bug/安全",
    "奖金与补贴",
    "市场投放",
    "人力与行政",
    "支付成本",
]

FORRA_ORDER = [
    "自家代理佣金收入",
    "其他收入",
    "代理渠道成本",
    "人力与行政",
    "技术与产品",
    "市场投放",
    "AI订阅",
    "税费",
]

MIND_ORDER = [
    "公司内部收入",
    "其他收入",
    "人力与行政",
    "技术与产品",
    "税费",
    "会计调整",
]


def init_company_buckets(order: List[str]) -> Dict[str, List[Dict[str, Any]]]:
    return {k: [] for k in order}


# =========================
# Entries 解析
# =========================

def parse_entries_sheet(excel_path: str) -> pd.DataFrame:
    raw = read_raw_sheet(excel_path, "📝 Entries")
    if len(raw) < 4:
        return pd.DataFrame()

    headers = [normalize_text(x) for x in raw.iloc[2].tolist()]
    data = raw.iloc[3:].copy()
    data.columns = headers
    data = data.reset_index(drop=True)

    rename_map = {}
    for col in data.columns:
        c = normalize_text(col).lower()
        if c == "date":
            rename_map[col] = "Date"
        elif c == "company":
            rename_map[col] = "Company"
        elif c == "category":
            rename_map[col] = "Category"
        elif c == "description":
            rename_map[col] = "Description"
        elif c == "person / entity":
            rename_map[col] = "Person / Entity"
        elif "amount" in c:
            rename_map[col] = "Amount"

    data = data.rename(columns=rename_map)

    for col in ["Company", "Category", "Description", "Person / Entity"]:
        if col in data.columns:
            data[col] = (
                data[col]
                .fillna("")
                .astype(str)
                .str.replace("\xa0", " ", regex=False)
                .str.replace("—", "-", regex=False)
                .str.strip()
            )

    if "Amount" in data.columns:
        data["Amount"] = pd.to_numeric(data["Amount"], errors="coerce")

    if "Company" in data.columns:
        data["Company"] = data["Company"].apply(normalize_company_name)

    return data


# =========================
# 通用映射 + 公司特例
# =========================

COMMON_EXPENSE_MAP = {
    "AWS Servers": ("技术与产品", "服务器费用（{m}月）"),
    "Tools / Platforms": ("人力与行政", "工具/软件（{m}月）"),
    "AI - Subscriptions": ("人力与行政", "AI订阅（{m}月）"),

    "Salaries": ("人力与行政", "工资（{m}月）"),
    "Accounting": ("人力与行政", "会计服务费（{m}月）"),
    "Accounting - Extra": ("人力与行政", "额外会计费"),
    "Office (Bills)": ("人力与行政", "办公室账单/日常费用"),
    "Office (Equipment)": ("人力与行政", "办公设备采购（{m}月）"),
    "Utensílios Escritório": ("人力与行政", "办公用品/耗材"),
    "Rent": ("人力与行政", "房租（{m}月）"),
    "Electricity": ("人力与行政", "电费（{m}月）"),
    "Training": ("人力与行政", "其他行政/域名/培训"),
    "HR": ("人力与行政", "其他行政/域名/培训"),
    "Plataformas - Assinaturas": ("人力与行政", "其他行政/域名/培训"),
    "Web Domains": ("市场投放", "网站域名费用"),
    "Staff Logistics": ("人力与行政", "人员后勤/差旅（{m}月）"),

    "Marketing - Paid Traffic": ("市场投放", "Facebook付费投放（{m}月合计）"),
    "Marketing - Influencers": ("市场投放", "网红/KOL投放（社媒推广合计）"),
    "Marketing - Events": ("市场投放", "线下活动/赛事物料（Copa Paulista等）"),
    "Marketing - Logistics": ("市场投放", "市场物流费用（{m}月）"),
    "Marketing - Merchandise": ("市场投放", "礼品/赠品/物料"),
    "Marketing - Videomaker": ("市场投放", "视频制作（{m}月）"),
    "Merchandise Suppliers": ("市场投放", "奖牌/奖杯/印刷供应商"),
    "Sponsorship": ("市场投放", "赞助费用"),
    "Sponsorship - Events": ("市场投放", "活动赞助/赛事费用（{m}月）"),
    "Sponsorship - Merchandise": ("市场投放", "赞助礼品"),
    "Sponsorship - Infrastructure": ("市场投放", "赞助基础设施"),

    "Taxes": ("税费", "税费（{m}月）"),
}

COMPANY_OVERRIDES = {
    "ABEMC": {
        "Affiliate Commissions": ("代理渠道成本", "所有代理分成/佣金"),
        "Technology Licensing": ("技术与产品", "技术授权费（30%总营收给MIND）"),

        "Security Reimbursement": ("人力与行政", "安全团队报销/补偿"),
        "Bonus - Bug's App": ("奖金与补贴", "Bug奖金（App Bug）"),

        "Bonus - Ranking": ("奖金与补贴", "{m}月排行榜奖励（以筹码形式发放）"),
        "Tournaments - Overlay": ("奖金与补贴", "比赛补贴/补贴池（Overlay）"),
        "Bonus - Remarketing": ("奖金与补贴", "再营销奖金"),

        "Gateway Fees (Paag)": ("支付成本", "Paag通道手续费"),
        "Gateway Fees (Trio)": ("支付成本", "Trio通道手续费"),
        "Gateway Fees (E2)": ("支付成本", "E2通道手续费"),
    },
    "FORRA": {
        "Affiliate Commissions": ("代理渠道成本", "代理分成/佣金"),
        "CRM": ("市场投放", "CRM（{m}月）"),
        "Computers & Peripherals": ("技术与产品", "电脑/外设采购（Meta眼镜等）"),
    },
    "MIND SPORTS": {
        "China Technology Team": ("人力与行政", "技术相关支出（向中国汇款）"),
        # Account Adjustment 在下面按正负动态处理
    },
}


def resolve_expense_mapping(company: str, category: str) -> Optional[Tuple[Optional[str], Optional[str]]]:
    company_map = COMPANY_OVERRIDES.get(company, {})
    if category in company_map:
        return company_map[category]
    return COMMON_EXPENSE_MAP.get(category)


def add_merged_row(
    merged: Dict[Tuple[str, str], Dict[str, float]],
    category_name: str,
    display_name: str,
    income: float = 0.0,
    expense: float = 0.0,
):
    key = (category_name, display_name)
    if key not in merged:
        merged[key] = {"income": 0.0, "expense": 0.0}
    merged[key]["income"] = round2(merged[key]["income"] + income)
    merged[key]["expense"] = round2(merged[key]["expense"] + expense)


def merged_to_buckets(
    order: List[str],
    merged: Dict[Tuple[str, str], Dict[str, float]],
) -> Dict[str, List[Dict[str, Any]]]:
    buckets = init_company_buckets(order)
    for category in order:
        for (cat, display_name), amounts in merged.items():
            if cat != category:
                continue

            income_val = round2(amounts.get("income", 0.0))
            expense_val = round2(amounts.get("expense", 0.0))

            buckets[category].append({
                "display_name": display_name,
                "income": income_val if income_val else "",
                "expense": expense_val if expense_val else "",
            })
    return buckets


def build_company_buckets_from_entries(entries_df: pd.DataFrame, year: int, month_num: int):
    abemc_merged: Dict[Tuple[str, str], Dict[str, float]] = {}
    forra_merged: Dict[Tuple[str, str], Dict[str, float]] = {}
    mind_merged: Dict[Tuple[str, str], Dict[str, float]] = {}

    abemc_buckets = init_company_buckets(ABEMC_ORDER)
    forra_buckets = init_company_buckets(FORRA_ORDER)
    mind_buckets = init_company_buckets(MIND_ORDER)

    unknown = {
        "ABEMC": [],
        "FORRA": [],
        "MIND SPORTS": [],
    }

    abemc_service_revenue = 0.0
    forra_service_revenue = 0.0
    mind_service_revenue = 0.0

    abemc_other_income = 0.0
    forra_other_income = 0.0
    mind_other_income = 0.0

    for _, row in entries_df.iterrows():
        company = normalize_company_name(row.get("Company", ""))
        category = normalize_text(row.get("Category", ""))
        amount = safe_num(row.get("Amount", None))

        if amount is None:
            continue

        # =========================
        # 正数（收入 / 入账）
        # =========================
        if amount > 0:
            if company == "ABEMC" and category == "Service Revenue":
                abemc_service_revenue += round2(amount)

            elif company == "FORRA" and category == "Service Revenue":
                forra_service_revenue += round2(amount)

            elif company == "MIND SPORTS" and category == "Service Revenue":
                mind_service_revenue += round2(amount)

            elif company == "ABEMC" and category in {"Security Withdrawal", "Office (Bills)"}:
                abemc_other_income += round2(amount)

            elif company == "FORRA" and category == "Refunds":
                forra_other_income += round2(amount)

            elif company == "MIND SPORTS" and category == "Account Adjustment":
                add_merged_row(
                    mind_merged,
                    "会计调整",
                    "会计调整(入账)",
                    income=round2(amount),
                    expense=0.0,
                )

            elif company == "MIND SPORTS":
                mind_other_income += 0.0

        # =========================
        # 负数（支出 / 出账）
        # =========================
        elif amount < 0:
            amount_abs = round2(abs(amount))

            if company == "MIND SPORTS" and category == "Account Adjustment":
                add_merged_row(
                    mind_merged,
                    "会计调整",
                    "会计调整(出账)",
                    income=0.0,
                    expense=amount_abs,
                )
                continue

            mapping = resolve_expense_mapping(company, category)

            if mapping is None:
                unknown[company].append({"raw_name": category, "amount": amount_abs})
                continue

            cat_name, display_tmpl = mapping
            if not cat_name or not display_tmpl:
                continue

            display_name = fmt_label(display_tmpl, month_num)

            if company == "ABEMC":
                add_merged_row(abemc_merged, cat_name, display_name, expense=amount_abs)
            elif company == "FORRA":
                add_merged_row(forra_merged, cat_name, display_name, expense=amount_abs)
            elif company == "MIND SPORTS":
                add_merged_row(mind_merged, cat_name, display_name, expense=amount_abs)
            else:
                unknown[company].append({"raw_name": category, "amount": amount_abs})

    if abemc_service_revenue:
        abemc_buckets["总营收"].append({
            "display_name": f"{year}年{month_num}月总营收",
            "income": round2(abemc_service_revenue),
            "expense": "",
        })

    if forra_service_revenue:
        forra_buckets["自家代理佣金收入"].append({
            "display_name": "FORRA佣金（来自于ABEMC）",
            "income": round2(forra_service_revenue),
            "expense": "",
        })

    if mind_service_revenue:
        mind_buckets["公司内部收入"].append({
            "display_name": "技术授权费（30%总抽水来自于ABEMC）",
            "income": round2(mind_service_revenue),
            "expense": "",
        })

    if abemc_other_income:
        abemc_buckets["其他收入"].append({
            "display_name": "安全回款/冲销（Security相关）",
            "income": round2(abemc_other_income),
            "expense": "",
        })

    if forra_other_income:
        forra_buckets["其他收入"].append({
            "display_name": "Mercado Livre退款/回款",
            "income": round2(forra_other_income),
            "expense": "",
        })

    if mind_other_income:
        mind_buckets["其他收入"].append({
            "display_name": "其他收入",
            "income": round2(mind_other_income),
            "expense": "",
        })

    for cat, items in merged_to_buckets(ABEMC_ORDER, abemc_merged).items():
        abemc_buckets[cat].extend(items)

    for cat, items in merged_to_buckets(FORRA_ORDER, forra_merged).items():
        forra_buckets[cat].extend(items)

    for cat, items in merged_to_buckets(MIND_ORDER, mind_merged).items():
        mind_buckets[cat].extend(items)

    return abemc_buckets, forra_buckets, mind_buckets, unknown


# =========================
# Distribution 解析
# =========================

def parse_distribution_sheet(excel_path: str) -> Dict[str, Any]:
    try:
        df = read_raw_sheet(excel_path, "💸 Distribution")
    except Exception:
        return {
            "distributed": 0.0,
            "remaining_balance": 0.0,
            "forra_mind_treasury": 0.0,
            "abemc_treasury": 0.0,
        }

    result = {
        "distributed": 0.0,
        "remaining_balance": 0.0,
        "forra_mind_treasury": 0.0,
        "abemc_treasury": 0.0,
    }

    for i in range(len(df)):
        row0 = normalize_text(df.iloc[i, 0])
        if row0.startswith("TOTAL ") or re.match(r"^[A-Za-z]+\s+\d{4}$", row0):
            distributed = safe_num(df.iloc[i, 4])
            remaining = safe_num(df.iloc[i, 5])
            if distributed is not None:
                result["distributed"] = round2(distributed)
            if remaining is not None:
                result["remaining_balance"] = round2(remaining)

    for i in range(len(df)):
        row0 = normalize_text(df.iloc[i, 0])
        row1 = safe_num(df.iloc[i, 1])

        if "Forra+Mind Treasury" in row0 and row1 is not None:
            result["forra_mind_treasury"] = round2(row1)
        elif "ABEMC Treasury" in row0 and row1 is not None:
            result["abemc_treasury"] = round2(row1)

    return result


# =========================
# 展平给 write_to_sheet 用
# =========================

def flatten_company_rows(company_buckets: Dict[str, List[Dict[str, Any]]], order: List[str]) -> List[List[Any]]:
    rows: List[List[Any]] = []

    for category in order:
        items = company_buckets.get(category, [])
        if not items:
            continue

        for idx, item in enumerate(items):
            rows.append([
                category if idx == 0 else "",
                item["display_name"],
                item["income"],
                item["expense"],
            ])

    return rows


# =========================
# 主解析函数
# =========================

def parse_finance_excel(excel_path: str) -> Dict[str, Any]:
    year, month_num = parse_period_from_cover(excel_path)
    title = make_title(year, month_num)

    entries_df = parse_entries_sheet(excel_path)
    distribution = parse_distribution_sheet(excel_path)

    abemc_buckets, forra_buckets, mind_buckets, unknown = build_company_buckets_from_entries(
        entries_df, year, month_num
    )

    def company_amount_sum(company_name: str) -> float:
        tmp = entries_df[entries_df["Company"] == company_name]
        return round2(tmp["Amount"].sum()) if not tmp.empty else 0.0

    def service_revenue_sum(company_name: str) -> float:
        tmp = entries_df[
            (entries_df["Company"] == company_name) &
            (entries_df["Category"] == "Service Revenue")
        ]
        return round2(tmp["Amount"].sum()) if not tmp.empty else 0.0

    def taxes_sum(company_name: str) -> float:
        tmp = entries_df[
            (entries_df["Company"] == company_name) &
            (entries_df["Category"] == "Taxes")
        ]
        return round2(abs(tmp["Amount"].sum())) if not tmp.empty else 0.0

    abemc_gross = service_revenue_sum("ABEMC")
    forra_gross = service_revenue_sum("FORRA")
    mind_gross = service_revenue_sum("MIND SPORTS")

    abemc_net = company_amount_sum("ABEMC")
    forra_net = company_amount_sum("FORRA")
    mind_net = company_amount_sum("MIND SPORTS")

    abemc_taxes = taxes_sum("ABEMC")
    forra_taxes = taxes_sum("FORRA")
    mind_taxes = taxes_sum("MIND SPORTS")

    def total_expense_from_buckets(buckets: Dict[str, List[Dict[str, Any]]]) -> float:
        total = 0.0
        for items in buckets.values():
            for item in items:
                total += round2(item["expense"] or 0)
        return round2(total)

    abemc_total_expense = total_expense_from_buckets(abemc_buckets)
    forra_total_expense = total_expense_from_buckets(forra_buckets)
    mind_total_expense = total_expense_from_buckets(mind_buckets)

    other_income_total = round2(
        sum(item["income"] or 0 for item in abemc_buckets["其他收入"]) +
        sum(item["income"] or 0 for item in forra_buckets["其他收入"]) +
        sum(item["income"] or 0 for item in mind_buckets["其他收入"])
    )

    total_balance = round2(abemc_net + forra_net + mind_net)
    total_expense_left = round2(abemc_gross + other_income_total - total_balance)
    distributable_90 = round2((forra_net + mind_net) * 0.9)

    pnl = {
        "ABEMC": {
            "company": "ABEMC",
            "gross_revenue": abemc_gross,
            "total_costs": abemc_total_expense,
            "taxes": abemc_taxes,
            "final_net_profit": abemc_net,
        },
        "FORRA": {
            "company": "FORRA",
            "gross_revenue": forra_gross,
            "total_costs": forra_total_expense,
            "taxes": forra_taxes,
            "final_net_profit": forra_net,
        },
        "MIND SPORTS": {
            "company": "MIND SPORTS",
            "gross_revenue": mind_gross,
            "total_costs": mind_total_expense,
            "taxes": mind_taxes,
            "final_net_profit": mind_net,
        },
    }

    summary_left = [
        [f"{month_num}月总营收", abemc_gross],
        [f"{month_num}月其他收入", other_income_total],
        [f"{month_num}月总支出", total_expense_left],
        [f"{month_num}月总净利(A+B+C)", total_balance],
    ]

    summary_right = [
        [f"{month_num}月总抽水(CRM)", ""],
        [f"{month_num}月平台抽水(CRM)", ""],
        [f"{month_num}月所有代理抽水(CRM)", ""],
        [f"可分配利润[(B+C)*90%]", distributable_90],
    ]

    result = {
        "meta": {
            "year": year,
            "month_num": month_num,
            "title": title,
        },
        "title": title,
        "summary_left": summary_left,
        "summary_right": summary_right,
        "companies": {
            "ABEMC": abemc_buckets,
            "FORRA": forra_buckets,
            "MIND SPORTS": mind_buckets,
        },
        "flat_rows": {
            "ABEMC": flatten_company_rows(abemc_buckets, ABEMC_ORDER),
            "FORRA": flatten_company_rows(forra_buckets, FORRA_ORDER),
            "MIND SPORTS": flatten_company_rows(mind_buckets, MIND_ORDER),
        },
        "pnl": pnl,
        "distribution": distribution,
        "entries_df": entries_df,
        "unknown": unknown,
    }

    return result


if __name__ == "__main__":
    EXCEL_PATH = "downloads/202511.xlsx"
    data = parse_finance_excel(EXCEL_PATH)

    print("\n=== 标题 ===")
    print(data["title"])

    print("\n=== 顶部汇总 ===")
    print(data["summary_left"])
    print(data["summary_right"])

    print("\n=== FORRA PNL ===")
    print(data["pnl"]["FORRA"])

    print("\n=== FORRA 明细 ===")
    for cat, items in data["companies"]["FORRA"].items():
        if items:
            print(cat, items)

    print("\n=== MIND SPORTS 明细 ===")
    for cat, items in data["companies"]["MIND SPORTS"].items():
        if items:
            print(cat, items)

    print("\n=== 未匹配项目 ===")
    print(data["unknown"])