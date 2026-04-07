from typing import Any, Dict, List

from parse_finance_excel import parse_finance_excel, round2


def _sum_bucket_income(company_bucket: Dict[str, List[Dict[str, Any]]]) -> float:
    total = 0.0
    for items in company_bucket.values():
        for item in items:
            total += round2(item.get("income") or 0)
    return round2(total)


def _sum_bucket_expense(company_bucket: Dict[str, List[Dict[str, Any]]]) -> float:
    total = 0.0
    for items in company_bucket.values():
        for item in items:
            total += round2(item.get("expense") or 0)
    return round2(total)


def _company_entry_metrics(entries_df, company: str) -> Dict[str, float]:
    tmp = entries_df[entries_df["Company"] == company].copy()

    if tmp.empty:
        return {
            "income": 0.0,
            "expense": 0.0,
            "net": 0.0,
            "taxes": 0.0,
            "service_revenue": 0.0,
        }

    income = round2(tmp.loc[tmp["Amount"] > 0, "Amount"].sum())
    expense = round2(abs(tmp.loc[tmp["Amount"] < 0, "Amount"].sum()))
    net = round2(tmp["Amount"].sum())
    taxes = round2(abs(tmp.loc[tmp["Category"] == "Taxes", "Amount"].sum()))
    service_revenue = round2(
        tmp.loc[tmp["Category"] == "Service Revenue", "Amount"].sum()
    )

    return {
        "income": income,
        "expense": expense,
        "net": net,
        "taxes": taxes,
        "service_revenue": service_revenue,
    }


def validate_finance_excel(excel_path: str) -> Dict[str, Any]:
    data = parse_finance_excel(excel_path)
    entries_df = data["entries_df"]

    result: Dict[str, Any] = {
        "ok": True,
        "errors": [],
        "warnings": [],
        "details": {
            "companies": {},
            "summary": {},
            "unknown": {},
        },
    }

    unknown = data.get("unknown", {}) or {}
    result["details"]["unknown"] = unknown

    for company, items in unknown.items():
        if items:
            for item in items:
                result["warnings"].append(
                    f"{company} 未映射项目类型: {item['raw_name']} | 金额: {item['amount']} | 类型: {item.get('type', '')}"
                )

    companies = ["ABEMC", "FORRA", "MIND SPORTS"]

    for company in companies:
        bucket = data["companies"][company]
        pnl = data["pnl"][company]
        entry_metrics = _company_entry_metrics(entries_df, company)

        bucket_income = _sum_bucket_income(bucket)
        bucket_expense = _sum_bucket_expense(bucket)
        bucket_net = round2(bucket_income - bucket_expense)

        parser_gross = round2(pnl["gross_revenue"])
        parser_total_costs = round2(pnl["total_costs"])
        parser_net = round2(pnl["final_net_profit"])
        parser_taxes = round2(pnl["taxes"])

        company_errors = []

        if bucket_income != entry_metrics["income"]:
            company_errors.append(
                f"总收入不一致: 表内={bucket_income}, entries={entry_metrics['income']}"
            )

        if bucket_expense != entry_metrics["expense"]:
            company_errors.append(
                f"总支出不一致: 表内={bucket_expense}, entries={entry_metrics['expense']}"
            )

        if bucket_net != entry_metrics["net"]:
            company_errors.append(
                f"净额不一致: 表内={bucket_net}, entries={entry_metrics['net']}"
            )

        if parser_taxes != entry_metrics["taxes"]:
            company_errors.append(
                f"税费不一致: parser={parser_taxes}, entries={entry_metrics['taxes']}"
            )

        if parser_gross != entry_metrics["service_revenue"]:
            company_errors.append(
                f"主营收入不一致: parser={parser_gross}, entries_service_revenue={entry_metrics['service_revenue']}"
            )

        if parser_total_costs != bucket_expense:
            company_errors.append(
                f"总成本不一致: parser={parser_total_costs}, 表内支出={bucket_expense}"
            )

        if parser_net != entry_metrics["net"]:
            company_errors.append(
                f"PNL净额不一致: parser={parser_net}, entries={entry_metrics['net']}"
            )

        result["details"]["companies"][company] = {
            "entries": entry_metrics,
            "bucket": {
                "income": bucket_income,
                "expense": bucket_expense,
                "net": bucket_net,
            },
            "pnl": {
                "gross_revenue": parser_gross,
                "total_costs": parser_total_costs,
                "taxes": parser_taxes,
                "final_net_profit": parser_net,
            },
            "errors": company_errors,
        }

        if company_errors:
            result["ok"] = False
            result["errors"].append(f"{company}: " + "；".join(company_errors))

    month_num = data["meta"]["month_num"]
    summary_left = data["summary_left"]

    summary_map = {}
    for row in summary_left:
        if len(row) >= 2:
            summary_map[str(row[0]).strip()] = round2(row[1])

    gross_key = f"{month_num}月总营收"
    other_income_key = f"{month_num}月其他收入"
    expense_key = f"{month_num}月总支出"
    balance_key = f"{month_num}月总净利(A+B+C)"

    expected_total_gross = round2(data["pnl"]["ABEMC"]["gross_revenue"])

    expected_other_income = round2(
        sum(item.get("income") or 0 for item in data["companies"]["ABEMC"].get("其他收入", [])) +
        sum(item.get("income") or 0 for item in data["companies"]["FORRA"].get("其他收入", [])) +
        sum(item.get("income") or 0 for item in data["companies"]["MIND SPORTS"].get("其他收入", [])) +
        sum(item.get("income") or 0 for item in data["companies"]["ABEMC"].get("未映射项目", [])) +
        sum(item.get("income") or 0 for item in data["companies"]["FORRA"].get("未映射项目", [])) +
        sum(item.get("income") or 0 for item in data["companies"]["MIND SPORTS"].get("未映射项目", []))
    )

    expected_total_balance = round2(
        data["pnl"]["ABEMC"]["final_net_profit"] +
        data["pnl"]["FORRA"]["final_net_profit"] +
        data["pnl"]["MIND SPORTS"]["final_net_profit"]
    )

    expected_total_expense = round2(
        expected_total_gross + expected_other_income - expected_total_balance
    )

    summary_errors = []

    if summary_map.get(gross_key) != expected_total_gross:
        summary_errors.append(
            f"{gross_key} 不一致: summary={summary_map.get(gross_key)}, expected={expected_total_gross}"
        )

    if summary_map.get(other_income_key) != expected_other_income:
        summary_errors.append(
            f"{other_income_key} 不一致: summary={summary_map.get(other_income_key)}, expected={expected_other_income}"
        )

    if summary_map.get(expense_key) != expected_total_expense:
        summary_errors.append(
            f"{expense_key} 不一致: summary={summary_map.get(expense_key)}, expected={expected_total_expense}"
        )

    if summary_map.get(balance_key) != expected_total_balance:
        summary_errors.append(
            f"{balance_key} 不一致: summary={summary_map.get(balance_key)}, expected={expected_total_balance}"
        )

    result["details"]["summary"] = {
        "summary_map": summary_map,
        "expected": {
            gross_key: expected_total_gross,
            other_income_key: expected_other_income,
            expense_key: expected_total_expense,
            balance_key: expected_total_balance,
        },
        "errors": summary_errors,
    }

    if summary_errors:
        result["ok"] = False
        result["errors"].append("顶部汇总异常: " + "；".join(summary_errors))

    if result["ok"]:
        result["warnings"].append("校验通过：总收入、总支出、净额、税费、顶部汇总均一致。")

    return result


if __name__ == "__main__":
    excel_path = "downloads/202511.xlsx"
    check = validate_finance_excel(excel_path)

    print("OK =", check["ok"])

    print("\n=== ERRORS ===")
    for err in check["errors"]:
        print("-", err)

    print("\n=== WARNINGS ===")
    for w in check["warnings"]:
        print("-", w)

    print("\n=== COMPANY DETAILS ===")
    for company, info in check["details"]["companies"].items():
        print(f"\n[{company}]")
        print(info)