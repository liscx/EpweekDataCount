# -*- coding: utf-8 -*-
"""日常综合统计：上周、本周、上月、本月、全量（均分供应商类型+分专区）"""
import pandas as pd
import os
import sys
from datetime import datetime, timedelta

from stats_utils import (
    zone_stats, supplier_stats,
    make_xlsx, save_xlsx, set_col_widths,
    write_section_with_supplier,
    RESULT_DIR,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_FILE = os.path.join(BASE_DIR, 'Data', 'source_data.xlsx')


def run():
    """执行综合统计，返回 (result, xlsx_path)"""
    df = pd.read_excel(SOURCE_FILE)
    df['订单创建时间'] = pd.to_datetime(df['订单创建时间'])

    now = datetime.now()
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    current_year = today.year
    current_month = today.month

    # 本月
    month_mask = (df['订单创建时间'].dt.year == current_year) & (df['订单创建时间'].dt.month == current_month)
    month_data = df[month_mask]

    # 上月
    if current_month == 1:
        last_year, last_month = current_year - 1, 12
    else:
        last_year, last_month = current_year, current_month - 1
    last_month_mask = (df['订单创建时间'].dt.year == last_year) & (df['订单创建时间'].dt.month == last_month)
    last_month_data = df[last_month_mask]

    # 上周（上周一 ~ 上周日，完整7天）
    days_since_monday = today.weekday()
    this_monday = today - timedelta(days=days_since_monday)
    last_monday = this_monday - timedelta(days=7)
    last_week_mask = (df['订单创建时间'] >= last_monday) & (df['订单创建时间'] < this_monday)
    last_week_data = df[last_week_mask]

    # 本周（本周一 ~ 当前时刻，精确到秒）
    current_week_mask = (df['订单创建时间'] >= this_monday) & (df['订单创建时间'] <= now)
    current_week_data = df[current_week_mask]

    result = {
        "type": "normal",
        "last_week": {
            "range": f"{last_monday.strftime('%Y-%m-%d')} ~ {(this_monday - timedelta(days=1)).strftime('%Y-%m-%d')}",
            "supplier": supplier_stats(last_week_data),
            "zones": zone_stats(last_week_data)
        },
        "current_week": {
            "range": f"{this_monday.strftime('%Y-%m-%d')} ~ {now.strftime('%Y-%m-%d %H:%M')}",
            "supplier": supplier_stats(current_week_data),
            "zones": zone_stats(current_week_data)
        },
        "last_month": {
            "month": f"{last_year}-{last_month:02d}",
            "supplier": supplier_stats(last_month_data),
            "zones": zone_stats(last_month_data)
        },
        "current_month": {
            "month": f"{current_year}-{current_month:02d}",
            "supplier": supplier_stats(month_data),
            "zones": zone_stats(month_data)
        },
        "total": {
            "supplier": supplier_stats(df),
            "zones": zone_stats(df)
        }
    }

    # 控制台输出
    for label, key in [("上周", "last_week"), ("本周", "current_week"),
                        ("上月", "last_month"), ("本月", "current_month"),
                        ("全量", "total")]:
        r = result[key]
        header = label
        if "range" in r:
            header += f"（{r['range']}）"
        elif "month" in r:
            header += f"（{r['month']}）"
        total_count = sum(s["order_count"] for s in r["supplier"].values())
        total_amount = round(sum(s["total_amount"] for s in r["supplier"].values()), 2)
        print(f"\n{header}: 订单数 {total_count}, 销售额 {total_amount}")
        for stype, s in r["supplier"].items():
            print(f"  {stype}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")
        print(f"  专区统计:")
        for zone, s in r["zones"].items():
            print(f"    {zone}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")

    # 导出 Excel
    wb, ws, xlsx_path = make_xlsx("NM", "综合统计")

    row = 1
    for label, key in [("上周", "last_week"), ("本周", "current_week"),
                        ("上月", "last_month"), ("本月", "current_month"),
                        ("全量", "total")]:
        r = result[key]
        title = label
        if "range" in r:
            title += f"（{r['range']}）"
        elif "month" in r:
            title += f"（{r['month']}）"
        row = write_section_with_supplier(ws, row, title, r["supplier"], r["zones"])

    set_col_widths(ws)
    save_xlsx(wb, xlsx_path)
    return result, xlsx_path


if __name__ == "__main__":
    run()
