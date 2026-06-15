# -*- coding: utf-8 -*-
"""周五统计：本周汇总+供应商类型，分专区（本周/本月/全量）"""
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
    """执行周五统计，返回 (result, xlsx_path)"""
    df = pd.read_excel(SOURCE_FILE)
    df['订单创建时间'] = pd.to_datetime(df['订单创建时间'])

    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    current_year = today.year
    current_month = today.month
    now = datetime.now()

    # 本周（本周一 00:00 ~ 当前执行时间）
    days_since_monday = today.weekday()
    this_monday = today - timedelta(days=days_since_monday)
    week_mask = (df['订单创建时间'] >= this_monday) & (df['订单创建时间'] <= now)
    week_data = df[week_mask]

    # 本月
    month_mask = (df['订单创建时间'].dt.year == current_year) & (df['订单创建时间'].dt.month == current_month)
    month_data = df[month_mask]

    result = {
        "type": "friday",
        "current_week": {
            "range": f"{this_monday.strftime('%Y-%m-%d')} ~ {now.strftime('%Y-%m-%d %H:%M')}",
            "order_count": int(week_data['订单号'].nunique()),
            "total_amount": round(float(week_data['订单金额（元）'].sum()), 2),
            "supplier": supplier_stats(week_data),
        },
        "zones": {
            "current_week": {
                "range": f"{this_monday.strftime('%Y-%m-%d')} ~ {now.strftime('%Y-%m-%d %H:%M')}",
                "supplier": supplier_stats(week_data),
                **zone_stats(week_data)
            },
            "current_month": {
                "month": f"{current_year}-{current_month:02d}",
                "supplier": supplier_stats(month_data),
                **zone_stats(month_data)
            },
            "total": {
                "supplier": supplier_stats(df),
                **zone_stats(df)
            }
        }
    }

    # 控制台输出
    for label, key in [("本周", "current_week"), ("本月", "current_month"), ("全量", "total")]:
        z = result["zones"][key]
        header = label
        if "range" in z:
            header += f"({z['range']})"
        elif "month" in z:
            header += f"({z['month']})"
        print(f"\n{header}:")
        for stype, s in z["supplier"].items():
            print(f"  {stype}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")
        print(f"  专区统计:")
        for zone, s in z.items():
            if isinstance(s, dict) and zone != "supplier":
                print(f"    {zone}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")

    # 导出 Excel
    wb, ws, xlsx_path = make_xlsx("F", "周五统计")

    row = 1
    for label, key in [("本周", "current_week"), ("本月", "current_month"), ("全量", "total")]:
        z = result["zones"][key]
        title = label
        if "range" in z:
            title += f"（{z['range']}）"
        elif "month" in z:
            title += f"（{z['month']}）"
        zone_items = {k: v for k, v in z.items() if isinstance(v, dict) and k != "supplier"}
        row = write_section_with_supplier(ws, row, title, z["supplier"], zone_items)

    set_col_widths(ws)
    save_xlsx(wb, xlsx_path)
    return result, xlsx_path


if __name__ == "__main__":
    run()
