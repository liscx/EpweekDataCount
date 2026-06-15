# -*- coding: utf-8 -*-
"""周一统计：上周、本月、总计（分专区）"""
import pandas as pd
import os
import sys
from datetime import datetime, timedelta

from stats_utils import (
    zone_stats, supplier_stats,
    make_xlsx, save_xlsx, set_col_widths,
    write_title, write_header, write_row, write_total_row,
    write_section_zone_only, write_section_with_supplier,
    RESULT_DIR,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_FILE = os.path.join(BASE_DIR, 'Data', 'source_data.xlsx')


def run():
    """执行周一统计，返回 (result, xlsx_path)"""
    df = pd.read_excel(SOURCE_FILE)
    df['订单创建时间'] = pd.to_datetime(df['订单创建时间'])

    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    current_year = today.year
    current_month = today.month

    # 本月
    month_mask = (df['订单创建时间'].dt.year == current_year) & (df['订单创建时间'].dt.month == current_month)
    month_data = df[month_mask]

    # 上周（周一到周日）
    days_since_monday = today.weekday()
    this_monday = today - timedelta(days=days_since_monday)
    last_monday = this_monday - timedelta(days=7)
    last_week_mask = (df['订单创建时间'] >= last_monday) & (df['订单创建时间'] < this_monday)
    last_week_data = df[last_week_mask]

    result = {
        "type": "monday",
        "total": {
            "order_count": int(df['订单号'].nunique()),
            "total_amount": round(float(df['订单金额（元）'].sum()), 2),
            "zones": zone_stats(df)
        },
        "last_week": {
            "range": f"{last_monday.strftime('%Y-%m-%d')} ~ {(this_monday - timedelta(days=1)).strftime('%Y-%m-%d')}",
            "order_count": int(last_week_data['订单号'].nunique()),
            "total_amount": round(float(last_week_data['订单金额（元）'].sum()), 2),
            "zones": zone_stats(last_week_data)
        },
        "current_month": {
            "month": f"{current_year}-{current_month:02d}",
            "order_count": int(month_data['订单号'].nunique()),
            "total_amount": round(float(month_data['订单金额（元）'].sum()), 2),
            "zones": zone_stats(month_data)
        },
        "total_all": {
            "supplier": supplier_stats(df),
            "zones": zone_stats(df)
        }
    }

    # 控制台输出
    for key, label in [("last_week", "上周"), ("current_month", "本月")]:
        r = result[key]
        header = label
        if "range" in r:
            header += f"({r['range']})"
        elif "month" in r:
            header += f"({r['month']})"
        print(f"{header}: 订单数 {r['order_count']}, 销售额 {r['total_amount']}")
        for zone, s in r["zones"].items():
            print(f"  {zone}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")

    if "total_all" in result:
        t = result["total_all"]
        print(f"\n全量:")
        for stype, s in t["supplier"].items():
            print(f"  {stype}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")
        print(f"  专区统计:")
        for zone, s in t["zones"].items():
            print(f"    {zone}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")

    # 导出 Excel
    wb, ws, xlsx_path = make_xlsx("M", "订单统计")

    row = 1
    for label, key in [("上周", "last_week"), ("本月", "current_month")]:
        r = result[key]
        title = label
        if "range" in r:
            title += f"（{r['range']}）专区统计"
        elif "month" in r:
            title += f"（{r['month']}）订单统计"
        row = write_section_zone_only(ws, row, title, r["zones"])

    # 全量统计（含供应商类型 + 专区）
    if "total_all" in result:
        t = result["total_all"]
        row = write_title(ws, row, "全量统计")

        # 供应商类型表
        row = write_header(ws, row, ["供应商类型", "订单数", "销售金额（元）"])
        data_start = row
        for stype, s in t["supplier"].items():
            row = write_row(ws, row, stype, s["order_count"], s["total_amount"])
        data_end = row - 1
        write_total_row(ws, row, data_start, data_end)
        row += 1

        # 专区表
        row = write_header(ws, row)
        data_start = row
        for zone, s in t["zones"].items():
            row = write_row(ws, row, zone, s["order_count"], s["total_amount"])
        data_end = row - 1
        write_total_row(ws, row, data_start, data_end)

    set_col_widths(ws)
    save_xlsx(wb, xlsx_path)
    return result, xlsx_path


if __name__ == "__main__":
    run()
