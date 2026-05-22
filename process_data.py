import pandas as pd
import json
import os
from datetime import datetime, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_FILE = os.path.join(BASE_DIR, 'Data', 'source_data.xlsx')
OUTPUT_FILE = os.path.join(BASE_DIR, 'analysis_results.json')


def zone_stats(data):
    """按专区名称统计，按订单号去重"""
    zones = {}
    for name, group in data.groupby('专区名称'):
        zones[name] = {
            "order_count": int(group['订单号'].nunique()),
            "total_amount": round(float(group['订单金额（元）'].sum()), 2)
        }
    return zones


def supplier_stats(data):
    """按供应商类型统计，按订单号去重"""
    result = {}
    for stype in ["本地供应商", "电商供应商"]:
        stype_data = data[data['供应商类型'] == stype]
        result[stype] = {
            "order_count": int(stype_data['订单号'].nunique()),
            "total_amount": round(float(stype_data['订单金额（元）'].sum()), 2)
        }
    return result


def process_monday():
    """周一统计：上周、本月、总计（分专区）"""
    df = pd.read_excel(SOURCE_FILE)
    df['订单日期'] = pd.to_datetime(df['订单日期'])

    today = datetime.now()
    current_year = today.year
    current_month = today.month

    # 本月
    month_mask = (df['订单日期'].dt.year == current_year) & (df['订单日期'].dt.month == current_month)
    month_data = df[month_mask]

    # 上周（周一到周日）
    days_since_monday = today.weekday()
    this_monday = today - timedelta(days=days_since_monday)
    last_monday = this_monday - timedelta(days=7)
    last_sunday = this_monday - timedelta(days=1)
    last_week_mask = (df['订单日期'] >= last_monday) & (df['订单日期'] <= last_sunday)
    last_week_data = df[last_week_mask]

    result = {
        "type": "monday",
        "total": {
            "order_count": int(df['订单号'].nunique()),
            "total_amount": round(float(df['订单金额（元）'].sum()), 2),
            "zones": zone_stats(df)
        },
        "last_week": {
            "range": f"{last_monday.strftime('%Y-%m-%d')} ~ {last_sunday.strftime('%Y-%m-%d')}",
            "order_count": int(last_week_data['订单号'].nunique()),
            "total_amount": round(float(last_week_data['订单金额（元）'].sum()), 2),
            "zones": zone_stats(last_week_data)
        },
        "current_month": {
            "month": f"{current_year}-{current_month:02d}",
            "order_count": int(month_data['订单号'].nunique()),
            "total_amount": round(float(month_data['订单金额（元）'].sum()), 2),
            "zones": zone_stats(month_data)
        }
    }

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=4)

    for key, label in [("last_week", "上周"), ("current_month", "本月"), ("total", "总计")]:
        r = result[key]
        header = label
        if "range" in r:
            header += f"({r['range']})"
        elif "month" in r:
            header += f"({r['month']})"
        print(f"{header}: 订单数 {r['order_count']}, 销售额 {r['total_amount']}")
        for zone, s in r["zones"].items():
            print(f"  {zone}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")

    print(f"\n结果已保存: {OUTPUT_FILE}")
    return result


def process_friday():
    """周五统计：本周（按供应商类型 + 分专区）"""
    df = pd.read_excel(SOURCE_FILE)
    df['订单日期'] = pd.to_datetime(df['订单日期'])

    today = datetime.now()

    # 本周（周一到周日）
    days_since_monday = today.weekday()
    this_monday = today - timedelta(days=days_since_monday)
    this_sunday = this_monday + timedelta(days=6)
    week_mask = (df['订单日期'] >= this_monday) & (df['订单日期'] <= this_sunday)
    week_data = df[week_mask]

    result = {
        "type": "friday",
        "current_week": {
            "range": f"{this_monday.strftime('%Y-%m-%d')} ~ {this_sunday.strftime('%Y-%m-%d')}",
            "order_count": int(week_data['订单号'].nunique()),
            "total_amount": round(float(week_data['订单金额（元）'].sum()), 2),
            "supplier": supplier_stats(week_data),
            "zones": zone_stats(week_data)
        }
    }

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=4)

    r = result["current_week"]
    print(f"本周({r['range']}): 订单数 {r['order_count']}, 销售额 {r['total_amount']}")
    for stype, s in r["supplier"].items():
        print(f"  {stype}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")
    for zone, s in r["zones"].items():
        print(f"  {zone}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")

    print(f"\n结果已保存: {OUTPUT_FILE}")
    return result


def process(mode="auto"):
    """
    mode: "monday" | "friday" | "auto"
      monday - 强制跑周一统计（上周/本月/总计 + 分专区）
      friday - 强制跑周五统计（本周 + 供应商类型 + 分专区）
      auto   - 根据今天是周几自动选择
    """
    if mode == "monday":
        return process_monday()
    elif mode == "friday":
        return process_friday()
    else:
        today = datetime.now()
        if today.weekday() == 0:
            return process_monday()
        else:
            return process_friday()


if __name__ == "__main__":
    import sys
    mode = sys.argv[1] if len(sys.argv) > 1 else "auto"
    process(mode)
