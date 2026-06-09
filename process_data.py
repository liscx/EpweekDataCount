import pandas as pd
import os
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_FILE = os.path.join(BASE_DIR, 'Data', 'source_data.xlsx')
RESULT_DIR = os.path.join(BASE_DIR, 'result')
os.makedirs(RESULT_DIR, exist_ok=True)


def _write_total_row(ws, row, data_start, data_end, thin_border):
    """写合计行，用静态值（从上方单元格求和）"""
    total_b = sum(ws.cell(row=r, column=2).value or 0 for r in range(data_start, data_end + 1))
    total_c = sum(ws.cell(row=r, column=3).value or 0 for r in range(data_start, data_end + 1))
    ws.cell(row=row, column=1, value="合计").font = Font(bold=True)
    ws.cell(row=row, column=1).border = thin_border
    ws.cell(row=row, column=2, value=total_b)
    ws.cell(row=row, column=2).border = thin_border
    ws.cell(row=row, column=2).font = Font(bold=True)
    ws.cell(row=row, column=3, value=round(total_c, 2))
    ws.cell(row=row, column=3).border = thin_border
    ws.cell(row=row, column=3).font = Font(bold=True)
    ws.cell(row=row, column=3).number_format = '#,##0.00'


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


def export_friday_xlsx(result):
    """将周五统计结果导出为xlsx，格式参照模板"""
    timestamp = datetime.now().strftime('%Y%m%d%H%M')
    xlsx_path = os.path.join(RESULT_DIR, f'analysis_results_{timestamp}_F.xlsx')
    wb = Workbook()
    ws = wb.active
    ws.title = "周五统计"

    # 样式
    title_font = Font(bold=True, size=18)
    header_font = Font(bold=True, size=11)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    def write_title(row, text):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c = ws.cell(row=row, column=1, value=text)
        c.font = title_font
        c.alignment = Alignment(horizontal='center')
        return row + 1

    def write_header(row):
        for col, val in enumerate(["专区", "订单数", "销售金额（元）"], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = header_font
            c.border = thin_border
        return row + 1

    def write_row(row, label, count, amount, bold=False):
        c1 = ws.cell(row=row, column=1, value=label)
        c2 = ws.cell(row=row, column=2, value=count)
        c3 = ws.cell(row=row, column=3, value=amount)
        c3.number_format = '#,##0.00'
        for c in (c1, c2, c3):
            c.border = thin_border
            if bold:
                c.font = Font(bold=True)
        return row + 1

    def write_section(row, title, supplier, zones):
        """写一个区块：标题 + 供应商表 + 专区表"""
        # 标题
        row = write_title(row, title)

        # 供应商类型表
        for col, val in enumerate(["供应商类型", "订单数", "销售金额（元）"], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = header_font
            c.border = thin_border
        row += 1
        data_start = row
        for stype, s in supplier.items():
            row = write_row(row, stype, s["order_count"], s["total_amount"])
        data_end = row - 1
        _write_total_row(ws, row, data_start, data_end, thin_border)
        row += 1

        # 专区表
        row = write_header(row)
        zone_items = {k: v for k, v in zones.items() if isinstance(v, dict) and k != "supplier"}
        data_start = row
        for zone, s in zone_items.items():
            row = write_row(row, zone, s["order_count"], s["total_amount"])
        data_end = row - 1
        _write_total_row(ws, row, data_start, data_end, thin_border)
        return row + 1

    row = 1
    for label, key in [("本周", "current_week"), ("本月", "current_month"), ("全量", "total")]:
        z = result["zones"][key]
        title = label
        if "range" in z:
            title += f"（{z['range']}）"
        elif "month" in z:
            title += f"（{z['month']}）"
        row = write_section(row, title, z["supplier"], z)

    # 列宽
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 20

    wb.save(xlsx_path)
    print(f"Excel已保存: {xlsx_path}")
    return xlsx_path


def export_monday_xlsx(result):
    """将周一统计结果导出为xlsx，格式参照模板"""
    timestamp = datetime.now().strftime('%Y%m%d%H%M')
    xlsx_path = os.path.join(RESULT_DIR, f'analysis_results_{timestamp}_M.xlsx')

    wb = Workbook()
    ws = wb.active
    ws.title = "订单统计"

    title_font = Font(bold=True, size=18)
    header_font = Font(bold=True, size=11)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    def write_title(row, text):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c = ws.cell(row=row, column=1, value=text)
        c.font = title_font
        c.alignment = Alignment(horizontal='center')
        return row + 1

    def write_header(row):
        for col, val in enumerate(["专区", "订单数", "销售金额（元）"], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = header_font
            c.border = thin_border
        return row + 1

    def write_row(row, label, count, amount, bold=False):
        ws.cell(row=row, column=1, value=label).border = thin_border
        ws.cell(row=row, column=2, value=count).border = thin_border
        c = ws.cell(row=row, column=3, value=amount)
        c.border = thin_border
        c.number_format = '#,##0.00'
        if bold:
            for cell in ws[row]:
                cell.font = Font(bold=True)
        return row + 1

    def write_section(row, title, zones):
        row = write_title(row, title)
        row = write_header(row)
        data_start = row
        for zone, s in zones.items():
            row = write_row(row, zone, s["order_count"], s["total_amount"])
        data_end = row - 1
        # 合计行用SUM公式
        _write_total_row(ws, row, data_start, data_end, thin_border)
        return row + 2  # +1 for next row, +1 for empty row

    row = 1
    for label, key in [("上周", "last_week"), ("本月", "current_month")]:
        r = result[key]
        title = label
        if "range" in r:
            title += f"（{r['range']}）专区统计"
        elif "month" in r:
            title += f"（{r['month']}）订单统计"
        row = write_section(row, title, r["zones"])

    # 全量统计（含供应商类型 + 专区）
    if "total_all" in result:
        t = result["total_all"]
        row = write_title(row, "全量统计")

        # 供应商类型表
        for col, val in enumerate(["供应商类型", "订单数", "销售金额（元）"], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = header_font
            c.border = thin_border
        row += 1
        data_start = row
        for stype, s in t["supplier"].items():
            row = write_row(row, stype, s["order_count"], s["total_amount"])
        data_end = row - 1
        _write_total_row(ws, row, data_start, data_end, thin_border)
        row += 1

        # 专区表
        row = write_header(row)
        data_start = row
        for zone, s in t["zones"].items():
            row = write_row(row, zone, s["order_count"], s["total_amount"])
        data_end = row - 1
        _write_total_row(ws, row, data_start, data_end, thin_border)
        row += 2

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 20

    wb.save(xlsx_path)
    print(f"Excel已保存: {xlsx_path}")
    return xlsx_path


def process_monday():
    """周一统计：上周、本月、总计（分专区）"""
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

    # 全量统计
    if "total_all" in result:
        t = result["total_all"]
        print(f"\n全量:")
        for stype, s in t["supplier"].items():
            print(f"  {stype}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")
        print(f"  专区统计:")
        for zone, s in t["zones"].items():
            print(f"    {zone}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")

    export_monday_xlsx(result)
    return result


def export_last_month_xlsx(result):
    """将上月统计结果导出为xlsx"""
    timestamp = datetime.now().strftime('%Y%m%d%H%M')
    xlsx_path = os.path.join(RESULT_DIR, f'analysis_results_{timestamp}_LM.xlsx')
    wb = Workbook()
    ws = wb.active
    ws.title = "上月统计"

    title_font = Font(bold=True, size=18)
    header_font = Font(bold=True, size=11)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    def write_title(row, text):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c = ws.cell(row=row, column=1, value=text)
        c.font = title_font
        c.alignment = Alignment(horizontal='center')
        return row + 1

    def write_header(row):
        for col, val in enumerate(["专区", "订单数", "销售金额（元）"], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = header_font
            c.border = thin_border
        return row + 1

    def write_row(row, label, count, amount, bold=False):
        c1 = ws.cell(row=row, column=1, value=label)
        c2 = ws.cell(row=row, column=2, value=count)
        c3 = ws.cell(row=row, column=3, value=amount)
        c3.number_format = '#,##0.00'
        for c in (c1, c2, c3):
            c.border = thin_border
            if bold:
                c.font = Font(bold=True)
        return row + 1

    row = 1
    month_label = result["month"]
    row = write_title(row, f"{month_label} 订单统计")

    # 供应商类型表
    for col, val in enumerate(["供应商类型", "订单数", "销售金额（元）"], 1):
        c = ws.cell(row=row, column=col, value=val)
        c.font = header_font
        c.border = thin_border
    row += 1
    data_start = row
    for stype, s in result["supplier"].items():
        row = write_row(row, stype, s["order_count"], s["total_amount"])
    data_end = row - 1
    _write_total_row(ws, row, data_start, data_end, thin_border)
    row += 1

    # 专区表
    row = write_header(row)
    data_start = row
    for zone, s in result["zones"].items():
        row = write_row(row, zone, s["order_count"], s["total_amount"])
    data_end = row - 1
    _write_total_row(ws, row, data_start, data_end, thin_border)

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 20

    wb.save(xlsx_path)
    print(f"Excel已保存: {xlsx_path}")
    return xlsx_path


def export_normal_xlsx(result):
    """将综合统计结果导出为xlsx"""
    timestamp = datetime.now().strftime('%Y%m%d%H%M')
    xlsx_path = os.path.join(RESULT_DIR, f'analysis_results_{timestamp}_NM.xlsx')
    wb = Workbook()
    ws = wb.active
    ws.title = "综合统计"

    title_font = Font(bold=True, size=18)
    header_font = Font(bold=True, size=11)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    def write_title(row, text):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c = ws.cell(row=row, column=1, value=text)
        c.font = title_font
        c.alignment = Alignment(horizontal='center')
        return row + 1

    def write_header(row):
        for col, val in enumerate(["专区", "订单数", "销售金额（元）"], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = header_font
            c.border = thin_border
        return row + 1

    def write_row(row, label, count, amount, bold=False):
        c1 = ws.cell(row=row, column=1, value=label)
        c2 = ws.cell(row=row, column=2, value=count)
        c3 = ws.cell(row=row, column=3, value=amount)
        c3.number_format = '#,##0.00'
        for c in (c1, c2, c3):
            c.border = thin_border
            if bold:
                c.font = Font(bold=True)
        return row + 1

    def write_section(row, title, supplier, zones):
        """写一个区块：标题 + 供应商表 + 专区表"""
        row = write_title(row, title)

        # 供应商类型表
        for col, val in enumerate(["供应商类型", "订单数", "销售金额（元）"], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = header_font
            c.border = thin_border
        row += 1
        data_start = row
        for stype, s in supplier.items():
            row = write_row(row, stype, s["order_count"], s["total_amount"])
        data_end = row - 1
        _write_total_row(ws, row, data_start, data_end, thin_border)
        row += 1

        # 专区表
        row = write_header(row)
        data_start = row
        for zone, s in zones.items():
            row = write_row(row, zone, s["order_count"], s["total_amount"])
        data_end = row - 1
        _write_total_row(ws, row, data_start, data_end, thin_border)
        return row + 2

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
        row = write_section(row, title, r["supplier"], r["zones"])

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 20

    wb.save(xlsx_path)
    print(f"Excel已保存: {xlsx_path}")
    return xlsx_path


def process_last_month():
    """上月统计：分供应商类型 + 分专区"""
    df = pd.read_excel(SOURCE_FILE)
    df['订单创建时间'] = pd.to_datetime(df['订单创建时间'])

    today = datetime.now()
    # 上个月：如果当前是1月，上个月是去年12月
    if today.month == 1:
        last_year, last_month = today.year - 1, 12
    else:
        last_year, last_month = today.year, today.month - 1

    mask = (df['订单创建时间'].dt.year == last_year) & (df['订单创建时间'].dt.month == last_month)
    month_data = df[mask]

    result = {
        "type": "last_month",
        "month": f"{last_year}-{last_month:02d}",
        "order_count": int(month_data['订单号'].nunique()),
        "total_amount": round(float(month_data['订单金额（元）'].sum()), 2),
        "supplier": supplier_stats(month_data),
        "zones": zone_stats(month_data)
    }

    print(f"{result['month']} 订单统计:")
    print(f"  总订单数: {result['order_count']}, 总销售额: {result['total_amount']}")
    print(f"  供应商类型:")
    for stype, s in result["supplier"].items():
        print(f"    {stype}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")
    print(f"  专区统计:")
    for zone, s in result["zones"].items():
        print(f"    {zone}: 订单数 {s['order_count']}, 销售额 {s['total_amount']}")

    export_last_month_xlsx(result)
    return result


def process_normal():
    """综合统计：上周、本周、上月、本月、全量（均分供应商类型+分专区）"""
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

    export_normal_xlsx(result)
    return result


def process_friday():
    """周五统计：本周汇总+供应商类型，分专区（本周/本月/全量）"""
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

    export_friday_xlsx(result)
    return result


def process(mode="auto"):
    """
    mode: "monday" | "friday" | "last_month" | "normal" | "auto"
      monday     - 强制跑周一统计（上周/本月/总计 + 分专区）
      friday     - 强制跑周五统计（本周 + 供应商类型 + 分专区）
      last_month - 上月统计（分供应商类型 + 分专区）
      normal     - 综合统计（上周/本周/上月/本月/全量，均分供应商+专区）
      auto       - 周一→monday，周五→friday，其他→normal
    """
    if mode == "monday":
        return process_monday()
    elif mode == "friday":
        return process_friday()
    elif mode == "last_month":
        return process_last_month()
    elif mode == "normal":
        return process_normal()
    else:
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        weekday = today.weekday()
        if weekday == 0:
            return process_monday()
        elif weekday == 4:
            return process_friday()
        else:
            return process_normal()


if __name__ == "__main__":
    import sys
    mode = sys.argv[1] if len(sys.argv) > 1 else "auto"
    process(mode)
