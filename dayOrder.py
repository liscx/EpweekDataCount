# -*- coding: utf-8 -*-
"""
综合统计（normal模式）独立脚本
功能：上周、本周、上月、本月、全量（均分供应商类型+分专区）统计
"""
import pandas as pd
import os
import smtplib
import yaml
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_FILE = os.path.join(BASE_DIR, 'Data', 'source_data.xlsx')
RESULT_DIR = os.path.join(BASE_DIR, 'result')
CONFIG_FILE = os.path.join(BASE_DIR, 'config.yaml')
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
    for label, key in [("本周", "current_week"), ("本月", "current_month"),
                        ("上周", "last_week"), ("上月", "last_month"),
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


def send_email(xlsx_path, result):
    """发送统计结果邮件，配置从 config.yaml 读取"""
    if not os.path.exists(CONFIG_FILE):
        print(f"[WARN] 配置文件不存在: {CONFIG_FILE}，跳过发送")
        return

    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)

    email_cfg = config.get("email", {})
    smtp_host = email_cfg.get("smtp_host", "smtp.qq.com")
    smtp_port = int(email_cfg.get("smtp_port", 465))
    smtp_user = email_cfg.get("smtp_user", "")
    smtp_pass = email_cfg.get("smtp_pass", "")
    to_list = email_cfg.get("to", [])

    if not all([smtp_user, smtp_pass, to_list]):
        print("[WARN] 邮件配置不完整（smtp_user/smtp_pass/to），跳过发送")
        return

    filename = os.path.basename(xlsx_path)
    today_str = datetime.now().strftime('%Y-%m-%d')

    # 构建邮件正文
    lines = [f"综合统计 {today_str}", "=" * 40]
    for label, key in [("本周", "current_week"), ("本月", "current_month"),
                        ("上周", "last_week"), ("上月", "last_month"),
                        ("全量", "total")]:
        r = result[key]
        header = label
        if "range" in r:
            header += f"（{r['range']}）"
        elif "month" in r:
            header += f"（{r['month']}）"
        total_count = sum(s["order_count"] for s in r["supplier"].values())
        total_amount = round(sum(s["total_amount"] for s in r["supplier"].values()), 2)
        lines.append(f"\n【{header}】订单数 {total_count}，销售额 {total_amount}")
        for stype, s in r["supplier"].items():
            lines.append(f"  {stype}: 订单数 {s['order_count']}，销售额 {s['total_amount']}")
        lines.append("  专区统计:")
        for zone, s in r["zones"].items():
            lines.append(f"    {zone}: 订单数 {s['order_count']}，销售额 {s['total_amount']}")
    lines.append("\n-- 自动发送")
    body = "\n".join(lines)

    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = ", ".join(to_list)
    msg["Subject"] = f"综合统计日报 {today_str}"

    msg.attach(MIMEText(body, "plain", "utf-8"))

    with open(xlsx_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)

    try:
        server = smtplib.SMTP_SSL(smtp_host, smtp_port)
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, to_list, msg.as_string())
        server.quit()
        print(f"[OK] 邮件已发送至: {', '.join(to_list)}")
    except Exception as e:
        print(f"[ERROR] 邮件发送失败: {e}")


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

    for label, key in [("本周", "current_week"), ("本月", "current_month"),
                        ("上周", "last_week"), ("上月", "last_month"),
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

    xlsx_path = export_normal_xlsx(result)
    send_email(xlsx_path, result)
    return result


if __name__ == "__main__":
    process_normal()
