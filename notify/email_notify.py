# -*- coding: utf-8 -*-
"""邮件通知模块"""
import os
import smtplib
import yaml
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_FILE = os.path.join(BASE_DIR, 'config.yaml')
ENV_FILE = os.path.join(BASE_DIR, '.env')

load_dotenv(ENV_FILE)


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
    smtp_pass = os.environ.get("SMTP_PASS", "")
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
