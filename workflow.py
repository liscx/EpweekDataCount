# -*- coding: utf-8 -*-
"""
订单数据统计 workflow — 纯编排入口
周一 → monday_stats + feishu_sheet
周五 → friday_stats + feishu_sheet
日常 → normal_stats + feishu_sheet + email_notify
"""
import sys
import os

from notify.feishu_sheet import resolve_mode, update_online_sheet


def main(mode="auto"):
    resolved = resolve_mode(mode)

    # Step 1: 登录导出数据
    print("=== Step 1: Login & Export ===")
    from xyt_export import main as login_export
    login_export()

    # Step 2: 按模式执行统计
    print(f"\n=== Step 2: Data Processing ({resolved}) ===")
    if resolved == "monday":
        from monday_stats import run
        result, xlsx_path = run()
    elif resolved == "friday":
        from friday_stats import run
        result, xlsx_path = run()
    else:
        from normal_stats import run
        result, xlsx_path = run()

    # Step 3: 邮件通知（仅日常/综合模式）
    if resolved == "normal":
        print("\n=== Step 3: Email Notify ===")
        from notify.email_notify import send_email
        send_email(xlsx_path, result)

    # Step 4: 更新飞书在线表格（所有模式都走）
    print("\n=== Step 4: Feishu Sheet ===")
    update_online_sheet(resolved, xlsx_path)

    print("\n=== All Done ===")


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "auto"
    main(mode)
