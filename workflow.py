# -*- coding: utf-8 -*-
"""
订单数据统计 workflow — 纯编排入口
周一 → monday_stats + feishu_sheet
周五 → friday_stats + feishu_sheet
日常 → normal_stats + feishu_sheet + email_notify + process_data
"""
import sys
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from notify.feishu_sheet import resolve_mode, update_online_sheet


def main(mode="auto"):
    resolved = resolve_mode(mode)

    # Step 1: 登录导出数据（失败重试最多2次，都失败则停止）
    print("=== Step 1: Login & Export ===")
    from xyt_export import main as login_export
    max_retries = 2
    for attempt in range(max_retries + 1):
        try:
            login_export()
            break  # 成功，跳出重试循环
        except Exception as e:
            if attempt < max_retries:
                print(f"[RETRY] 数据导出失败（第{attempt+1}次），正在重试...")
                print(f"[RETRY] 错误: {e}")
            else:
                print(f"[FATAL] 数据导出连续失败{max_retries+1}次，停止后续流程")
                print(f"[FATAL] 最后错误: {e}")
                sys.exit(1)

    # Step 1.5: 过滤测试数据
    print("\n=== Step 1.5: Filter Test Data ===")
    from filter_test_data import filter_test_data
    filter_test_data()

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

    # Step 4: 生成 dashboard.json（仅日常/综合模式）
    if resolved == "normal":
        print("\n=== Step 4: Process Data ===")
        from _ext.loader import run as process_data_run
        process_data_run()

    # Step 5: 更新飞书在线表格（所有模式都走）
    print("\n=== Step 5: Feishu Sheet ===")
    update_online_sheet(resolved, xlsx_path)

    print("\n=== All Done ===")


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "auto"
    main(mode)
