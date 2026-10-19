# -*- coding: utf-8 -*-
"""
订单数据统计 workflow — 纯编排入口
周一 → monday_stats + feishu_sheet
日常 → normal_stats + feishu_sheet + email_notify + process_data
       （当天为周五时额外执行 gys_count）

注：friday 模式已弃用（friday_stats 不再调用），周五按 normal 处理，
    供应商统计 gys_count 由 normal 分支内的周五判断触发。
"""
import sys
import os
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from notify.feishu_sheet import resolve_mode, update_online_sheet


def main(mode="auto"):
    resolved = resolve_mode(mode)

    # friday 模式已弃用：统一按综合统计（normal）处理
    if resolved == "friday":
        print("[INFO] friday 模式已弃用，转为 normal（综合统计）")
        resolved = "normal"

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
    gys_xlsx_path = None
    if resolved == "monday":
        from monday_stats import run
        result, xlsx_path = run()
    # [弃用] friday 模式，改为统一走 normal
    # elif resolved == "friday":
    #     from friday_stats import run
    #     result, xlsx_path = run()
    #     # 供应商统计 Top10（周五额外执行）
    #     print(f"\n=== Step 2.5: GYS Count ===")
    #     from gys_count import run as gys_run
    #     gys_result, gys_xlsx_path = gys_run()
    else:
        from normal_stats import run
        result, xlsx_path = run()
        # 供应商统计 Top10（综合模式 + 当天为周五时额外执行）
        if datetime.now().weekday() == 4:
            print(f"\n=== Step 2.5: GYS Count ===")
            try:
                from gys_count import run as gys_run
                gys_result, gys_xlsx_path = gys_run()
            except Exception as e:
                # 附属统计失败不阻断主流程（邮件/dashboard/飞书仍要正常走）
                gys_xlsx_path = None
                print(f"[WARN] 供应商统计执行失败，跳过（不影响主流程）: {e}")

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
    # 供应商统计表（本次跑过 gys_count 才更新）
    if gys_xlsx_path:
        update_online_sheet("gys_count", gys_xlsx_path)

    print("\n=== All Done ===")


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "auto"
    main(mode)
