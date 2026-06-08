# -*- coding: utf-8 -*-
from xyt_export import main as login_export
from process_data import process
import sys
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RESULT_DIR = os.path.join(BASE_DIR, 'result')

MODE_SHEET_MAP = {
    "monday": ("周一统计", "M"),
    "friday": ("周五统计", "F"),
    "last_month": ("上月统计", "LM"),
    "normal": ("综合统计", "NM"),
}


def find_latest_result(mode):
    """根据 mode 找到最新的结果文件"""
    if mode == "auto":
        from datetime import datetime
        weekday = datetime.now().weekday()
        mode = "monday" if weekday == 0 else "friday" if weekday == 4 else "normal"

    if mode not in MODE_SHEET_MAP:
        return None

    _, suffix = MODE_SHEET_MAP[mode]
    pattern = f"_{suffix}.xlsx"
    matches = sorted([f for f in os.listdir(RESULT_DIR) if f.endswith(pattern)])
    if matches:
        return os.path.join(RESULT_DIR, matches[-1])
    return None


def update_online_sheet(mode, xlsx_path):
    """将单个模式的结果更新到飞书在线表格对应sheet"""
    if mode == "auto":
        from datetime import datetime
        weekday = datetime.now().weekday()
        mode = "monday" if weekday == 0 else "friday" if weekday == 4 else "normal"

    if mode not in MODE_SHEET_MAP:
        print(f"[SKIP] 未知模式: {mode}")
        return

    sheet_label, _ = MODE_SHEET_MAP[mode]
    print(f"\n=== 更新在线表格: {sheet_label} ===")

    try:
        from feishu_notify import _load_env, _get_tenant_token
        from run_all_modes import (
            load_or_create_spreadsheet, list_sheets, add_sheet,
            write_excel_to_sheet_formatted, send_file_to_feishu,
        )

        chat_id = os.environ.get("FEISHU_NOTIFY_CHAT_ID", "")

        app_id, app_secret = _load_env()
        token = _get_tenant_token(app_id, app_secret)

        # 加载或创建在线表格
        ss_token = load_or_create_spreadsheet(token)
        existing_sheets = list_sheets(token, ss_token)

        # 获取或创建对应sheet
        if sheet_label in existing_sheets:
            sheet_id = existing_sheets[sheet_label]
            print(f"  使用已有sheet: {sheet_id}")
        else:
            sheet_id = add_sheet(token, ss_token, sheet_label)
            print(f"  新建sheet: {sheet_id}")

        # 写入数据
        write_excel_to_sheet_formatted(token, ss_token, sheet_id, xlsx_path, mode)
        print(f"  [OK] {sheet_label} 已更新")

        # 发送结果文件
        if chat_id:
            send_file_to_feishu(xlsx_path, f"{sheet_label}结果", chat_id)

        # 发送在线表格链接
        sheet_url = f"https://icnqngo556na.feishu.cn/sheets/{ss_token}"
        if chat_id:
            from feishu_notify import send_text
            send_text(f"在线表格已更新（{sheet_label}）: {sheet_url}", chat_id)
        print(f"  在线表格: {sheet_url}")

    except Exception as e:
        print(f"[WARN] 更新在线表格失败: {e}")
        import traceback
        traceback.print_exc()


def main(mode="auto"):
    print("=== Step 1: Login & Export ===")
    login_export()

    print("\n=== Step 2: Data Processing ===")
    process(mode)

    # 找到结果文件并更新在线表格
    xlsx_path = find_latest_result(mode)
    if xlsx_path:
        resolved = mode
        if mode == "auto":
            from datetime import datetime
            weekday = datetime.now().weekday()
            resolved = "monday" if weekday == 0 else "friday" if weekday == 4 else "normal"
        update_online_sheet(resolved, xlsx_path)
    else:
        print("[WARN] 未找到结果文件，跳过在线表格更新")

    print("\n=== All Done ===")


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "auto"
    main(mode)
