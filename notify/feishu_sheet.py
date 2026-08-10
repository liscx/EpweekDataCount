# -*- coding: utf-8 -*-
"""飞书在线表格更新 + 消息通知模块"""
import os
import sys

MODE_SHEET_MAP = {
    "monday": ("周一统计", "M"),
    "friday": ("周五统计", "F"),
    "last_month": ("上月统计", "LM"),
    "normal": ("综合统计", "NM"),
    "gys_count": ("供应商统计", "GYS"),
}


def resolve_mode(mode):
    """将 auto 模式解析为具体模式"""
    if mode == "auto":
        from datetime import datetime
        weekday = datetime.now().weekday()
        return "monday" if weekday == 0 else "friday" if weekday == 4 else "normal"
    return mode


def update_online_sheet(mode, xlsx_path):
    """将单个模式的结果更新到飞书在线表格对应sheet"""
    mode = resolve_mode(mode)

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
