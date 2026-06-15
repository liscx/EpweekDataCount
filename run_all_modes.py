# -*- coding: utf-8 -*-
"""
全模式执行脚本：
1. 运行四种统计模式（monday/friday/last_month/normal）
2. 发送结果文件和源文件到飞书
3. 创建/更新飞书在线表格（每种模式一个sheet，带格式美化）
"""
import os
import sys
import json
import pandas as pd
import urllib.request
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from monday_stats import run as monday_run
from friday_stats import run as friday_run
from normal_stats import run as normal_run
from feishu_notify import _load_env, _get_tenant_token, send_text
from stats_utils import (
    zone_stats, supplier_stats,
    make_xlsx, save_xlsx, set_col_widths,
    write_section_with_supplier,
)

SOURCE_FILE = os.path.join(BASE_DIR, 'Data', 'source_data.xlsx')
RESULT_DIR = os.path.join(BASE_DIR, 'result')
CHAT_ID = os.environ.get("FEISHU_NOTIFY_CHAT_ID", "")
DOMAIN = "https://open.feishu.cn"
SPREADSHEET_TOKEN_FILE = os.path.join(BASE_DIR, 'spreadsheet_token.json')


def last_month_run():
    """上月统计：仅上月数据，返回 (result, xlsx_path)"""
    import pandas as pd
    from datetime import datetime

    df = pd.read_excel(SOURCE_FILE)
    df['订单创建时间'] = pd.to_datetime(df['订单创建时间'])

    now = datetime.now()
    current_year, current_month = now.year, now.month
    if current_month == 1:
        last_year, last_month = current_year - 1, 12
    else:
        last_year, last_month = current_year, current_month - 1

    mask = (df['订单创建时间'].dt.year == last_year) & (df['订单创建时间'].dt.month == last_month)
    data = df[mask]

    result = {
        "type": "last_month",
        "last_month": {
            "month": f"{last_year}-{last_month:02d}",
            "supplier": supplier_stats(data),
            "zones": zone_stats(data)
        }
    }

    wb, ws, xlsx_path = make_xlsx("LM", "上月统计")
    r = result["last_month"]
    title = f"上月（{r['month']}）"
    row = write_section_with_supplier(ws, 1, title, r["supplier"], r["zones"])
    set_col_widths(ws)
    save_xlsx(wb, xlsx_path)

    total_count = sum(s["order_count"] for s in r["supplier"].values())
    total_amount = round(sum(s["total_amount"] for s in r["supplier"].values()), 2)
    print(f"上月（{r['month']}）: 订单数 {total_count}, 销售额 {total_amount}")

    return result, xlsx_path

MODES = [
    ("monday", "周一统计"),
    ("friday", "周五统计"),
    ("last_month", "上月统计"),
    ("normal", "综合统计"),
]

# 样式常量
STYLE_TITLE = {
    "font": {"bold": False},
    "hAlign": 1,
    "borderType": "FULL_BORDER",
    "backColor": "#E8F0FE",
}
STYLE_HEADER = {
    "font": {"bold": True},
    "hAlign": 1,
    "borderType": "FULL_BORDER",
    "backColor": "#FFF9C4",
}
STYLE_DATA = {
    "hAlign": 2,
    "borderType": "FULL_BORDER",
}
STYLE_TOTAL = {
    "font": {"bold": True},
    "hAlign": 2,
    "borderType": "FULL_BORDER",
}
STYLE_AMOUNT = {
    "hAlign": 2,
    "borderType": "FULL_BORDER",
    "formatter": "#,##0.00",
}


def _api_request(url, token, data=None, method="GET"):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    if data is not None:
        body = json.dumps(data).encode()
        req = urllib.request.Request(url, data=body, headers=headers, method=method)
    else:
        req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())


def col_letter(n):
    """数字转列字母，1->A, 2->B, ..."""
    result = ""
    while n > 0:
        n -= 1
        result = chr(65 + n % 26) + result
        n //= 26
    return result


def create_spreadsheet(token, title, folder_token=None):
    url = f"{DOMAIN}/open-apis/sheets/v3/spreadsheets"
    data = {"title": title}
    if folder_token:
        data["folder_token"] = folder_token
    result = _api_request(url, token, data, method="POST")
    if result.get("code") != 0:
        raise RuntimeError(f"创建表格失败: {result}")
    return result["data"]["spreadsheet"]["spreadsheet_token"]


def add_sheet(token, spreadsheet_token, title):
    url = f"{DOMAIN}/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/sheets_batch_update"
    data = {"requests": [{"addSheet": {"properties": {"title": title}}}]}
    result = _api_request(url, token, data, method="POST")
    if result.get("code") != 0:
        raise RuntimeError(f"添加sheet失败: {result}")
    return result["data"]["replies"][0]["addSheet"]["properties"]["sheetId"]


def list_sheets(token, spreadsheet_token):
    url = f"{DOMAIN}/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query"
    result = _api_request(url, token)
    if result.get("code") != 0:
        raise RuntimeError(f"查询sheets失败: {result}")
    return {s["title"]: s["sheet_id"] for s in result["data"]["sheets"]}


def apply_style(token, spreadsheet_token, range_str, style):
    """对指定范围应用样式"""
    url = f"{DOMAIN}/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/style"
    data = {"appendStyle": {"range": range_str, "style": style}}
    result = _api_request(url, token, data, method="PUT")
    if result.get("code") != 0:
        print(f"  [WARN] 样式设置失败: {result.get('msg')}")


def set_column_width(token, spreadsheet_token, sheet_id, col_start, col_end, width):
    """设置列宽"""
    url = f"{DOMAIN}/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/dimension_range"
    data = {
        "dimension": {
            "sheetId": sheet_id,
            "majorDimension": "COLUMNS",
            "startIndex": col_start,
            "endIndex": col_end,
        },
        "dimensionProperties": {"fixedSize": width},
    }
    _api_request(url, token, data, method="PUT")


def merge_cells(token, spreadsheet_token, sheet_id, start_row, start_col, end_row, end_col):
    """合并单元格"""
    url = f"{DOMAIN}/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/sheets_batch_update"
    data = {
        "requests": [
            {
                "mergeCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_row,
                        "endRowIndex": end_row,
                        "startColumnIndex": start_col,
                        "endColumnIndex": end_col,
                    },
                    "mergeType": "MERGE_ALL",
                }
            }
        ]
    }
    _api_request(url, token, data, method="POST")


def write_values(token, spreadsheet_token, range_str, values):
    """写入数据"""
    url = f"{DOMAIN}/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values_batch_update"
    data = {"valueRanges": [{"range": range_str, "values": values}]}
    result = _api_request(url, token, data, method="POST")
    if result.get("code") != 0:
        raise RuntimeError(f"写入失败: {result}")


def write_section(token, ss_token, sheet_id, row, title, df, has_supplier=False):
    """
    写入一个统计区块：标题行 + 表头 + 数据行 + 合计行
    返回下一个可用行号
    """
    ncols = len(df.columns)
    end_col_letter = col_letter(ncols)

    # 1. 标题行（合并单元格）
    merge_cells(token, ss_token, sheet_id, row, 0, row + 1, ncols)
    write_values(token, ss_token, f"{sheet_id}!A{row+1}:{end_col_letter}{row+1}", [[title] + [""] * (ncols - 1)])
    apply_style(token, ss_token, f"{sheet_id}!A{row+1}:{end_col_letter}{row+1}", STYLE_TITLE)
    row += 2

    # 2. 表头行
    headers = [list(df.columns)]
    write_values(token, ss_token, f"{sheet_id}!A{row+1}:{end_col_letter}{row+1}", headers)
    apply_style(token, ss_token, f"{sheet_id}!A{row+1}:{end_col_letter}{row+1}", STYLE_HEADER)
    row += 1

    # 3. 数据行
    data_start = row
    for _, r in df.iterrows():
        values_row = []
        for v in r:
            if pd.notna(v):
                values_row.append(str(v) if not isinstance(v, (int, float)) else v)
            else:
                values_row.append("")
        write_values(token, ss_token, f"{sheet_id}!A{row+1}:{end_col_letter}{row+1}", [values_row])

        # 样式：第一列左对齐，数字列右对齐
        apply_style(token, ss_token, f"{sheet_id}!A{row+1}:A{row+1}", STYLE_DATA)
        if ncols >= 3:
            apply_style(token, ss_token, f"{sheet_id}!B{row+1}:{end_col_letter}{row+1}", STYLE_AMOUNT)
        row += 1

    # 4. 合计行
    total_row_values = ["合计"] + [""] * (ncols - 1)
    write_values(token, ss_token, f"{sheet_id}!A{row+1}:{end_col_letter}{row+1}", [total_row_values])
    # 用 SUM 公式
    for col_idx in range(1, ncols):
        c = col_letter(col_idx + 1)
        formula = f"=SUM({c}{data_start+1}:{c}{row})"
        write_values(token, ss_token, f"{sheet_id}!{c}{row+1}:{c}{row+1}", [[formula]])
    apply_style(token, ss_token, f"{sheet_id}!A{row+1}:{end_col_letter}{row+1}", STYLE_TOTAL)
    row += 2  # 空一行

    return row


def write_excel_to_sheet_formatted(token, ss_token, sheet_id, excel_path, mode):
    """
    将 Excel 结果文件按原始结构写入在线表格，保留数据类型和公式。
    批量操作减少 API 调用次数。
    """
    import openpyxl

    wb = openpyxl.load_workbook(excel_path)  # 保留公式
    ws = wb.active

    ncols = 3
    end_col = col_letter(ncols)

    # 设置列宽
    set_column_width(token, ss_token, sheet_id, 1, 2, 240)
    set_column_width(token, ss_token, sheet_id, 2, 3, 130)
    set_column_width(token, ss_token, sheet_id, 3, 4, 170)

    # 读取原始 Excel 数据
    all_rows = []
    for row in ws.iter_rows(min_col=1, max_col=ncols, values_only=True):
        all_rows.append(list(row))

    if not all_rows:
        print(f"  [WARN] Excel 为空")
        return

    # 写入数据 - 保留原始类型
    write_data = []
    for row in all_rows:
        out_row = []
        for v in row:
            if v is None:
                out_row.append("")
            elif isinstance(v, str) and v.startswith("="):
                out_row.append(v)
            elif isinstance(v, (int, float)):
                out_row.append(v)
            else:
                out_row.append(str(v))
        write_data.append(out_row)

    total_rows = len(write_data)

    # 先清除旧数据（删除多余行）
    try:
        info_url = f"{DOMAIN}/open-apis/sheets/v2/spreadsheets/{ss_token}/values/{sheet_id}"
        info_resp = _api_request(info_url, token)
        if info_resp.get("code") == 0:
            old_rows = len(info_resp["data"]["valueRange"].get("values", []))
            if old_rows > total_rows:
                # 删除多余行
                del_url = f"{DOMAIN}/open-apis/sheets/v2/spreadsheets/{ss_token}/dimension_range"
                del_data = {
                    "dimension": {
                        "sheetId": sheet_id,
                        "majorDimension": "ROWS",
                        "startIndex": total_rows,
                        "endIndex": old_rows
                    }
                }
                _api_request(del_url, token, del_data, method="DELETE")
                print(f"  [OK] 已删除旧数据 {old_rows - total_rows} 行")
    except Exception as e:
        print(f"  [WARN] 清除旧数据失败: {e}")

    write_values(token, ss_token, f"{sheet_id}!A1:{end_col}{total_rows}", write_data)
    print(f"  [OK] 已写入 {total_rows} 行数据")

    # 先清除整个数据区域的旧样式，防止残留
    clear_style = {
        "font": {"bold": False},
        "hAlign": 0,
        "backColor": "#FFFFFF",
        "borderType": "FULL_BORDER",
        "borderColor": "#FFFFFF",
    }
    apply_style(token, ss_token, f"{sheet_id}!A1:{end_col}{total_rows + 5}", clear_style)

    # 收集需要样式的行范围
    title_ranges = []   # 标题行
    header_ranges = []  # 表头行
    total_ranges = []   # 合计行
    data_col_b = []     # 数据行 B 列
    data_col_c = []     # 数据行 C 列

    for row_idx, row_data in enumerate(all_rows):
        r = row_idx + 1  # 1-indexed
        first_cell = str(row_data[0]).strip() if row_data[0] else ""

        if not first_cell and all(v is None for v in row_data):
            continue

        # 标题行
        if first_cell and (len(row_data) < 2 or row_data[1] is None):
            title_ranges.append(r)
            continue

        # 表头行
        if first_cell in ("专区", "供应商类型"):
            header_ranges.append(r)
            continue

        # 合计行
        if first_cell == "合计":
            total_ranges.append(r)
            continue

        # 普通数据行
        data_col_b.append(r)
        data_col_c.append(r)

    # 批量应用样式（每种样式一次调用）
    for r in title_ranges:
        apply_style(token, ss_token, f"{sheet_id}!A{r}:{end_col}{r}", STYLE_TITLE)
    for r in header_ranges:
        apply_style(token, ss_token, f"{sheet_id}!A{r}:{end_col}{r}", STYLE_HEADER)
    for r in total_ranges:
        apply_style(token, ss_token, f"{sheet_id}!A{r}:{end_col}{r}", STYLE_TOTAL)
    # 数据行：整行应用数据样式，金额列额外应用数字格式
    for r in data_col_b:
        apply_style(token, ss_token, f"{sheet_id}!A{r}:B{r}", STYLE_DATA)
    for r in data_col_c:
        apply_style(token, ss_token, f"{sheet_id}!C{r}:C{r}", STYLE_AMOUNT)

    print(f"  [OK] 样式已应用（标题{len(title_ranges)}行, 表头{len(header_ranges)}行, 合计{len(total_ranges)}行）")


def _upload_file(token, file_path):
    url = f"{DOMAIN}/open-apis/im/v1/files"
    boundary = "----HermesBoundary"
    filename = os.path.basename(file_path)
    with open(file_path, "rb") as f:
        file_data = f.read()
    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file_type"\r\n\r\n'
        f"stream\r\n"
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file_name"\r\n\r\n'
        f"{filename}\r\n"
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
        f"Content-Type: application/octet-stream\r\n\r\n"
    ).encode() + file_data + f"\r\n--{boundary}--\r\n".encode()
    req = urllib.request.Request(url, data=body, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": f"multipart/form-data; boundary={boundary}",
    })
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = json.loads(resp.read())
    if data.get("code") != 0:
        raise RuntimeError(f"上传文件失败: {data}")
    return data["data"]["file_key"]


def _send_msg(token, receive_id, msg_type, content, id_type="chat_id"):
    url = f"{DOMAIN}/open-apis/im/v1/messages?receive_id_type={id_type}"
    body = json.dumps({"receive_id": receive_id, "msg_type": msg_type, "content": content}).encode()
    req = urllib.request.Request(url, data=body, headers={
        "Authorization": f"Bearer {token}", "Content-Type": "application/json",
    })
    with urllib.request.urlopen(req, timeout=10) as resp:
        data = json.loads(resp.read())
    if data.get("code") != 0:
        raise RuntimeError(f"发送消息失败: {data}")


def send_file_to_feishu(file_path, text="", chat_id=""):
    app_id, app_secret = _load_env()
    token = _get_tenant_token(app_id, app_secret)
    if chat_id.startswith("ou_"):
        receive_id, id_type = chat_id, "open_id"
    elif chat_id:
        receive_id, id_type = chat_id, "chat_id"
    else:
        from feishu_notify import _resolve_target
        receive_id, id_type = _resolve_target("")
    if text:
        _send_msg(token, receive_id, "text", json.dumps({"text": text}), id_type)
    file_key = _upload_file(token, file_path)
    _send_msg(token, receive_id, "file", json.dumps({"file_key": file_key}), id_type)
    print(f"[飞书] 文件已发送: {os.path.basename(file_path)}")


def set_public_permission(token, spreadsheet_token):
    """设置表格为任何人可编辑"""
    url = f"{DOMAIN}/open-apis/drive/v1/permissions/{spreadsheet_token}/public?type=sheet"
    data = {
        "external_access_entity": "open",
        "security_entity": "anyone_can_view",
        "comment_entity": "anyone_can_view",
        "share_entity": "anyone",
        "manage_collaborator_entity": "collaborator_can_view",
        "link_share_entity": "anyone_editable",
    }
    try:
        _api_request(url, token, data, method="PATCH")
        print("[权限] 已设置为任何人可编辑")
    except Exception as e:
        print(f"[WARN] 设置权限失败: {e}")


def load_or_create_spreadsheet(token):
    if os.path.exists(SPREADSHEET_TOKEN_FILE):
        with open(SPREADSHEET_TOKEN_FILE, 'r') as f:
            info = json.load(f)
            print(f"复用已有在线表格: {info['token']}")
            ss_token = info['token']
    else:
        title = f"阳光优采订单统计_{datetime.now().strftime('%Y%m%d')}"
        ss_token = create_spreadsheet(token, title)
        print(f"创建新在线表格: {ss_token}")
        with open(SPREADSHEET_TOKEN_FILE, 'w') as f:
            json.dump({"token": ss_token, "title": title, "created": datetime.now().isoformat()}, f)
    # 确保权限为任何人可编辑
    set_public_permission(token, ss_token)
    return ss_token


def main():
    print("=" * 60)
    print("阳光优采订单统计 - 全模式执行")
    print("=" * 60)

    if not os.path.exists(SOURCE_FILE):
        print(f"[ERROR] 源文件不存在: {SOURCE_FILE}")
        sys.exit(1)

    app_id, app_secret = _load_env()
    token = _get_tenant_token(app_id, app_secret)

    # 1. 运行四种模式
    results = {}
    for mode, label in MODES:
        print(f"\n{'='*40}")
        print(f"[执行] {label} ({mode})")
        print(f"{'='*40}")
        try:
            if mode == "monday":
                result, _ = monday_run()
            elif mode == "friday":
                result, _ = friday_run()
            elif mode == "last_month":
                result, _ = last_month_run()
            else:
                result, _ = normal_run()
            results[mode] = result
            print(f"[OK] {label} 完成")
        except Exception as e:
            print(f"[FAIL] {label} 失败: {e}")
            results[mode] = None

    # 2. 发送文件
    print(f"\n{'='*40}")
    print("[发送] 文件到飞书")
    print(f"{'='*40}")

    if os.path.exists(SOURCE_FILE):
        send_file_to_feishu(SOURCE_FILE, "源数据文件", CHAT_ID)

    # 只发送最新4个结果文件
    suffix_map = {"monday": "M", "friday": "F", "last_month": "LM", "normal": "NM"}
    latest_files = []
    for mode, label in MODES:
        suffix = suffix_map[mode]
        pattern = f"_{suffix}.xlsx"
        matches = sorted([f for f in os.listdir(RESULT_DIR) if f.endswith(pattern)])
        if matches:
            latest_files.append((matches[-1], label))

    for f, mode_label in latest_files:
        fp = os.path.join(RESULT_DIR, f)
        send_file_to_feishu(fp, f"{mode_label}结果", CHAT_ID)

    # 3. 更新在线表格（带格式美化）
    print(f"\n{'='*40}")
    print("[更新] 飞书在线表格（格式美化）")
    print(f"{'='*40}")

    ss_token = load_or_create_spreadsheet(token)
    existing_sheets = list_sheets(token, ss_token)
    print(f"现有sheets: {list(existing_sheets.keys())}")

    for mode, label in MODES:
        print(f"\n[处理] {label}...")
        if label in existing_sheets:
            sheet_id = existing_sheets[label]
            print(f"  使用已有sheet: {sheet_id}")
        else:
            sheet_id = add_sheet(token, ss_token, label)
            print(f"  新建sheet: {sheet_id}")
            existing_sheets[label] = sheet_id

        # 找最新结果文件
        suffix = suffix_map[mode]
        result_file = None
        matches = sorted([f for f in os.listdir(RESULT_DIR) if f.endswith(f"_{suffix}.xlsx")])
        if matches:
            result_file = os.path.join(RESULT_DIR, matches[-1])

        if not result_file:
            print(f"  [WARN] 未找到结果文件，跳过")
            continue

        try:
            write_excel_to_sheet_formatted(token, ss_token, sheet_id, result_file, mode)
        except Exception as e:
            print(f"  [FAIL] 写入失败: {e}")
            import traceback
            traceback.print_exc()

    # 删除默认 Sheet1（如果存在）
    if "Sheet1" in existing_sheets:
        try:
            url = f"{DOMAIN}/open-apis/sheets/v2/spreadsheets/{ss_token}/sheets_batch_update"
            data = {"requests": [{"deleteSheet": {"sheetId": existing_sheets["Sheet1"]}}]}
            _api_request(url, token, data, method="POST")
            print("[清理] 已删除默认 Sheet1")
        except Exception:
            pass

    sheet_url = f"https://icnqngo556na.feishu.cn/sheets/{ss_token}"
    # 发送链接
    app_id2, app_secret2 = _load_env()
    token2 = _get_tenant_token(app_id2, app_secret2)
    if CHAT_ID.startswith("ou_"):
        _send_msg(token2, CHAT_ID, "text", json.dumps({"text": f"在线表格已更新（含格式美化）: {sheet_url}"}), "open_id")
    else:
        send_text(f"在线表格已更新（含格式美化）: {sheet_url}", CHAT_ID)

    print(f"\n{'='*60}")
    print("[DONE] 全部完成！")
    print(f"在线表格: {sheet_url}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
