# -*- coding: utf-8 -*-
"""过滤测试订单数据：将测试数据移到单独的sheet，返回干净数据"""
import pandas as pd
import os
from openpyxl import Workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_FILE = os.path.join(BASE_DIR, 'Data', 'source_data.xlsx')


def is_test_order(row):
    """判断是否为测试订单"""
    # 采购企业包含"测试"
    if pd.notna(row.get('采购企业')) and '测试' in str(row['采购企业']):
        return True

    # 采购部门包含"测试"
    if pd.notna(row.get('采购部门')) and '测试' in str(row['采购部门']):
        return True

    # 采购部门为"系统管理部"
    if pd.notna(row.get('采购部门')) and str(row['采购部门']).strip() == '系统管理部':
        return True

    # 收货地址包含"测试"、"国泰测试"、"测试地址"
    if pd.notna(row.get('收货地址')):
        addr = str(row['收货地址'])
        if '测试' in addr or '国泰测试' in addr or '测试地址' in addr:
            return True

    return False


def filter_test_data():
    """过滤测试数据，将测试数据写入"测试数据"sheet，返回干净的DataFrame"""
    print("=== 过滤测试订单 ===")

    # 读取源数据
    df = pd.read_excel(SOURCE_FILE)
    total_before = len(df)
    print(f"原始数据行数: {total_before}")

    # 标记测试订单
    test_mask = df.apply(is_test_order, axis=1)
    test_data = df[test_mask]
    clean_data = df[~test_mask]

    test_count = len(test_data)
    clean_count = len(clean_data)
    print(f"测试订单数: {test_count}")
    print(f"有效订单数: {clean_count}")

    # 创建新的工作簿
    wb = Workbook()

    # 写入干净数据到Sheet1
    ws_clean = wb.active
    ws_clean.title = 'Sheet1'

    # 写入表头
    for col_idx, col_name in enumerate(clean_data.columns, 1):
        ws_clean.cell(row=1, column=col_idx, value=col_name)

    # 写入干净数据
    for row_idx, (_, row) in enumerate(clean_data.iterrows(), 2):
        for col_idx, value in enumerate(row, 1):
            ws_clean.cell(row=row_idx, column=col_idx, value=value)

    # 如果有测试数据，写入"测试数据"sheet
    if test_count > 0:
        ws_test = wb.create_sheet('测试数据')

        # 写入表头
        for col_idx, col_name in enumerate(test_data.columns, 1):
            ws_test.cell(row=1, column=col_idx, value=col_name)

        # 写入测试数据
        for row_idx, (_, row) in enumerate(test_data.iterrows(), 2):
            for col_idx, value in enumerate(row, 1):
                ws_test.cell(row=row_idx, column=col_idx, value=value)

        print(f"测试数据已保存到 '测试数据' sheet")

    # 保存文件
    wb.save(SOURCE_FILE)
    print(f"source_data.xlsx 已更新，仅保留有效数据")

    return clean_data


if __name__ == "__main__":
    filter_test_data()
