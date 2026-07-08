import pandas as pd
import json
import os
import yaml
from datetime import datetime
from pathlib import Path

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 读取配置
with open(os.path.join(BASE_DIR, 'config.yaml'), 'r', encoding='utf-8') as f:
    _config = yaml.safe_load(f)
OUTPUT_DIR = _config.get('dataview', os.path.join(BASE_DIR, 'result'))


def run():
    # 读取 Excel（来自 xyt_export 导出的数据）
    df = pd.read_excel(os.path.join(BASE_DIR, 'Data', 'source_data.xlsx'))

    # 按订单号去重，取订单首行金额
    orders = df.drop_duplicates(subset='订单号', keep='first').copy()

    # 基础时间
    now = datetime.now()
    data_end_date = orders['订单创建时间'].max()[:10] if len(orders) > 0 else now.strftime('%Y-%m-%d')

    # ============ KPI 指标 ============
    total_amount = round(orders['订单金额（元）'].sum(), 2)
    total_orders = len(orders)
    avg_amount = round(total_amount / total_orders, 2) if total_orders > 0 else 0

    # 本周数据（近7天）
    if len(orders) > 0:
        orders['日期'] = pd.to_datetime(orders['订单创建时间'])
        week_ago = orders['日期'].max() - pd.Timedelta(days=7)
        week_orders = orders[orders['日期'] >= week_ago]
        week_amount = round(week_orders['订单金额（元）'].sum(), 2)
        week_count = len(week_orders)
    else:
        week_amount = 0
        week_count = 0

    # 活跃专区数
    zone_count = orders['专区名称'].nunique()

    # 采购企业数
    buyer_count = orders['采购企业'].nunique()

    # 供应商数
    supplier_count = orders['供应商'].nunique()

    # ============ 专区排行 ============
    zone_rank = orders.groupby('专区名称').agg(
        count=('订单号', 'count'),
        amount=('订单金额（元）', 'sum')
    ).reset_index()
    zone_rank['amount'] = zone_rank['amount'].round(2)
    zone_rank = zone_rank.sort_values('amount', ascending=False)
    zone_rank_list = zone_rank.rename(columns={'专区名称': 'name'}).to_dict('records')

    # ============ 供应商类型 ============
    supplier_types = orders.groupby('供应商类型').agg(
        count=('订单号', 'count'),
        amount=('订单金额（元）', 'sum')
    ).reset_index()
    supplier_types['amount'] = supplier_types['amount'].round(2)
    supplier_types_list = supplier_types.rename(columns={'供应商类型': 'name'}).to_dict('records')

    # ============ 月度趋势 ============
    if len(orders) > 0:
        orders['月份'] = orders['日期'].dt.month.apply(lambda x: f"{int(x)}月" if pd.notna(x) else "")
        month_trend = orders.groupby('月份').agg(
            count=('订单号', 'count'),
            amount=('订单金额（元）', 'sum')
        ).reset_index()
        month_trend['amount'] = month_trend['amount'].round(2)
        # 按月份排序
        month_order = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']
        month_trend['sort_key'] = month_trend['月份'].apply(lambda x: month_order.index(x) if x in month_order else 99)
        month_trend = month_trend.sort_values('sort_key')
        month_trend_list = month_trend.rename(columns={'月份': 'label'}).drop(columns=['sort_key']).to_dict('records')
    else:
        month_trend_list = []

    # ============ 近一周趋势 ============
    if len(orders) > 0:
        max_date = orders['日期'].max()
        week_start = max_date - pd.Timedelta(days=6)
        week_data = orders[orders['日期'] >= week_start].copy()
        week_data['日期标签'] = week_data['日期'].dt.month.astype(str) + '/' + week_data['日期'].dt.day.astype(str)
        week_trend = week_data.groupby('日期标签').agg(
            count=('订单号', 'count'),
            amount=('订单金额（元）', 'sum')
        ).reset_index()
        week_trend['amount'] = week_trend['amount'].round(2)
        # 补全7天
        all_days = pd.date_range(week_start, max_date)
        all_days_labels = [f"{d.month}/{d.day}" for d in all_days]
        week_trend = week_trend.set_index('日期标签').reindex(all_days_labels, fill_value=0).reset_index()
        week_trend.columns = ['label', 'count', 'amount']
        week_trend_list = week_trend.to_dict('records')
    else:
        week_trend_list = []

    # ============ 订单状态分布 ============
    status_rank = orders.groupby('订单状态').agg(
        count=('订单号', 'count'),
        amount=('订单金额（元）', 'sum')
    ).reset_index()
    status_rank['amount'] = status_rank['amount'].round(2)
    status_rank = status_rank.sort_values('count', ascending=False)
    status_rank_list = status_rank.rename(columns={'订单状态': 'name'}).to_dict('records')

    # ============ 订单金额分层 ============
    def amount_band(amount):
        if amount < 1000: return '1千以下'
        elif amount < 5000: return '1千-5千'
        elif amount < 10000: return '5千-1万'
        elif amount < 50000: return '1万-5万'
        else: return '5万以上'

    orders['金额区间'] = orders['订单金额（元）'].apply(amount_band)
    amount_bands = orders.groupby('金额区间').agg(
        count=('订单号', 'count'),
        amount=('订单金额（元）', 'sum')
    ).reset_index()
    amount_bands['amount'] = amount_bands['amount'].round(2)
    band_order = ['1千以下', '1千-5千', '5千-1万', '1万-5万', '5万以上']
    amount_bands['sort_key'] = amount_bands['金额区间'].apply(lambda x: band_order.index(x))
    amount_bands = amount_bands.sort_values('sort_key')
    amount_bands_list = amount_bands.rename(columns={'金额区间': 'name'}).drop(columns=['sort_key']).to_dict('records')

    # ============ 采购企业排行 ============
    buyer_rank = orders.groupby('采购企业').agg(
        count=('订单号', 'count'),
        amount=('订单金额（元）', 'sum')
    ).reset_index()
    buyer_rank['amount'] = buyer_rank['amount'].round(2)
    buyer_rank = buyer_rank.sort_values('amount', ascending=False)
    buyer_rank_list = buyer_rank.rename(columns={'采购企业': 'name'}).to_dict('records')

    # ============ 供应商排行 ============
    supplier_rank = orders.groupby('供应商').agg(
        count=('订单号', 'count'),
        amount=('订单金额（元）', 'sum')
    ).reset_index()
    supplier_rank['amount'] = supplier_rank['amount'].round(2)
    supplier_rank = supplier_rank.sort_values('amount', ascending=False)
    supplier_rank_list = supplier_rank.rename(columns={'供应商': 'name'}).to_dict('records')

    # ============ 电商供应商 ============
    ecommerce = orders[orders['供应商类型'] == '电商供应商']
    ecommerce_suppliers = ecommerce.groupby('供应商').agg(
        count=('订单号', 'count'),
        amount=('订单金额（元）', 'sum')
    ).reset_index()
    ecommerce_suppliers['amount'] = ecommerce_suppliers['amount'].round(2)
    ecommerce_suppliers = ecommerce_suppliers.sort_values('amount', ascending=False)
    ecommerce_suppliers_list = ecommerce_suppliers.rename(columns={'供应商': 'name'}).to_dict('records')

    # ============ 本地供应商 ============
    local = orders[orders['供应商类型'] == '本地供应商']
    local_suppliers = local.groupby('供应商').agg(
        count=('订单号', 'count'),
        amount=('订单金额（元）', 'sum')
    ).reset_index()
    local_suppliers['amount'] = local_suppliers['amount'].round(2)
    local_suppliers = local_suppliers.sort_values('amount', ascending=False)
    local_suppliers_list = local_suppliers.rename(columns={'供应商': 'name'}).to_dict('records')

    # ============ 集中度指标 ============
    buyer_top10_share = round(buyer_rank.head(10)['amount'].sum() / total_amount, 6) if total_amount > 0 else 0
    supplier_top10_share = round(supplier_rank.head(10)['amount'].sum() / total_amount, 6) if total_amount > 0 else 0

    # ============ 组装结果 ============
    result = {
        "generatedAt": now.strftime('%Y-%m-%d %H:%M'),
        "dataEndDate": data_end_date,
        "source": "source_data.xlsx",
        "kpis": {
            "totalAmount": total_amount,
            "totalOrders": total_orders,
            "avgAmount": avg_amount,
            "zoneCount": zone_count,
            "weekAmount": week_amount,
            "weekOrders": week_count,
            "buyerCount": buyer_count,
            "supplierCount": supplier_count,
            "buyerTop10Share": buyer_top10_share,
            "supplierTop10Share": supplier_top10_share
        },
        "zoneRank": zone_rank_list,
        "supplierTypes": supplier_types_list,
        "monthTrend": month_trend_list,
        "weekTrend": week_trend_list,
        "statusRank": status_rank_list,
        "amountBands": amount_bands_list,
        "buyerRank": buyer_rank_list,
        "supplierRank": supplier_rank_list,
        "ecommerceSuppliers": ecommerce_suppliers_list,
        "localSuppliers": local_suppliers_list
    }

    # 保存到 JSON
    output_path = Path(OUTPUT_DIR) / 'dashboard.json'
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f'数据已生成: {output_path}')
    print(f'总订单数: {total_orders}')
    print(f'总金额: {total_amount:,.2f}')
    print(f'数据截至: {data_end_date}')

    return str(output_path)


if __name__ == "__main__":
    run()
