import pandas as pd
import json
from datetime import datetime
from pathlib import Path

# 读取 Excel
df = pd.read_excel('src/assets/阳光优采交易订单.xlsx')

# 按订单号去重，取订单首行金额
orders = df.drop_duplicates(subset='订单号', keep='first').copy()

# 基础时间
now = datetime.now()
data_end_date = orders['订单创建时间'].max()[:10] if len(orders) > 0 else now.strftime('%Y-%m-%d')

# ============ KPI 指标 ============
total_amount = round(orders['订单金额（元）'].sum(), 2)
total_orders = len(orders)
avg_amount = round(total_amount / total_orders, 2) if total_orders > 0 else 0

# 本周数据（自然周，周一到周日）
if len(orders) > 0:
    orders['日期'] = pd.to_datetime(orders['订单创建时间'])
    max_date = orders['日期'].max()
    # 获取 max_date 所在周的周一 0点
    week_start = max_date - pd.Timedelta(days=max_date.weekday())
    week_start = week_start.replace(hour=0, minute=0, second=0, microsecond=0)
    week_orders = orders[orders['日期'] >= week_start]
    week_amount = round(week_orders['订单金额（元）'].sum(), 2)
    week_count = len(week_orders)

    # 今日数据
    today_start = max_date.replace(hour=0, minute=0, second=0, microsecond=0)
    today_orders = orders[orders['日期'] >= today_start]
    today_amount = round(today_orders['订单金额（元）'].sum(), 2)
    today_count = len(today_orders)

    # 月环比同期（本月1号到今天 vs 上月1号到上月同日）
    cur_month_start = max_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    cur_month_orders = orders[(orders['日期'] >= cur_month_start) & (orders['日期'] <= max_date)]
    cur_month_amount = round(cur_month_orders['订单金额（元）'].sum(), 2)
    cur_month_count = len(cur_month_orders)

    # 上月同期截止日（处理月份跨年和月末天数问题）
    prev_month_end_day = max_date.day
    if max_date.month == 1:
        prev_month = 12
        prev_year = max_date.year - 1
    else:
        prev_month = max_date.month - 1
        prev_year = max_date.year
    import calendar
    max_day_in_prev = calendar.monthrange(prev_year, prev_month)[1]
    prev_month_end_day = min(prev_month_end_day, max_day_in_prev)
    prev_month_start = max_date.replace(year=prev_year, month=prev_month, day=1, hour=0, minute=0, second=0, microsecond=0)
    prev_month_end = max_date.replace(year=prev_year, month=prev_month, day=prev_month_end_day, hour=23, minute=59, second=59, microsecond=999999)
    prev_month_orders = orders[(orders['日期'] >= prev_month_start) & (orders['日期'] <= prev_month_end)]
    prev_month_amount = round(prev_month_orders['订单金额（元）'].sum(), 2)

    if prev_month_amount > 0:
        mom_rate = round((cur_month_amount - prev_month_amount) / prev_month_amount * 100, 1)
    else:
        mom_rate = None
else:
    week_amount = 0
    week_count = 0
    today_amount = 0
    today_count = 0
    cur_month_amount = 0
    cur_month_count = 0
    prev_month_amount = 0
    mom_rate = None

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

# ============ 月度趋势（最近6个月，按专区拆分） ============
if len(orders) > 0:
    max_date = orders['日期'].max()
    six_months_ago = max_date - pd.DateOffset(months=5)
    six_months_ago = six_months_ago.replace(day=1)
    month_data = orders[orders['日期'] >= six_months_ago].copy()
    month_data['月份'] = month_data['日期'].apply(lambda x: f"{x.year}年{x.month}月" if pd.notna(x) else "")

    # 保留所有专区，不再归为"其他"
    month_data['专区分组'] = month_data['专区名称']

    # 按月+专区汇总
    month_zone = month_data.groupby(['月份', '专区分组']).agg(
        amount=('订单金额（元）', 'sum'),
        count=('订单号', 'count')
    ).reset_index()
    month_zone['amount'] = month_zone['amount'].round(2)

    # 获取所有月份（排序）
    month_dates = month_data.groupby('月份')['日期'].min().sort_values()
    months_sorted = month_dates.index.tolist()

    # 获取所有专区（按金额排序）
    zones_sorted = month_zone.groupby('专区分组')['amount'].sum().sort_values(ascending=False).index.tolist()

    # 构建输出格式
    month_trend_list = []
    for month in months_sorted:
        month_zone_data = month_zone[month_zone['月份'] == month]
        zone_dict = dict(zip(month_zone_data['专区分组'], month_zone_data['amount']))
        item = {'label': month}
        item['count'] = int(month_zone_data['count'].sum())
        for zone in zones_sorted:
            item[zone] = zone_dict.get(zone, 0)
        month_trend_list.append(item)

    # 保存专区列表到全局数据
    month_trend_zones = zones_sorted
else:
    month_trend_list = []
    month_trend_zones = []

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
    elif amount < 30000: return '1万-3万'
    elif amount < 50000: return '3万-5万'
    else: return '5万以上'

orders['金额区间'] = orders['订单金额（元）'].apply(amount_band)
amount_bands = orders.groupby('金额区间').agg(
    count=('订单号', 'count'),
    amount=('订单金额（元）', 'sum')
).reset_index()
amount_bands['amount'] = amount_bands['amount'].round(2)
band_order = ['1千以下', '1千-5千', '5千-1万', '1万-3万', '3万-5万', '5万以上']
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
    "source": "阳光优采交易订单.xlsx",
    "kpis": {
        "totalAmount": total_amount,
        "totalOrders": total_orders,
        "avgAmount": avg_amount,
        "zoneCount": zone_count,
        "weekAmount": week_amount,
        "weekOrders": week_count,
        "todayAmount": today_amount,
        "todayOrders": today_count,
        "buyerCount": buyer_count,
        "supplierCount": supplier_count,
        "curMonthAmount": cur_month_amount,
        "curMonthOrders": cur_month_count,
        "prevMonthAmount": prev_month_amount,
        "momRate": mom_rate,
        "buyerTop10Share": buyer_top10_share,
        "supplierTop10Share": supplier_top10_share
    },
    "zoneRank": zone_rank_list,
    "supplierTypes": supplier_types_list,
    "monthTrend": month_trend_list,
    "monthTrendZones": month_trend_zones,
    "weekTrend": week_trend_list,
    "statusRank": status_rank_list,
    "amountBands": amount_bands_list,
    "buyerRank": buyer_rank_list,
    "supplierRank": supplier_rank_list,
    "ecommerceSuppliers": ecommerce_suppliers_list,
    "localSuppliers": local_suppliers_list
}

# 保存到 JSON
output_path = Path('src/data/dashboard.json')
output_path.parent.mkdir(parents=True, exist_ok=True)
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f'数据已生成: {output_path}')
print(f'总订单数: {total_orders}')
print(f'总金额: {total_amount:,.2f}')
print(f'数据截至: {data_end_date}')
