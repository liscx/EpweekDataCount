# 订单数据自动统计

## 功能
自动登录阳光优采平台（xyt.etrading.cn），导出订单数据并统计分析。

## 项目结构
```
weekDataCount/
├── workflow.py          # 主入口，串联导出和数据处理
├── xyt_export.py        # 登录平台、导航、导出订单数据
├── process_data.py      # 数据统计分析
├── Data/
│   └── source_data.xlsx # 导出的原始订单数据
└── analysis_results.json # 统计结果输出
```

## 两种统计模式

### 周一统计（每周一执行）
1. 上周（周一至周日）订单总数、总金额，分专区
2. 本月订单总数、总金额，分专区
3. 总计订单总数、总金额，分专区

### 周五统计（每周五执行）
1. 本周（周一至周日）订单总数、总金额
2. 按供应商类型统计（本地供应商、电商供应商）
3. 分专区统计

## 执行命令
```bash
# 自动运行（默认：周一跑周一统计，其他时间跑周五统计）
cd D:\AutoWorkSkill\cronJob\weekDataCount && .venv\Scripts\python workflow.py

# 手动指定模式（跳过导出，直接处理已有数据）
.venv\Scripts\python process_data.py monday   # 强制周一统计
.venv\Scripts\python process_data.py friday   # 强制周五统计
```

## 定时任务配置
- 周一 09:00 执行 `workflow.py`（自动走 monday 模式）
- 周五 09:00 执行 `workflow.py`（自动走 friday 模式）

## 输出结果示例（analysis_results.json）

### 周一统计
```json
{
  "type": "monday",
  "total": {
    "order_count": 246,
    "total_amount": 361233.54,
    "zones": {
      "中国煤地电子商城": { "order_count": 206, "total_amount": 320156.69 }
    }
  },
  "last_week": {
    "range": "2026-05-11 ~ 2026-05-17",
    "order_count": 13,
    "total_amount": 31274.57,
    "zones": {}
  },
  "current_month": {
    "month": "2026-05",
    "order_count": 61,
    "total_amount": 98610.08,
    "zones": {}
  }
}
```

### 周五统计
```json
{
  "type": "friday",
  "current_week": {
    "range": "2026-05-18 ~ 2026-05-24",
    "order_count": 18,
    "total_amount": 48797.05,
    "supplier": {
      "本地供应商": { "order_count": 13, "total_amount": 41239.54 },
      "电商供应商": { "order_count": 5, "total_amount": 7557.51 }
    },
    "zones": {
      "中国煤地电子商城": { "order_count": 14, "total_amount": 45455.05 }
    }
  }
}
```

## 前置依赖
- Python 3.12（虚拟环境：.venv）
- Google Chrome 已安装
- 依赖包：selenium, webdriver-manager, pandas, openpyxl

## 注意事项
- 需要有桌面环境（Chrome 以非 headless 模式运行）
- 登录账号已内置在 xyt_export.py 中
- ChromeDriver 自动管理，优先使用本地缓存
- 一个订单可能包含多个商品（多行），统计时按订单号去重
- 超时统一 180s，所有等待均为自适应检测
