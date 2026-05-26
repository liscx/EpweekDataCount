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
└── result/
    └── analysis_results_{时间戳}_{M/F}.xlsx  # 统计结果（Excel）
```

## 两种统计模式

### 周一统计（每周一执行）
1. 上周（周一至周日）订单总数、总金额，分专区
2. 本月订单总数、总金额，分专区

### 周五统计（每周五执行）
1. 本周（周一至周日）订单总数、总金额
2. 按供应商类型统计（本地供应商、电商供应商）
3. 分专区统计（本周/本月/全量三个维度）

## 输出文件

文件名格式：`analysis_results_{YYYYMMDDHHmm}_{M/F}.xlsx`
- `_M`：周一统计
- `_F`：周五统计
- 时间戳精确到分钟，如 `analysis_results_202605260930_M.xlsx`

#### 周一 Excel 格式
| 区块 | 内容 |
|------|------|
| 上周（日期）专区统计 | 标题合并居中 → 表头 → 各专区 → 合计（SUM公式） |
| 本月（月份）订单统计 | 标题合并居中 → 表头 → 各专区 → 合计（SUM公式） |

#### 周五 Excel 格式
| 区块 | 内容 |
|------|------|
| 本周（日期） | 供应商类型表 → 专区表 + 合计 |
| 本月（月份） | 供应商类型表 → 专区表 + 合计 |
| 全量 | 供应商类型表 → 专区表 + 合计 |

## 执行命令
```bash
cd D:\AutoWorkSkill\cronJob\weekDataCount

# 自动运行（默认：周一跑周一统计，其他时间跑周五统计）
.venv\Scripts\python workflow.py

# 手动指定模式
.venv\Scripts\python workflow.py monday   # 强制周一统计
.venv\Scripts\python workflow.py friday   # 强制周五统计
```

## 定时任务配置
- 周一 09:00 执行 `workflow.py`（自动走 monday 模式）
- 周五 09:00 执行 `workflow.py`（自动走 friday 模式）

## 前置依赖
- Python 3.12（虚拟环境：.venv）
- Google Chrome 已安装
- 依赖包：selenium, webdriver-manager, pandas, openpyxl

## 注意事项
- 需要有桌面环境（Chrome 以非 headless 模式运行）
- 浏览器最大化打开，默认缩放 75%
- 登录账号已内置在 xyt_export.py 中
- ChromeDriver 自动管理，优先使用本地缓存
- 一个订单可能包含多个商品（多行），统计时按订单号去重
- 超时统一 180s，所有等待均为自适应检测
