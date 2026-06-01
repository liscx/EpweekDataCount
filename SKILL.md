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
    └── analysis_results_{时间戳}_{M/F/LM/NM}.xlsx  # 统计结果（Excel）
```

## 四种统计模式

### 周一统计（每周一执行）
1. 上周（周一至周日）订单总数、总金额，分专区
2. 本月订单总数、总金额，分专区
3. 全量：按供应商类型统计 + 分专区统计

### 周五统计（每周五执行）
1. 本周（上周五至本周四）订单总数、总金额
2. 按供应商类型统计（本地供应商、电商供应商）
3. 分专区统计（本周/本月/全量三个维度）

### 上月统计
1. 上个月全月订单总数、总金额
2. 按供应商类型统计（本地供应商、电商供应商）
3. 分专区统计

### 综合统计（其他日期自动执行）
1. 上周（上周一 ~ 上周日，完整7天）
2. 本周（本周一 ~ 当前时刻，精确到分钟）
3. 上月全月
4. 本月至今
5. 全量
6. 每个维度均按供应商类型 + 分专区统计

## 输出文件

文件名格式：`analysis_results_{YYYYMMDDHHmm}_{后缀}.xlsx`
- `_M`：周一统计
- `_F`：周五统计
- `_LM`：上月统计
- `_NM`：综合统计
- 时间戳精确到分钟，如 `analysis_results_202605260930_M.xlsx`

#### 周一 Excel 格式
| 区块 | 内容 |
|------|------|
| 上周（日期）专区统计 | 标题合并居中 → 表头 → 各专区 → 合计（SUM公式） |
| 本月（月份）订单统计 | 标题合并居中 → 表头 → 各专区 → 合计（SUM公式） |
| 全量统计 | 供应商类型表 + 合计 → 专区表 + 合计（SUM公式） |

#### 周五 Excel 格式
| 区块 | 内容 |
|------|------|
| 本周（日期） | 供应商类型表 → 专区表 + 合计 |
| 本月（月份） | 供应商类型表 → 专区表 + 合计 |
| 全量 | 供应商类型表 → 专区表 + 合计 |

#### 上月 Excel 格式
| 区块 | 内容 |
|------|------|
| 上月（月份）订单统计 | 供应商类型表 + 合计 → 专区表 + 合计（SUM公式） |

#### 综合 Excel 格式
| 区块 | 内容 |
|------|------|
| 上周（日期范围） | 供应商类型表 + 合计 → 专区表 + 合计（SUM公式） |
| 本周（日期 ~ 当前时间） | 供应商类型表 + 合计 → 专区表 + 合计（SUM公式） |
| 上月（月份） | 供应商类型表 + 合计 → 专区表 + 合计（SUM公式） |
| 本月（月份） | 供应商类型表 + 合计 → 专区表 + 合计（SUM公式） |
| 全量 | 供应商类型表 + 合计 → 专区表 + 合计（SUM公式） |

## 执行命令
```bash
cd D:\AutoWorkSkill\cronJob\weekDataCount

# 自动运行（周一→周一统计，周五→周五统计，其他→综合统计）
.venv\Scripts\python workflow.py

# 手动指定模式
.venv\Scripts\python workflow.py monday       # 强制周一统计
.venv\Scripts\python workflow.py friday       # 强制周五统计
.venv\Scripts\python workflow.py last_month   # 上月统计
.venv\Scripts\python workflow.py normal       # 综合统计
```

## 定时任务配置
- 周一 09:00 执行 `workflow.py`（自动走 monday 模式）
- 周五 09:00 执行 `workflow.py`（自动走 friday 模式）
- 其他日期 09:00 执行 `workflow.py`（自动走 normal 模式）
- 月初可手动执行 `workflow.py last_month` 统计上月数据

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
