# 订单数据自动统计

## 功能
自动登录阳光优采平台（xyt.etrading.cn），导出订单数据并统计分析。

## 项目结构
```
weekDataCount/
├── workflow.py              # 主入口，纯编排：按日期调用对应统计脚本+通知
├── xyt_export.py            # 登录平台、导航、导出订单数据
├── stats_utils.py           # 共享工具：数据统计函数 + Excel写入辅助
├── monday_stats.py          # 周一统计脚本
├── friday_stats.py          # 周五统计脚本
├── normal_stats.py          # 日常综合统计脚本
├── notify/
│   ├── email_notify.py      # 邮件通知模块
│   └── feishu_sheet.py      # 飞书在线表格更新 + 消息通知
├── config.yaml              # 邮件配置
├── Data/
│   └── source_data.xlsx     # 导出的原始订单数据
└── result/
    └── analysis_results_{时间戳}_{M/F/LM/NM}.xlsx  # 统计结果
```

## 四种统计模式

### 周一统计（每周一执行）— monday_stats.py
1. 上周（周一至周日）订单总数、总金额，分专区
2. 本月订单总数、总金额，分专区
3. 全量：按供应商类型统计 + 分专区统计

### 周五统计（每周五执行）— friday_stats.py
1. 本周（周一至今）订单总数、总金额
2. 按供应商类型统计（本地供应商、电商供应商）
3. 分专区统计（本周/本月/全量三个维度）

### 上月统计
1. 上个月全月订单总数、总金额
2. 按供应商类型统计（本地供应商、电商供应商）
3. 分专区统计

### 综合统计（其他日期自动执行）— normal_stats.py
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
| 上周（日期）专区统计 | 标题合并居中 → 表头 → 各专区 → 合计 |
| 本月（月份）订单统计 | 标题合并居中 → 表头 → 各专区 → 合计 |
| 全量统计 | 供应商类型表 + 合计 → 专区表 + 合计 |

#### 周五 Excel 格式
| 区块 | 内容 |
|------|------|
| 本周（日期） | 供应商类型表 → 专区表 + 合计 |
| 本月（月份） | 供应商类型表 → 专区表 + 合计 |
| 全量 | 供应商类型表 → 专区表 + 合计 |

#### 上月 Excel 格式
| 区块 | 内容 |
|------|------|
| 上月（月份）订单统计 | 供应商类型表 + 合计 → 专区表 + 合计 |

#### 综合 Excel 格式
| 区块 | 内容 |
|------|------|
| 上周（日期范围） | 供应商类型表 + 合计 → 专区表 + 合计 |
| 本周（日期 ~ 当前时间） | 供应商类型表 + 合计 → 专区表 + 合计 |
| 上月（月份） | 供应商类型表 + 合计 → 专区表 + 合计 |
| 本月（月份） | 供应商类型表 + 合计 → 专区表 + 合计 |
| 全量 | 供应商类型表 + 合计 → 专区表 + 合计 |

## workflow.py 编排逻辑

```
模式解析（auto → 周一/周五/日常）
  ↓
Step 1: xyt_export 登录导出数据
  ↓
Step 2: 按模式调用对应统计脚本
  周一 → monday_stats.run()
  周五 → friday_stats.run()
  日常 → normal_stats.run()
  ↓
Step 3: email_notify 发送邮件（仅日常模式）
  ↓
Step 4: feishu_sheet 更新飞书在线表格（所有模式）
```

## 执行命令
```bash
cd D:\AutoWorkSkill\cronJob\weekDataCount

# 自动运行（周一→周一统计，周五→周五统计，其他→综合统计）
python workflow.py

# 手动指定模式
python workflow.py monday       # 强制周一统计
python workflow.py friday       # 强制周五统计
python workflow.py last_month   # 上月统计
python workflow.py normal       # 综合统计

# 单独运行某个统计脚本
python monday_stats.py
python friday_stats.py
python normal_stats.py
```


## 前置依赖
- Python 3.12（系统环境）
- Chromium 浏览器（Playwright 自动管理）
- 依赖包：playwright, pandas, openpyxl, pyyaml
- 首次使用需安装浏览器：`playwright install chromium`

## 飞书在线表格
- 第一次运行会自动创建在线表格
- 表格 token 保存在 `spreadsheet_token.json`，后续运行复用
- 4 个 sheet：周一统计、周五统计、上月统计、综合统计
- 每次运行会覆盖写入最新数据
- 需要环境变量 `FEISHU_NOTIFY_CHAT_ID`（open_id 或 chat_id）

## 注意事项
- 需要有桌面环境（Chromium 以非 headless 模式运行）
- 浏览器窗口 1920x1080，默认缩放 75%
- 登录账号已内置在 xyt_export.py 中
- Chromium 浏览器由 Playwright 自动管理
- 首次使用需运行 `playwright install chromium` 安装浏览器
- 一个订单可能包含多个商品（多行），统计时按订单号去重
- 超时统一 180s，所有等待均为自适应检测
- 统计时间基于**订单创建时间**字段（而非订单日期）
