# 发货建议工具（Shipment Planner）

基于订单和销售 `.xlsx` 数据生成发货建议结果，输出报表，不回写 ERP。

## 文档导航

- `README.md`：快速上手（本文件）
- `AGENTS.md`：仓库协作与开发规范

## 快速开始（CLI）

### 1) 准备输入目录

将输入文件放在 `data/input/`：

- 订单文件（`.xlsx`）
- 销售文件（`.xlsx`）
- 可选约束文件：`shipment_constraints.json`

### 2) 自动识别输入并运行

```bash
python3 src/main.py --input-dir data/input --out-dir data/output
```

说明：

- 当未显式传入 `--orders` / `--sales` 时，程序会在 `--input-dir` 中扫描 `.xlsx`，按修改时间从新到旧查找，并基于“必填列是否齐全”自动识别订单/销售文件。

### 3) 显式指定文件运行

```bash
python3 src/main.py \
  --orders data/input/<orders>.xlsx \
  --sales data/input/<sales>.xlsx \
  --out-dir data/output
```

## 常用参数

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `--input-dir` | `data/input` | 自动识别输入文件目录 |
| `--out-dir` | `data/output` | 输出目录 |
| `--orders` | 空 | 指定订单文件（跳过自动识别） |
| `--sales` | 空 | 指定销售文件（跳过自动识别） |
| `--constraints` | 空 | 指定约束 JSON；不传时会尝试读取 `--input-dir/shipment_constraints.json` |
| `--min-order-ship-qty` | `10` | 订单级最小起发阈值 |
| `--global-gap-multiplier` | `1.0` | 全局缺口倍率 |
| `--sold30-weight` | `0.2` | 近30日销量权重（会与 7 日权重归一化） |
| `--sold7-weight` | `0.8` | 近7日销量权重（会与 30 日权重归一化） |

## 输出文件

CLI 默认输出到 `data/output/`（或 `--out-dir` 指定目录）：

- `发货建议明细.csv`
- `数据质量报告.csv`
- `运行摘要.json`

## 开发环境

建议用可编辑安装方式进行后续开发：

```bash
python3 -m venv .venv
python3 -m pip install -e ".[dev]"
```

如需运行桌面 UI 或打包：

```bash
python3 -m pip install -e ".[ui]"
python3 -m pip install -e ".[build]"
```

常用校验命令：

```bash
python3 -m pytest
python3 -m compileall src
```

## 桌面 UI（PyQt6）

安装依赖（如未安装）：

```bash
python3 -m pip install PyQt6
```

启动：

```bash
python3 src/ui_main.py
```

或：

```bash
PYTHONPATH=src python3 -m planner_ui
```

UI 说明：

- 可在界面设置 `近7日销量占比`、`近30日销量占比`、`全局增长系数`
- 每次运行会在你选择的输出目录下生成一个时间戳子目录（如 `output_20260211_120102`）
- 点击 `配置目录` 可快速打开约束文件目录

## 约束文件（shipment_constraints.json）

支持三类规则：

- `sku_order_max_qty`：按 `(内部订单号, SKU)` 限制建议数量
- `exclude_skc`：命中 `店铺款式编码` 时强制建议量为 0
- `exclude_skuid`：命中 `店铺商品编码` 时强制建议量为 0

示例：

```json
{
  "sku_order_max_qty": {
    "ZX45": 100,
    "ZX44": 80
  },
  "exclude_skc": [
    "86164486235"
  ],
  "exclude_skuid": [
    "43401042798"
  ]
}
```

约束加载行为：

- CLI：未传 `--constraints` 时，自动尝试读取 `data/input/shipment_constraints.json`
- UI（源码运行）：使用仓库内 `data/input/shipment_constraints.json`
- EXE：使用可执行文件所在目录的 `data/input/shipment_constraints.json`；首次启动若不存在会自动创建模板

## 业务规则与字段要求

关键输入规则：

- 输入必须是 `.xlsx`
- 订单文件必须包含如 `内部订单号`、`下单时间`、`店铺款式编码`、`店铺商品编码`、`数量`、`状态`、`标签` 等必填列
- 销售文件必须包含如 `平台商品基本信息-skc`、`平台商品基本信息-平台SKUID`、`平台商品基本信息-SKU货号`、`销售数据-近30日销量`、`销售数据-近7日销量`、库存相关列等必填列
- `下单时间` 格式要求：`YYYY-MM-DD HH:MM:SS`
- 仅处理 `标签` 包含 `今日可发货` 的订单行
- `状态 = 发货中` 且 `地址` 非空的行只计入 `发货中数量`，不参与建议分配

核心计算规则：

- 按 `(店铺款式编码, 店铺商品编码)` 汇总订单需求并匹配销售数据
- 建议需求按 `近30日销量` 与 `近7日销量` 加权日均销量乘以 `备货逻辑天数` 得出
- 可用库存使用 `平台仓内库存 + 平台待收货库存 + 发货中数量`
- 热销款缺口上浮 `1.2` 倍，再应用 `--global-gap-multiplier`
- 单行建议量在同一 `(SKC, SKUID)` 内按状态、下单时间、源行号顺序分配
- 30% 内的小变动会保留原订单行数量
- 订单级 `--min-order-ship-qty` 阈值会拦截低于起发量的订单，符合零 7 日销量豁免规则的订单除外

## Windows 打包 EXE

仓库根目录执行：

```bat
build.bat
```

清理后重建：

```bat
build.bat clean
```

产物输出到 `dist/`。
