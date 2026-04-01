---
name: daily-order-process
description: "当天订单梳理。将当天订单.xls经过空值填充、日期格式转换后，按购货单位是否包含BK自动分流追加到系统销售订单汇总文件的不同工作表中。Keywords: 订单梳理, 订单处理, 当天订单, 空值填充, 数据分流, 销售订单."
---

# 当天订单梳理

将"当天订单.xls"经过数据处理后追加到"系统销售订单汇总"文件中。

## 处理流程

1. 读取源文件 (.xls)，检查第3列列名是否为"购货单位"，不是则自动修正
2. 空值填充：所有列向前/向后填充，**M列"客户订单号"不填充**
3. 日期格式转换为 YYYY/M/D
4. 数据分流：购货单位包含"BK" → "其他销售订单"，其余 → "开票销售订单"
5. 追加到目标文件最后有数据的行之后，复制目标格式（边框、字号），字体统一用**宋体**
6. 开票销售订单的K列(总金额)使用公式 `=H{row}*J{row}`

## 执行方式

有两种方式：

### 方式一：直接执行 Python 脚本（交互式输入路径）

```powershell
python scripts/process_order.py
```

用户需输入源文件和目标文件路径，也可直接回车使用默认路径。

### 方式二：通过 AI Agent 非交互执行

当用户通过对话要求执行时，直接用 Python 内联代码运行，将源文件和目标文件路径硬编码为变量，跳过 `input()` 交互。核心逻辑：

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, Font
from datetime import datetime

SOURCE_FILE = r'<用户提供的源文件路径>'
TARGET_FILE = r'<用户提供的目标文件路径>'
SKIP_FILL_COLUMNS = ['客户订单号']
BK_KEYWORD = 'BK'

# 读取 → 修正列名 → 空值填充 → 日期转换 → 分离BK/非BK → 追加到两个工作表
```

完整代码见 `scripts/process_order.py`。

## 源文件结构

| 列 | 列名 |
|----|------|
| A | 订单日期 |
| B | 发货方式 |
| C | 购货单位 |
| D | 部门 |
| E | 业务员 |
| F | 产品名称 |
| G | 规格型号 |
| H | 数量 |
| I | 单位 |
| J | 含税单价 |
| K | 总金额 |
| L | 摘要 |
| M | 客户订单号（不填充空值） |

## 目标文件结构

两个工作表：
- **开票销售订单**：有边框，K列为公式 `=H*J`
- **其他销售订单**：无边框

## 注意事项

- 运行前目标文件必须关闭，否则保存会报 PermissionError
- 默认路径可在 `scripts/process_order.py` 中修改
