# -*- coding: utf-8 -*-
"""
当天订单梳理 - 自动处理Excel表格
功能：
1. 用户输入源文件和目标文件路径
2. 空值填充（指定列除外）
3. 日期格式转换
4. 按条件分离数据并追加到目标文件的不同工作表
5. 复制目标文件的格式
"""
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, Font
from datetime import datetime
import os

print("=" * 50)
print("    当天订单梳理 - 自动化处理工具")
print("=" * 50)
print()

# 用户输入源文件路径
print("请输入源文件路径 (当天订单.xls):")
print("示例: D:\\共享\\当天订单.xls")
SOURCE_FILE = input("路径: ").strip()
if not SOURCE_FILE:
    SOURCE_FILE = r"D:\共享\当天订单.xls"
    print(f"使用默认路径: {SOURCE_FILE}")

# 用户输入目标文件路径
print()
print("请输入目标文件路径 (系统销售订单汇总.xlsx):")
print("示例: D:\\共享\\02-系统销售订单汇总-2026.03.31（系统版）.xlsx")
TARGET_FILE = input("路径: ").strip()
if not TARGET_FILE:
    TARGET_FILE = r"D:\共享\02-系统销售订单汇总-2026.03.31（系统版）.xlsx"
    print(f"使用默认路径: {TARGET_FILE}")

# 检查文件是否存在
if not os.path.exists(SOURCE_FILE):
    print(f"\n错误: 源文件不存在: {SOURCE_FILE}")
    input("按回车键退出...")
    exit(1)

if not os.path.exists(TARGET_FILE):
    print(f"\n错误: 目标文件不存在: {TARGET_FILE}")
    input("按回车键退出...")
    exit(1)

print()
print("-" * 50)

# 配置参数
SKIP_FILL_COLUMNS = ['客户订单号']  # 不填充空值的列
BK_KEYWORD = 'BK'  # 购货单位包含此关键词的数据追加到"其他销售订单"

# 1. 读取源文件
print(f"\n[1] 读取源文件: {SOURCE_FILE}")
df = pd.read_excel(SOURCE_FILE)
print(f"    源文件总行数: {len(df)}")

# 2. 空值填充
print("\n[2] 执行空值填充...")

# 检查C1（第3列列名）是否为"购货单位"，如果不是则修改
if df.columns[2] != '购货单位':
    old_name = df.columns[2]
    df.columns = list(df.columns[:2]) + ['购货单位'] + list(df.columns[3:])
    print(f"    已将第3列列名从 '{old_name}' 改为 '购货单位'")

for col in df.columns:
    if col.strip() not in [c.strip() for c in SKIP_FILL_COLUMNS]:
        df[col] = df[col].ffill().bfill()
print("    空值填充完成 (M列'客户订单号'不填充)")

# 3. 日期格式改为 YYYY/M/D
print("\n[3] 转换日期格式为 YYYY/M/D...")
def format_date(val):
    if pd.isna(val):
        return val
    if isinstance(val, datetime):
        return val.strftime('%Y/%m/%d')
    return val

df['订单日期'] = df['订单日期'].apply(format_date)
print("    日期格式转换完成")

# 4. 分离BK和非BK数据
print("\n[4] 分离数据...")
df_bk = df[df['购货单位'].astype(str).str.contains(BK_KEYWORD, na=False)].copy()
df_other = df[~df['购货单位'].astype(str).str.contains(BK_KEYWORD, na=False)].copy()
print(f"    BK数据 (追加到'其他销售订单'): {len(df_bk)}行")
print(f"    其他数据 (追加到'开票销售订单'): {len(df_other)}行")

# 5. 加载目标文件
print(f"\n[5] 加载目标文件: {TARGET_FILE}")
wb = load_workbook(TARGET_FILE)

# 6. 处理"其他销售订单" - 追加BK数据
print("\n[6] 处理'其他销售订单'工作表...")
ws_other = wb['其他销售订单']
last_row_other = 0
for row in range(ws_other.max_row, 0, -1):
    if any(ws_other.cell(row, col).value is not None for col in range(1, 14)):
        last_row_other = row
        break
print(f"    最后数据行: {last_row_other}")

src_row_other = last_row_other if last_row_other > 1 else 1

for idx, row_data in df_bk.iterrows():
    new_row = last_row_other + 1 + df_bk.index.get_loc(idx)
    ws_other.cell(new_row, 1).value = row_data['订单日期']
    ws_other.cell(new_row, 2).value = row_data['发货方式']
    ws_other.cell(new_row, 3).value = row_data['购货单位']
    ws_other.cell(new_row, 4).value = row_data['部门']
    ws_other.cell(new_row, 5).value = row_data['业务员']
    ws_other.cell(new_row, 6).value = row_data['产品名称']
    ws_other.cell(new_row, 7).value = row_data['规格型号']
    ws_other.cell(new_row, 8).value = row_data['数量']
    ws_other.cell(new_row, 9).value = row_data['单位']
    ws_other.cell(new_row, 10).value = row_data['含税单价']
    ws_other.cell(new_row, 11).value = row_data['总金额']
    ws_other.cell(new_row, 12).value = row_data['摘要']
    ws_other.cell(new_row, 13).value = row_data['客户订单号']

    for col in range(1, 14):
        src_cell = ws_other.cell(src_row_other, col)
        dst_cell = ws_other.cell(new_row, col)
        dst_cell.font = Font(name='宋体', size=src_cell.font.size if src_cell.font else 11)
        if src_cell.alignment:
            dst_cell.alignment = Alignment(horizontal=src_cell.alignment.horizontal, vertical=src_cell.alignment.vertical)

    # 复制行高
    if src_row_other in ws_other.row_dimensions and ws_other.row_dimensions[src_row_other].height:
        ws_other.row_dimensions[new_row].height = ws_other.row_dimensions[src_row_other].height

print(f"    已追加 {len(df_bk)} 行")

# 7. 处理"开票销售订单" - 追加非BK数据
print("\n[7] 处理'开票销售订单'工作表...")
ws_kp = wb['开票销售订单']
last_row_kp = 0
for row in range(ws_kp.max_row, 0, -1):
    if any(ws_kp.cell(row, col).value is not None for col in range(1, 14)):
        last_row_kp = row
        break
print(f"    最后数据行: {last_row_kp}")

src_row_kp = last_row_kp if last_row_kp > 1 else 1
sample_cell = ws_kp.cell(src_row_kp, 1)
border_style = sample_cell.border.left.style if sample_cell.border and sample_cell.border.left else 'thin'

for idx, row_data in df_other.iterrows():
    new_row = last_row_kp + 1 + df_other.index.get_loc(idx)
    ws_kp.cell(new_row, 1).value = row_data['订单日期']
    ws_kp.cell(new_row, 2).value = row_data['发货方式']
    ws_kp.cell(new_row, 3).value = row_data['购货单位']
    ws_kp.cell(new_row, 4).value = row_data['部门']
    ws_kp.cell(new_row, 5).value = row_data['业务员']
    ws_kp.cell(new_row, 6).value = row_data['产品名称']
    ws_kp.cell(new_row, 7).value = row_data['规格型号']
    ws_kp.cell(new_row, 8).value = row_data['数量']
    ws_kp.cell(new_row, 9).value = row_data['单位']
    ws_kp.cell(new_row, 10).value = row_data['含税单价']
    ws_kp.cell(new_row, 11).value = f'=H{new_row}*J{new_row}'
    ws_kp.cell(new_row, 12).value = row_data['摘要']
    ws_kp.cell(new_row, 13).value = row_data['客户订单号']

    for col in range(1, 14):
        src_cell = ws_kp.cell(src_row_kp, col)
        dst_cell = ws_kp.cell(new_row, col)
        side = Side(style=border_style, color='000000')
        dst_cell.border = Border(left=side, right=side, top=side, bottom=side)
        dst_cell.font = Font(name='宋体', size=src_cell.font.size if src_cell.font else 11)
        if src_cell.alignment:
            dst_cell.alignment = Alignment(horizontal=src_cell.alignment.horizontal, vertical=src_cell.alignment.vertical)

    # 复制行高
    if src_row_kp in ws_kp.row_dimensions and ws_kp.row_dimensions[src_row_kp].height:
        ws_kp.row_dimensions[new_row].height = ws_kp.row_dimensions[src_row_kp].height

print(f"    已追加 {len(df_other)} 行")

# 8. 保存
print("\n[8] 保存文件...")
wb.save(TARGET_FILE)

print()
print("=" * 50)
print("    处理完成!")
print("=" * 50)
print(f"  - 其他销售订单: {len(df_bk)}行 (从第{last_row_other+1}行开始)")
print(f"  - 开票销售订单: {len(df_other)}行 (从第{last_row_kp+1}行开始)")
print()
input("按回车键退出...")
