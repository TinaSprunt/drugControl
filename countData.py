#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
递归读取 data 文件夹下所有表格文件（包含.xlsx .xls .et），
统计每个 sheet 的行数（包含表头），并记录文件数、sheet 数和总行数，输出到 log.txt。
依赖：openpyxl, xlrd
安装：pip install openpyxl xlrd
"""

import os
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("缺少 openpyxl 库，请执行：pip install openpyxl")
try:
    import xlrd
except ImportError:
    sys.exit("缺少 xlrd 库，请执行：pip install xlrd")

# 支持的文件后缀（大小写不敏感）
SUFFIXES = {'.xlsx', '.xls', '.et'}

def count_rows_xlsx(filepath: str) -> dict:
    """使用 openpyxl 只读模式统计 .xlsx 文件每个 sheet 的行数。"""
    sheet_rows = {}
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            row_count = sum(1 for _ in ws.rows)   # 包含表头
            sheet_rows[name] = row_count
    finally:
        wb.close()
    return sheet_rows

def count_rows_xls_et(filepath: str) -> dict:
    """使用 xlrd 统计 .xls 和 .et 文件每个 sheet 的行数。"""
    sheet_rows = {}
    wb = xlrd.open_workbook(filepath, on_demand=True)
    try:
        for name in wb.sheet_names():
            sheet = wb.sheet_by_name(name)
            sheet_rows[name] = sheet.nrows   # nrows 包含表头
    finally:
        wb.release_resources()
    return sheet_rows

def process_file(filepath: str) -> dict:
    """根据扩展名调用对应的统计函数，返回 {sheet_name: row_count} 或抛出异常。"""
    ext = Path(filepath).suffix.lower()
    if ext == '.xlsx':
        return count_rows_xlsx(filepath)
    elif ext in ('.xls', '.et'):
        return count_rows_xls_et(filepath)
    else:
        raise ValueError(f"不支持的文件类型: {ext}")

def main():
    data_dir = Path('data')
    if not data_dir.exists():
        sys.exit(f"错误：文件夹 '{data_dir}' 不存在")

    # 递归查找所有匹配的表格文件
    files = []
    for root, dirs, filenames in os.walk(data_dir):
        for fname in filenames:
            if Path(fname).suffix.lower() in SUFFIXES:
                files.append(Path(root) / fname)

    total_files = len(files)   # 找到的表格文件总数
    if not files:
        sys.exit("未找到任何表格文件（.xlsx .xls .et），程序退出。")

    total_sheets = 0           # 成功读取的 sheet 总数
    total_rows = 0
    log_lines = []
    log_lines.append("===== 表格行数统计 =====")

    for fp in files:
        file_str = str(fp)
        print(f"正在处理: {file_str}")
        try:
            sheet_info = process_file(file_str)
        except Exception as e:
            log_lines.append(f"[错误] 文件: {file_str}  原因: {e}")
            continue

        file_total = sum(sheet_info.values())
        total_rows += file_total
        total_sheets += len(sheet_info)   # 累加读取到的 sheet 数量

        log_lines.append(f"\n文件: {file_str}")
        for sheet_name, rows in sheet_info.items():
            log_lines.append(f"  {sheet_name}: {rows} 行")
        log_lines.append(f"  文件小计: {file_total} 行")

    # 添加汇总信息（包含新增统计）
    log_lines.append(f"\n===== 汇总 =====")
    log_lines.append(f"表格文件总数: {total_files} 个")
    log_lines.append(f"读取 sheet 总数: {total_sheets} 个")
    log_lines.append(f"全部表格总行数: {total_rows} 行")

    # 写入日志文件
    log_content = '\n'.join(log_lines)
    with open('log.txt', 'w', encoding='utf-8') as f:
        f.write(log_content)

    print(f"\n统计完成!共找到 {total_files} 个表格文件，"
          f"读取 {total_sheets} 个 sheet，总计 {total_rows} 行。详情见 log.txt")

if __name__ == '__main__':
    main()