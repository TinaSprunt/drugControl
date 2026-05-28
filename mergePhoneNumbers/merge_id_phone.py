#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将 GA.xlsx 和 MT.xlsx 中的身份证号合并去重后写入 id.xlsx，
并匹配对应的电话号码和姓名。

功能：
1. 读取 GA.xlsx / MT.xlsx 的 A 列（身份证号），合并去重后写入 id.xlsx 的 A 列。
2. 遍历去重后的身份证号，分别在 GA.xlsx / MT.xlsx 的 A 列中匹配，
   提取对应 B 列（电话号码）和 C 列（姓名），去重后用中文逗号（，）拼接，
   写入 id.xlsx 的 B 列和 C 列。
"""

import os
import re
import pandas as pd


def read_column(file_path, col_index):
    """读取 Excel 文件指定列的数据（无表头，从第一行开始读取），返回清洗后的列表"""
    try:
        df = pd.read_excel(file_path, header=None, dtype=str)
        if df.empty:
            return []
        data = df.iloc[:, col_index]  # 无表头，直接从第 0 行开始读取
        data = data.dropna().str.strip()
        return [x for x in data.tolist() if x]
    except FileNotFoundError:
        print(f"[警告] 文件不存在: {file_path}")
        return []
    except ValueError as e:
        print(f"[错误] 读取列失败，可能文件结构异常: {file_path} —— {e}")
        return []
    except Exception as e:
        print(f"[错误] 读取文件 {file_path} 失败: {e}")
        return []


def read_all_data(file_path):
    """读取 Excel 文件的 A/B/C 列所有数据（无表头，从第一行开始读取），返回清洗后的 DataFrame"""
    try:
        df = pd.read_excel(file_path, header=None, dtype=str)
        if df.empty:
            return pd.DataFrame(columns=['id', 'phone', 'name'])
        # 无表头，直接从第 0 行开始读取
        data = df.iloc[:, :].copy()
        # 补全到至少 3 列（兼容部分文件只有 A/B 两列的情况）
        while data.shape[1] < 3:
            data[data.shape[1]] = None
        # 取前 3 列
        data = data.iloc[:, :3]
        data.columns = ['id', 'phone', 'name']
        # 清洗：去除首尾空白
        for col in data.columns:
            data[col] = data[col].astype(str).str.strip()
        # 过滤身份证号为空的无效行
        data = data.dropna(subset=['id'])
        data = data[data['id'] != '']
        data = data[data['id'] != 'nan']
        data = data[data['id'] != 'nan']  # 二次确认（astype(str) 后 'nan' 可能出现变体）
        return data
    except FileNotFoundError:
        print(f"[警告] 文件不存在: {file_path}")
        return pd.DataFrame(columns=['id', 'phone', 'name'])
    except Exception as e:
        print(f"[错误] 读取文件 {file_path} 失败: {e}")
        return pd.DataFrame(columns=['id', 'phone', 'name'])


def normalize_phone(phone):
    """
    标准化单个电话号码，返回纯数字形式。
    - 去除 +86 / 0086 等中国区号前缀
    - 去除空格、短横线(-)、括号等所有非数字字符
    """
    phone = phone.strip()
    # 去除区号前缀：+86、86、0086（允许中间有空格或短横线）
    phone = re.sub(r'^(?:\+?86|0086)[\s\-]?', '', phone)
    # 去除所有非数字字符
    digits = re.sub(r'\D', '', phone)
    return digits


def deduplicate_phones(phone_list):
    """
    对电话号码列表去重，保持首次出现顺序。
    兼容 +86/0086 前缀、空格、短横线等格式差异，
    将实质相同的号码视为重复项并合并。
    返回去重后的纯数字号码列表。
    """
    seen = set()      # 存储已见过的标准化（纯数字）号码
    result = []       # 保持首次出现顺序的结果列表
    for phone in phone_list:
        normalized = normalize_phone(phone)
        if not normalized:
            continue
        if normalized not in seen:
            seen.add(normalized)
            result.append(normalized)   # 返回标准化的纯数字格式
    return result


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    ga_path = os.path.join(base_dir, 'GA.xlsx')
    mt_path = os.path.join(base_dir, 'MT.xlsx')
    output_path = os.path.join(base_dir, 'id.xlsx')

    # ==================== 步骤 1：读取并合并去重身份证号 ====================
    print("=" * 50)
    print("步骤 1：读取身份证号并去重")
    print("=" * 50)
    ids_ga = read_column(ga_path, 0)
    ids_mt = read_column(mt_path, 0)

    all_ids = sorted(set(ids_ga + ids_mt))
    print(f"  GA.xlsx → {len(ids_ga)} 条身份证号")
    print(f"  MT.xlsx → {len(ids_mt)} 条身份证号")
    print(f"  合并去重后共 {len(all_ids)} 条身份证号")

    if not all_ids:
        print("[错误] 未读取到任何有效的身份证号，程序退出。")
        return

    # ==================== 步骤 2：读取完整数据用于匹配 ====================
    print("\n" + "=" * 50)
    print("步骤 2：读取电话号码和姓名数据")
    print("=" * 50)
    ga_data = read_all_data(ga_path)
    mt_data = read_all_data(mt_path)
    combined = pd.concat([ga_data, mt_data], ignore_index=True)
    print(f"  GA.xlsx → {len(ga_data)} 条记录")
    print(f"  MT.xlsx → {len(mt_data)} 条记录")
    print(f"  合并后共 {len(combined)} 条记录")

    # ==================== 步骤 3：匹配电话号码和姓名 ====================
    print("\n" + "=" * 50)
    print("步骤 3：匹配电话号码和姓名")
    print("=" * 50)
    result_rows = []
    matched_count = 0
    for uid in all_ids:
        matches = combined[combined['id'] == uid]

        # 提取电话号码：先拆分（兼容中文逗号和英文逗号），再去重
        raw_phones = matches['phone'].dropna()
        raw_phones = raw_phones[(raw_phones != '') & (raw_phones != 'nan')]
        all_phones = []
        for phone_str in raw_phones:
            for part in re.split(r'[，,]', phone_str):
                part = part.strip()
                if part and part.lower() != 'nan':
                    all_phones.append(part)
        phones = deduplicate_phones(all_phones)
        phones_str = '，'.join(phones) if phones else ''

        # 提取唯一的姓名（排除空值、NaN、'nan' 字符串）
        names = matches['name'].dropna()
        names = names[(names != '') & (names != 'nan')].unique().tolist()
        names_str = '，'.join(names) if names else ''

        if phones_str or names_str:
            matched_count += 1

        result_rows.append([uid, phones_str, names_str])

    print(f"  成功匹配 {matched_count} / {len(all_ids)} 条身份证号")

    # ==================== 步骤 4：写入 id.xlsx ====================
    print("\n" + "=" * 50)
    print("步骤 4：写入 id.xlsx")
    print("=" * 50)
    result_df = pd.DataFrame(result_rows)

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, header=False)
        print(f"  输出文件: {output_path}")
        print(f"  共写入 {len(result_rows)} 行数据 (A列:身份证号, B列:电话号码, C列:姓名)")
    except PermissionError:
        # 文件被占用时，尝试写入备用文件名
        alt_path = output_path.replace('.xlsx', '_new.xlsx')
        try:
            with pd.ExcelWriter(alt_path, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, header=False)
            print(f"  [提示] id.xlsx 正被占用，已写入备用文件: {alt_path}")
            print(f"  [提示] 请关闭 id.xlsx 后将 {os.path.basename(alt_path)} 重命名为 id.xlsx")
        except Exception as e2:
            print(f"[错误] 写入备用文件也失败: {e2}")
            return
    except Exception as e:
        print(f"[错误] 写入文件 {output_path} 失败: {e}")
        return

    print("\n处理完成！")


if __name__ == '__main__':
    main()
