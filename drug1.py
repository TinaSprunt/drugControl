import os
import pandas as pd
from pathlib import Path

# ===== 配置 =====
TARGET_FILE = "target.xlsx"      # 身份证号来源表
DATA_DIR = "data"                # 待搜索的Excel文件夹
OUTPUT_DIR = "output"            # 输出文件夹

# 创建输出目录
Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

# ---------- 1. 读取目标身份证号 ----------
try:
    df_target = pd.read_excel(TARGET_FILE, header=None, usecols=[0], dtype=str)
except Exception as e:
    print(f"读取 {TARGET_FILE} 失败: {e}")
    exit(1)

# 提取A列所有非空身份证号，去重，去除首尾空格
id_series = df_target.iloc[:, 0].dropna().astype(str).str.strip()
id_list = id_series.unique().tolist()
print(f"共读取到 {len(id_list)} 个唯一身份证号")

if not id_list:
    print("target.xlsx 的A列无有效身份证号，程序退出。")
    exit(0)

# ---------- 2. 收集 data 文件夹下所有 Excel 文件 ----------
data_path = Path(DATA_DIR)
if not data_path.exists():
    print(f"文件夹 {DATA_DIR} 不存在")
    exit(1)

excel_files = list(data_path.rglob("*.xlsx")) + list(data_path.rglob("*.xls"))
# 排除临时文件（~$开头）
excel_files = [f for f in excel_files if not f.name.startswith("~$")]
print(f"在 {DATA_DIR} 中找到 {len(excel_files)} 个 Excel 文件")

# ---------- 3. 为每个身份证号准备一个空 DataFrame 列表 ----------
from collections import defaultdict
id_to_rows = defaultdict(list)  # key: 身份证号, value: 匹配行DataFrame列表

# ---------- 4. 遍历每个 Excel 文件，搜索所有工作表 ----------
for excel_file in excel_files:
    print(f"正在处理: {excel_file.name}")
    try:
        # 读取所有工作表，sheet_name=None 返回字典 {sheet_name: DataFrame}
        sheets_dict = pd.read_excel(excel_file, sheet_name=None, dtype=str, header=None)
    except Exception as e:
        print(f"  警告: 无法读取文件 {excel_file}，跳过。错误: {e}")
        continue

    for sheet_name, df in sheets_dict.items():
        # 跳过空 DataFrame
        if df.empty:
            continue
        # 将所有单元格转为字符串，并去除前后空格（使匹配更鲁棒）
        df_str = df.astype(str).apply(lambda col: col.str.strip())
        # 对每个身份证号，查找匹配的行
        for id_str in id_list:
            # 判断每行是否包含完全等于该身份证号的单元格
            mask = (df_str == id_str).any(axis=1)
            if mask.any():
                matched_rows = df[mask].copy()
                # 记录来源信息（可选）
                matched_rows.insert(0, "来源文件", excel_file.name)
                matched_rows.insert(1, "来源工作表", sheet_name)
                id_to_rows[id_str].append(matched_rows)

# ---------- 5. 输出结果 ----------
output_count = 0
for id_str, df_list in id_to_rows.items():
    if df_list:
        # 合并所有来源的匹配行
        result_df = pd.concat(df_list, ignore_index=True)
        # 文件名使用身份证号，注意身份证号可能含'X'等，Windows允许
        safe_filename = f"{id_str}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, safe_filename)
        try:
            result_df.to_excel(output_path, index=False, header=False)
            output_count += 1
            print(f"已生成: {output_path}，共 {len(result_df)} 行数据")
        except Exception as e:
            print(f"保存文件 {output_path} 失败: {e}")
    else:
        print(f"身份证号 {id_str} 未在任何文件中找到匹配行，跳过生成文件。")

print(f"处理完成，共生成 {output_count} 个结果文件。")