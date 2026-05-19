import os
import pandas as pd
from pathlib import Path
from collections import defaultdict


# 3.0版本优化执行效率

# ===== 配置 =====
TARGET_FILE = "newTarget.xlsx"
DATA_DIR = "data"
OUTPUT_DIR = "output"

Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

# ---------- 1. 读取身份证号，转为 set ----------
print("正在读取 newTarget.xlsx ...")
df_target = pd.read_excel(TARGET_FILE, header=None, usecols=[0], dtype=str)
id_series = df_target.iloc[:, 0].dropna().astype(str).str.strip()
id_set = set(id_series.unique())                # 核心：转为集合
print(f"共读入 {len(id_set)} 个唯一身份证号")

if not id_set:
    print("无有效身份证号，退出。")
    exit(0)

# ---------- 2. 收集 Excel 文件 ----------
data_path = Path(DATA_DIR)
if not data_path.exists():
    print(f"文件夹 {DATA_DIR} 不存在")
    exit(1)

excel_files = list(data_path.rglob("*.xlsx")) + list(data_path.rglob("*.xls"))
excel_files = [f for f in excel_files if not f.name.startswith("~$")]
print(f"在 {DATA_DIR} 中找到 {len(excel_files)} 个 Excel 文件")

# ---------- 3. 高效遍历匹配 ----------
# 数据结构：id_to_rows[身份证号] = [行列表1, 行列表2, ...]
# 每个行列表 = [来源文件, 来源工作表, 原始单元格1, 原始单元格2, ...]
id_to_rows = defaultdict(list)

for idx, excel_file in enumerate(excel_files, 1):
    print(f"处理文件 {idx}/{len(excel_files)}: {excel_file.name}")

    try:
        # 统一使用 calamine 引擎，完美支持 .xls 和 .xlsx
        sheets = pd.read_excel(
            excel_file,
            sheet_name=None,
            dtype=str,
            header=None,
            engine='calamine'
        )
    except Exception as e:
        print(f"  警告: 读取失败，跳过 ({e})")
        continue

    for sheet_name, df in sheets.items():
        if df.empty:
            continue

        # 去除所有单元格首尾空格（确保匹配准确）
        df = df.apply(lambda col: col.str.strip())

        # ---------- 关键优化：一次性找出包含任何目标身份证号的行 ----------
        match_mask = df.isin(id_set).any(axis=1)
        if not match_mask.any():
            continue

        # 只处理候选行（数量极少）
        candidate_df = df[match_mask]

        # 遍历候选行，确定该行匹配了哪些身份证号
        for row_idx, row in candidate_df.iterrows():
            # 收集该行中出现的所有目标身份证号（去重）
            matched_ids = {val for val in row.values if val in id_set}
            if not matched_ids:
                continue

            # 构造带来源信息的行（插入两列）
            row_data = [excel_file.name, sheet_name] + row.tolist()

            # 将该行加入每一个匹配身份证号的结果列表
            for mid in matched_ids:
                id_to_rows[mid].append(row_data)

# ---------- 4. 输出结果文件 ----------
print("开始生成结果文件...")
output_count = 0
for id_val, rows in id_to_rows.items():
    if not rows:
        continue

    result_df = pd.DataFrame(rows)
    save_path = os.path.join(OUTPUT_DIR, f"{id_val}.xlsx")
    try:
        result_df.to_excel(save_path, index=False, header=False)
        output_count += 1
        if output_count % 1000 == 0:
            print(f"已生成 {output_count} 个文件...")
    except Exception as e:
        print(f"  保存失败: {save_path} - {e}")

print(f"全部完成!共生成 {output_count} 个结果文件。")