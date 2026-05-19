import os
import pandas as pd
from pathlib import Path
from collections import defaultdict

# ===== 配置 =====
TARGET_FILE = "newTarget.xlsx"
DATA_DIR = "data"
OUTPUT_DIR = "output"
LOG_FILE = os.path.join(OUTPUT_DIR, "result_log.txt")   # 日志文件路径

Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

# ---------- 1. 读取身份证号 ----------
print("正在读取 newTarget.xlsx ...")
df_target = pd.read_excel(TARGET_FILE, header=None, usecols=[0], dtype=str)
id_series = df_target.iloc[:, 0].dropna().astype(str).str.strip()
id_set = set(id_series.unique())
print(f"共读入 {len(id_set)} 个唯一身份证号")

if not id_set:
    print("无有效身份证号，退出。")
    exit(0)

# ---------- 2. 收集文件 ----------
data_path = Path(DATA_DIR)
if not data_path.exists():
    print(f"文件夹 {DATA_DIR} 不存在")
    exit(1)

excel_files = list(data_path.rglob("*.xlsx")) + list(data_path.rglob("*.xls"))
excel_files = [f for f in excel_files if not f.name.startswith("~$")]
print(f"在 {DATA_DIR} 中找到 {len(excel_files)} 个 Excel 文件")

# ---------- 3. 高效匹配 ----------
id_to_rows = defaultdict(list)

for idx, excel_file in enumerate(excel_files, 1):
    print(f"处理文件 {idx}/{len(excel_files)}: {excel_file.name}")

    try:
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

        # 去除空格
        df = df.apply(lambda col: col.str.strip())

        # 一次性找出包含任何目标ID的行
        match_mask = df.isin(id_set).any(axis=1)
        if not match_mask.any():
            continue

        candidate_df = df[match_mask]

        # 遍历候选行，确定具体匹配了哪些ID
        for row_idx, row in candidate_df.iterrows():
            matched_ids = {val for val in row.values if val in id_set}
            if not matched_ids:
                continue

            row_data = [excel_file.name, sheet_name] + row.tolist()
            for mid in matched_ids:
                id_to_rows[mid].append(row_data)

# ---------- 4. 生成结果并写日志 ----------
print("开始生成结果文件及日志...")

output_count = 0
# 以追加模式打开日志文件，防止中途卡死丢失记录
with open(LOG_FILE, 'w', encoding='utf-8') as log:
    log.write("结果文件\t行数\n")   # 表头

    for id_val, rows in id_to_rows.items():
        if not rows:
            continue

        result_df = pd.DataFrame(rows)
        row_count = len(result_df)
        save_path = os.path.join(OUTPUT_DIR, f"{id_val}.xlsx")

        try:
            result_df.to_excel(save_path, index=False, header=False)
            output_count += 1
            # 写入日志
            log.write(f"{id_val}.xlsx\t{row_count}\n")
            if output_count % 1000 == 0:
                print(f"已生成 {output_count} 个文件...")
        except Exception as e:
            print(f"  保存失败: {save_path} - {e}")

print(f"全部完成!共生成 {output_count} 个结果文件。")
print(f"日志文件保存在: {LOG_FILE}")