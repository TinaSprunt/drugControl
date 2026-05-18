import os
import pandas as pd
from pathlib import Path
from collections import defaultdict




# ===== 配置 =====
TARGET_FILE = "newTarget.xlsx"
DATA_DIR = "data"
OUTPUT_DIR = "output"

Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

# ---------- 1. 读取身份证号 ----------
try:
    df_target = pd.read_excel(TARGET_FILE, header=None, usecols=[0], dtype=str)
except Exception as e:
    print(f"读取 {TARGET_FILE} 失败: {e}")
    exit(1)

id_series = df_target.iloc[:, 0].dropna().astype(str).str.strip()
id_list = id_series.unique().tolist()
print(f"共读取到 {len(id_list)} 个唯一身份证号")
if not id_list:
    print("newTarget.xlsx 的A列无有效身份证号，程序退出。")
    exit(0)

# ---------- 2. 收集所有 Excel 文件 ----------
data_path = Path(DATA_DIR)
if not data_path.exists():
    print(f"文件夹 {DATA_DIR} 不存在")
    exit(1)

excel_files = list(data_path.rglob("*.xlsx")) + list(data_path.rglob("*.xls"))
excel_files = [f for f in excel_files if not f.name.startswith("~$")]
print(f"在 {DATA_DIR} 中找到 {len(excel_files)} 个 Excel 文件")

# ---------- 3. 准备存储 ----------
id_to_rows = defaultdict(list)

# ---------- 4. 遍历文件，统一使用 calamine 引擎 ----------
for excel_file in excel_files:
    print(f"正在处理: {excel_file.name}")
    try:
        # calamine 引擎支持 .xls 和 .xlsx，无需区分扩展名
        sheets_dict = pd.read_excel(
            excel_file,
            sheet_name=None,
            dtype=str,
            header=None,
            engine='calamine'   # 关键参数
        )
    except Exception as e:
        print(f"  警告: 无法读取文件 {excel_file}，跳过。错误: {e}")
        continue

    for sheet_name, df in sheets_dict.items():
        if df.empty:
            continue
        df_str = df.astype(str).apply(lambda col: col.str.strip())
        for id_str in id_list:
            mask = (df_str == id_str).any(axis=1)
            if mask.any():
                matched_rows = df[mask].copy()
                matched_rows.insert(0, "来源文件", excel_file.name)
                matched_rows.insert(1, "来源工作表", sheet_name)
                id_to_rows[id_str].append(matched_rows)

# ---------- 5. 输出结果 ----------
output_count = 0
for id_str, df_list in id_to_rows.items():
    if df_list:
        result_df = pd.concat(df_list, ignore_index=True)
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