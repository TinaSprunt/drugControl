import os
import sys
import pandas as pd
from pathlib import Path
from collections import defaultdict
import traceback

"""
5.0版本新增对wps表格 .et 文件的支持

根据newTarget.xlsx文件中的身份证号, 递归遍历data文件夹下的所有表格文件( 包括excel文件.xlsx和.xls, 包括wps表格文件 .et)
获取每个人所有的购买记录,按照身份证号生成xlsx文件,将提取到的购买记录写入

"""


# ===== 配置 =====
TARGET_FILE = "newTarget.xlsx"
DATA_DIR = "data"
OUTPUT_DIR = "output"
LOG_FILE = os.path.join(OUTPUT_DIR, "result_log.txt")

Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

# ---------- 检测 pywin32 是否可用 ----------
HAS_WPS_COM = False
try:
    import win32com.client
    # 尝试创建 WPS 对象检测是否真的可用（不启动界面）
    try:
        wps_test = win32com.client.Dispatch("ET.Application")
        wps_test.Quit()
        HAS_WPS_COM = True
        print("✅ WPS COM 组件可用，将优先使用 WPS 读取 .et 文件。")
    except:
        print("⚠️ 未找到 WPS COM 组件（pywin32 已安装但 WPS 未注册），将使用 xlrd 读取 .et 文件。")
except ImportError:
    print("⚠️ 未安装 pywin32，将直接使用 xlrd 读取 .et 文件。")

# ---------- 1. 读取身份证号 ----------
print("正在读取 newTarget.xlsx ...")
try:
    df_target = pd.read_excel(TARGET_FILE, header=None, usecols=[0], dtype=str)
except Exception as e:
    print(f"❌ 读取 {TARGET_FILE} 失败: {e}")
    sys.exit(1)

id_series = df_target.iloc[:, 0].dropna().astype(str).str.strip()
id_set = set(id_series.unique())
print(f"✅ 共读入 {len(id_set)} 个唯一身份证号")
if not id_set:
    print("⚠️ 无有效身份证号，程序退出。")
    sys.exit(0)

# ---------- 2. 收集文件 ----------
data_path = Path(DATA_DIR)
if not data_path.exists():
    print(f"❌ 文件夹 {DATA_DIR} 不存在")
    sys.exit(1)

excel_files = []
for ext in ['*.xlsx', '*.xls', '*.et']:
    excel_files.extend(list(data_path.rglob(ext)))
excel_files = [f for f in excel_files if not f.name.startswith("~$")]
print(f"✅ 在 {DATA_DIR} 中找到 {len(excel_files)} 个表格文件")

# ---------- 3. 各格式读取函数 ----------
def read_xls_manual(file_path):
    """用 xlrd 手动读取 .xls / .et 文件，返回 {sheet: DataFrame}"""
    import xlrd
    wb = xlrd.open_workbook(file_path, formatting_info=False)
    sheets = {}
    for name in wb.sheet_names():
        ws = wb.sheet_by_name(name)
        data = []
        for r in range(ws.nrows):
            row = []
            for c in range(ws.ncols):
                cell = ws.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_EMPTY:
                    row.append('')
                elif cell.ctype == xlrd.XL_CELL_NUMBER:
                    # 避免身份证号变成浮点科学计数法
                    if cell.value == int(cell.value):
                        row.append(str(int(cell.value)))
                    else:
                        row.append(str(cell.value))
                else:
                    row.append(str(cell.value))
            data.append(row)
        sheets[name] = pd.DataFrame(data, dtype=str)
    return sheets

def read_et_with_wps_com(file_path):
    """通过 WPS COM 打开 .et 文件，另存为临时 .xlsx 再读取"""
    wps = None
    temp_xlsx = None
    try:
        # 启动 WPS 表格（ProgID 必须正确，通常为 ET.Application）
        wps = win32com.client.Dispatch("ET.Application")
        wps.Visible = False
        wps.DisplayAlerts = False

        abs_path = str(Path(file_path).resolve())
        workbook = wps.Workbooks.Open(abs_path, UpdateLinks=False, ReadOnly=True)

        # 另存为临时 xlsx
        temp_dir = Path(OUTPUT_DIR) / "_temp_wps"
        temp_dir.mkdir(exist_ok=True)
        temp_xlsx = temp_dir / (Path(file_path).stem + "_tmp.xlsx")
        workbook.SaveAs(str(temp_xlsx), FileFormat=51)  # 51 = xlsx
        workbook.Close(SaveChanges=False)

        # 用 calamine 高效读取
        result = pd.read_excel(temp_xlsx, sheet_name=None, dtype=str, header=None, engine='calamine')
        return result

    except Exception as e:
        raise Exception(f"WPS COM 处理失败: {e}")
    finally:
        # 关闭 WPS 进程
        if wps is not None:
            try:
                wps.Quit()
            except:
                pass
        # 删除临时文件
        if temp_xlsx and Path(temp_xlsx).exists():
            try:
                os.remove(temp_xlsx)
            except:
                pass

def read_et_file(file_path):
    """读取 .et 文件：优先 WPS COM，失败回退到 xlrd"""
    if HAS_WPS_COM:
        try:
            print(f"  尝试用 WPS COM 读取 .et 文件...")
            return read_et_with_wps_com(file_path)
        except Exception as e:
            print(f"  WPS COM 失败: {e}")
            print(f"  回退到 xlrd 方式读取...")
    else:
        print(f"  使用 xlrd 读取 .et 文件...")

    # 回退方式：使用 xlrd
    try:
        return read_xls_manual(file_path)
    except Exception as e:
        raise Exception(f"xlrd 也无法读取: {e}")

def read_excel_flexible(file_path):
    """根据扩展名选择引擎"""
    suffix = file_path.suffix.lower()
    if suffix == '.xlsx':
        return pd.read_excel(file_path, sheet_name=None, dtype=str, header=None, engine='calamine')
    elif suffix == '.xls':
        return read_xls_manual(file_path)
    elif suffix == '.et':
        return read_et_file(file_path)
    else:
        raise Exception(f"不支持的文件类型: {suffix}")

# ---------- 4. 遍历匹配 ----------
id_to_rows = defaultdict(list)
failed_files = []

for idx, excel_file in enumerate(excel_files, 1):
    print(f"处理文件 [{idx}/{len(excel_files)}]: {excel_file.name}")
    try:
        sheets = read_excel_flexible(excel_file)
    except Exception as e:
        print(f"  ❌ 读取失败，跳过此文件: {e}")
        traceback.print_exc()
        failed_files.append((excel_file.name, str(e)))
        continue

    for sheet_name, df in sheets.items():
        if df.empty:
            continue

        # 去除首尾空格
        df = df.apply(lambda col: col.str.strip())

        # 一次性找出包含任意目标身份证号的行
        match_mask = df.isin(id_set).any(axis=1)
        if not match_mask.any():
            continue

        candidate_df = df[match_mask]

        for row_idx, row in candidate_df.iterrows():
            matched_ids = {val for val in row.values if val in id_set}
            if not matched_ids:
                continue

            # 构造带来源信息的行
            row_data = [excel_file.name, sheet_name] + row.tolist()
            for mid in matched_ids:
                id_to_rows[mid].append(row_data)

# ---------- 5. 生成结果文件与日志 ----------
print("开始生成结果文件及日志...")
output_count = 0
with open(LOG_FILE, 'w', encoding='utf-8') as log:
    log.write("结果文件\t行数\n")
    for id_val, rows in id_to_rows.items():
        if not rows:
            continue

        result_df = pd.DataFrame(rows)
        row_count = len(result_df)
        save_path = os.path.join(OUTPUT_DIR, f"{id_val}.xlsx")

        try:
            result_df.to_excel(save_path, index=False, header=False)
            output_count += 1
            log.write(f"{id_val}.xlsx\t{row_count}\n")
            if output_count % 500 == 0:
                print(f"已生成 {output_count} 个文件...")
        except Exception as e:
            print(f"  ❌ 保存失败: {save_path} - {e}")

# 清理临时文件夹
temp_dir = Path(OUTPUT_DIR) / "_temp_wps"
if temp_dir.exists():
    try:
        import shutil
        shutil.rmtree(temp_dir)
    except:
        pass

# ---------- 6. 汇总报告 ----------
print("\n" + "="*60)
print(f"✅ 任务完成!共为 {output_count} 个身份证号生成了结果文件。")
print(f"📄 日志文件: {LOG_FILE}")
if failed_files:
    print(f"⚠️ 有 {len(failed_files)} 个文件读取失败：")
    for fname, reason in failed_files:
        print(f"   - {fname}: {reason}")
else:
    print("✅ 所有文件均已成功处理。")
print("="*60)