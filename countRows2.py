import os
import shutil
import pandas as pd
import openpyxl
import xlrd
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

# 计算output文件夹下的所有表格文件内部的行数，并将各个表格名称、内部内容行数写入日志文件log.xlsx
# 新建文件夹并根据表格文件的行数，从output复制过来存入对应的文件夹



# ---------- 配置 ----------
INPUT_DIR = "output"          # 源表格文件夹（相对路径，会自动转为绝对路径）
OUTPUT_FILE = "log.xlsx"      # 统计结果文件
KEY_DIR = "keyPersonnel"      # 分级存储根目录
MAX_WORKERS_STAT = 32         # 统计行数时的线程数
MAX_WORKERS_COPY = 32         # 复制文件时的线程数
EXCEL_EXTS = ('.xlsx', '.xls', '.xlsm')  # 支持格式
# -------------------------

# 区间定义：按顺序判断，满足条件即存入对应子文件夹（先大后小）
CATEGORIES = [
    ("1000行以上", lambda r: r > 1000),
    ("大于500行小于等于1000行", lambda r: 500 < r <= 1000),
    ("大于300行小于等于500行", lambda r: 300 < r <= 500),
    ("大于200行小于等于300行", lambda r: 200 < r <= 300),
    ("大于100行小于等于200行", lambda r: 100 < r <= 200),
    ("大于等于50行小于等于100行", lambda r: 50 <= r <= 100),
]

def count_rows(filepath: str) -> int:
    """快速获取 Excel 文件所有工作表的总行数（含表头），失败返回 -1"""
    ext = os.path.splitext(filepath)[1].lower()
    total = 0
    try:
        if ext in ('.xlsx', '.xlsm'):
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            for ws in wb.worksheets:
                total += ws.max_row or 0
            wb.close()
        elif ext == '.xls':
            wb = xlrd.open_workbook(filepath, on_demand=True)
            for sheet in wb.sheets():
                total += sheet.nrows
            wb.release_resources()
        else:
            sheets = pd.read_excel(filepath, sheet_name=None)
            total = sum(df.shape[0] for df in sheets.values())
    except Exception as e:
        # 静默处理，由上层统计错误数量
        total = -1
    return total

def categorize(rows: int) -> str:
    """根据行数返回子文件夹名，若不归入任何区间返回空字符串"""
    if rows <= 0:          # 包括 -1（读取失败）
        return ""
    for folder, condition in CATEGORIES:
        if condition(rows):
            return folder
    return ""

def ensure_dirs(base: str):
    """一次性创建基文件夹及其子文件夹"""
    os.makedirs(base, exist_ok=True)
    for folder, _ in CATEGORIES:
        os.makedirs(os.path.join(base, folder), exist_ok=True)

def copy_file(args):
    """单个文件复制任务，返回 (src, dest, success)"""
    src, dest = args
    try:
        shutil.copy2(src, dest)
        return (src, dest, True)
    except Exception as e:
        print(f"[复制失败] {src} -> {dest}: {e}")
        return (src, dest, False)

def main():
    # 1. 路径预处理与检查
    input_dir = os.path.abspath(INPUT_DIR)   # 转为绝对路径，便于定位
    if not os.path.exists(input_dir):
        print(f"错误：文件夹不存在 -> {input_dir}")
        print("请确认 'output' 文件夹与脚本在同一目录，或修改 INPUT_DIR 为正确路径。")
        return
    if not os.path.isdir(input_dir):
        print(f"错误：路径并非文件夹 -> {input_dir}")
        return

    print(f"源文件夹: {input_dir}")
    print(f"统计结果: {os.path.abspath(OUTPUT_FILE)}")
    print(f"分级目录: {os.path.abspath(KEY_DIR)}")

    # 2. 创建目标文件夹结构
    ensure_dirs(KEY_DIR)

    # 3. 收集所有待处理的 Excel 文件
    try:
        all_files = os.listdir(input_dir)
    except PermissionError:
        print("错误：无权限访问该文件夹，请以管理员身份运行或检查权限。")
        return

    files = [f for f in all_files if f.lower().endswith(EXCEL_EXTS)]
    total_files = len(files)
    if total_files == 0:
        print("未找到任何符合扩展名的 Excel 文件，请检查扩展名列表或文件夹内容。")
        return

    print(f"发现 {total_files} 个 Excel 文件，开始统计行数（线程数={MAX_WORKERS_STAT}）...")

    results = []          # 存储统计信息
    copy_tasks = []       # 需要复制的任务: (src_path, dest_path)
    errors = 0

    # 4. 多线程统计行数
    with ThreadPoolExecutor(max_workers=MAX_WORKERS_STAT) as executor:
        future_to_file = {
            executor.submit(count_rows, os.path.join(input_dir, f)): f
            for f in files
        }
        with tqdm(total=total_files, desc="统计行数", unit="file") as pbar:
            for future in as_completed(future_to_file):
                fname = future_to_file[future]
                src_path = os.path.join(input_dir, fname)
                try:
                    rows = future.result()
                except Exception as e:
                    print(f"[严重异常] {fname}: {e}")
                    rows = -1

                if rows == -1:
                    errors += 1
                    category = ""
                    rows_for_log = 0   # 日志中显示0，表示失败
                else:
                    category = categorize(rows)
                    rows_for_log = rows

                results.append({
                    "文件名称": fname,
                    "总行数": rows_for_log,
                    "分类": category if category else "未分类",
                    "完整路径": src_path
                })

                if category:
                    dest_dir = os.path.join(KEY_DIR, category)
                    dest_path = os.path.join(dest_dir, fname)
                    copy_tasks.append((src_path, dest_path))

                pbar.update(1)

    # 5. 写入统计结果 log.xlsx
    df_out = pd.DataFrame(results)
    df_out.to_excel(OUTPUT_FILE, index=False)
    print(f"\n统计完成：成功 {total_files - errors} 个，失败 {errors} 个，结果保存至 {OUTPUT_FILE}")

    # 6. 并行复制文件到分类文件夹
    if copy_tasks:
        print(f"开始复制 {len(copy_tasks)} 个文件到 '{KEY_DIR}' 子文件夹...")
        copy_success = 0
        copy_fail = 0
        with ThreadPoolExecutor(max_workers=MAX_WORKERS_COPY) as executor:
            futures = [executor.submit(copy_file, task) for task in copy_tasks]
            with tqdm(total=len(copy_tasks), desc="复制进度", unit="file") as pbar:
                for future in as_completed(futures):
                    _, _, ok = future.result()
                    if ok:
                        copy_success += 1
                    else:
                        copy_fail += 1
                    pbar.update(1)
        print(f"复制完成：成功 {copy_success} 个，失败 {copy_fail} 个")
    else:
        print("没有文件符合分类条件，未执行复制。")

if __name__ == "__main__":
    main()