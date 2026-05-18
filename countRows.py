import os
import pandas as pd
import openpyxl
import xlrd
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

# 计算output文件夹下的所有表格文件内部的行数，并将各个表格名称、内部内容行数写入日志文件log.xlsx

# ---------- 配置 ----------
INPUT_DIR = "output"          # 表格所在文件夹
OUTPUT_FILE = "log.xlsx"      # 结果文件
MAX_WORKERS = 32              # 线程数，可根据磁盘性能调整（HDD 建议 32~64，SSD 可 64~128）
EXCEL_EXTS = ('.xlsx', '.xls', '.xlsm')  # 处理格式，忽略 .xlsb / .et 等
# -------------------------

def count_rows(filepath: str) -> int:
    """快速获取 Excel 文件所有工作表的总行数（含表头）"""
    ext = os.path.splitext(filepath)[1].lower()
    total = 0
    
    try:
        if ext in ('.xlsx', '.xlsm'):
            # openpyxl 只读模式，仅访问元数据
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            for ws in wb.worksheets:
                # 只读模式下 max_row 可能为 None（空表），用 0 代替
                total += ws.max_row or 0
            wb.close()
            
        elif ext == '.xls':
            # xlrd 直接获取 nrows
            wb = xlrd.open_workbook(filepath, on_demand=True)
            for sheet in wb.sheets():
                total += sheet.nrows
            wb.release_resources()
            
        else:
            # 其他格式（如 .xlsb）回退到 pandas，速度较慢但保证兼容
            sheets = pd.read_excel(filepath, sheet_name=None)
            total = sum(df.shape[0] for df in sheets.values())
            
    except Exception as e:
        print(f"[错误] {filepath}: {e}")
        return -1   # 返回 -1 表示读取失败，最终写入时记为 0 或标记错误
    
    return total


def main():
    if not os.path.exists(INPUT_DIR):
        print(f"文件夹 '{INPUT_DIR}' 不存在")
        return

    # 收集所有待处理文件
    files = [
        f for f in os.listdir(INPUT_DIR)
        if f.lower().endswith(EXCEL_EXTS)
    ]
    total_files = len(files)
    print(f"发现 {total_files} 个 Excel 文件，开始统计（线程数={MAX_WORKERS}）...")

    # 结果列表：字典形式，方便转 DataFrame
    results = []
    errors = 0

    # 多线程处理
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        # 提交所有任务
        future_to_file = {
            executor.submit(count_rows, os.path.join(INPUT_DIR, f)): f
            for f in files
        }

        # 带进度条收集结果
        with tqdm(total=total_files, desc="处理进度", unit="file") as pbar:
            for future in as_completed(future_to_file):
                fname = future_to_file[future]
                try:
                    rows = future.result()
                except Exception as e:
                    print(f"[异常] {fname}: {e}")
                    rows = -1

                if rows == -1:
                    errors += 1
                    rows = 0   # 或保留为 -1 便于筛选
                results.append({"文件名称": fname, "总行数": rows})
                pbar.update(1)

    # 写入结果 Excel
    df_out = pd.DataFrame(results)
    df_out.to_excel(OUTPUT_FILE, index=False)
    print(f"\n完成！成功 {total_files - errors} 个文件，失败 {errors} 个，结果保存至 {OUTPUT_FILE}")


if __name__ == "__main__":
    main()