# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import os
import re
import glob
import warnings

# 忽略 pandas 的未來警告 (例如 to_datetime 的 UserWarning)
warnings.filterwarnings("ignore", category=UserWarning)

# 嘗試匯入 tabula-py 和 xlsxwriter，這在雲端環境中是必要的
try:
    import tabula
    print("tabula-py 模組已載入。")
except ImportError:
    print("tabula-py 模組未安裝，將嘗試安裝...")
    try:
        os.system('apt-get update -qq && apt-get install openjdk-17-jre -y -qq')
        os.environ["JAVA_HOME"] = "/usr/lib/jvm/java-17-openjdk-amd64"
        print("Java Runtime Environment (JRE) 安裝完成。")
        os.system('pip install tabula-py')
        import tabula
    except Exception as e:
        print(f"Java 或 tabula-py 安裝失敗，無法處理 PDF 檔案。錯誤: {e}")
        tabula = None

try:
    import xlsxwriter
    print("xlsxwriter 模組已載入。")
except ImportError:
    print("xlsxwriter 模組未安裝，將嘗試安裝...")
    os.system('pip install xlsxwriter')
    import xlsxwriter

# ====================================================================
# 【 全 域 設 定 】
# ====================================================================
# 輸入資料夾名稱：所有原始 Excel 和 PDF 檔案都必須放在此資料夾內
INPUT_FOLDER_NAME = 'input_files'

# 輸出報告檔案名稱
OUTPUT_FILE_NAME = '案件報告統計表_合併批次處理.xlsx'

# 報告客製化資訊
REPORT_CONFIGURATION = {
    "REPORT_NAME": "Ekran System側錄監控統計表",
    "REPORT_TIME_RANGE": "114年8月份", # 統計期間：請自行修改
    "REPORT_DATE": "114年9月22日", # 建置日期：請自行修改
    "SYSTEM_NAME": "跳板機群組", # 系統別：請自行修改
    "MONITORING_SERVER": "監控跳板機", # 監控跳板機：請自行修改
}

# *** 核心變動: 只需要這兩個欄位 ***
REQUIRED_COLS = [
    '使用者名稱',         # 對應 UserClient Identifier (Client name/User name)
    '活動時間'            # 對應 Activity Time (時長字串, e.g., '1h 16m 8s')
]

# PDF 表格掃描區域 (Top, Left, Bottom, Right) - 調整 Top 跳過報告頂部的 Details/Filter 文字區塊
PDF_AREA = [180, 0, 800, 612] # 假設 A4 頁面寬度 612 點

# ====================================================================
# 【 輔 助 函 數 】
# ====================================================================
def convert_seconds_to_dhms(total_seconds):
    """將總秒數轉換為 '0d 0h 0m 0s' 格式"""
    if pd.isna(total_seconds) or total_seconds <= 0:
        return '0d 0h 0m 0s'

    total_seconds = int(total_seconds)
    days = total_seconds // (24 * 3600)
    total_seconds %= (24 * 3600)
    hours = total_seconds // 3600
    total_seconds %= 3600
    minutes = total_seconds // 60
    seconds = total_seconds % 60

    return f"{days}d {hours}h {minutes}m {seconds}s"

def convert_dhms_to_seconds(dhms_str):
    """將 '0d 0h 0m 0s' 或 '1h 16m 8s' 格式的時長字串轉換為總秒數"""
    if pd.isna(dhms_str) or not isinstance(dhms_str, str):
        return 0

    # 清理並統一格式，確保空格分隔
    dhms_str = dhms_str.replace('\n', ' ').strip()
    total_seconds = 0

    # 匹配所有數字及其後面的單位 (d, h, m, s)
    matches = re.findall(r'(\d+)\s*([dhms])', dhms_str, re.IGNORECASE)

    for value, unit in matches:
        value = int(value)
        unit = unit.lower()
        if unit == 'd':
            total_seconds += value * 86400
        elif unit == 'h':
            total_seconds += value * 3600
        elif unit == 'm':
            total_seconds += value * 60
        elif unit == 's':
            total_seconds += value

    return total_seconds

# ====================================================================
# 【 主 要 處 理 函 數 】
# ====================================================================
def read_file_to_df(filepath):
    """根據檔案類型讀取檔案並返回 DataFrame"""
    filename = os.path.basename(filepath)
    print(f"--- 正在處理檔案: {filename} ---")

    df = pd.DataFrame()
    N_REQUIRED_COLS = len(REQUIRED_COLS)

    try:
        if filepath.lower().endswith(('.xlsx', '.xls')):
            # *** 處理 Excel 檔案 (核心精進: 讀取所有工作表，強制映射前 2 欄) ***
            all_sheets = pd.read_excel(filepath, sheet_name=None, header=None)

            df_list = []
            for sheet_name, sheet_df in all_sheets.items():
                print(f"DEBUG: 正在處理工作表: {sheet_name}")
                if sheet_df.shape[0] >= 6:
                    df_data = sheet_df[5:].copy()

                    if df_data.shape[1] >= N_REQUIRED_COLS:
                        df_temp = df_data.iloc[:, :N_REQUIRED_COLS].copy()
                        df_temp.columns = REQUIRED_COLS
                        df_list.append(df_temp)
                    else:
                        print(f"警告: 工作表 '{sheet_name}' 實際欄位數 ({df_data.shape[1]}) 少於所需欄位數 ({N_REQUIRED_COLS})，跳過。")
                else:
                    print(f"警告: 工作表 '{sheet_name}' 行數不足 (小於 6 行數據)，跳過。")

            if df_list:
                df = pd.concat(df_list, ignore_index=True)
                print(f"檔案讀取與合併完成。總行數: {len(df)}")
                print(f"Excel 欄位已透過位置 (前 {N_REQUIRED_COLS} 欄) 映射完成。")
            else:
                print(f"檔案 {filename} 讀取失敗或內容為空。跳過。")
                return None

        elif filepath.lower().endswith('.pdf'):
            if tabula is None:
                print("錯誤: tabula-py 模組未能成功載入，無法處理 PDF 檔案。")
                return None

            print(f"嘗試從 PDF 檔案 {filename} 提取表格...")
            dfs = tabula.read_pdf(
                filepath,
                pages='all',
                stream=True,
                area=PDF_AREA,
                pandas_options={'header': None}
            )

            if not dfs:
                print(f"錯誤: 從 PDF 檔案 {filename} 提取表格失敗。未找到任何表格。")
                return None

            df = pd.concat(dfs, ignore_index=True)
            print(f"PDF 提取完成。總行數: {len(df)}")

            N_COLS_ACTUAL = df.shape[1]
            if N_COLS_ACTUAL >= N_REQUIRED_COLS:
                print(f"偵測到 {N_COLS_ACTUAL} 個欄位，將取前 {N_REQUIRED_COLS} 欄進行映射。")
                col_rename_map = {i: col_name for i, col_name in enumerate(REQUIRED_COLS)}
                actual_cols_to_rename = {k: v for k, v in col_rename_map.items() if k < N_COLS_ACTUAL}
                df.rename(columns=actual_cols_to_rename, inplace=True)
                df = df.reindex(columns=REQUIRED_COLS)
            else:
                print(f"錯誤: PDF 提取的欄位數 ({N_COLS_ACTUAL}) 少於所需欄位數 ({N_REQUIRED_COLS})。無法處理。")
                return None
        else:
            print(f"警告: 檔案 {filename} 格式不受支援。跳過。")
            return None

    except Exception as e:
        print(f"處理檔案 {filename} 過程中發生錯誤: {e}")
        return None

    if df.empty:
        print(f"檔案 {filename} 讀取失敗或內容為空。跳過。")
        return None

    initial_row_count = len(df)
    df['使用者名稱'] = df['使用者名稱'].ffill()

    df_clean = df.copy()
    df_clean['活動時間'] = df_clean['活動時間'].astype(str).str.replace(r'[^\x00-\x7F]+', ' ', regex=True).str.replace(r'\s+', ' ', regex=True).str.strip()
    
    date_like_filter = df_clean['活動時間'].str.contains(r'[/:-]', na=False)
    if date_like_filter.any():
        df_clean.loc[date_like_filter, '活動時間'] = np.nan

    df_clean.dropna(subset=['活動時間'], how='any', inplace=True)
    df_clean['Total Seconds'] = df_clean['活動時間'].apply(convert_dhms_to_seconds)
    
    rows_lost_in_seconds_filter = len(df_clean)
    df_clean = df_clean[df_clean['Total Seconds'] > 0]
    rows_lost_in_seconds_filter -= len(df_clean)

    if rows_lost_in_seconds_filter > 0 and len(df_clean) == 0:
        return None

    def clean_user_client_identifier(identifier_str):
        if pd.isna(identifier_str) or not isinstance(identifier_str, str):
            return identifier_str
        cleaned = re.sub(r'(\s+[\d\s]*[dhms]\s*)+$', '', identifier_str, flags=re.IGNORECASE).strip()
        return cleaned.strip()

    df_clean['Group Key'] = df_clean['使用者名稱'].apply(clean_user_client_identifier)
    df_clean.dropna(subset=['Group Key'], how='any', inplace=True)
    df_clean['Source File'] = filename

    return df_clean

def generate_report(file_data_list, config):
    """根據每個檔案的數據列表生成 Excel 報告，每個檔案一個分頁"""
    if not file_data_list:
        print("沒有有效的數據可以生成報告。")
        return

    output_filepath = config['OUTPUT_FILE_NAME']
    writer = pd.ExcelWriter(output_filepath, engine='xlsxwriter')
    workbook = writer.book

    bold_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    left_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
    title_format = workbook.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'font_size': 12})
    note_format = workbook.add_format({'text_wrap': True, 'align': 'left', 'valign': 'top'})

    for df_data in file_data_list:
        source_filename = df_data['Source File'].iloc[0]
        merged_df = df_data.copy()

        sheet_name = source_filename.replace('附件-', '').replace('.pdf', '').replace('.xlsx', '').replace('.xls', '')[:31]
        worksheet = workbook.add_worksheet(sheet_name)
        print(f"--- 正在生成分頁: {sheet_name} ---")

        stats_grouped = merged_df.groupby('Group Key').agg(
            Count=('Total Seconds', 'count'),
            Total_Seconds=('Total Seconds', 'sum')
        ).reset_index()

        stats_grouped['總時長'] = stats_grouped['Total_Seconds'].apply(convert_seconds_to_dhms)
        stats_grouped['次數'] = stats_grouped['Count'].astype(int)
        stats_grouped.rename(columns={'Group Key': '使用者名稱'}, inplace=True)
        stats_df = stats_grouped[['使用者名稱', '次數', '總時長']]

        worksheet.merge_range('A1:C1', config['REPORT_NAME'], title_format)
        row = 1
        col_width = [15, 25]

        worksheet.write(row, 0, '統計期間', left_format)
        worksheet.write(row, 1, config['REPORT_TIME_RANGE'], left_format)
        worksheet.write(row, 2, '建置日期', left_format)
        worksheet.write(row, 3, config['REPORT_DATE'], left_format)

        worksheet.set_column(0, 0, col_width[0])
        worksheet.set_column(1, 1, col_width[1])
        worksheet.set_column(2, 2, col_width[0])
        worksheet.set_column(3, 3, col_width[1])

        worksheet.write(row + 1, 0, '系統別', left_format)
        worksheet.write(row + 1, 1, config['SYSTEM_NAME'], left_format)
        worksheet.write(row + 1, 2, '資料來源', left_format)
        worksheet.write(row + 1, 3, source_filename, left_format)

        total_seconds_all = merged_df['Total Seconds'].sum()
        total_time_dhms = convert_seconds_to_dhms(total_seconds_all)
        worksheet.write(row + 2, 0, '總時長 (所有用戶)', left_format)
        worksheet.merge_range(row + 2, 1, row + 2, 3, total_time_dhms, center_format)

        start_row = 6
        worksheet.write(start_row - 1, 0, '使用者名稱', bold_format)
        worksheet.write(start_row - 1, 1, '次數', bold_format)
        worksheet.write(start_row - 1, 2, '總時長', bold_format)

        worksheet.set_column(0, 0, 25)
        worksheet.set_column(1, 1, 10, center_format)
        worksheet.set_column(2, 2, 20, center_format)

        for row_idx, row_data in stats_df.iterrows():
            worksheet.write(start_row + row_idx, 0, row_data['使用者名稱'], left_format)
            worksheet.write(start_row + row_idx, 1, row_data['次數'], center_format)
            worksheet.write(start_row + row_idx, 2, row_data['總時長'], center_format)

        note_start_row = start_row + len(stats_df) + 2
        note_text = (
            "(1).連線次數及使用時長皆包含測試連線(可能僅連線數秒)\n"
            "(2).用戶說明(我在自己擴充)\n"
            f"(3).本頁數據僅來自檔案: {source_filename}"
        )
        worksheet.merge_range(note_start_row, 0, note_start_row + 2, 3, note_text, note_format)

    writer.close()
    print(f"\n✅ 統計報告已成功生成至: {output_filepath}")

# ====================================================================
# 【 主 程式 執 行 區 塊 】
# ====================================================================
if __name__ == "__main__":
    # 使用 exist_ok=True 確保資料夾存在時不會報錯或衝突
    os.makedirs(INPUT_FOLDER_NAME, exist_ok=True)

    all_files = glob.glob(os.path.join(INPUT_FOLDER_NAME, '*.xlsx')) + \
                glob.glob(os.path.join(INPUT_FOLDER_NAME, '*.xls')) + \
                glob.glob(os.path.join(INPUT_FOLDER_NAME, '*.pdf'))

    if not all_files:
        print(f"錯誤: 在資料夾 '{INPUT_FOLDER_NAME}' 中未找到任何 Excel (.xlsx/.xls) 或 PDF (.pdf) 檔案。")
        print(f"請確認你的 Excel 檔案是否確實放在 GitHub 專案的 '{INPUT_FOLDER_NAME}' 資料夾內！")
    else:
        print(f"找到 {len(all_files)} 個檔案，開始批次處理...")
        file_data_list = []

        for filepath in all_files:
            df_clean = read_file_to_df(filepath)
            if df_clean is not None and not df_clean.empty:
                file_data_list.append(df_clean)

        if not file_data_list:
            print("所有檔案處理後皆無有效數據。報告生成失敗。")
        else:
            report_config = REPORT_CONFIGURATION.copy()
            report_config['OUTPUT_FILE_NAME'] = OUTPUT_FILE_NAME
            generate_report(file_data_list, report_config)
