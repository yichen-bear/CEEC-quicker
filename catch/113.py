import requests
import pdfplumber
import io
import json
from pathlib import Path

# --- 設定參數 ---
pdf_url = "https://www.uac.edu.tw/113data/113_04.pdf"
output_filename = "113.json" # 輸出檔案名稱（簡化版）

# --- 主要執行流程 ---
def download_and_parse_pdf(url):
    """下載PDF並解析表格"""
    print(f"正在從 {url} 下載資料...")
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()  # 檢查下載是否成功
        print("下載完成。")
    except requests.exceptions.RequestException as e:
        print(f"下載失敗: {e}")
        return None

    # 使用pdfplumber解析PDF
    print("正在解析PDF表格，這可能需要幾秒鐘...")
    try:
        with pdfplumber.open(io.BytesIO(response.content)) as pdf:
            all_rows = []
            # 遍歷每一頁
            for page_num, page in enumerate(pdf.pages):
                print(f"  處理第 {page_num + 1} 頁...")
                # 嘗試用線條策略擷取表格
                table = page.extract_table({
                    "vertical_strategy": "lines", 
                    "horizontal_strategy": "lines"
                })
                
                # 如果失敗，改用文字策略
                if table is None:
                    table = page.extract_table({
                        "vertical_strategy": "text", 
                        "horizontal_strategy": "text"
                    })

                if table:
                    # 將當前頁的表格加入總列表
                    all_rows.extend(table)
                else:
                    print(f"    警告: 在第 {page_num + 1} 頁找不到表格。")

            if not all_rows:
                print("錯誤：無法從PDF中提取任何表格。")
                return None

            print(f"表格解析完成，共處理 {len(pdf.pages)} 頁，提取 {len(all_rows)} 行原始資料。")
            return all_rows

    except Exception as e:
        print(f"解析PDF時發生錯誤: {e}")
        return None

def clean_and_structure_data(rows):
    """清理並將原始表格資料轉為結構化字典列表（只保留需要的欄位）"""
    if not rows:
        return []

    # 我們只需要前面7個欄位（索引0-6）
    # 系組代碼(0), 校名(1), 系組名(2), 採計及加權(3), 
    # 錄取人數(含外加)(4), 普通生錄取分數(5), 普通生同分參酌(6)
    
    # 定義我們要保留的欄位名稱
    selected_columns = [
        '系組代碼', '校名', '系組名', '採計及加權',
        '錄取人數(含外加)', '普通生錄取分數', '普通生同分參酌'
    ]

    data_list = []
    # 從第二行開始是資料（跳過原始表頭）
    for row_num, row in enumerate(rows[1:]):
        # 跳過明顯的分頁標題行或空行
        if not row or not any(row):
            continue
        # 跳過可能包含表頭文字的列（例如頁首重複的標題）
        first_cell = str(row[0]) if row[0] else ''
        if '系組代碼' in first_cell or '校名' in first_cell:
            continue

        # 確保此行至少有7個欄位，不足則補空字串
        if len(row) < 7:
            # 如果欄位不足，跳過這行（可能是無效資料）
            continue
            
        # 只取前7個欄位，並建立字典
        cleaned_row = {}
        for i, col_name in enumerate(selected_columns):
            # 取得對應欄位的值，清理空白和換行
            cell_value = row[i].strip().replace('\n', ' ') if row[i] and row[i].strip() else None
            
            # 處理 '-----' 代表無資料的情況
            if cell_value and '-----' in cell_value:
                cell_value = None
                
            # 對數值欄位嘗試轉為數字
            if cell_value and col_name in ['錄取人數(含外加)', '普通生錄取分數']:
                try:
                    if '.' in cell_value:
                        cell_value = float(cell_value)
                    else:
                        # 錄取人數應該是整數
                        if col_name == '錄取人數(含外加)':
                            cell_value = int(cell_value)
                        else:
                            # 分數可能有小數
                            cell_value = float(cell_value) if '.' in cell_value else int(cell_value)
                except ValueError:
                    pass # 保留原始字串

            cleaned_row[col_name] = cell_value

        # 確保至少有基本的資料（系組代碼存在）
        if cleaned_row.get('系組代碼'):
            data_list.append(cleaned_row)

    print(f"資料清理完成，只保留7個核心欄位，共取得 {len(data_list)} 筆有效系組資料。")
    return data_list

def save_to_json(data, filename):
    """將資料儲存為JSON檔案"""
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"資料已成功儲存為 {filename}")
        file_path = Path(filename).absolute()
        print(f"檔案絕對路徑: {file_path}")
        print(f"檔案大小: {file_path.stat().st_size / 1024:.2f} KB")
    except Exception as e:
        print(f"儲存JSON檔案時發生錯誤: {e}")

# --- 執行主程式 ---
if __name__ == "__main__":
    print("開始爬取大學分發入學資料（簡化版，只保留核心欄位）...")
    raw_table = download_and_parse_pdf(pdf_url)
    
    if raw_table:
        structured_data = clean_and_structure_data(raw_table)
        if structured_data:
            save_to_json(structured_data, output_filename)
            # 顯示前幾筆資料作為範例
            print("\n資料範例 (前3筆):")
            print(json.dumps(structured_data[:3], ensure_ascii=False, indent=2))
            
            # 顯示欄位列表
            print(f"\n輸出的JSON檔案包含以下欄位：")
            for col in structured_data[0].keys():
                print(f"  - {col}")
        else:
            print("無法結構化資料，程式結束。")
    else:
        print("無法取得PDF資料，程式結束。")