"""
消防及義消子女獎學金 AI 智慧審核系統 - 自動儲存與檔案持久化管理模組
負責將同仁上傳照片存入 ./uploads/，並將審核資料自動追加 (append) 寫入 ./data/獎學金總表.xlsx 與 ./data/records.json
"""

import os
import io
import json
import shutil
import zipfile
from datetime import datetime
from typing import List, Dict, Any, Tuple
from PIL import Image

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads")
DATA_DIR = os.path.join(BASE_DIR, "data")
EXCEL_FILE = os.path.join(DATA_DIR, "獎學金總表.xlsx")
JSON_FILE = os.path.join(DATA_DIR, "records.json")
BACKUP_DIR = os.path.join(DATA_DIR, "backup")

def ensure_directories():
    """確保 uploads 與 data 資料夾存在"""
    os.makedirs(UPLOADS_DIR, exist_ok=True)
    os.makedirs(DATA_DIR, exist_ok=True)

def save_case_to_storage(case_dict: Dict[str, Any]) -> Tuple[bool, str]:
    """
    1. 將照片儲存至 ./uploads/{case_id}_{index}.jpg
    2. 將紀錄更新/追加至 ./data/records.json
    3. 自動追加 (append) 至 ./data/獎學金總表.xlsx
    """
    ensure_directories()
    case_id = case_dict.get("id", f"TTFD-{datetime.now().strftime('%Y%m%d%H%M%S')}")
    
    # 1. 儲存圖片至 ./uploads/
    saved_paths = []
    images = case_dict.get("images", [])
    for idx, img in enumerate(images, 1):
        filename = f"{case_id}_page{idx}.jpg"
        full_path = os.path.join(UPLOADS_DIR, filename)
        try:
            if isinstance(img, Image.Image):
                # 轉為 RGB 存檔
                rgb_img = img.convert("RGB") if img.mode in ("RGBA", "P") else img
                rgb_img.save(full_path, "JPEG", quality=92)
                saved_paths.append(filename)
        except Exception as e:
            print(f"Error saving image {filename}: {e}")
            
    case_dict["image_paths"] = saved_paths
    
    # 2. 讀取並追加至 records.json
    all_records = []
    if os.path.exists(JSON_FILE):
        try:
            with open(JSON_FILE, "r", encoding="utf-8") as f:
                all_records = json.load(f)
        except Exception:
            all_records = []
            
    # 序列化處理 (去除無法轉 json 的 PIL 物件)
    json_record = {k: v for k, v in case_dict.items() if k not in ("images",)}
    
    # 檢查是否已存在同案號 (若存在則更新，否則 append)
    existing_idx = next((i for i, r in enumerate(all_records) if r.get("id") == case_id), None)
    if existing_idx is not None:
        all_records[existing_idx] = json_record
    else:
        all_records.append(json_record)
        
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(all_records, f, ensure_ascii=False, indent=2)
        
    # 3. 依「獎學金類別」重新產生 ./data/獎學金總表.xlsx (自動分頁：義消聯合總會獎助學金 / 本局津芳冰城陳慶銳先生獎學金)
    try:
        append_case_to_excel(all_records)
    except Exception as e:
        print(f"Error appending to Excel: {e}")

    return True, f"已成功將 {len(saved_paths)} 張照片存入 ./uploads/，並將數據追加至 ./data/獎學金總表.xlsx！"

def append_case_to_excel(all_records: List[Dict[str, Any]]):
    """依目前完整紀錄清單，重新產生 ./data/獎學金總表.xlsx (依獎學金類別自動分頁)"""
    from excel_exporter import export_scholarship_excel

    ensure_directories()
    excel_bytes = export_scholarship_excel(all_records)
    with open(EXCEL_FILE, "wb") as f:
        f.write(excel_bytes.getvalue())

def load_stored_cases() -> List[Dict[str, Any]]:
    """從 ./data/records.json 與 ./uploads/ 載入所有歷史儲存案件"""
    ensure_directories()
    if not os.path.exists(JSON_FILE):
        return []
        
    try:
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            records = json.load(f)
            
        for r in records:
            images = []
            image_labels = []
            for path_fn in r.get("image_paths", []):
                full_p = os.path.join(UPLOADS_DIR, path_fn)
                if os.path.exists(full_p):
                    try:
                        img = Image.open(full_p)
                        images.append(img)
                        image_labels.append(path_fn)
                    except Exception:
                        pass
            r["images"] = images
            r["image_labels"] = image_labels
            
        return records
    except Exception as e:
        print(f"Error loading stored cases: {e}")
        return []

def package_uploads_zip() -> io.BytesIO:
    """將整個 ./uploads/ 目錄打包為 zip 檔案供承辦人下載"""
    ensure_directories()
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(UPLOADS_DIR):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, UPLOADS_DIR)
                z.write(full_path, rel_path)
    zip_buf.seek(0)
    return zip_buf

def load_records_json() -> List[Dict[str, Any]]:
    """僅讀取 ./data/records.json 的案件資料 (不載入照片)，作為最新、最完整的資料來源"""
    ensure_directories()
    if not os.path.exists(JSON_FILE):
        return []
    try:
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception as e:
        print(f"Error reading records.json: {e}")
        return []

def delete_cases_from_storage(case_ids: List[str]) -> Tuple[List[str], str]:
    """
    依案件編號刪除指定案件；只會動到被指定的編號，其餘案件原封不動。
    1. 先把整份 records.json 備份到 ./data/backup/
    2. 從 records.json 移除指定案件 (先寫暫存檔再替換，避免寫到一半中斷而損毀)
    3. 刪除這些案件的原始照片
    4. 依剩餘案件重新產生 ./data/獎學金總表.xlsx
    返回: (實際刪除的案件編號清單, 訊息)
    """
    ensure_directories()
    target_ids = {str(c) for c in case_ids}
    if not target_ids:
        return [], "未選取任何案件"
    if not os.path.exists(JSON_FILE):
        return [], "找不到案件資料檔，未刪除任何案件"

    try:
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            all_records = json.load(f)
    except Exception as e:
        return [], f"讀取案件資料檔失敗，未刪除任何案件：{e}"

    to_delete = [r for r in all_records if str(r.get("id", "")) in target_ids]
    if not to_delete:
        return [], "選取的案件已不存在，未刪除任何案件"
    keep = [r for r in all_records if str(r.get("id", "")) not in target_ids]

    # 1. 備份
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    shutil.copyfile(JSON_FILE, os.path.join(BACKUP_DIR, f"records_{stamp}.json"))

    # 2. 寫回剩餘案件
    tmp_path = JSON_FILE + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(keep, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, JSON_FILE)

    # 3. 刪除被刪案件的原始照片 (只刪該案件自己記錄的檔名)
    for r in to_delete:
        for fn in r.get("image_paths", []):
            fp = os.path.join(UPLOADS_DIR, os.path.basename(fn))
            if os.path.isfile(fp):
                try:
                    os.remove(fp)
                except Exception:
                    pass

    # 4. 重新產生 Excel 總表
    try:
        append_case_to_excel(keep)
    except Exception as e:
        print(f"Error regenerating Excel: {e}")

    deleted_ids = [str(r.get("id", "")) for r in to_delete]
    return deleted_ids, f"已刪除 {len(deleted_ids)} 筆案件，其餘 {len(keep)} 筆案件未受影響"

def load_all_known_case_ids() -> List[str]:
    """回傳現有案件與所有刪除前備份中出現過的案件編號，讓已刪除的編號不會被重複使用"""
    ids = [str(r.get("id", "")) for r in load_records_json()]
    if os.path.isdir(BACKUP_DIR):
        for fn in os.listdir(BACKUP_DIR):
            if fn.startswith("records_") and fn.endswith(".json"):
                try:
                    with open(os.path.join(BACKUP_DIR, fn), "r", encoding="utf-8") as f:
                        ids.extend(str(r.get("id", "")) for r in json.load(f))
                except Exception:
                    pass
    return ids
