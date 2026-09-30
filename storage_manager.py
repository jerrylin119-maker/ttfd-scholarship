"""
消防及義消子女獎學金 AI 智慧審核系統 - 自動儲存與檔案持久化管理模組
負責將同仁上傳照片存入 ./uploads/，並將審核資料自動追加 (append) 寫入 ./data/獎學金總表.xlsx 與 ./data/records.json
"""

import os
import io
import json
import re
import shutil
import zipfile
from datetime import datetime
from typing import List, Dict, Any, Tuple, Optional
from PIL import Image

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads")
DATA_DIR = os.path.join(BASE_DIR, "data")
EXCEL_FILE = os.path.join(DATA_DIR, "獎學金總表.xlsx")
JSON_FILE = os.path.join(DATA_DIR, "records.json")
BACKUP_DIR = os.path.join(DATA_DIR, "backup")
DELETE_LOG_FILE = os.path.join(DATA_DIR, "delete_log.json")

def ensure_directories():
    """確保 uploads 與 data 資料夾存在"""
    os.makedirs(UPLOADS_DIR, exist_ok=True)
    os.makedirs(DATA_DIR, exist_ok=True)

def save_case_to_storage(case_dict: Dict[str, Any], skip_excel: bool = False) -> Tuple[bool, str]:
    """
    1. 將照片儲存至 ./uploads/{case_id}_{index}.jpg
    2. 將紀錄更新/追加至 ./data/records.json
    3. 自動追加 (append) 至 ./data/獎學金總表.xlsx

    skip_excel=True：批次處理多筆案件時使用，先略過 Excel 重新產生 (每次呼叫都會用「全部案件」
    重建整份活頁簿，在迴圈裡逐筆呼叫的話等於重複做好幾十次一樣的事)，等迴圈結束後由呼叫端
    自行呼叫一次 append_case_to_excel(load_records_json())，避免批次操作耗時暴增、拖垮伺服器。
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
    if not skip_excel:
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
    """
    從 ./data/records.json 載入所有歷史儲存案件的文字資料。
    【效能】不會預先把每一筆案件的照片都讀進記憶體——每個瀏覽者一開啟系統，若把全部人的全部照片
    都一次載入，案件與照片數量增加後很容易把伺服器記憶體撐爆而當機。改成「哪一筆案件真的要顯示，
    才用 load_case_images() 現場載入那一筆」，見 app.py 的呼叫方式。
    """
    ensure_directories()
    if not os.path.exists(JSON_FILE):
        return []

    try:
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            records = json.load(f)

        for r in records:
            r.setdefault("images", [])
            r.setdefault("image_labels", [])

        return records
    except Exception as e:
        print(f"Error loading stored cases: {e}")
        return []

def load_case_images(case: Dict[str, Any]) -> List[Any]:
    """依單一案件的 image_paths，從 ./uploads/ 現場載入該筆案件的照片 (PIL Image 清單)。
    只在畫面真的要顯示/使用某一筆案件的照片時才呼叫，不要在迴圈裡對整批案件呼叫。"""
    images = []
    for path_fn in case.get("image_paths", []):
        full_p = os.path.join(UPLOADS_DIR, path_fn)
        if os.path.exists(full_p):
            try:
                img = Image.open(full_p)
                img.load()
                images.append(img)
            except Exception:
                pass
    return images

def package_uploads_zip(only_files=None) -> io.BytesIO:
    """將 ./uploads/ 目錄打包為 zip 檔案供承辦人下載；only_files 有指定時，只打包這些檔名 (用於各大隊只下載自己的照片)"""
    only = None if only_files is None else {os.path.basename(str(f)) for f in only_files}
    ensure_directories()
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(UPLOADS_DIR):
            for file in files:
                if only is not None and file not in only:
                    continue
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

def delete_cases_from_storage(case_ids: List[str], allowed_unit: Optional[str] = None,
                              actor: str = "") -> Tuple[List[str], str]:
    """
    依案件編號刪除指定案件；只會動到被指定的編號，其餘案件原封不動。
    1. 先把整份 records.json 備份到 ./data/backup/
    2. 從 records.json 移除指定案件 (先寫暫存檔再替換，避免寫到一半中斷而損毀)
    3. 刪除這些案件的原始照片
    4. 依剩餘案件重新產生 ./data/獎學金總表.xlsx
    allowed_unit: 若指定 (例如某大隊承辦人)，則只允許刪除該大隊的案件，其他大隊案件即使被選取也不會刪除。
    actor: 執行刪除的帳號 (記錄於刪除紀錄)。
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

    def _allowed(r):
        return allowed_unit is None or r.get("unit_level1") == allowed_unit

    to_delete = [r for r in all_records if str(r.get("id", "")) in target_ids and _allowed(r)]
    if not to_delete:
        return [], "選取的案件不存在，或不屬於您可管理的單位，未刪除任何案件"
    delete_ids = {str(r.get("id", "")) for r in to_delete}
    keep = [r for r in all_records if str(r.get("id", "")) not in delete_ids]

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

    # 3.5 記錄刪除紀錄 (誰、何時、刪了哪一件)
    try:
        log = []
        if os.path.exists(DELETE_LOG_FILE):
            with open(DELETE_LOG_FILE, "r", encoding="utf-8") as f:
                log = json.load(f)
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for r in to_delete:
            log.append({
                "time": now_str, "by": actor or "未記錄",
                "id": r.get("id", ""), "scholarship_type": r.get("scholarship_type", ""),
                "unit_level1": r.get("unit_level1", ""), "unit_level2": r.get("unit_level2", ""),
                "applicant_name": r.get("applicant_name", ""), "child_name": r.get("child_name", ""),
                "submitted_at": r.get("submitted_at", ""),
            })
        with open(DELETE_LOG_FILE, "w", encoding="utf-8") as f:
            json.dump(log, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Error writing delete log: {e}")

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

def _norm(v: Any) -> str:
    return re.sub(r"\s+", "", str(v or "")).lower()

def case_dup_keys(r: Dict[str, Any]) -> List[tuple]:
    """回傳可用來判斷「同一件申請」的識別鍵：(獎學金類別, 子女姓名, 家長姓名) 或 (類別, 子女姓名, 家長身分證)"""
    t, child = _norm(r.get("scholarship_type")), _norm(r.get("child_name"))
    if not child:
        return []
    keys = []
    if _norm(r.get("applicant_name")):
        keys.append((t, child, "name", _norm(r.get("applicant_name"))))
    if _norm(r.get("applicant_id")):
        keys.append((t, child, "id", _norm(r.get("applicant_id"))))
    return keys

def find_duplicate_cases(new_case: Dict[str, Any]) -> List[Dict[str, Any]]:
    """在已存檔案件中，找出與 new_case 疑似為同一件申請的案件 (同類別、同子女，且家長姓名或身分證相同)"""
    new_keys = set(case_dup_keys(new_case))
    if not new_keys:
        return []
    return [r for r in load_records_json() if new_keys & set(case_dup_keys(r))]

def load_delete_log(unit: Optional[str] = None) -> List[Dict[str, Any]]:
    """讀取刪除紀錄 (最新的在前)；指定 unit 時只回傳該大隊的紀錄"""
    if not os.path.exists(DELETE_LOG_FILE):
        return []
    try:
        with open(DELETE_LOG_FILE, "r", encoding="utf-8") as f:
            log = json.load(f)
    except Exception:
        return []
    if unit:
        log = [x for x in log if x.get("unit_level1") == unit]
    return list(reversed(log))

def mark_all_as_paper_review() -> int:
    """把目前所有案件標記為紙本審核 (review_mode = "paper")，回傳異動筆數"""
    ensure_directories()
    if not os.path.exists(JSON_FILE):
        return 0
    with open(JSON_FILE, "r", encoding="utf-8") as f:
        all_records = json.load(f)
    changed = 0
    for r in all_records:
        if r.get("review_mode") != "paper":
            r["review_mode"] = "paper"
            changed += 1
    if changed:
        with open(JSON_FILE, "w", encoding="utf-8") as f:
            json.dump(all_records, f, ensure_ascii=False, indent=2)
        try:
            append_case_to_excel(all_records)
        except Exception as e:
            print(f"Error regenerating Excel: {e}")
    return changed
