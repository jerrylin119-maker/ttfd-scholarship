"""
消防及義消子女獎學金 AI 智慧審核系統 - 雲端試算表同步模組 (含兩階層組織單位)
"""

import base64
import io
import json
import re
from typing import List, Dict, Any, Optional, Tuple
import requests
from PIL import Image

GOOGLE_APPS_SCRIPT_TEMPLATE = """
// === 請將以下程式碼貼入 Google 試算表的「擴充功能」->「Apps Script」並部署為「網路應用程式」===
// 本版本會依「獎學金類別」自動分別寫入不同分頁 (工作表)：
//   1. 義消聯合總會獎助學金
//   2. 本局津芳冰城陳慶銳先生獎學金
// 若日後新增其他類別，會自動以該類別名稱建立新分頁，無需修改程式碼。
// 支援 GET ?action=export：匯出所有分頁資料為 JSON，供系統端「從雲端試算表復原資料」功能使用。
// 支援 POST {action:"upload_photos"}：把新案件的原始照片存進 Google 雲端硬碟指定資料夾（依大隊分子資料夾），
// 讓照片不受 Streamlit 伺服器重啟影響、能長期保存；回傳的雲端硬碟連結會一併寫回試算表欄位。
// 【安全性修正】同步已改為「依案件編號比對更新」(upsert)，不會再清空整個分頁重寫：
// 即使本機系統資料因伺服器重啟而不完整，也不會誤刪試算表裡既有的其他案件。
// 真正要刪除案件時，系統會另外送出 {action:"delete_records"} 明確指定要刪除的案件編號。

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (action === "export") {
    return exportAllSheets_();
  }
  if (action === "list_photos") {
    return listAllPhotos_();
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    message: "🚒 臺東縣消防局 獎學金同步 Webhook 連線正常！"
  })).setMimeType(ContentService.MimeType.JSON);
}

// 匯出所有分頁資料為 JSON，供系統資料遺失時復原使用 (只有文字欄位，不含原始照片)
function exportAllSheets_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var result = {};
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      var values = sheet.getDataRange().getValues();
      var rows = [];
      if (values.length >= 2) {
        var headers = values[0];
        var idIdx = headers.indexOf("案件編號");
        for (var r = 1; r < values.length; r++) {
          var row = values[r];
          if (idIdx >= 0 && (!row[idIdx] || String(row[idIdx]).trim() === "")) continue;
          var obj = {};
          for (var c = 0; c < headers.length; c++) {
            obj[headers[c]] = row[c];
          }
          rows.push(obj);
        }
      }
      result[sheet.getName()] = rows;
    }
    return ContentService.createTextOutput(JSON.stringify({status: "ok", sheets: result}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

var SHEET_HEADERS_ = [
  "序號", "案件編號", "獎學金類別", "大隊/局本部", "分隊/科室", "申請人姓名", "身分證字號", "子女姓名", "申請組別",
  "學期總平均", "操行成績", "附件檢核(5項)", "審核結果", "判定理由說明", "雲端硬碟連結", "最後同步時間"
];
var ID_COL_ = 2; // 「案件編號」是第 2 欄

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(SHEET_HEADERS_);
    sheet.getRange(1, 1, 1, SHEET_HEADERS_.length).setBackground("#1F4E79").setFontColor("#FFFFFF").setFontWeight("bold");
  }
  return sheet;
}

// 依案件編號比對更新：既有的更新該列，新的加到最後面；不存在於這批資料中的既有列「原封不動保留」，
// 避免因本機資料不完整而誤刪試算表裡其他案件 (例如伺服器重啟後本機只剩剛送出的 1 筆時)
// 【效能】全部先在記憶體裡合併好，最後只用「一次」讀取 + 「一次」寫入處理整張表，
// 不逐列呼叫 Sheets API，避免案件數變多後同步逾時 (逐列讀寫在案件數上升後會慢到超過用戶端逾時設定)。
function upsertRows_(sheet, rows) {
  var numCols = SHEET_HEADERS_.length;
  var lastRow = sheet.getLastRow();
  var existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, numCols).getValues() : [];
  var idToIndex = {};
  for (var i = 0; i < existing.length; i++) {
    var existingId = String(existing[i][ID_COL_ - 1] || "").trim();
    if (existingId) idToIndex[existingId] = i;
  }
  var now = new Date().toLocaleString("zh-TW", {timeZone: "Asia/Taipei"});
  for (var j = 0; j < rows.length; j++) {
    var r = rows[j];
    var id = String(r.id || "").trim();
    var rowValues = [
      "", // 序號留待最後統一重新編號
      id,
      r.scholarship_type || "未分類",
      r.unit_level1 || "",
      r.unit_level2 || "",
      r.applicant_name || "",
      r.applicant_id || "",
      r.child_name || "",
      r.category || "",
      r.semester_gpa !== null && r.semester_gpa !== undefined ? r.semester_gpa : "-",
      r.conduct || "",
      r.attachment_desc || "",
      r.review_status || "",
      r.review_reason || "",
      r.drive_links || "",
      now
    ];
    if (id && idToIndex.hasOwnProperty(id)) {
      existing[idToIndex[id]] = rowValues;
    } else {
      existing.push(rowValues);
      if (id) idToIndex[id] = existing.length - 1;
    }
  }
  for (var k = 0; k < existing.length; k++) {
    existing[k][0] = k + 1; // 重新編排序號欄
  }
  if (existing.length > 0) {
    sheet.getRange(2, 1, existing.length, numCols).setValues(existing);
  }
}

// 重新編排「序號」欄 (第 1 欄)，讓畫面上的序號維持連續，不影響案件編號等其他欄位
function renumberSeq_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var seq = [];
  for (var i = 1; i <= lastRow - 1; i++) seq.push([i]);
  sheet.getRange(2, 1, lastRow - 1, 1).setValues(seq);
}

// 明確刪除指定案件編號的列 (用於後台「刪除指定案件」時，同步移除試算表對應資料)
function deleteRecords_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ids = {};
  (data.ids || []).forEach(function (id) { ids[String(id).trim()] = true; });
  var deleted = [];
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    var idVals = sheet.getRange(2, ID_COL_, lastRow - 1, 1).getValues();
    // 由下往上刪，避免刪除後列號位移影響尚未處理的列
    for (var r = idVals.length - 1; r >= 0; r--) {
      var id = String(idVals[r][0] || "").trim();
      if (id && ids[id]) {
        sheet.deleteRow(r + 2);
        deleted.push(id);
      }
    }
    renumberSeq_(sheet);
  }
  return ContentService.createTextOutput(JSON.stringify({status: "ok", deleted: deleted}))
    .setMimeType(ContentService.MimeType.JSON);
}

// 獎學金申請附件存放的雲端硬碟資料夾 ID (若要換一個資料夾，把下面這串換成新資料夾網址裡的 ID 即可)
var PHOTO_ROOT_FOLDER_ID = "1ykZyvsQ3wqjt5Mfh0-ITBdkudDot0guV";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === "upload_photos") {
      return uploadPhotos_(data);
    }
    if (data.action === "delete_records") {
      return deleteRecords_(data);
    }
    if (data.action === "download_photos") {
      return downloadPhotos_(data);
    }
    return syncAll_(data);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function syncAll_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = data.records;

  // 依「獎學金類別」將資料分組，各類別各自寫入獨立分頁 (只新增/更新，不清除既有其他列)
  var grouped = {};
  var order = [];
  for (var i = 0; i < rows.length; i++) {
    var t = rows[i].scholarship_type || "未分類";
    if (!grouped[t]) {
      grouped[t] = [];
      order.push(t);
    }
    grouped[t].push(rows[i]);
  }

  for (var j = 0; j < order.length; j++) {
    var typeName = order[j];
    var sheet = getOrCreateSheet_(ss, typeName);
    upsertRows_(sheet, grouped[typeName]);
  }

  return ContentService.createTextOutput(JSON.stringify({status: "success", count: rows.length, types: order}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSubfolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

// 把一筆案件的照片存進「PHOTO_ROOT_FOLDER_ID / 大隊名稱 /」底下，回傳每張照片的雲端硬碟連結
// 檔案沿用試算表擁有者（貴局業務科）帳號的預設權限，不會另外公開分享，權限請自行透過雲端硬碟的「共用」設定給需要查看的同仁
function uploadPhotos_(data) {
  var root = DriveApp.getFolderById(PHOTO_ROOT_FOLDER_ID);
  var unitFolder = getOrCreateSubfolder_(root, data.unit_level1 || "未分類單位");
  var links = [];
  var files = data.files || [];
  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    var bytes = Utilities.base64Decode(f.data);
    var blob = Utilities.newBlob(bytes, f.mimeType || "image/jpeg", f.filename || ("photo_" + i + ".jpg"));
    var file = unitFolder.createFile(blob);
    links.push(file.getUrl());
  }
  return ContentService.createTextOutput(JSON.stringify({status: "ok", links: links}))
    .setMimeType(ContentService.MimeType.JSON);
}

// 依雲端硬碟檔案 ID 清單，把原始照片下載回來 (用於本機資料遺失後，把雲端硬碟裡的照片還原回系統)
// 【效能】改用 UrlFetchApp.fetchAll 平行下載，而不是用 DriveApp 一個一個檔案排隊抓：
// 逐一呼叫 DriveApp.getFileById().getBlob() 每個檔案大約要 3-4 秒的網路延遲，7 張照片就要
// 快 30 秒，容易讓 Streamlit 那端等到逾時、被平台判定沒回應而重啟。平行送出後全部檔案
// 加起來的時間跟抓「一張」差不多，不會隨檔案數量線性累加。
function downloadPhotos_(data) {
  var ids = data.file_ids || [];
  var files = [];
  if (ids.length === 0) {
    return ContentService.createTextOutput(JSON.stringify({status: "ok", files: files}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var token = ScriptApp.getOAuthToken();
  var contentReqs = ids.map(function (id) {
    return {
      url: "https://www.googleapis.com/drive/v3/files/" + id + "?alt=media",
      headers: {Authorization: "Bearer " + token},
      muteHttpExceptions: true
    };
  });
  var metaReqs = ids.map(function (id) {
    return {
      url: "https://www.googleapis.com/drive/v3/files/" + id + "?fields=name,mimeType",
      headers: {Authorization: "Bearer " + token},
      muteHttpExceptions: true
    };
  });
  var contentResps = UrlFetchApp.fetchAll(contentReqs);
  var metaResps = UrlFetchApp.fetchAll(metaReqs);
  for (var i = 0; i < ids.length; i++) {
    try {
      var code = contentResps[i].getResponseCode();
      if (code !== 200) {
        files.push({file_id: ids[i], error: "下載失敗，HTTP " + code});
        continue;
      }
      var meta = {};
      try { meta = JSON.parse(metaResps[i].getContentText()); } catch (e) {}
      var blob = contentResps[i].getBlob();
      files.push({
        file_id: ids[i],
        filename: meta.name || (ids[i] + ".jpg"),
        mimeType: meta.mimeType || blob.getContentType() || "image/jpeg",
        data: Utilities.base64Encode(blob.getBytes())
      });
    } catch (err) {
      files.push({file_id: ids[i], error: err.toString()});
    }
  }
  return ContentService.createTextOutput(JSON.stringify({status: "ok", files: files}))
    .setMimeType(ContentService.MimeType.JSON);
}

// 掃描 PHOTO_ROOT_FOLDER_ID 底下所有大隊子資料夾的照片檔案，依檔名 (案件編號_pageN.jpg) 歸類，
// 回傳「案件編號 -> 連結清單」對照表 (files_by_case)。用於批次補回因本機資料遺失、連結記錄跟著不見，
// 但照片其實仍安好存在雲端硬碟裡的案件，不需要逐筆手動搜尋。
// 另外，若承辦人已手動把某張照片的檔名加上「申請表」三個字做標記 (例如 115-30_page3_申請表.jpg)，
// 一併整理成 form_by_case (案件編號 -> {id, url})，讓系統可以只精準下載那一張，不用整批照片都抓。
function listAllPhotos_() {
  try {
    var root = DriveApp.getFolderById(PHOTO_ROOT_FOLDER_ID);
    var byCase = {};
    var formByCase = {};
    var subfolders = root.getFolders();
    while (subfolders.hasNext()) {
      var folder = subfolders.next();
      var files = folder.getFiles();
      while (files.hasNext()) {
        var f = files.next();
        var name = f.getName();
        var m = name.match(/^(.+?)_page\\d+/);
        if (m) {
          var caseId = m[1];
          if (!byCase[caseId]) byCase[caseId] = [];
          byCase[caseId].push(f.getUrl());
          if (name.indexOf("申請表") !== -1) {
            formByCase[caseId] = {id: f.getId(), url: f.getUrl()};
          }
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "ok", files_by_case: byCase, form_by_case: formByCase}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
"""

def test_webhook_connection(webhook_url: str) -> Tuple[bool, str]:
    """測試 Google Apps Script Webhook 是否可連通"""
    if not webhook_url or not webhook_url.startswith("http"):
        return False, "未輸入有效的 Webhook 網址 (開頭需為 https://script.google.com/...)"
    try:
        resp = requests.get(webhook_url.strip(), timeout=10, allow_redirects=True)
        if resp.status_code == 200:
            return True, "✅ 連線測試成功！Google 試算表 Webhook 回應正常。"
        else:
            return False, f"⚠️ 連線回應異常，狀態碼：{resp.status_code}"
    except Exception as e:
        return False, f"❌ 無法連線至該網址：{str(e)}"

def sync_to_google_sheets(webhook_url: str, records: List[Dict[str, Any]]) -> Tuple[bool, str]:
    """
    將目前的審核紀錄同步至指定的 Google Sheets Webhook URL
    """
    clean_url = webhook_url.strip() if webhook_url else ""
    if not clean_url or not clean_url.startswith("http"):
        return False, "未設定有效的 Google 試算表 Webhook 網址"
        
    formatted_records = []
    for r in records:
        att = r.get("attachments", {})
        att_keys = [
            ("application_form", "申請表"),
            ("student_id_or_enrollment", "在學證明"),
            ("transcript", "成績單"),
            ("household_registration", "戶籍謄本"),
            ("service_certificate", "服務證明")
        ]
        present_cnt = sum(1 for k, _ in att_keys if att.get(k, False))
        missing_items = [name for k, name in att_keys if not att.get(k, False)]
        att_desc = "齊全 (5/5)" if present_cnt == 5 else f"缺: {','.join(missing_items)} ({present_cnt}/5)"
        
        formatted_records.append({
            "id": r.get("id", ""),
            "scholarship_type": r.get("scholarship_type", "未分類"),
            "unit_level1": r.get("unit_level1", "未指定"),
            "unit_level2": r.get("unit_level2", "未指定"),
            "applicant_name": r.get("applicant_name", ""),
            "applicant_id": r.get("applicant_id", ""),
            "child_name": r.get("child_name", ""),
            "category": r.get("category", ""),
            "semester_gpa": r.get("semester_gpa"),
            "conduct": r.get("conduct", ""),
            "attachment_desc": att_desc,
            "review_status": r.get("review_status", ""),
            "review_reason": r.get("review_reason", r.get("notes", "")),
            "drive_links": " | ".join(r.get("drive_photo_links", []) or [])
        })
        
    payload = {
        "action": "sync_all",
        "records": formatted_records
    }
    
    try:
        resp = requests.post(
            clean_url,
            data=json.dumps(payload),
            headers={"Content-Type": "application/json"},
            timeout=60,
            allow_redirects=True
        )
        if resp.status_code in (200, 302):
            return True, f"成功同步 {len(formatted_records)} 筆審核資料至 Google 雲端試算表！"
        else:
            return False, f"同步失敗，伺服器回應代碼：{resp.status_code}"
    except Exception as e:
        return False, f"連線至 Google 試算表時發生錯誤: {str(e)}"


def _parse_attachment_desc(desc: Any) -> Dict[str, bool]:
    """把附件檢核欄位的描述文字 (如「齊全 (5/5)」或「缺: 申請表,成績單 (3/5)」) 還原成 5 項附件的布林值"""
    keys = [
        ("application_form", "申請表"),
        ("student_id_or_enrollment", "在學證明"),
        ("transcript", "成績單"),
        ("household_registration", "戶籍謄本"),
        ("service_certificate", "服務證明"),
    ]
    text = str(desc or "").strip()
    if text.startswith("缺"):
        missing_part = text.split("(")[0]
        missing_part = missing_part.replace("缺:", "").replace("缺：", "")
        missing_names = {x.strip() for x in missing_part.split(",") if x.strip()}
        return {k: (name not in missing_names) for k, name in keys}
    # 「齊全 (5/5)」或格式不明時，預設視為齊全 (避免復原後全部誤判為待補件；請大隊複核時人工確認)
    return {k: True for k, _ in keys}


def fetch_cases_from_google_sheets(webhook_url: str) -> Tuple[bool, Any]:
    """
    從 Google 試算表讀回所有分頁的資料，還原成系統可用的案件清單。
    僅適用於本地資料遺失時的緊急復原：只能救回文字欄位 (姓名、身分證字號、成績、審核結果等)，
    原始照片與確切的原始送件時間無法復原。
    成功時回傳 (True, 案件清單)；失敗時回傳 (False, 錯誤訊息字串)。
    """
    clean_url = webhook_url.strip() if webhook_url else ""
    if not clean_url or not clean_url.startswith("http"):
        return False, "未設定有效的 Google 試算表 Webhook 網址"
    try:
        resp = requests.get(clean_url, params={"action": "export"}, timeout=30, allow_redirects=True)
    except Exception as e:
        return False, f"連線至 Google 試算表時發生錯誤: {e}"
    if resp.status_code != 200:
        return False, f"讀取失敗，伺服器回應代碼：{resp.status_code}"
    try:
        data = resp.json()
    except Exception:
        return False, "回應格式無法解析，請確認 Apps Script 已更新為最新版本並重新部署"
    if data.get("status") != "ok":
        return False, f"讀取失敗：{data.get('error', '未知錯誤')}"

    restored: List[Dict[str, Any]] = []
    for rows in (data.get("sheets") or {}).values():
        for row in rows:
            # 只採用目前系統格式的分頁 (含「獎學金類別」欄)，略過試算表中更早期版本
            # 遺留的舊分頁 (例如以大隊命名、或沒有類別欄的「總表」)，避免同一案件重複匯入、
            # 或被舊格式的殘缺資料覆蓋回去。
            if "獎學金類別" not in row:
                continue
            case_id = str(row.get("案件編號", "")).strip()
            if not case_id:
                continue
            gpa_raw = row.get("學期總平均", "")
            try:
                gpa = round(float(gpa_raw), 2)
            except (TypeError, ValueError):
                gpa = None
            status = str(row.get("審核結果", "") or "").strip()
            sync_time = str(row.get("最後同步時間", "") or "").strip()
            restored.append({
                "id": case_id,
                "scholarship_type": str(row.get("獎學金類別", "") or "未分類"),
                "unit_level1": str(row.get("大隊/局本部", "") or "未指定"),
                "unit_level2": str(row.get("分隊/科室", "") or "未指定"),
                "applicant_name": str(row.get("申請人姓名", "") or ""),
                "applicant_id": str(row.get("身分證字號", "") or ""),
                "child_name": str(row.get("子女姓名", "") or ""),
                "category": str(row.get("申請組別", "") or "大專院校"),
                "semester_gpa": gpa,
                "conduct": str(row.get("操行成績", "") or ""),
                "attachments": _parse_attachment_desc(row.get("附件檢核(5項)", "")),
                "review_status": status or "待審核",
                "review_reason": str(row.get("判定理由說明", "") or ""),
                "is_eligible": status == "符合資格",
                "notes": f"⚠️ 本案件由雲端試算表復原，原始照片與確切送件時間已遺失。試算表最後同步時間：{sync_time or '未知'}",
                "review_mode": "paper",  # 復原案件缺少原始照片，一律列為紙本審核
                "drive_photo_links": [u.strip() for u in str(row.get("雲端硬碟連結", "") or "").split("|") if u.strip()],
                "images": [],
                "image_labels": [],
                "submitted_at": sync_time or "未知（復原資料）",
            })
    return True, restored


def upload_photos_to_drive(webhook_url: str, case_id: str, unit_level1: str, images: List[Any]) -> Tuple[bool, Any]:
    """
    把一筆案件的原始照片上傳到 Google 雲端硬碟 (透過 Apps Script Webhook)，讓照片不受 Streamlit
    伺服器重啟影響。成功時回傳 (True, 雲端硬碟連結清單)；失敗時回傳 (False, 錯誤訊息字串)。
    """
    clean_url = webhook_url.strip() if webhook_url else ""
    if not clean_url or not clean_url.startswith("http"):
        return False, "未設定有效的 Google 試算表 Webhook 網址"
    if not images:
        return True, []

    files = []
    for idx, img in enumerate(images, 1):
        try:
            buf = io.BytesIO()
            rgb_img = img.convert("RGB") if img.mode in ("RGBA", "P") else img
            rgb_img.save(buf, "JPEG", quality=92)
            files.append({
                "filename": f"{case_id}_page{idx}.jpg",
                "mimeType": "image/jpeg",
                "data": base64.b64encode(buf.getvalue()).decode("ascii"),
            })
        except Exception as e:
            return False, f"照片編碼失敗 (第 {idx} 張): {e}"

    payload = {"action": "upload_photos", "case_id": case_id, "unit_level1": unit_level1 or "未分類單位", "files": files}
    try:
        resp = requests.post(
            clean_url,
            data=json.dumps(payload),
            headers={"Content-Type": "application/json"},
            timeout=60,
            allow_redirects=True,
        )
    except Exception as e:
        return False, f"連線至 Google 試算表時發生錯誤: {e}"
    if resp.status_code not in (200, 302):
        return False, f"上傳失敗，伺服器回應代碼：{resp.status_code}"
    try:
        data = resp.json()
    except Exception:
        return False, "回應格式無法解析，請確認 Apps Script 已更新為最新版本並重新部署"
    if data.get("status") != "ok":
        return False, f"上傳失敗：{data.get('error', '未知錯誤')}"
    return True, data.get("links", [])


def _extract_drive_file_id(url: str) -> Optional[str]:
    """從 Google 雲端硬碟連結 (如 .../file/d/FILE_ID/view) 擷取檔案 ID"""
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", str(url or ""))
    return m.group(1) if m else None


def download_photos_from_drive(webhook_url: str, drive_links: List[str]) -> Tuple[bool, Any]:
    """
    依雲端硬碟連結清單，把原始照片下載回來還原成 PIL Image 清單。
    用於本機資料遺失（伺服器重啟）後，把已備份在雲端硬碟的照片重新補回案件裡，
    不需要請分隊重新拍照上傳。成功時回傳 (True, PIL Image 清單)；失敗時回傳 (False, 錯誤訊息)。
    """
    clean_url = webhook_url.strip() if webhook_url else ""
    if not clean_url or not clean_url.startswith("http"):
        return False, "未設定有效的 Google 試算表 Webhook 網址"
    file_ids = [_extract_drive_file_id(u) for u in (drive_links or [])]
    file_ids = [f for f in file_ids if f]
    if not file_ids:
        return False, "沒有可用的雲端硬碟連結"

    payload = {"action": "download_photos", "file_ids": file_ids}
    try:
        resp = requests.post(
            clean_url,
            data=json.dumps(payload),
            headers={"Content-Type": "application/json"},
            timeout=90,
            allow_redirects=True,
        )
    except Exception as e:
        return False, f"連線至 Google 試算表時發生錯誤: {e}"
    if resp.status_code not in (200, 302):
        return False, f"下載失敗，伺服器回應代碼：{resp.status_code}"
    try:
        data = resp.json()
    except Exception:
        return False, "回應格式無法解析，請確認 Apps Script 已更新為最新版本並重新部署"
    if data.get("status") != "ok":
        return False, f"下載失敗：{data.get('error', '未知錯誤')}"

    images = []
    errors = []
    for f in data.get("files", []):
        if f.get("error"):
            errors.append(f.get("error"))
            continue
        try:
            img_bytes = base64.b64decode(f["data"])
            img = Image.open(io.BytesIO(img_bytes))
            img.load()
            images.append(img)
        except Exception as e:
            errors.append(str(e))
    if not images:
        return False, f"下載失敗：{'; '.join(errors) if errors else '未知錯誤'}"
    return True, images


def list_drive_photos_by_case(webhook_url: str) -> Tuple[bool, Any]:
    """
    掃描雲端硬碟整個照片資料夾，依檔名 (案件編號_pageN.jpg) 建立「案件編號 -> 連結清單」對照表。
    用於批次補回「照片其實已備份在雲端硬碟、但系統本機記錄的連結遺失」的案件，
    不需要逐筆手動搜尋。
    若承辦人已手動把某張照片的檔名標記「申請表」三個字 (例如 115-30_page3_申請表.jpg)，
    也會一併整理出 form_by_case (案件編號 -> {id, url})，可用來只精準下載那一張。
    成功時回傳 (True, {"files_by_case": {...}, "form_by_case": {...}})；失敗時回傳 (False, 錯誤訊息)。
    """
    clean_url = webhook_url.strip() if webhook_url else ""
    if not clean_url or not clean_url.startswith("http"):
        return False, "未設定有效的 Google 試算表 Webhook 網址"
    try:
        resp = requests.get(clean_url, params={"action": "list_photos"}, timeout=60, allow_redirects=True)
    except Exception as e:
        return False, f"連線至 Google 試算表時發生錯誤: {e}"
    if resp.status_code != 200:
        return False, f"讀取失敗，伺服器回應代碼：{resp.status_code}"
    try:
        data = resp.json()
    except Exception:
        return False, "回應格式無法解析，請確認 Apps Script 已更新為最新版本並重新部署"
    if data.get("status") != "ok":
        return False, f"讀取失敗：{data.get('error', '未知錯誤')}"
    return True, {
        "files_by_case": data.get("files_by_case", {}),
        "form_by_case": data.get("form_by_case", {}),
    }


def delete_from_google_sheets(webhook_url: str, ids: List[str]) -> Tuple[bool, str]:
    """明確通知 Google 試算表刪除指定案件編號的列 (用於後台刪除案件時，讓試算表與本機保持一致)"""
    clean_url = webhook_url.strip() if webhook_url else ""
    if not clean_url or not clean_url.startswith("http"):
        return False, "未設定有效的 Google 試算表 Webhook 網址"
    ids = [str(i) for i in (ids or []) if str(i).strip()]
    if not ids:
        return True, "沒有需要從試算表刪除的案件"
    payload = {"action": "delete_records", "ids": ids}
    try:
        resp = requests.post(
            clean_url,
            data=json.dumps(payload),
            headers={"Content-Type": "application/json"},
            timeout=60,
            allow_redirects=True,
        )
    except Exception as e:
        return False, f"連線至 Google 試算表時發生錯誤: {e}"
    if resp.status_code not in (200, 302):
        return False, f"刪除失敗，伺服器回應代碼：{resp.status_code}"
    try:
        data = resp.json()
    except Exception:
        return False, "回應格式無法解析，請確認 Apps Script 已更新為最新版本並重新部署"
    if data.get("status") != "ok":
        return False, f"刪除失敗：{data.get('error', '未知錯誤')}"
    deleted = data.get("deleted", [])
    return True, f"已從雲端試算表刪除 {len(deleted)} 筆案件"
