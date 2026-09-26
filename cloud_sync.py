"""
消防及義消子女獎學金 AI 智慧審核系統 - 雲端試算表同步模組 (含兩階層組織單位)
"""

import json
from typing import List, Dict, Any, Tuple
import requests

GOOGLE_APPS_SCRIPT_TEMPLATE = """
// === 請將以下程式碼貼入 Google 試算表的「擴充功能」->「Apps Script」並部署為「網路應用程式」===
// 本版本會依「獎學金類別」自動分別寫入不同分頁 (工作表)：
//   1. 義消聯合總會獎助學金
//   2. 本局津芳冰城陳慶銳先生獎學金
// 若日後新增其他類別，會自動以該類別名稱建立新分頁，無需修改程式碼。
// 支援 GET ?action=export：匯出所有分頁資料為 JSON，供系統端「從雲端試算表復原資料」功能使用。

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (action === "export") {
    return exportAllSheets_();
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

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function writeHeaderAndRows_(sheet, rows) {
  // 清除現有內容並寫入 15 欄標準表頭 (含獎學金類別)
  sheet.clear();
  var headers = [
    "序號", "案件編號", "獎學金類別", "大隊/局本部", "分隊/科室", "申請人姓名", "身分證字號", "子女姓名", "申請組別",
    "學期總平均", "操行成績", "附件檢核(5項)", "審核結果", "判定理由說明", "最後同步時間"
  ];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#1F4E79").setFontColor("#FFFFFF").setFontWeight("bold");

  var now = new Date().toLocaleString("zh-TW", {timeZone: "Asia/Taipei"});
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    sheet.appendRow([
      i + 1,
      r.id || "",
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
      now
    ]);
  }
  sheet.autoResizeColumns(1, headers.length);
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);
    var rows = data.records;

    // 依「獎學金類別」將資料分組，各類別各自寫入獨立分頁
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
      writeHeaderAndRows_(sheet, grouped[typeName]);
    }

    return ContentService.createTextOutput(JSON.stringify({status: "success", count: rows.length, types: order}))
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
            "review_reason": r.get("review_reason", r.get("notes", ""))
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
            timeout=20,
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
                "images": [],
                "image_labels": [],
                "submitted_at": sync_time or "未知（復原資料）",
            })
    return True, restored
