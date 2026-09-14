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

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    message: "🚒 臺東縣消防局 獎學金同步 Webhook 連線正常！"
  })).setMimeType(ContentService.MimeType.JSON);
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
