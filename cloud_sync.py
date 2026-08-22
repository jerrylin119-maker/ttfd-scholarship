"""
消防及義消子女獎學金 AI 智慧審核系統 - 雲端試算表同步模組 (含兩階層組織單位)
"""

import json
from typing import List, Dict, Any, Tuple
import requests

GOOGLE_APPS_SCRIPT_TEMPLATE = """
// === 請將以下程式碼貼入 Google 試算表的「擴充功能」->「Apps Script」並部署為「網路應用程式」===
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  var rows = data.records;
  
  sheet.clear();
  var headers = [
    "序號", "案件編號", "大隊/局本部", "分隊/科室", "申請人姓名", "身分證字號", "子女姓名", "申請組別",
    "學期總平均", "操行成績", "附件檢核(5項)", "審核結果", "判定理由說明", "同步時間"
  ];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#1F4E79").setFontColor("#FFFFFF").setFontWeight("bold");
  
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    sheet.appendRow([
      i + 1,
      r.id || "",
      r.unit_level1 || "",
      r.unit_level2 || "",
      r.applicant_name || "",
      r.applicant_id || "",
      r.child_name || "",
      r.category || "",
      r.semester_gpa !== null ? r.semester_gpa : "-",
      r.conduct || "",
      r.attachment_desc || "",
      r.review_status || "",
      r.review_reason || "",
      new Date().toLocaleString("zh-TW", {timeZone: "Asia/Taipei"})
    ]);
  }
  
  sheet.autoResizeColumns(1, headers.length);
  return ContentService.createTextOutput(JSON.stringify({status: "success", count: rows.length}))
    .setMimeType(ContentService.MimeType.JSON);
}
"""

def sync_to_google_sheets(webhook_url: str, records: List[Dict[str, Any]]) -> Tuple[bool, str]:
    """
    將目前的審核紀錄同步至指定的 Google Sheets Webhook URL
    """
    if not webhook_url or not webhook_url.startswith("http"):
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
            webhook_url,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=15,
            allow_redirects=True
        )
        if resp.status_code in (200, 302):
            return True, f"成功同步 {len(formatted_records)} 筆審核資料至 Google 雲端試算表！"
        else:
            return False, f"同步失敗，伺服器回應代碼：{resp.status_code}"
    except Exception as e:
        return False, f"連線至 Google 試算表時發生錯誤: {str(e)}"