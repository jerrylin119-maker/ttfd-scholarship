"""
消防及義消子女獎學金 AI 智慧審核系統 - Google GenAI 多模態解析模組
使用 Google GenAI SDK (Gemini 3.6 Flash / Gemini 3.5 Flash)
"""

import os
import json
import re
from typing import List, Dict, Any, Optional, Tuple
from io import BytesIO
from PIL import Image

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None

ATTACHMENT_NAMES = {
    "application_form": "1. 獎學金申請表",
    "student_id_or_enrollment": "2. 學生證或在學證明",
    "transcript": "3. 前學期成績證明單",
    "household_registration": "4. 戶口名簿影本或戶籍謄本",
    "service_certificate": "5. 消防/義消在職或服務證明文件"
}

SYSTEM_INSTRUCTION = """
你是一位專業的「消防及義消人員子女獎學金」審查 AI 助理。
你的任務是從使用者上傳的申請文件照片（可能包含多張：申請表、成績單、學生證、在學證明、戶口名簿、在職或義消服務證明等）中，精確擷取申請資訊，並檢核 5 項必備附件是否存在與齊全。

請嚴格依照下列規格輸出 JSON 格式：
{
  "applicant_name": "申請人家長姓名 (字串，找不到填空字串)",
  "applicant_id": "申請人家長身分證字號 (字串，如 A123456789，找不到填空字串)",
  "child_name": "子女姓名 (字串，找不到填空字串)",
  "category": "申請組別，只能是：大專院校、高中職、國中、國小 其中之一",
  "semester_gpa": 85.50, // 浮點數，學期總平均成績，精確至小數點後兩位。若無則填 null
  "conduct": "操行成績 (如 88分、甲等、優等、85 等字串)",
  "attachments": {
    "application_form": true, // 是否有包含「獎學金申請表」
    "student_id_or_enrollment": true, // 是否有包含「學生證」或「在學證明」
    "transcript": true, // 是否有包含「前學期成績證明單」
    "household_registration": true, // 是否有包含「戶口名簿」或「戶籍謄本」
    "service_certificate": true // 是否有包含「消防員在職證明」或「義消服務證明」
  },
  "detected_documents": ["申請表", "成績單"], // 識別到的文件名稱清單
  "notes": "備註說明 (例如：成績單蓋有學校戳章、身分證字號清晰、缺在學證明等)"
}

注意事項：
1. 學期總平均 (semester_gpa) 請提取智育/學業/學期總平均分數，必須是 0~100 之間的數值（浮點數）。
2. 若成績單為五等第制或 GPA (4.0/4.3)，請換算為百分制或直接擷取百分制欄位。
3. 附件檢查必須客觀根據圖片中實際呈現的文件內容進行判斷。
4. 請僅輸出符合 JSON 格式的內容，不要包含多餘 Markdown 代碼區塊外的文字。
"""

def get_client(api_key: Optional[str] = None):
    """取得 Google GenAI Client 實例"""
    if not genai:
        raise RuntimeError("未安裝 google-genai 套件，請執行 pip install google-genai")
    
    key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not key:
        raise ValueError("未提供 Gemini API Key，請在介面輸入或設定環境變數 GEMINI_API_KEY")
    
    return genai.Client(api_key=key)

def evaluate_eligibility(semester_gpa: Optional[float], attachments: Dict[str, bool]) -> Tuple[str, str, bool]:
    """
    審查邏輯判定：
    1. 學期總平均 >= 80.0
    2. 5 項必備附件皆齊全 (無缺件)
    返回: (狀態文字, 判定理由, 是否符合)
    """
    missing = []
    for key, label in ATTACHMENT_NAMES.items():
        if not attachments.get(key, False):
            missing.append(label)
    
    all_attachments_present = (len(missing) == 0)
    gpa_valid = (semester_gpa is not None and semester_gpa >= 80.0)
    
    if gpa_valid and all_attachments_present:
        return "符合資格", f"學期平均 {semester_gpa:.2f} (>=80.0)，5 項必備附件齊全", True
    elif not gpa_valid and not all_attachments_present:
        gpa_str = f"{semester_gpa:.2f}" if semester_gpa is not None else "未填/無成績"
        return "不符資格", f"學期平均 {gpa_str} (<80.0) 且缺件: {', '.join(missing)}", False
    elif not gpa_valid:
        gpa_str = f"{semester_gpa:.2f}" if semester_gpa is not None else "未填/無成績"
        return "不符資格", f"學期平均 {gpa_str} 未達 80.0 分申請門檻", False
    else:
        return "待補件", f"成績達標 (平均 {semester_gpa:.2f})，但缺件: {', '.join(missing)}", False

def analyze_scholarship_documents(
    images: List[Image.Image],
    api_key: Optional[str] = None,
    model_name: str = "gemini-3.6-flash"
) -> Dict[str, Any]:
    """
    使用 Gemini 3.6 Flash 分析一組申請文件照片，具備自動容錯與多模型備援
    """
    client = get_client(api_key)
    
    # 準備多模態輸入內容
    contents = []
    
    for i, img in enumerate(images):
        buffered = BytesIO()
        img_format = img.format if img.format else 'JPEG'
        if img_format.upper() not in ['JPEG', 'PNG', 'WEBP']:
            img_format = 'JPEG'
        
        # 轉換為 RGB 以防 RGBA 儲存 JPEG 出錯
        if img.mode in ('RGBA', 'P') and img_format.upper() == 'JPEG':
            rgb_img = img.convert('RGB')
            rgb_img.save(buffered, format='JPEG', quality=90)
            mime_type = 'image/jpeg'
        else:
            img.save(buffered, format=img_format)
            mime_type = f'image/{img_format.lower()}'
            
        img_bytes = buffered.getvalue()
        contents.append(types.Part.from_bytes(data=img_bytes, mime_type=mime_type))
    
    prompt_text = "請詳細檢閱以上上傳之消防/義消子女獎學金申請資料照片，擷取申請人姓名、身分證字號、子女姓名、組別、學期總平均、操行，並逐一檢核 5 項附件（申請表、學生證/在學證明、成績單、戶籍資料、服務證明）。嚴格輸出指定 JSON 結構。"
    contents.append(prompt_text)
    
    # 嘗試的模型清單（優先使用指定模型，若 404 則自動依序嘗試其他可用模型）
    candidate_models = [model_name]
    fallback_pool = [
        "gemini-3.6-flash",
        "gemini-3.5-flash",
        "gemini-3.6-pro",
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash"
    ]
    for m in fallback_pool:
        if m not in candidate_models:
            candidate_models.append(m)
            
    last_error = None
    response_text = None
    used_model = None
    
    for model_to_try in candidate_models:
        try:
            try:
                response = client.models.generate_content(
                    model=model_to_try,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_INSTRUCTION,
                        response_mime_type="application/json",
                        temperature=0.1
                    )
                )
                response_text = response.text
            except Exception:
                response = client.models.generate_content(
                    model=model_to_try,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_INSTRUCTION,
                        temperature=0.1
                    )
                )
                response_text = response.text
                
            if response_text:
                used_model = model_to_try
                break
        except Exception as e:
            last_error = e
            continue
            
    if not response_text:
        raise RuntimeError(f"Gemini 模型呼叫失敗，最後錯誤: {str(last_error)}")
    
    # 解析 JSON
    result = parse_gemini_json_response(response_text)
    result["used_model"] = used_model
    
    # 計算資格
    status, reason, is_eligible = evaluate_eligibility(
        result.get("semester_gpa"),
        result.get("attachments", {})
    )
    result["review_status"] = status
    result["review_reason"] = reason
    result["is_eligible"] = is_eligible
    
    return result

def parse_gemini_json_response(text: str) -> Dict[str, Any]:
    """解析並修復可能的 JSON 回應"""
    cleaned = text.strip()
    if cleaned.startswith("```json"):
        cleaned = cleaned[7:]
    if cleaned.startswith("```"):
        cleaned = cleaned[3:]
    if cleaned.endswith("```"):
        cleaned = cleaned[:-3]
    cleaned = cleaned.strip()
    
    try:
        data = json.loads(cleaned)
    except Exception:
        match = re.search(r'\{.*\}', cleaned, re.DOTALL)
        if match:
            try:
                data = json.loads(match.group(0))
            except Exception:
                data = {}
        else:
            data = {}
    
    semester_gpa = data.get("semester_gpa")
    if semester_gpa is not None:
        try:
            semester_gpa = round(float(semester_gpa), 2)
        except (ValueError, TypeError):
            semester_gpa = None
            
    attachments = data.get("attachments", {})
    if not isinstance(attachments, dict):
        attachments = {}
        
    normalized_attachments = {
        "application_form": bool(attachments.get("application_form", False)),
        "student_id_or_enrollment": bool(attachments.get("student_id_or_enrollment", False)),
        "transcript": bool(attachments.get("transcript", False)),
        "household_registration": bool(attachments.get("household_registration", False)),
        "service_certificate": bool(attachments.get("service_certificate", False))
    }
    
    return {
        "applicant_name": str(data.get("applicant_name", "")).strip(),
        "applicant_id": str(data.get("applicant_id", "")).strip(),
        "child_name": str(data.get("child_name", "")).strip(),
        "category": str(data.get("category", "大專院校")).strip(),
        "semester_gpa": semester_gpa,
        "conduct": str(data.get("conduct", "")).strip(),
        "attachments": normalized_attachments,
        "detected_documents": data.get("detected_documents", []),
        "notes": str(data.get("notes", "")).strip()
    }