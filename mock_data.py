"""
消防及義消子女獎學金 AI 智慧審核系統 - 示範資料與測試樣張產生模組 (含臺東縣消防局單位)
"""

from io import BytesIO
from typing import List, Dict, Any
from PIL import Image, ImageDraw, ImageFont

def create_mock_document_image(title: str, subtitle: str, fields: List[tuple], badge: str = "政府公務用印") -> Image.Image:
    """使用 PIL 繪製擬真公務申請文件/成績單圖片"""
    width, height = 750, 950
    img = Image.new("RGB", (width, height), color=(252, 252, 250))
    draw = ImageDraw.Draw(img)
    
    # 邊框與裝飾線
    draw.rectangle([(20, 20), (width - 20, height - 20)], outline=(180, 190, 205), width=2)
    draw.rectangle([(25, 25), (width - 25, height - 25)], outline=(200, 210, 225), width=1)
    
    # 頂部公文橫幅
    draw.rectangle([(30, 30), (width - 30, 110)], fill=(240, 244, 250), outline=(210, 220, 235))
    
    # 標題文字
    draw.text((width // 2 - 180, 45), title, fill=(30, 60, 110))
    draw.text((width // 2 - 140, 80), subtitle, fill=(100, 110, 125))
    
    # 表格區域
    start_y = 135
    row_height = 42
    
    draw.rectangle([(40, start_y), (width - 40, start_y + len(fields) * row_height)], fill=(255, 255, 255), outline=(160, 175, 195), width=1)
    
    for i, (label, val) in enumerate(fields):
        y = start_y + i * row_height
        draw.rectangle([(40, y), (200, y + row_height)], fill=(245, 248, 252), outline=(190, 205, 225), width=1)
        draw.rectangle([(200, y), (width - 40, y + row_height)], outline=(190, 205, 225), width=1)
        
        draw.text((55, y + 12), label, fill=(50, 65, 85))
        draw.text((215, y + 12), str(val), fill=(20, 30, 45))
        
    # 底部防偽戳章與印鑑模擬
    stamp_y = height - 160
    draw.ellipse([(width - 180, stamp_y), (width - 60, stamp_y + 110)], outline=(210, 45, 45), width=3)
    draw.text((width - 160, stamp_y + 35), "審核查驗章", fill=(210, 45, 45))
    draw.text((width - 165, stamp_y + 60), "113年度核定", fill=(210, 45, 45))
    
    # 附件檢核備註區
    draw.rectangle([(40, stamp_y), (width - 200, stamp_y + 110)], fill=(248, 249, 250), outline=(210, 220, 230))
    draw.text((50, stamp_y + 10), "【審查備註說明】", fill=(80, 90, 100))
    draw.text((50, stamp_y + 35), "本件為標準獎學金審查存查聯，各項證明文件已電子封存。", fill=(110, 120, 130))
    draw.text((50, stamp_y + 60), "請審查人員確認身分證字號與前學期總平均無誤。", fill=(110, 120, 130))
    draw.text((50, stamp_y + 85), f"文件編號：TTFD-SCHOLAR-{badge}-2026", fill=(130, 140, 150))
    
    return img

def get_mock_cases() -> List[Dict[str, Any]]:
    """產生 4 組包含臺東縣消防局所屬單位與圖片的示範案例"""
    
    # 案例 1: 臺東大隊 / 特種搜救分隊
    img1_form = create_mock_document_image(
        "消防及義消人員子女獎學金 申請表",
        "臺東縣消防局 113學年度第一學期",
        [
            ("申請人家長", "陳大明"),
            ("家長身分證", "V123456789"),
            ("服務單位", "臺東縣消防局 臺東大隊 特種搜救分隊"),
            ("職稱身分", "消防隊員 (在職)"),
            ("子女姓名", "陳小美"),
            ("申請組別", "大專院校"),
            ("就讀學校", "國立臺灣大學 資訊工程學系 三年級"),
            ("連絡電話", "0912-345-678")
        ],
        badge="CASE01-APP"
    )
    img1_grade = create_mock_document_image(
        "國立臺灣大學 學生學期成績證明單",
        "112學年度第二學期 官方核發正本",
        [
            ("學生姓名", "陳小美"),
            ("學號", "B10902001"),
            ("就讀系所", "資訊工程學系"),
            ("學期智育/總平均", "88.50 分"),
            ("操行成績", "88 分 (甲等)"),
            ("班級排名", "05 / 58"),
            ("系所蓋章", "國立臺灣大學教務處 註冊組戳章")
        ],
        badge="CASE01-GRD"
    )
    
    # 案例 2: 關山大隊 / 關山分隊
    img2_form = create_mock_document_image(
        "消防及義消人員子女獎學金 申請表",
        "臺東縣消防局 113學年度第一學期",
        [
            ("申請人家長", "林志偉"),
            ("家長身分證", "V198765432"),
            ("服務單位", "臺東縣消防局 關山大隊 關山分隊"),
            ("職稱身分", "義消隊員"),
            ("子女姓名", "林冠宇"),
            ("申請組別", "高中職"),
            ("就讀學校", "國立臺東高級中學 二年級"),
            ("連絡電話", "0922-888-999")
        ],
        badge="CASE02-APP"
    )
    img2_grade = create_mock_document_image(
        "國立臺東高級中學 成績通知單",
        "112學年度第二學期 成績證明",
        [
            ("學生姓名", "林冠宇"),
            ("學號", "1110325"),
            ("學期學業總平均", "84.00 分"),
            ("德行評量(操行)", "86 分 (良好)"),
            ("獎懲紀錄", "嘉獎三次，無曠課紀錄"),
            ("教務處核章", "臺東高中教務處 註冊組")
        ],
        badge="CASE02-GRD"
    )
    
    # 案例 3: 成功大隊 / 成功分隊
    img3_form = create_mock_document_image(
        "消防及義消人員子女獎學金 申請表",
        "臺東縣消防局 113學年度第一學期",
        [
            ("申請人家長", "張建國"),
            ("家長身分證", "V135792468"),
            ("服務單位", "臺東縣消防局 成功大隊 成功分隊"),
            ("職稱身分", "消防小隊長"),
            ("子女姓名", "張宇軒"),
            ("申請組別", "國中"),
            ("就讀學校", "臺東縣立新港國民中學 八年級"),
            ("連絡電話", "0933-111-222")
        ],
        badge="CASE03-APP"
    )
    img3_grade = create_mock_document_image(
        "臺東縣立新港國民中學 學期成績單",
        "112學年度第二學期 定期評量總表",
        [
            ("學生姓名", "張宇軒"),
            ("就讀班級", "八年三班 15號"),
            ("學期領域總平均", "76.50 分"),
            ("日常生活表現(操行)", "80 分 (乙等)"),
            ("導師評語", "學習態度認真，宜再加強數理理解"),
            ("教務處核章", "新港國中教務處戳印")
        ],
        badge="CASE03-GRD"
    )
    
    # 案例 4: 大武大隊 / 大武分隊
    img4_form = create_mock_document_image(
        "消防及義消人員子女獎學金 申請表",
        "臺東縣消防局 113學年度第一學期",
        [
            ("申請人家長", "黃敏華"),
            ("家長身分證", "V246813579"),
            ("服務單位", "臺東縣消防局 大武大隊 大武分隊"),
            ("職稱身分", "義消分隊長"),
            ("子女姓名", "黃子芸"),
            ("申請組別", "國小"),
            ("就讀學校", "臺東縣大武鄉大武國民小學 五年級"),
            ("連絡電話", "0955-666-777")
        ],
        badge="CASE04-APP"
    )
    img4_grade = create_mock_document_image(
        "臺東縣大武國民小學 學習評量成績證明",
        "112學年度第二學期 學生定期評量表",
        [
            ("學生姓名", "黃子芸"),
            ("就讀班級", "五年一班 08號"),
            ("學期總平均成績", "92.00 分"),
            ("日常生活表現(操行)", "優等 (95分)"),
            ("出缺席情況", "全勤"),
            ("教務處核章", "大武國小教務處 註冊章")
        ],
        badge="CASE04-GRD"
    )
    
    cases = [
        {
            "id": "CASE-001",
            "unit_level1": "臺東大隊",
            "unit_level2": "特種搜救分隊",
            "applicant_name": "陳大明",
            "applicant_id": "V123456789",
            "child_name": "陳小美",
            "category": "大專院校",
            "semester_gpa": 88.50,
            "conduct": "88分 (甲等)",
            "attachments": {
                "application_form": True,
                "student_id_or_enrollment": True,
                "transcript": True,
                "household_registration": True,
                "service_certificate": True
            },
            "review_status": "符合資格",
            "review_reason": "學期平均 88.50 (>=80.0)，5 項必備附件齊全",
            "is_eligible": True,
            "notes": "特搜分隊同仁子女，台大資工系成績優秀，證件齊全",
            "images": [img1_form, img1_grade],
            "image_labels": ["獎學金申請表", "前學期成績證明單"]
        },
        {
            "id": "CASE-002",
            "unit_level1": "關山大隊",
            "unit_level2": "關山分隊",
            "applicant_name": "林志偉",
            "applicant_id": "V198765432",
            "child_name": "林冠宇",
            "category": "高中職",
            "semester_gpa": 84.00,
            "conduct": "86分",
            "attachments": {
                "application_form": True,
                "student_id_or_enrollment": True,
                "transcript": True,
                "household_registration": True,
                "service_certificate": False
            },
            "review_status": "待補件",
            "review_reason": "成績達標 (平均 84.00)，但缺件: 5. 消防/義消在職或服務證明文件",
            "is_eligible": False,
            "notes": "未見關山分隊義消在職服務證明，已通知申請人於一週內補正",
            "images": [img2_form, img2_grade],
            "image_labels": ["獎學金申請表", "高中學期成績單"]
        },
        {
            "id": "CASE-003",
            "unit_level1": "成功大隊",
            "unit_level2": "成功分隊",
            "applicant_name": "張建國",
            "applicant_id": "V135792468",
            "child_name": "張宇軒",
            "category": "國中",
            "semester_gpa": 76.50,
            "conduct": "80分 (乙等)",
            "attachments": {
                "application_form": True,
                "student_id_or_enrollment": True,
                "transcript": True,
                "household_registration": True,
                "service_certificate": True
            },
            "review_status": "不符資格",
            "review_reason": "學期平均 76.50 未達 80.0 分申請門檻",
            "is_eligible": False,
            "notes": "附件均備齊，惟學期總平均 76.50 分未達規定 80 分",
            "images": [img3_form, img3_grade],
            "image_labels": ["獎學金申請表", "國中學期成績單"]
        },
        {
            "id": "CASE-004",
            "unit_level1": "大武大隊",
            "unit_level2": "大武分隊",
            "applicant_name": "黃敏華",
            "applicant_id": "V246813579",
            "child_name": "黃子芸",
            "category": "國小",
            "semester_gpa": 92.00,
            "conduct": "優等",
            "attachments": {
                "application_form": True,
                "student_id_or_enrollment": True,
                "transcript": True,
                "household_registration": True,
                "service_certificate": True
            },
            "review_status": "符合資格",
            "review_reason": "學期平均 92.00 (>=80.0)，5 項必備附件齊全",
            "is_eligible": True,
            "notes": "大武義消分隊長子女，國小成績優異",
            "images": [img4_form, img4_grade],
            "image_labels": ["獎學金申請表", "國小定期評量成績單"]
        }
    ]
    return cases