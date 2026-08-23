"""
臺東縣消防局 消防及義消子女獎學金 AI 智慧審核網頁系統
前後台分流權限設計 ‧ 磁碟自動儲存 (./uploads/ & ./data/獎學金總表.xlsx) ‧ 雲端試算表自動同步
"""

import os
import io
import time
import socket
from datetime import datetime
import streamlit as st
import pandas as pd
from PIL import Image
import pypdfium2 as pdfium
import qrcode
import importlib

# 匯入自訂模組
import excel_exporter
importlib.reload(excel_exporter)
from excel_exporter import export_scholarship_excel

import storage_manager
importlib.reload(storage_manager)
from storage_manager import (
    save_case_to_storage,
    load_stored_cases,
    package_uploads_zip,
    EXCEL_FILE,
    UPLOADS_DIR,
    DATA_DIR
)

from org_structure import TAITUNG_FIRE_ORG, get_level1_units, get_level2_units
from gemini_analyzer import (
    analyze_scholarship_documents,
    evaluate_eligibility,
    ATTACHMENT_NAMES
)
from mock_data import get_mock_cases
from cloud_sync import sync_to_google_sheets, GOOGLE_APPS_SCRIPT_TEMPLATE

# 頁面配置
st.set_page_config(
    page_title="臺東縣消防局 消防及義消子女獎學金 智慧申請審核系統",
    page_icon="🚒",
    layout="wide",
    initial_sidebar_state="auto"
)

# 自訂樣式
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 22px 28px;
        border-radius: 12px;
        margin-bottom: 24px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    }
    .main-header h1 {
        color: white;
        font-size: 25px;
        font-weight: 700;
        margin: 0 0 6px 0;
    }
    .main-header p {
        color: #e0e8f5;
        font-size: 14px;
        margin: 0;
    }
    .receipt-card {
        background: #ffffff;
        border: 2px solid #2563eb;
        border-radius: 12px;
        padding: 24px 28px;
        margin: 20px 0;
        box-shadow: 0 4px 15px rgba(37, 99, 235, 0.08);
    }
    .receipt-title {
        font-size: 20px;
        font-weight: 700;
        color: #1e3c72;
        border-bottom: 2px dashed #cbd5e1;
        padding-bottom: 12px;
        margin-bottom: 16px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .stat-card {
        background: white;
        border-radius: 10px;
        padding: 16px 20px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        text-align: center;
    }
    .stat-val {
        font-size: 26px;
        font-weight: 700;
        margin-top: 4px;
    }
    .stat-label {
        font-size: 13px;
        color: #64748b;
        font-weight: 600;
    }
    .badge-eligible {
        background-color: #d1fae5;
        color: #065f46;
        padding: 6px 14px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 15px;
        display: inline-block;
    }
    .badge-pending {
        background-color: #fef3c7;
        color: #92400e;
        padding: 6px 14px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 15px;
        display: inline-block;
    }
    .badge-ineligible {
        background-color: #fee2e2;
        color: #991b1b;
        padding: 6px 14px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 15px;
        display: inline-block;
    }
    .section-title {
        font-size: 18px;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .unit-tag {
        background-color: #e0f2fe;
        color: #0369a1;
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 12px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "10.41.51.44"

def process_uploaded_file(file) -> list[tuple[Image.Image, str]]:
    results = []
    filename = file.name
    file_ext = filename.split(".")[-1].lower()
    
    if file_ext == "pdf":
        file_bytes = file.read()
        pdf = pdfium.PdfDocument(file_bytes)
        total_pages = len(pdf)
        for page_idx in range(total_pages):
            page = pdf[page_idx]
            pix = page.render(scale=2.0)
            pil_img = pix.to_pil()
            page_label = f"{filename} (第 {page_idx + 1}/{total_pages} 頁)"
            results.append((pil_img, page_label))
    else:
        img = Image.open(file)
        results.append((img, filename))
        
    return results

SECRETS_FILE = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
DEFAULT_ADMIN_PASSWORD = "ttfd888"

def load_persistent_api_key() -> str:
    if os.path.exists(SECRETS_FILE):
        try:
            with open(SECRETS_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    if line.startswith("GEMINI_API_KEY") and "=" in line:
                        val = line.split("=", 1)[1].strip().strip('"').strip("'")
                        if val:
                            return val
        except Exception:
            pass
    return os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or ""

def save_persistent_api_key(key: str):
    os.makedirs(os.path.dirname(SECRETS_FILE), exist_ok=True)
    with open(SECRETS_FILE, "w", encoding="utf-8") as f:
        f.write('GEMINI_API_KEY = "' + key.strip() + '"\n')
    os.environ["GEMINI_API_KEY"] = key.strip()

# 檢查網址參數是否帶有管理員/審核專用直通金鑰 (如 ?admin=ttfd888)
try:
    if "admin" in st.query_params:
        if st.query_params["admin"] == DEFAULT_ADMIN_PASSWORD:
            st.session_state.is_admin = True
except Exception:
    pass

# 初始化 Session State (優先從 ./data/records.json 載入歷史儲存紀錄)
if "records" not in st.session_state:
    stored = load_stored_cases()
    st.session_state.records = stored if stored else get_mock_cases()

if "selected_case_id" not in st.session_state:
    st.session_state.selected_case_id = st.session_state.records[0]["id"] if st.session_state.records else None

if "api_key" not in st.session_state:
    st.session_state.api_key = load_persistent_api_key()

if "camera_photos" not in st.session_state:
    st.session_state.camera_photos = []

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

if "last_submitted_case" not in st.session_state:
    st.session_state.last_submitted_case = None

# ----------------- 側邊欄 -----------------
with st.sidebar:
    st.markdown("## 🚒 臺東縣消防局\n### 獎學金申請審核系統")
    
    if not st.session_state.is_admin:
        st.info("📍 **當前身分：申請同仁專區**\n\n(已啟用個資防護，僅可交件)")
        
        with st.expander("🔐 各大隊及業務科 審核人員登入", expanded=False):
            st.markdown("各大隊、分隊承辦人及業務科同仁請輸入通關密碼以檢視完整清冊與複核：")
            pwd_input = st.text_input("審核人員密碼", type="password", key="admin_pwd_input")
            if st.button("🔑 登入審核管理後台", use_container_width=True, type="primary"):
                if pwd_input == DEFAULT_ADMIN_PASSWORD:
                    st.session_state.is_admin = True
                    st.success("✅ 驗證通過，已切換至審核人員模式！")
                    st.rerun()
                else:
                    st.error("❌ 密碼錯誤！請洽業務科獎學金承辦人。")
    else:
        st.success("👑 **當前身分：各大隊及業務科 審核管理員**")
        if st.button("🚪 登出審核後台（切換回同仁申請模式）", use_container_width=True):
            st.session_state.is_admin = False
            try:
                if "admin" in st.query_params:
                    del st.query_params["admin"]
            except Exception:
                pass
            st.rerun()
            
        with st.expander("🔗 審核後台專屬直通網址", expanded=False):
            admin_direct_url = f"https://means-brandon-tip-snap.trycloudflare.com/?admin={DEFAULT_ADMIN_PASSWORD}"
            st.text_input("專屬直通網址 (可加入書籤)", value=admin_direct_url, help="此網址可直接進入審核後台，免手動輸入密碼")
            st.caption("提示：各大隊與業務科審核承辦人可將此網址加入書籤。")
            
        st.markdown("---")
        st.subheader("🔑 Google GenAI 設定")
        
        has_global_key = bool(load_persistent_api_key())
        if has_global_key:
            st.success("✅ 伺服器全域 API Key 已就緒\n\n(全體同仁免輸入)")
            with st.expander("🔧 更換 / 管理 API Key", expanded=False):
                api_key_input = st.text_input(
                    "更新 API Key",
                    value=st.session_state.api_key,
                    type="password",
                    help="輸入新的 Google Gemini API Key"
                )
                if st.button("💾 儲存並套用新金鑰", use_container_width=True):
                    save_persistent_api_key(api_key_input)
                    st.session_state.api_key = api_key_input
                    st.success("已更新全域金鑰！")
                    st.rerun()
        else:
            api_key_input = st.text_input(
                "Gemini API Key",
                value=st.session_state.api_key,
                type="password",
                help="輸入您的 Google Gemini API Key"
            )
            if api_key_input != st.session_state.api_key:
                st.session_state.api_key = api_key_input
                
            if st.button("💾 儲存為全域金鑰 (全體免再輸入)", use_container_width=True):
                if api_key_input:
                    save_persistent_api_key(api_key_input)
                    st.session_state.api_key = api_key_input
                    st.success("已儲存為全域金鑰！全體同仁的手機與電腦均免再輸入。")
                    st.rerun()
                else:
                    st.error("請先輸入有效的 API Key！")
                    
        model_choice = st.selectbox(
            "AI 視覺辨識模型",
            ["gemini-3.6-flash (最新推薦)", "gemini-3.5-flash", "gemini-3.6-pro", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"],
            index=0
        )
        active_model_name = model_choice.split(" ")[0]
        
        st.markdown("---")
        st.subheader("⚡ 管理快捷操作")
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            if st.button("🔄 載入示範資料", use_container_width=True):
                st.session_state.records = get_mock_cases()
                if st.session_state.records:
                    st.session_state.selected_case_id = st.session_state.records[0]["id"]
                st.success("已載入臺東縣消防局 4 筆擬真示範案例！")
                st.rerun()
                
        with col_s2:
            if st.button("🗑️ 清空所有資料", use_container_width=True):
                st.session_state.records = []
                st.session_state.selected_case_id = None
                if os.path.exists(os.path.join(DATA_DIR, "records.json")):
                    try:
                        os.remove(os.path.join(DATA_DIR, "records.json"))
                    except Exception:
                        pass
                st.warning("已清空所有資料！")
                st.rerun()

    st.markdown("---")
    st.subheader("📌 審查標準門檻")
    st.markdown(r"""
    - **學業成績**：學期總平均 $\ge 80.0$ 分
    - **必備附件 (5項)**：
      1. 獎學金申請表
      2. 學生證或在學證明
      3. 前學期成績證明單
      4. 戶口名簿影本 / 戶籍謄本
      5. 消防 / 義消服務證明
    """)

# =========================================================================
# 模式 A：同仁線上申請交件專區 (Applicant Mode - 個資保護，只能交件與看自己收據)
# =========================================================================
if not st.session_state.is_admin:
    st.markdown("""
    <div class="main-header">
        <h1>🚒 臺東縣消防局 消防及義消子女獎學金 線上交件申請系統</h1>
        <p>歡迎同仁申請！請先選擇所屬大隊/分隊，拍照或上傳申請文件，AI 將立即協助檢核 5 項附件齊全度與學期成績。</p>
    </div>
    """, unsafe_allow_html=True)
    
    # 若剛剛有送出案件，優先在頂部顯示收執聯回饋卡片
    if st.session_state.last_submitted_case:
        rec = st.session_state.last_submitted_case
        status = rec.get("review_status", "")
        
        st.markdown(f"""
        <div class="receipt-card">
            <div class="receipt-title">
                <span>📋 獎學金交件確認收執聯</span>
                <span style="font-size:16px;color:#2563eb;">案件編號：<b>{rec.get('id')}</b></span>
            </div>
        """, unsafe_allow_html=True)
        
        col_r1, col_r2, col_r3 = st.columns(3)
        with col_r1:
            st.write(f"🏢 **所屬單位**：{rec.get('unit_level1')} / {rec.get('unit_level2')}")
            st.write(f"👤 **申請人家長**：{rec.get('applicant_name') or '（依證明文件查驗）'}")
        with col_r2:
            st.write(f"🎓 **子女姓名**：{rec.get('child_name') or '（依成績單查驗）'}")
            st.write(f"📚 **申請組別**：{rec.get('category')}")
        with col_r3:
            gpa_disp = f"{rec.get('semester_gpa'):.2f} 分" if rec.get('semester_gpa') is not None else "未識別"
            st.write(f"📊 **學期總平均**：{gpa_disp}")
            st.write(f"⭐ **操行成績**：{rec.get('conduct') or '正常'}")
            
        st.markdown("##### 📋 5 項必備附件檢核狀況：")
        att = rec.get("attachments", {})
        att_items = [
            ("application_form", "1. 獎學金申請表"),
            ("student_id_or_enrollment", "2. 學生證或在學證明"),
            ("transcript", "3. 前學期成績證明單"),
            ("household_registration", "4. 戶口名簿影本或戶籍謄本"),
            ("service_certificate", "5. 消防/義消在職或服務證明")
        ]
        
        att_c1, att_c2 = st.columns(2)
        for i, (k, label) in enumerate(att_items):
            target_col = att_c1 if i < 3 else att_c2
            with target_col:
                if att.get(k, False):
                    st.markdown(f"✅ **{label}**：已備妥")
                else:
                    st.markdown(f"❌ <span style='color:#dc2626;font-weight:bold;'>{label}：缺漏（需補件）</span>", unsafe_allow_html=True)
                    
        st.markdown("---")
        if status == "符合資格":
            st.markdown(f'<div class="badge-eligible">🟢 初步審查：符合資格</div> &nbsp; <b>{rec.get("review_reason")}</b>', unsafe_allow_html=True)
            st.success("🎉 您上傳的文件齊全且成績達標！資料已自動永久存檔並送出至大隊與業務科，請靜候後續核定通知。")
        elif status == "待補件":
            st.markdown(f'<div class="badge-pending">🟡 初步審查：需補件</div> &nbsp; <b>{rec.get("review_reason")}</b>', unsafe_allow_html=True)
            st.warning(f"⚠️ 您的申請資料已收件存檔，但請注意：**{rec.get('review_reason')}**。請於收件截止日前補正檔案以利核發。")
        else:
            st.markdown(f'<div class="badge-ineligible">🔴 初步審查：未達門檻</div> &nbsp; <b>{rec.get("review_reason")}</b>', unsafe_allow_html=True)
            st.info("ℹ️ 您的資料已成功送出並自動存檔，承辦人員將於複核階段再次人工核對。")
            
        st.markdown("</div>", unsafe_allow_html=True)
        
        if st.button("✨ 繼續提交下一筆申請案件", use_container_width=True):
            st.session_state.last_submitted_case = None
            st.rerun()
            
        st.markdown("---")

    # 申請交件表單區
    st.markdown('<div class="section-title">📤 線上申請交件與照片上傳</div>', unsafe_allow_html=True)
    
    # 兩階層組織單位選擇區
    st.markdown("##### 🏢 第一步：請選擇您的所屬單位")
    col_sel_l1, col_sel_l2 = st.columns(2)
    with col_sel_l1:
        upload_level1 = st.selectbox("1. 大隊 / 局本部：", get_level1_units(), key="pub_upload_unit_l1")
    with col_sel_l2:
        level2_choices = get_level2_units(upload_level1)
        upload_level2 = st.selectbox("2. 分隊 / 科室：", level2_choices, key="pub_upload_unit_l2")
        
    st.caption(f"📍 您選擇的申請單位：**臺東縣消防局 {upload_level1} ➔ {upload_level2}**")
    st.markdown("---")
    
    st.markdown("##### 📁 第二步：選擇上傳照片方式 (可一次傳多張)")
    input_mode = st.radio("請選擇上傳方式：", ["📁 檔案 / PDF / 圖片上傳", "📷 即時相機拍照 (手機專用)"], horizontal=True, key="pub_input_mode")
    
    all_prepared_images = []
    all_prepared_labels = []
    
    if input_mode == "📁 檔案 / PDF / 圖片上傳":
        col_u1, col_u2 = st.columns([1.2, 0.8])
        with col_u1:
            uploaded_files = st.file_uploader(
                "請選取申請表、成績單與證明文件 (支援 PDF, JPG, PNG, WEBP)：",
                type=["pdf", "jpg", "jpeg", "png", "webp"],
                accept_multiple_files=True,
                key="pub_scholarship_uploader"
            )
            
        with col_u2:
            st.markdown("##### 📋 上傳預覽")
            if uploaded_files:
                for f in uploaded_files:
                    extracted = process_uploaded_file(f)
                    for p_img, p_label in extracted:
                        all_prepared_images.append(p_img)
                        all_prepared_labels.append(p_label)
                        
                st.success(f"已載入 {len(all_prepared_images)} 個頁面/圖檔")
                preview_cols = st.columns(min(len(all_prepared_images), 3))
                for i, p_img in enumerate(all_prepared_images[:3]):
                    with preview_cols[i]:
                        st.image(p_img, caption=all_prepared_labels[i][:20], use_container_width=True)
            else:
                st.write("尚未選取檔案。")
                
    else:  # 即時相機拍照模式
        col_c1, col_c2 = st.columns([1.2, 0.8])
        with col_c1:
            camera_file = st.camera_input("📷 請將鏡頭對準申請文件拍照：")
            if camera_file is not None:
                cam_img = Image.open(camera_file)
                if st.button("➕ 將此照片加入交件清單", use_container_width=True, key="pub_cam_add"):
                    st.session_state.camera_photos.append(cam_img)
                    st.success(f"已加入第 {len(st.session_state.camera_photos)} 張照片！可繼續拍下一頁。")
                    
        with col_c2:
            st.markdown("##### 📸 已拍照片")
            if st.session_state.camera_photos:
                st.info(f"已拍 {len(st.session_state.camera_photos)} 張照片")
                cam_cols = st.columns(min(len(st.session_state.camera_photos), 3))
                for i, p_img in enumerate(st.session_state.camera_photos):
                    with cam_cols[i % 3]:
                        st.image(p_img, caption=f"照片 {i+1}", use_container_width=True)
                        
                if st.button("🗑️ 清空重拍", use_container_width=True, key="pub_cam_clear"):
                    st.session_state.camera_photos = []
                    st.rerun()
                    
                for i, p_img in enumerate(st.session_state.camera_photos):
                    all_prepared_images.append(p_img)
                    all_prepared_labels.append(f"相機拍攝照片 {i+1}")
            else:
                st.write("尚未拍攝照片。")

    # 送出交件與 AI 審核按鈕
    if all_prepared_images:
        st.markdown("---")
        start_ai_btn = st.button(
            f"🚀 確認交件並開始 AI 智慧審查（{upload_level1}/{upload_level2}）",
            use_container_width=True,
            type="primary",
            key="pub_start_ai_btn"
        )
        
        if start_ai_btn:
            api_key_to_use = st.session_state.api_key or load_persistent_api_key()
            if not api_key_to_use:
                st.error("❌ 系統尚未設定 API Key，請通知業務科管理員於後台儲存金鑰！")
            else:
                progress_bar = st.progress(0, text=f"正在整理 {len(all_prepared_images)} 個影像檔案...")
                try:
                    progress_bar.progress(35, text="正在進行 AI 智慧多模態辨識與資料擷取...")
                    
                    ai_result = analyze_scholarship_documents(
                        images=all_prepared_images,
                        api_key=api_key_to_use,
                        model_name="gemini-3.6-flash"
                    )
                    
                    progress_bar.progress(75, text="正在比對審查標準與 5 項必備附件...")
                    
                    new_case_id = f"TTFD-2026-{len(st.session_state.records)+1:03d}"
                    
                    new_case = {
                        "id": new_case_id,
                        "unit_level1": upload_level1,
                        "unit_level2": upload_level2,
                        "applicant_name": ai_result.get("applicant_name", ""),
                        "applicant_id": ai_result.get("applicant_id", ""),
                        "child_name": ai_result.get("child_name", ""),
                        "category": ai_result.get("category", "大專院校"),
                        "semester_gpa": ai_result.get("semester_gpa"),
                        "conduct": ai_result.get("conduct", ""),
                        "attachments": ai_result.get("attachments", {}),
                        "review_status": ai_result.get("review_status", "待審核"),
                        "review_reason": ai_result.get("review_reason", ""),
                        "is_eligible": ai_result.get("is_eligible", False),
                        "notes": ai_result.get("notes", ""),
                        "images": all_prepared_images,
                        "image_labels": all_prepared_labels,
                        "submitted_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    
                    # 💾 自動存入 ./uploads/ 與追加寫入 ./data/獎學金總表.xlsx
                    progress_bar.progress(90, text="正在儲存原始照片至 ./uploads/ 並追加至總表 Excel...")
                    save_case_to_storage(new_case)
                    
                    st.session_state.records.append(new_case)
                    st.session_state.selected_case_id = new_case["id"]
                    st.session_state.last_submitted_case = new_case
                    st.session_state.camera_photos = []
                    
                    # 若有設定 Google Sheets Webhook，同時自動同步雲端試算表
                    webhook_url = st.session_state.get("google_sheet_webhook", "")
                    if webhook_url:
                        try:
                            sync_to_google_sheets(webhook_url, st.session_state.records)
                        except Exception:
                            pass
                            
                    progress_bar.progress(100, text="✅ 交件成功並已永久存檔！")
                    time.sleep(0.5)
                    st.balloons()
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ 處理發生錯誤: {str(e)}")

# =========================================================================
# 模式 B：各大隊及業務科 審核管理後台 (Reviewer / Admin Mode - 完整功能)
# =========================================================================
else:
    st.markdown("""
    <div class="main-header">
        <h1>🚒 臺東縣消防局 獎學金審核管理系統【各大隊及業務科專用】</h1>
        <p>管理員後台 ‧ 自動儲存 (./uploads/ & 獎學金總表.xlsx) ‧ 左圖右表人工複核 ‧ Google 雲端試算表即時同步</p>
    </div>
    """, unsafe_allow_html=True)
    
    records = st.session_state.records
    total_count = len(records)
    eligible_count = sum(1 for r in records if r.get("review_status") == "符合資格")
    pending_count = sum(1 for r in records if r.get("review_status") == "待補件")
    ineligible_count = sum(1 for r in records if r.get("review_status") == "不符資格")
    pass_rate = f"{(eligible_count / total_count * 100):.1f}%" if total_count > 0 else "0.0%"

    col_kpi1, col_kpi2, col_kpi3, col_kpi4, col_kpi5 = st.columns(5)
    with col_kpi1:
        st.markdown(f'<div class="stat-card"><div class="stat-label">總申請件數</div><div class="stat-val" style="color:#1e3c72;">{total_count} <span style="font-size:14px;font-weight:normal;">件</span></div></div>', unsafe_allow_html=True)
    with col_kpi2:
        st.markdown(f'<div class="stat-card"><div class="stat-label">符合資格</div><div class="stat-val" style="color:#059669;">{eligible_count} <span style="font-size:14px;font-weight:normal;">人</span></div></div>', unsafe_allow_html=True)
    with col_kpi3:
        st.markdown(f'<div class="stat-card"><div class="stat-label">待補件</div><div class="stat-val" style="color:#d97706;">{pending_count} <span style="font-size:14px;font-weight:normal;">人</span></div></div>', unsafe_allow_html=True)
    with col_kpi4:
        st.markdown(f'<div class="stat-card"><div class="stat-label">不符資格</div><div class="stat-val" style="color:#dc2626;">{ineligible_count} <span style="font-size:14px;font-weight:normal;">人</span></div></div>', unsafe_allow_html=True)
    with col_kpi5:
        st.markdown(f'<div class="stat-card"><div class="stat-label">核定通過率</div><div class="stat-val" style="color:#2563eb;">{pass_rate}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs([
        "🔍 「左圖右表」案件複核工作台",
        "📤 承辦人代為上傳新案件",
        "📊 全局審核總表清冊與 Excel / 雲端匯出"
    ])

    # ----------------- 後台 TAB 1: 左圖右表複核 -----------------
    with tab1:
        if not st.session_state.records:
            st.info("目前尚無任何案件資料。")
        else:
            case_options = {
                r["id"]: f"【{r.get('id')}】[{r.get('unit_level1', '未定')} / {r.get('unit_level2', '未定')}] {r.get('applicant_name', '未命名')} / 子女: {r.get('child_name', '未命名')} ({r.get('category', '未定')}) - {r.get('review_status', '待審')}"
                for r in st.session_state.records
            }
            
            if st.session_state.selected_case_id not in case_options:
                st.session_state.selected_case_id = list(case_options.keys())[0]
                
            selected_id = st.selectbox(
                "📁 請選擇要審核複核的申請案件：",
                options=list(case_options.keys()),
                format_func=lambda x: case_options[x],
                index=list(case_options.keys()).index(st.session_state.selected_case_id),
                key="admin_case_selector"
            )
            st.session_state.selected_case_id = selected_id
            
            curr_case_idx = next(i for i, r in enumerate(st.session_state.records) if r["id"] == selected_id)
            curr_case = st.session_state.records[curr_case_idx]
            
            st.markdown("---")
            col_left, col_right = st.columns([1.1, 0.9], gap="large")
            
            # 左側圖片區
            with col_left:
                st.markdown(f'<div class="section-title">🖼️ 原始申請資料與成績單檢視 &nbsp; <span class="unit-tag">{curr_case.get("unit_level1", "")} / {curr_case.get("unit_level2", "")}</span></div>', unsafe_allow_html=True)
                images = curr_case.get("images", [])
                image_labels = curr_case.get("image_labels", [])
                
                if not images:
                    st.warning("⚠️ 此案件無附加圖片檔案。")
                else:
                    if len(images) > 1:
                        labels = [image_labels[i] if i < len(image_labels) else f"照片 {i+1}" for i in range(len(images))]
                        selected_img_idx = st.radio(
                            "選擇檢視頁面：",
                            range(len(images)),
                            format_func=lambda i: f"📄 {labels[i]}",
                            horizontal=True,
                            key="admin_img_select"
                        )
                    else:
                        selected_img_idx = 0
                        
                    target_img = images[selected_img_idx]
                    st.image(target_img, use_container_width=True, caption=f"原檔 - {image_labels[selected_img_idx] if selected_img_idx < len(image_labels) else f'照片 {selected_img_idx+1}'}")
                    
            # 右側表單區
            with col_right:
                st.markdown('<div class="section-title">✍️ 承辦人審核複核與狀態判定</div>', unsafe_allow_html=True)
                
                with st.form(key=f"admin_review_form_{curr_case['id']}"):
                    st.markdown("##### 🏢 所屬單位")
                    col_u1, col_u2 = st.columns(2)
                    
                    level1_list = get_level1_units()
                    curr_l1 = curr_case.get("unit_level1", level1_list[0])
                    idx_l1 = level1_list.index(curr_l1) if curr_l1 in level1_list else 0
                    
                    with col_u1:
                        edit_l1 = st.selectbox("大隊 / 局本部", level1_list, index=idx_l1, key="admin_edit_l1")
                    
                    level2_list = get_level2_units(edit_l1)
                    curr_l2 = curr_case.get("unit_level2", level2_list[0] if level2_list else "")
                    idx_l2 = level2_list.index(curr_l2) if curr_l2 in level2_list else 0
                    
                    with col_u2:
                        edit_l2 = st.selectbox("分隊 / 科室", level2_list, index=idx_l2, key="admin_edit_l2")
                    
                    st.markdown("---")
                    st.markdown("##### 👤 申請人與成績資訊")
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        applicant_name = st.text_input("申請人家長姓名", value=curr_case.get("applicant_name", ""))
                        applicant_id = st.text_input("身分證字號", value=curr_case.get("applicant_id", ""))
                        category = st.selectbox(
                            "申請組別",
                            ["大專院校", "高中職", "國中", "國小"],
                            index=["大專院校", "高中職", "國中", "國小"].index(curr_case.get("category", "大專院校")) if curr_case.get("category") in ["大專院校", "高中職", "國中", "國小"] else 0
                        )
                    with col_f2:
                        child_name = st.text_input("子女姓名", value=curr_case.get("child_name", ""))
                        current_gpa = curr_case.get("semester_gpa")
                        gpa_val = float(current_gpa) if current_gpa is not None else 0.0
                        semester_gpa = st.number_input(
                            "學期總平均 (小數點兩位)",
                            min_value=0.0,
                            max_value=100.0,
                            value=gpa_val,
                            step=0.01,
                            format="%.2f"
                        )
                        conduct = st.text_input("操行成績 (等第/分數)", value=curr_case.get("conduct", ""))
                        
                    st.markdown("##### 📋 5 項必備附件檢核清單")
                    att = curr_case.get("attachments", {})
                    
                    att_c1, att_c2 = st.columns(2)
                    with att_c1:
                        chk_app = st.checkbox("1. 獎學金申請表", value=att.get("application_form", True))
                        chk_student = st.checkbox("2. 學生證或在學證明", value=att.get("student_id_or_enrollment", True))
                        chk_trans = st.checkbox("3. 前學期成績證明單", value=att.get("transcript", True))
                    with att_c2:
                        chk_house = st.checkbox("4. 戶口名簿/戶籍謄本", value=att.get("household_registration", True))
                        chk_service = st.checkbox("5. 消防/義消服務證明", value=att.get("service_certificate", True))
                        
                    notes = st.text_area("審核備註說明", value=curr_case.get("notes", ""), height=80)
                    
                    new_attachments = {
                        "application_form": chk_app,
                        "student_id_or_enrollment": chk_student,
                        "transcript": chk_trans,
                        "household_registration": chk_house,
                        "service_certificate": chk_service
                    }
                    calc_status, calc_reason, is_elig = evaluate_eligibility(semester_gpa, new_attachments)
                    
                    st.markdown("---")
                    st.markdown("##### 🎯 即時資格試算結果")
                    if calc_status == "符合資格":
                        st.markdown(f'<div class="badge-eligible">🟢 符合資格</div> &nbsp; <b>{calc_reason}</b>', unsafe_allow_html=True)
                    elif calc_status == "待補件":
                        st.markdown(f'<div class="badge-pending">🟡 待補件</div> &nbsp; <b>{calc_reason}</b>', unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="badge-ineligible">🔴 不符資格</div> &nbsp; <b>{calc_reason}</b>', unsafe_allow_html=True)
                        
                    st.markdown("<br>", unsafe_allow_html=True)
                    submit_btn = st.form_submit_button("💾 儲存並更新審核結果 (同步存檔)", use_container_width=True)
                    
                    if submit_btn:
                        curr_case["unit_level1"] = edit_l1
                        curr_case["unit_level2"] = edit_l2
                        curr_case["applicant_name"] = applicant_name
                        curr_case["applicant_id"] = applicant_id
                        curr_case["child_name"] = child_name
                        curr_case["category"] = category
                        curr_case["semester_gpa"] = semester_gpa
                        curr_case["conduct"] = conduct
                        curr_case["attachments"] = new_attachments
                        curr_case["notes"] = notes
                        curr_case["review_status"] = calc_status
                        curr_case["review_reason"] = calc_reason
                        curr_case["is_eligible"] = is_elig
                        
                        st.session_state.records[curr_case_idx] = curr_case
                        save_case_to_storage(curr_case)
                        st.success(f"已成功儲存【{curr_case['id']}】的複核結果並更新總表！")
                        st.rerun()

    # ----------------- 後台 TAB 2: 承辦人代為上傳 -----------------
    with tab2:
        st.markdown('<div class="section-title">📤 承辦人代為上傳申請文件與 AI 解析</div>', unsafe_allow_html=True)
        col_sel_l1, col_sel_l2 = st.columns(2)
        with col_sel_l1:
            admin_upload_l1 = st.selectbox("1. 選擇大隊 / 局本部：", get_level1_units(), key="admin_up_l1")
        with col_sel_l2:
            admin_l2_choices = get_level2_units(admin_upload_l1)
            admin_upload_l2 = st.selectbox("2. 選擇分隊 / 科室：", admin_l2_choices, key="admin_up_l2")
            
        admin_uploaded_files = st.file_uploader(
            "選擇檔案 (PDF/圖片)：",
            type=["pdf", "jpg", "jpeg", "png", "webp"],
            accept_multiple_files=True,
            key="admin_file_uploader"
        )
        
        if admin_uploaded_files and st.button("🚀 啟動 AI 辨識並加入總表存檔", use_container_width=True, type="primary"):
            admin_imgs, admin_lbls = [], []
            for f in admin_uploaded_files:
                for p_img, p_label in process_uploaded_file(f):
                    admin_imgs.append(p_img)
                    admin_lbls.append(p_label)
                    
            with st.spinner("AI 解析並存檔中..."):
                ai_res = analyze_scholarship_documents(admin_imgs, st.session_state.api_key, "gemini-3.6-flash")
                case_id = f"TTFD-2026-{len(st.session_state.records)+1:03d}"
                new_c = {
                    "id": case_id,
                    "unit_level1": admin_upload_l1,
                    "unit_level2": admin_upload_l2,
                    "applicant_name": ai_res.get("applicant_name", ""),
                    "applicant_id": ai_res.get("applicant_id", ""),
                    "child_name": ai_res.get("child_name", ""),
                    "category": ai_res.get("category", "大專院校"),
                    "semester_gpa": ai_res.get("semester_gpa"),
                    "conduct": ai_res.get("conduct", ""),
                    "attachments": ai_res.get("attachments", {}),
                    "review_status": ai_res.get("review_status", "待審核"),
                    "review_reason": ai_res.get("review_reason", ""),
                    "is_eligible": ai_res.get("is_eligible", False),
                    "notes": ai_res.get("notes", ""),
                    "images": admin_imgs,
                    "image_labels": admin_lbls,
                    "submitted_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                save_case_to_storage(new_c)
                st.session_state.records.append(new_c)
                st.session_state.selected_case_id = new_c["id"]
                st.success(f"✅ 成功加入案件【{new_c['id']}】並存入 ./uploads/ 與總表！")
                st.rerun()

    # ----------------- 後台 TAB 3: 總表清冊與匯出 -----------------
    with tab3:
        st.markdown('<div class="section-title">📊 臺東縣消防局 獎學金審核彙總名冊 (完整個資)</div>', unsafe_allow_html=True)
        
        if not st.session_state.records:
            st.info("尚無審核紀錄。")
        else:
            table_rows = []
            for idx, r in enumerate(st.session_state.records, 1):
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
                
                gpa = r.get("semester_gpa")
                gpa_disp = f"{gpa:.2f}" if gpa is not None else "-"
                
                table_rows.append({
                    "序號": idx,
                    "案件編號": r.get("id", ""),
                    "大隊 / 局本部": r.get("unit_level1", "未指定"),
                    "分隊 / 科室": r.get("unit_level2", "未指定"),
                    "申請人姓名": r.get("applicant_name", ""),
                    "身分證字號": r.get("applicant_id", ""),
                    "子女姓名": r.get("child_name", ""),
                    "組別": r.get("category", ""),
                    "學期總平均": gpa_disp,
                    "操行": r.get("conduct", ""),
                    "附件檢核": att_desc,
                    "審核結果": r.get("review_status", ""),
                    "判定理由 / 備註": r.get("review_reason", r.get("notes", ""))
                })
                
            df = pd.DataFrame(table_rows)
            
            # 多維度篩選
            col_flt1, col_flt2, col_flt3 = st.columns(3)
            with col_flt1:
                all_l1 = ["全部"] + list(df["大隊 / 局本部"].unique())
                filter_l1 = st.selectbox("依大隊/局本部篩選：", all_l1, key="admin_flt_l1")
            with col_flt2:
                status_filter = st.multiselect("依審核結果篩選：", ["符合資格", "待補件", "不符資格"], default=["符合資格", "待補件", "不符資格"], key="admin_flt_stat")
            with col_flt3:
                cat_filter = st.multiselect("依組別篩選：", ["大專院校", "高中職", "國中", "國小"], default=["大專院校", "高中職", "國中", "國小"], key="admin_flt_cat")
                
            filtered_df = df[df["審核結果"].isin(status_filter) & df["組別"].isin(cat_filter)]
            if filter_l1 != "全部":
                filtered_df = filtered_df[filtered_df["大隊 / 局本部"] == filter_l1]
                
            st.dataframe(
                filtered_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "大隊 / 局本部": st.column_config.TextColumn("大隊/局本部", width="medium"),
                    "分隊 / 科室": st.column_config.TextColumn("分隊/科室", width="medium"),
                    "審核結果": st.column_config.TextColumn("審核結果", help="符合資格: GPA >= 80 且無缺件"),
                    "學期總平均": st.column_config.TextColumn("學期總平均", help="最低申請門檻 80.00 分")
                }
            )
            
            st.markdown("---")
            col_e1, col_e2 = st.columns([1.5, 1])
            with col_e1:
                st.markdown("##### 📥 匯出標準格式 Excel 審核清冊")
                excel_bytes = export_scholarship_excel(st.session_state.records)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                st.download_button(
                    label="📥 立即下載 Excel 審核總表 (.xlsx)",
                    data=excel_bytes,
                    file_name=f"臺東縣消防局_消防及義消子女獎學金審核名冊_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                
                # 📦 一鍵下載所有分隊上傳的原始照片壓縮包
                st.markdown("<br>", unsafe_allow_html=True)
                zip_bytes = package_uploads_zip()
                st.download_button(
                    label="📦 一鍵下載所有分隊上傳原始照片壓縮包 (.zip)",
                    data=zip_bytes,
                    file_name=f"臺東縣消防局_獎學金申請原始檔案打包_{timestamp}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
            with col_e2:
                st.markdown("##### 📋 本地儲存與清冊規格說明")
                st.markdown(f"""
                - **檔案自動歸檔**：
                  - 原始照片：`./uploads/`
                  - 數據總表：`./data/獎學金總表.xlsx`
                - **清冊欄位**：完整 14 欄（含案件編號、雙階層組織、5項檢核、簽章）
                - **照片打包**：可隨時一鍵打包下載所有分隊上傳的佐證照片
                """)
                
            st.markdown("---")
            st.markdown('<div class="section-title">☁️ 雲端試算表 (Google Sheets) 一鍵同步</div>', unsafe_allow_html=True)
            col_g1, col_g2 = st.columns([1.4, 1.1])
            with col_g1:
                webhook_url = st.text_input(
                    "Google 試算表 Webhook 網址 (Google Apps Script Web App URL)",
                    value=st.session_state.get("google_sheet_webhook", ""),
                    help="請貼上您的 Google Apps Script 網路應用程式部署網址 (https://script.google.com/macros/s/.../exec)"
                )
                if webhook_url != st.session_state.get("google_sheet_webhook", ""):
                    st.session_state.google_sheet_webhook = webhook_url
                    
                if st.button("☁️ 立即同步審核名冊至 Google 雲端試算表", use_container_width=True, type="primary"):
                    if not webhook_url:
                        st.error("請先在上方輸入 Google 試算表 Webhook 網址！")
                    else:
                        with st.spinner("正在將審核清冊同步至 Google 雲端試算表..."):
                            success, msg = sync_to_google_sheets(webhook_url, st.session_state.records)
                            if success:
                                st.success(f"🎉 {msg}")
                            else:
                                st.error(f"❌ {msg}")
                                
            with col_g2:
                with st.expander("📖 1 分鐘建立 Google Sheets 雲端連線教學", expanded=False):
                    st.markdown("""
                    **三步驟快速設定 Google 試算表自動同步：**
                    1. 建立一個新的 [Google 試算表](https://sheets.new)。
                    2. 點擊上方選單 **「擴充功能」 $\\rightarrow$ 「Apps Script」**。
                    3. 貼上標準同步腳本並發布為 **「網路應用程式」** (所有人可存取)。
                    """)
                    st.code(GOOGLE_APPS_SCRIPT_TEMPLATE, language="javascript")
