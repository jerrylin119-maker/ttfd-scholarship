"""
臺東縣消防局 消防及義消子女獎學金 AI 智慧審核網頁系統
前後台分流權限設計 ‧ 磁碟自動儲存 (./uploads/ & ./data/獎學金總表.xlsx) ‧ 雲端試算表自動同步
"""

import os
import io
import re
import hmac
import hashlib
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
    load_records_json,
    load_delete_log,
    find_duplicate_cases,
    case_dup_keys,
    JSON_FILE,
    load_all_known_case_ids,
    delete_cases_from_storage,
    package_uploads_zip,
    EXCEL_FILE,
    UPLOADS_DIR,
    DATA_DIR
)

from org_structure import TAITUNG_FIRE_ORG, get_level1_units, get_level2_units, get_scholarship_types
from gemini_analyzer import (
    analyze_scholarship_documents,
    evaluate_eligibility,
    ATTACHMENT_NAMES
)
import cloud_sync
importlib.reload(cloud_sync)
from cloud_sync import sync_to_google_sheets, test_webhook_connection, GOOGLE_APPS_SCRIPT_TEMPLATE

@st.cache_resource
def get_global_server_config():
    """全伺服器所有連線共享之全域設定 (Webhook 網址與 API Key)"""
    return {
        "webhook_url": "",
        "api_key": ""
    }

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

def build_new_case(ai_result, images, labels, stype, l1, l2) -> dict:
    return {
        "id": generate_case_id(),
        "scholarship_type": stype,
        "unit_level1": l1,
        "unit_level2": l2,
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
        "images": images,
        "image_labels": labels,
        "submitted_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

def finalize_submission(new_case: dict):
    """存檔 → 更新畫面狀態 → 自動同步雲端試算表"""
    save_case_to_storage(new_case)
    st.session_state.records.append(new_case)
    st.session_state.selected_case_id = new_case["id"]
    st.session_state.last_submitted_case = new_case
    st.session_state.pending_submission = None
    st.session_state.camera_photos = []
    webhook_url = st.session_state.get("google_sheet_webhook", "") or load_persistent_webhook()
    if webhook_url:
        try:
            sync_latest_to_google_sheets(webhook_url)
        except Exception:
            pass

def sync_latest_to_google_sheets(webhook_url: str):
    """以 ./data/records.json 內最新、最完整的資料同步雲端試算表 (不使用可能過期的畫面暫存)"""
    latest = load_records_json()
    if not latest:
        return False, "目前系統內沒有任何案件資料，為避免清空雲端試算表，已略過同步"
    return sync_to_google_sheets(webhook_url, latest)

def generate_case_id() -> str:
    """
    產生簡化版案件編號：{民國年}-{序號}，如 115-01、115-02...
    序號依當年度已存在的案件編號自動接續編下去（與其他年度或舊格式編號互不影響）。
    """
    roc_year = datetime.now().year - 1911
    prefix = f"{roc_year}-"
    pattern = re.compile(rf"^{roc_year}-(\d+)$")

    max_seq = 0
    # 同時參考檔案內最新資料、刪除前備份與目前畫面資料：避免多人同時送件時編號重複，
    # 也避免已刪除案件的編號被重複使用 (紙本申請表上已註記該編號)
    known_ids = load_all_known_case_ids() + [str(r.get("id", "")) for r in st.session_state.records]
    for cid in known_ids:
        m = pattern.match(cid)
        if m:
            max_seq = max(max_seq, int(m.group(1)))

    return f"{prefix}{max_seq + 1:02d}"

SECRETS_FILE = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
DEFAULT_ADMIN_PASSWORD = "ttfd888"

# ---------------------------------------------------------------------
# 帳號與權限：各大隊各一組帳號 (只看得到、只管理自己大隊的案件) + 業務科帳號 (全部案件)
# 密碼一律放在 Streamlit Cloud 的 Secrets (本專案程式碼公開於 GitHub，不可把密碼寫在程式碼裡)：
#   HQ_PASSWORD = "業務科密碼"
#   [UNIT_PASSWORDS]
#   "臺東大隊" = "..."   "關山大隊" = "..."   "成功大隊" = "..."   "大武大隊" = "..."
# 一旦設定了 UNIT_PASSWORDS，舊版共用密碼即自動失效。
# ---------------------------------------------------------------------
HQ_UNIT_LABEL = "業務科（民力及訓練科）"
HQ_LEVEL1_NAME = "局本部及業務科室"
UNIT_ACCOUNTS = [u for u in get_level1_units() if u != HQ_LEVEL1_NAME]

# 業務科備援密碼：僅存加鹽雜湊值 (未設定 Secrets 時使用)；設定 HQ_PASSWORD 後優先使用 Secrets
SUPER_ADMIN_SALT = "ec983ba2e6b1b552851e021e13762f0e"
SUPER_ADMIN_HASH = "032e9ab341e29c579ca5df5ae2f1e77e650a39a8c9224f66448318d4c01c6717"

def _get_secret(name, default=None):
    try:
        if hasattr(st, "secrets") and name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return default

def get_unit_passwords() -> dict:
    raw = _get_secret("UNIT_PASSWORDS")
    if not raw:
        return {}
    try:
        return {str(k): str(v).strip() for k, v in dict(raw).items() if str(v).strip()}
    except Exception:
        return {}

def _safe_equal(a: str, b: str) -> bool:
    return hmac.compare_digest(a.encode("utf-8"), b.encode("utf-8"))

def verify_hq_password(pwd: str) -> bool:
    pwd = (pwd or "").strip()
    if not pwd:
        return False
    for key in ("HQ_PASSWORD", "SUPER_ADMIN_PASSWORD"):
        secret_pwd = str(_get_secret(key, "") or "").strip()
        if secret_pwd:
            return _safe_equal(pwd, secret_pwd)
    calc = hashlib.pbkdf2_hmac("sha256", pwd.encode("utf-8"), bytes.fromhex(SUPER_ADMIN_SALT), 200000).hex()
    return hmac.compare_digest(calc, SUPER_ADMIN_HASH)

def attempt_login(unit_choice: str, pwd: str):
    """驗證登入；成功時寫入 role / scope_unit / is_admin。返回 (是否成功, 訊息)"""
    pwd = (pwd or "").strip()
    if not pwd:
        return False, "請輸入密碼"
    if unit_choice == HQ_UNIT_LABEL:
        if verify_hq_password(pwd):
            st.session_state.role, st.session_state.scope_unit = "hq", None
        elif not get_unit_passwords() and _safe_equal(pwd, DEFAULT_ADMIN_PASSWORD):
            # 過渡期：尚未設定各大隊密碼時，舊版共用密碼暫時可用 (僅可檢視/複核，不可刪除)
            st.session_state.role, st.session_state.scope_unit = "legacy", None
        else:
            return False, "密碼錯誤"
    else:
        expected = get_unit_passwords().get(unit_choice)
        if not expected:
            return False, f"{unit_choice} 尚未設定密碼，請洽業務科"
        if not _safe_equal(pwd, expected):
            return False, "密碼錯誤"
        st.session_state.role, st.session_state.scope_unit = "unit", unit_choice
    st.session_state.is_admin = True
    return True, "登入成功"

def logout_admin():
    st.session_state.is_admin = False
    st.session_state.role = None
    st.session_state.scope_unit = None

def current_role():
    return st.session_state.get("role")

def current_scope():
    """各大隊帳號回傳該大隊名稱；業務科/舊版帳號回傳 None (代表全部)"""
    return st.session_state.get("scope_unit")

def can_manage_system() -> bool:
    """API Key、雲端試算表等系統設定：僅業務科與過渡期舊版帳號"""
    return current_role() in ("hq", "legacy")

def can_delete() -> bool:
    return current_role() in ("hq", "unit")

def visible(records):
    """依登入帳號過濾可見案件：各大隊只看得到自己大隊的案件"""
    scope = current_scope()
    return [r for r in records if r.get("unit_level1") == scope] if scope else list(records)

def actor_label() -> str:
    if current_role() == "unit":
        return f"{current_scope()} 承辦人"
    return "業務科" if current_role() == "hq" else "舊版共用帳號"

def load_persistent_api_key() -> str:
    # 1. 優先從 Streamlit Cloud 內建 Secrets 讀取 (支援雲端發布)
    try:
        if hasattr(st, "secrets") and "GEMINI_API_KEY" in st.secrets:
            val = str(st.secrets["GEMINI_API_KEY"]).strip()
            if val:
                return val
    except Exception:
        pass

    # 2. 從本機 .streamlit/secrets.toml 讀取
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
            
    # 3. 從環境變數讀取
    return os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or ""

def save_persistent_api_key(key: str):
    clean_key = key.strip()
    os.environ["GEMINI_API_KEY"] = clean_key
    st.session_state.api_key = clean_key
    
    # 嘗試寫入本機檔案 (若在 Streamlit 雲端唯讀環境則安全跳過，避免 OSError)
    try:
        os.makedirs(os.path.dirname(SECRETS_FILE), exist_ok=True)
        with open(SECRETS_FILE, "w", encoding="utf-8") as f:
            f.write('GEMINI_API_KEY = "' + clean_key + '"\n')
    except Exception:
        pass

def load_persistent_webhook() -> str:
    # 1. 優先從全域伺服器共享記憶體讀取
    g_conf = get_global_server_config()
    if g_conf.get("webhook_url"):
        return g_conf["webhook_url"]
        
    # 2. 從 Streamlit Secrets 讀取
    try:
        if hasattr(st, "secrets") and "GOOGLE_SHEET_WEBHOOK" in st.secrets:
            val = str(st.secrets["GOOGLE_SHEET_WEBHOOK"]).strip()
            if val:
                g_conf["webhook_url"] = val
                return val
    except Exception:
        pass
        
    # 3. 從本地檔案讀取
    webhook_f = os.path.join(DATA_DIR, "webhook.txt")
    if os.path.exists(webhook_f):
        try:
            with open(webhook_f, "r", encoding="utf-8") as f:
                val = f.read().strip()
                if val:
                    g_conf["webhook_url"] = val
                    return val
        except Exception:
            pass
            
    env_val = os.environ.get("GOOGLE_SHEET_WEBHOOK") or ""
    if env_val:
        g_conf["webhook_url"] = env_val
    return env_val

def save_persistent_webhook(url: str):
    clean_url = url.strip()
    st.session_state.google_sheet_webhook = clean_url
    os.environ["GOOGLE_SHEET_WEBHOOK"] = clean_url
    get_global_server_config()["webhook_url"] = clean_url
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        webhook_f = os.path.join(DATA_DIR, "webhook.txt")
        with open(webhook_f, "w", encoding="utf-8") as f:
            f.write(clean_url)
    except Exception:
        pass

# 初始化 Session State (優先從 ./data/records.json 載入歷史儲存紀錄)
if "records" not in st.session_state:
    st.session_state.records = load_stored_cases()
    try:
        st.session_state.records_mtime = os.path.getmtime(JSON_FILE)
    except OSError:
        st.session_state.records_mtime = None

if "google_sheet_webhook" not in st.session_state:
    st.session_state.google_sheet_webhook = load_persistent_webhook()

if "selected_case_id" not in st.session_state:
    st.session_state.selected_case_id = st.session_state.records[0]["id"] if st.session_state.records else None

if "api_key" not in st.session_state:
    st.session_state.api_key = load_persistent_api_key()

if "camera_photos" not in st.session_state:
    st.session_state.camera_photos = []

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

if "role" not in st.session_state:
    st.session_state.role = None
    st.session_state.scope_unit = None

if "login_fail" not in st.session_state:
    st.session_state.login_fail = 0

if "upload_ver" not in st.session_state:
    st.session_state.upload_ver = 0

if "pending_submission" not in st.session_state:
    st.session_state.pending_submission = None

def refresh_records_if_changed():
    """案件資料檔有變動 (例如其他分隊剛送件) 時，重新載入，讓審核人員看到最新案件"""
    try:
        mtime = os.path.getmtime(JSON_FILE)
    except OSError:
        mtime = None
    if st.session_state.get("records_mtime") != mtime:
        st.session_state.records = load_stored_cases()
        st.session_state.records_mtime = mtime

if st.session_state.is_admin:
    refresh_records_if_changed()

if "last_submitted_case" not in st.session_state:
    st.session_state.last_submitted_case = None

# ----------------- 側邊欄 -----------------
with st.sidebar:
    st.markdown("## 🚒 臺東縣消防局\n### 獎學金申請審核系統")

    if not st.session_state.is_admin:
        st.info("📍 **當前身分：申請同仁專區**\n\n(已啟用個資防護，僅可交件)")

        with st.expander("🔐 各大隊及業務科 審核人員登入", expanded=False):
            st.markdown("請選擇您的**單位帳號**，並輸入該單位的密碼（各大隊只看得到、也只能管理自己大隊的案件）：")
            locked = st.session_state.login_fail >= 5
            if locked:
                st.error("❌ 密碼錯誤次數過多，本次連線已鎖定，請重新整理頁面後再試。")
            else:
                login_unit = st.selectbox("單位帳號", UNIT_ACCOUNTS + [HQ_UNIT_LABEL], key="login_unit")
                pwd_input = st.text_input("密碼", type="password", key="admin_pwd_input")
                if st.button("🔑 登入審核管理後台", use_container_width=True, type="primary"):
                    ok_login, login_msg = attempt_login(login_unit, pwd_input)
                    if ok_login:
                        st.session_state.login_fail = 0
                        st.rerun()
                    else:
                        st.session_state.login_fail += 1
                        st.error(f"❌ {login_msg}！請洽業務科獎學金承辦人。（已錯 {st.session_state.login_fail}/5 次）")
    else:
        role_now = current_role()
        if role_now == "unit":
            st.success(f"👑 **當前身分：{current_scope()} 審核承辦人**\n\n僅可檢視、複核、刪除本大隊案件")
        elif role_now == "hq":
            st.success("👑 **當前身分：業務科（民力及訓練科）管理員**\n\n可檢視、管理全部案件")
        else:
            st.warning("⚠️ **目前使用舊版共用密碼登入**\n\n業務科尚未設定各大隊密碼，此帳號可看到全部案件、但不可刪除。請盡速改用各單位專屬帳號。")
        if role_now == "hq" and not get_unit_passwords():
            st.info("ℹ️ 尚未設定各大隊密碼，各大隊帳號目前無法登入。請至 Streamlit Cloud 的 Secrets 設定 UNIT_PASSWORDS（設定後，舊版共用密碼會自動失效）。")
        if st.button("🚪 登出審核後台（切換回同仁申請模式）", use_container_width=True):
            logout_admin()
            st.rerun()

        if can_manage_system():
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

        st.markdown("---")
        st.subheader("⚡ 資料管理")
        st.caption("🗑️ 如有重複或誤送的案件，請至「📊 全局審核總表清冊」分頁最下方的「刪除指定案件」，只會刪除您勾選的案件，其餘案件不受影響。")

    st.markdown("---")
    st.subheader("📌 審查標準門檻")
    st.markdown(r"""
    - **學業成績**：學期總平均 $\ge 80.0$ 分
    - **必備附件 (5項)**：
      1. 獎學金申請表
      2. 學生證或在學證明
      3. 前學期成績證明單
      4. 戶口名簿影本 / 戶籍謄本
      5. 消防 / 義消服務證明（服務證或派令皆可）
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
            st.write(f"🏆 **獎學金類別**：{rec.get('scholarship_type', '未指定')}")
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
            ("service_certificate", "5. 消防/義消在職或服務證明（服務證或派令皆可）")
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
            st.session_state.upload_ver += 1  # 換一個全新的上傳欄位，避免同一批檔案被誤送第二次
            st.rerun()
            
        # 已成功送出：只顯示收執聯，隱藏上傳表單，避免同仁以為沒送出而重複上傳
        st.info("✅ 本案已成功送出並存檔，**請勿重複上傳**。如需送下一筆，請按上方「繼續提交下一筆申請案件」。")
        st.stop()

    # 疑似重複送件：先請同仁確認，避免同一案件重複出現
    pend = st.session_state.pending_submission
    if pend:
        st.warning("⚠️ **這件申請看起來已經送出過了！** 請先確認，避免重複上傳。")
        for d in pend["dups"]:
            st.markdown(
                f"- 已存在案件編號 **{d.get('id')}**｜{d.get('scholarship_type', '')}｜"
                f"{d.get('unit_level1', '')}/{d.get('unit_level2', '')}｜"
                f"家長：{d.get('applicant_name') or '—'}｜子女：{d.get('child_name') or '—'}｜"
                f"送件時間：{d.get('submitted_at', '—')}｜狀態：{d.get('review_status', '—')}"
            )
        st.caption("若您是為了「補件或更正資料」而重新送件，請按「仍要送出」，並通知大隊承辦人刪除舊的案件；若只是重複上傳，請按「取消」。")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            if st.button("✅ 這是補件／更正，仍要送出", use_container_width=True, type="primary", key="pend_confirm"):
                new_case = build_new_case(pend["ai_result"], pend["images"], pend["labels"], pend["type"], pend["l1"], pend["l2"])
                finalize_submission(new_case)
                st.balloons()
                st.rerun()
        with col_p2:
            if st.button("❌ 取消，不重複送出", use_container_width=True, key="pend_cancel"):
                st.session_state.pending_submission = None
                st.session_state.upload_ver += 1
                st.rerun()
        st.stop()

    # 申請交件表單區
    st.markdown('<div class="section-title">📤 線上申請交件與照片上傳</div>', unsafe_allow_html=True)
    
    # 獎學金類別選擇區
    st.markdown("##### 🏆 第一步：請選擇欲申請的獎學金類別")
    upload_scholarship_type = st.selectbox(
        "獎學金類別：",
        get_scholarship_types(),
        key="pub_scholarship_type"
    )
    st.markdown("---")

    # 兩階層組織單位選擇區
    st.markdown("##### 🏢 第二步：請選擇您的所屬單位")
    col_sel_l1, col_sel_l2 = st.columns(2)
    with col_sel_l1:
        upload_level1 = st.selectbox("1. 大隊 / 局本部：", get_level1_units(), key="pub_upload_unit_l1")
    with col_sel_l2:
        level2_choices = get_level2_units(upload_level1)
        upload_level2 = st.selectbox("2. 分隊 / 科室：", level2_choices, key="pub_upload_unit_l2")

    st.caption(f"📍 您選擇的申請類別與單位：**{upload_scholarship_type}** ｜ **臺東縣消防局 {upload_level1} ➔ {upload_level2}**")
    st.markdown("---")

    st.markdown("##### 📁 第三步：選擇上傳照片方式 (可一次傳多張)")
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
                key=f"pub_scholarship_uploader_{st.session_state.upload_ver}"
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
            camera_file = st.camera_input("📷 請將鏡頭對準申請文件拍照：", key=f"pub_camera_{st.session_state.upload_ver}")
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
            f"🚀 確認交件並開始 AI 智慧審查（{upload_scholarship_type} - {upload_level1}/{upload_level2}）",
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
                    
                    probe_case = {
                        "scholarship_type": upload_scholarship_type,
                        "applicant_name": ai_result.get("applicant_name", ""),
                        "applicant_id": ai_result.get("applicant_id", ""),
                        "child_name": ai_result.get("child_name", ""),
                    }
                    dups = find_duplicate_cases(probe_case)
                    if dups:
                        # 疑似重複：暫不存檔，先請同仁確認
                        st.session_state.pending_submission = {
                            "ai_result": ai_result, "images": all_prepared_images, "labels": all_prepared_labels,
                            "type": upload_scholarship_type, "l1": upload_level1, "l2": upload_level2, "dups": dups,
                        }
                        progress_bar.empty()
                        st.rerun()

                    new_case = build_new_case(ai_result, all_prepared_images, all_prepared_labels,
                                              upload_scholarship_type, upload_level1, upload_level2)

                    # 💾 自動存入 ./uploads/ 與追加寫入 ./data/獎學金總表.xlsx，並同步雲端試算表
                    progress_bar.progress(90, text="正在儲存原始照片並同步總表...")
                    finalize_submission(new_case)

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
    admin_title = f"{current_scope()} 專用" if current_scope() else "業務科專用" if current_role() == "hq" else "各大隊及業務科專用"
    st.markdown(f"""
    <div class="main-header">
        <h1>🚒 臺東縣消防局 獎學金審核管理系統【{admin_title}】</h1>
        <p>管理員後台 ‧ 自動儲存 (./uploads/ & 獎學金總表.xlsx) ‧ 左圖右表人工複核 ‧ Google 雲端試算表即時同步</p>
    </div>
    """, unsafe_allow_html=True)

    my_records = visible(st.session_state.records)   # 各大隊帳號只看得到自己大隊的案件
    records = my_records
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
        if not my_records:
            st.info("目前尚無任何案件資料。")
        else:
            case_options = {
                r["id"]: f"【{r.get('id')}】[{r.get('scholarship_type', '未分類')}] [{r.get('unit_level1', '未定')} / {r.get('unit_level2', '未定')}] {r.get('applicant_name', '未命名')} / 子女: {r.get('child_name', '未命名')} ({r.get('category', '未定')}) - {r.get('review_status', '待審')}"
                for r in my_records
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
                st.markdown(f'<div class="section-title">🖼️ 原始申請資料與成績單檢視 &nbsp; <span class="unit-tag">{curr_case.get("scholarship_type", "未分類")}</span> &nbsp; <span class="unit-tag">{curr_case.get("unit_level1", "")} / {curr_case.get("unit_level2", "")}</span></div>', unsafe_allow_html=True)
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
                    st.markdown("##### 🏆 獎學金類別")
                    type_list = get_scholarship_types()
                    curr_type = curr_case.get("scholarship_type", type_list[0])
                    idx_type = type_list.index(curr_type) if curr_type in type_list else 0
                    edit_scholarship_type = st.selectbox("獎學金類別", type_list, index=idx_type, key=f"admin_edit_type_{curr_case['id']}")

                    st.markdown("##### 🏢 所屬單位")
                    col_u1, col_u2 = st.columns(2)
                    
                    level1_list = [current_scope()] if current_scope() else get_level1_units()
                    curr_l1 = curr_case.get("unit_level1", level1_list[0])
                    idx_l1 = level1_list.index(curr_l1) if curr_l1 in level1_list else 0
                    
                    with col_u1:
                        edit_l1 = st.selectbox("大隊 / 局本部", level1_list, index=idx_l1, key=f"admin_edit_l1_{curr_case['id']}")
                    
                    level2_list = get_level2_units(edit_l1)
                    curr_l2 = curr_case.get("unit_level2", level2_list[0] if level2_list else "")
                    idx_l2 = level2_list.index(curr_l2) if curr_l2 in level2_list else 0
                    
                    with col_u2:
                        edit_l2 = st.selectbox("分隊 / 科室", level2_list, index=idx_l2, key=f"admin_edit_l2_{curr_case['id']}")
                    
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
                        chk_service = st.checkbox("5. 消防/義消服務證明（服務證或派令皆可）", value=att.get("service_certificate", True))
                        
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
                        curr_case["scholarship_type"] = edit_scholarship_type
                        curr_case["unit_level1"] = current_scope() or edit_l1
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
        admin_upload_scholarship_type = st.selectbox("0. 選擇獎學金類別：", get_scholarship_types(), key="admin_up_type")
        col_sel_l1, col_sel_l2 = st.columns(2)
        with col_sel_l1:
            admin_upload_l1 = st.selectbox("1. 選擇大隊 / 局本部：", [current_scope()] if current_scope() else get_level1_units(), key="admin_up_l1")
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
                case_id = generate_case_id()
                new_c = {
                    "id": case_id,
                    "scholarship_type": admin_upload_scholarship_type,
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
        
        if not my_records:
            st.info("尚無審核紀錄。")
        else:
            table_rows = []
            for idx, r in enumerate(my_records, 1):
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
                    "獎學金類別": r.get("scholarship_type", "未分類"),
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
            col_flt0, col_flt1, col_flt2, col_flt3 = st.columns(4)
            with col_flt0:
                # 含「未分類」選項，避免舊資料 (更新前無獎學金類別欄位) 被篩選條件預設隱藏
                type_options = list(dict.fromkeys(get_scholarship_types() + list(df["獎學金類別"].unique())))
                type_filter = st.multiselect("依獎學金類別篩選：", type_options, default=type_options, key="admin_flt_type")
            with col_flt1:
                all_l1 = ["全部"] + list(df["大隊 / 局本部"].unique())
                filter_l1 = st.selectbox("依大隊/局本部篩選：", all_l1, key="admin_flt_l1")
            with col_flt2:
                status_filter = st.multiselect("依審核結果篩選：", ["符合資格", "待補件", "不符資格"], default=["符合資格", "待補件", "不符資格"], key="admin_flt_stat")
            with col_flt3:
                cat_filter = st.multiselect("依組別篩選：", ["大專院校", "高中職", "國中", "國小"], default=["大專院校", "高中職", "國中", "國小"], key="admin_flt_cat")

            filtered_df = df[df["獎學金類別"].isin(type_filter) & df["審核結果"].isin(status_filter) & df["組別"].isin(cat_filter)]
            if filter_l1 != "全部":
                filtered_df = filtered_df[filtered_df["大隊 / 局本部"] == filter_l1]
                
            st.dataframe(
                filtered_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "獎學金類別": st.column_config.TextColumn("獎學金類別", width="medium"),
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

                # 匯出內容跟著畫面篩選走：只匯出目前篩選條件下、畫面上看得到的案件，
                # 避免大隊承辦人下載後誤以為只有自己大隊卻夾帶了其他大隊的個資。
                matched_ids = set(filtered_df["案件編號"])
                filtered_records = [r for r in my_records if r.get("id", "") in matched_ids]

                scope_label = filter_l1 if filter_l1 != "全部" else "全局"
                safe_scope = scope_label.replace("/", "_").replace(" ", "")
                st.caption(f"📌 目前匯出範圍：**{scope_label}**（依畫面篩選條件，共 {len(filtered_records)} 筆）")

                excel_bytes = export_scholarship_excel(filtered_records)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

                st.download_button(
                    label="📥 立即下載 Excel 審核總表 (.xlsx)",
                    data=excel_bytes,
                    file_name=f"臺東縣消防局_消防及義消子女獎學金審核名冊_{safe_scope}_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                
                # 📦 一鍵下載所有分隊上傳的原始照片壓縮包
                st.markdown("<br>", unsafe_allow_html=True)
                if current_scope():
                    zip_bytes = package_uploads_zip(only_files=[f for r in my_records for f in r.get("image_paths", [])])
                else:
                    zip_bytes = package_uploads_zip()
                st.download_button(
                    label=("📦 一鍵下載本大隊分隊上傳原始照片壓縮包 (.zip)" if current_scope() else "📦 一鍵下載所有分隊上傳原始照片壓縮包 (.zip)"),
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
                - **分頁方式**：Excel 依「獎學金類別」自動分開為獨立工作表（義消聯合總會獎助學金 / 本局津芳冰城陳慶銳先生獎學金）
                - **匯出範圍**：跟著上方篩選條件走，選哪個大隊就只匯出該大隊的資料，避免夾帶其他大隊個資
                - **照片打包**：可隨時一鍵打包下載所有分隊上傳的佐證照片（此項不受篩選影響，固定為您可檢視的全部案件）
                """)
                
            if can_manage_system():
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
                        save_persistent_webhook(webhook_url)
                    
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        if st.button("🧪 測試 Webhook 連線", use_container_width=True):
                            if not webhook_url:
                                st.warning("請先輸入 Webhook 網址！")
                            else:
                                with st.spinner("連線測試中..."):
                                    ok, test_msg = test_webhook_connection(webhook_url)
                                    if ok:
                                        st.success(test_msg)
                                    else:
                                        st.error(test_msg)
                                    
                    with col_btn2:
                        if st.button("☁️ 立即同步至 Google 試算表", use_container_width=True, type="primary"):
                            if not webhook_url:
                                st.error("請先在上方輸入 Google 試算表 Webhook 網址！")
                            else:
                                with st.spinner("正在將審核清冊同步至 Google 雲端試算表..."):
                                    success, msg = sync_latest_to_google_sheets(webhook_url)
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

                        ⚠️ 若您先前已部署過舊版腳本，請改貼下方最新版程式碼並**重新部署**，
                        同步後會依「獎學金類別」自動分別建立 **義消聯合總會獎助學金** 與
                        **本局津芳冰城陳慶銳先生獎學金** 兩個工作表分頁。
                        """)
                        st.code(GOOGLE_APPS_SCRIPT_TEMPLATE, language="javascript")

            # ----------------- 刪除指定案件 (重複或誤送資料) -----------------
            st.markdown("---")
            st.markdown('<div class="section-title">🗑️ 刪除指定案件（重複或誤送資料）</div>', unsafe_allow_html=True)
            with st.expander("展開刪除工具（只會刪除您勾選的案件，其餘案件不受影響）", expanded=False):
                if not can_delete():
                    st.info("🔒 舊版共用密碼帳號沒有刪除權限，請改用「業務科」或「各大隊」專屬帳號登入。")
                else:
                    scope_now = current_scope()
                    stored_records = [r for r in load_records_json() if not scope_now or r.get("unit_level1") == scope_now]

                    # 找出疑似重複的案件 (同類別、同子女、同家長)，在選單上標註，方便挑出重複上傳的那一筆
                    key_groups = {}
                    for r in stored_records:
                        for k in case_dup_keys(r):
                            key_groups.setdefault(k, set()).add(str(r.get("id", "")))

                    def _dup_mark(r):
                        me = str(r.get("id", ""))
                        others = set()
                        for k in case_dup_keys(r):
                            others |= key_groups.get(k, set())
                        others.discard(me)
                        return f"｜⚠️疑似重複，另有 {'、'.join(sorted(others))}" if others else ""

                    del_options = {}
                    for r in stored_records:
                        del_options[str(r.get("id", ""))] = (
                            f"{r.get('id', '')}｜{r.get('scholarship_type', '未分類')}｜"
                            f"{r.get('unit_level1', '')}/{r.get('unit_level2', '')}｜"
                            f"家長:{r.get('applicant_name', '') or '—'}｜子女:{r.get('child_name', '') or '—'}｜"
                            f"送件:{r.get('submitted_at', '—')}｜{r.get('review_status', '—')}" + _dup_mark(r)
                        )

                    if st.session_state.get("del_result"):
                        kind, text = st.session_state.pop("del_result")
                        (st.success if kind == "ok" else st.error)(text)

                    def _do_delete():
                        if not can_delete():
                            return
                        ids = list(st.session_state.get("admin_del_select", []))
                        deleted_ids, msg = delete_cases_from_storage(ids, allowed_unit=current_scope(), actor=actor_label())
                        if not deleted_ids:
                            st.session_state["del_result"] = ("err", f"❌ {msg}")
                            return
                        # 以檔案內最新資料重新載入畫面，並清除選取狀態
                        st.session_state.records = load_stored_cases()
                        try:
                            st.session_state.records_mtime = os.path.getmtime(JSON_FILE)
                        except OSError:
                            pass
                        remain = visible(st.session_state.records)
                        if st.session_state.get("selected_case_id") in deleted_ids:
                            st.session_state.selected_case_id = remain[0]["id"] if remain else None
                        st.session_state["admin_del_select"] = []
                        st.session_state["admin_del_confirm"] = False
                        text = f"✅ {msg}（已自動備份刪除前的資料，並記錄於刪除紀錄）"
                        hook = st.session_state.get("google_sheet_webhook", "") or load_persistent_webhook()
                        if hook:
                            try:
                                ok_s, msg_s = sync_latest_to_google_sheets(hook)
                                text += f"　☁️ 雲端試算表：{msg_s}"
                            except Exception as e:
                                text += f"　⚠️ 雲端試算表同步失敗，請洽業務科手動同步：{e}"
                        st.session_state["del_result"] = ("ok", text)

                    if not del_options:
                        st.info("目前沒有可刪除的案件。")
                    else:
                        if scope_now:
                            st.caption(f"您只能看到、也只能刪除 **{scope_now}** 的案件。")
                        st.multiselect(
                            "請選擇要刪除的案件（可多選）：",
                            options=list(del_options.keys()),
                            format_func=lambda cid: del_options[cid],
                            key="admin_del_select",
                            placeholder="點選要刪除的重複／誤送案件…"
                        )
                        picked = st.session_state.get("admin_del_select", [])
                        if picked:
                            st.warning(
                                f"⚠️ 即將永久刪除 **{len(picked)}** 筆案件（含其原始照片）：{'、'.join(picked)}。"
                                f"其餘 **{len(stored_records) - len(picked)}** 筆案件不會被更動。"
                            )
                            st.checkbox("我已確認上列案件為重複上傳或誤送資料，同意永久刪除", key="admin_del_confirm")
                            st.button(
                                "🗑️ 確認刪除選取的案件",
                                type="primary",
                                disabled=not st.session_state.get("admin_del_confirm", False),
                                on_click=_do_delete,
                                use_container_width=True
                            )

                    log_rows = load_delete_log(unit=current_scope())
                    if log_rows:
                        st.markdown("##### 📜 最近刪除紀錄")
                        st.dataframe(
                            pd.DataFrame(log_rows[:30]).rename(columns={
                                "time": "刪除時間", "by": "執行帳號", "id": "案件編號", "scholarship_type": "獎學金類別",
                                "unit_level1": "大隊/局本部", "unit_level2": "分隊/科室",
                                "applicant_name": "家長", "child_name": "子女", "submitted_at": "原送件時間"
                            }),
                            use_container_width=True, hide_index=True
                        )
