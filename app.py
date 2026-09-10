import mammoth
import base64
import hashlib
import html
import io
import json
import os
import random
import re
import smtplib
import tempfile
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from fpdf import FPDF
from PIL import Image
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

import google.generativeai as genai
import streamlit as st
from streamlit_local_storage import LocalStorage
from streamlit_pdf_viewer import pdf_viewer

# --- CẤU HÌNH TRANG ĐẦU TIÊN (Phải luôn nằm trên cùng) ---
st.set_page_config(page_title="Bệnh án Lâm sàng", layout="wide")

# ==============================================================================
# HÀM ĐIỀU PHỐI API KEY (CHỐNG RATE LIMIT)
# ==============================================================================
def get_feature_model(feature_key_name, model_name="gemini-3.1-flash-lite"):
    """Lấy model AI với API key chuyên biệt cho từng tác vụ."""
    api_key = st.secrets.get(feature_key_name) or st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        return None
    genai.configure(api_key=api_key)
    return genai.GenerativeModel(model_name)

# ==============================================================================
# BẢO MẬT & XÁC THỰC DANH TÍNH (OTP + GMAIL + THIẾT BỊ)
# ==============================================================================
AUTH_STORAGE_KEY = "clinical_user_auth_token"
local_storage = LocalStorage()

def generate_auth_token(email):
    secret = st.secrets.get("AUTH_SECRET_KEY", "default_secret_medical_key_2026")
    raw_str = f"{email}_{secret}"
    return hashlib.sha256(raw_str.encode("utf-8")).hexdigest()

def send_otp_email(target_email, otp_code):
    sender_mail = st.secrets.get("SENDER_EMAIL")
    sender_pass = st.secrets.get("SENDER_APP_PASSWORD")
    if not (sender_mail and sender_pass):
        st.error("⚠️ Hệ thống chưa cấu hình SENDER_EMAIL hoặc SENDER_APP_PASSWORD trong Secrets!")
        return False
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_mail
        msg['To'] = target_email
        msg['Subject'] = f"🔑 Mã xác thực truy cập Bệnh án Lâm sàng: {otp_code}"
        body = f"Xin chào,\n\nMã xác thực (OTP) dùng để đăng nhập vào Ứng dụng Bệnh án Lâm sàng của bạn là:\n\n👉  {otp_code}  👈\n\nMã có hiệu lực trong phiên đăng nhập này."
        msg.attach(MIMEText(body, 'plain'))
        server = smtplib.SMTP('smtp.gmail.com', 587, timeout=10)
        server.starttls()
        server.login(sender_mail, sender_pass)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Lỗi kết nối gửi email xác thực: {e}")
        return False

def send_login_notification(user_email):
    admin_mail = st.secrets.get("ADMIN_EMAIL")
    sender_mail = st.secrets.get("SENDER_EMAIL")
    sender_pass = st.secrets.get("SENDER_APP_PASSWORD")
    if not (admin_mail and sender_mail and sender_pass): return
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_mail
        msg['To'] = admin_mail
        msg['Subject'] = f"🔔 [Bệnh Án Lâm Sàng] Người dùng mới đăng nhập: {user_email}"
        login_time = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        body = f"Hệ thống Bệnh Án Lâm Sàng ghi nhận lượt truy cập thành công:\n- Người dùng: {user_email}\n- Thời gian: {login_time}"
        msg.attach(MIMEText(body, 'plain'))
        server = smtplib.SMTP('smtp.gmail.com', 587, timeout=10)
        server.starttls()
        server.login(sender_mail, sender_pass)
        server.send_message(msg)
        server.quit()
    except Exception:
        pass

def check_password():
    admin_token_secret = str(st.secrets.get("ADMIN_BYPASS_TOKEN", "")).strip()
    url_admin_key = str(st.query_params.get("nam", "")).strip()
    is_admin_access = False
    if admin_token_secret and url_admin_key == admin_token_secret: is_admin_access = True
    elif "nam" in st.query_params and not admin_token_secret: is_admin_access = True

    if is_admin_access:
        st.session_state["password_correct"] = True
        st.session_state["is_admin"] = True
        if "logged_in_user" not in st.session_state: st.session_state["logged_in_user"] = "Admin"
        return True

    if st.session_state.get("password_correct"): return True

    try:
        saved_auth_raw = local_storage.getItem(AUTH_STORAGE_KEY)
        if saved_auth_raw:
            auth_data = json.loads(saved_auth_raw) if isinstance(saved_auth_raw, str) else saved_auth_raw
            if isinstance(auth_data, dict):
                saved_email = auth_data.get("email", "")
                saved_token = auth_data.get("token", "")
                if saved_email and saved_token == generate_auth_token(saved_email):
                    st.session_state["password_correct"] = True
                    st.session_state["logged_in_user"] = saved_email
                    if not st.session_state.get("sinh_vien"): st.session_state["sinh_vien"] = saved_email.split("@")[0]
                    st.rerun()
    except Exception: pass

    GMAIL_REGEX = r"^[a-zA-Z0-9](\.?[a-zA-Z0-9_-]){5,29}@gmail\.com$"
    with st.container():
        st.markdown("### 🔒 Ứng dụng Bệnh án Lâm sàng (Nội bộ)")
        st.caption("Vui lòng xác thực tài khoản Gmail chính chủ. Sau khi xác thực, thiết bị này sẽ được tự động ghi nhớ:")
        col_form, _ = st.columns([1.5, 1])
        with col_form:
            input_email = st.text_input("Địa chỉ Gmail của bạn:", placeholder="tenban@gmail.com", key="login_email_input").strip().lower()
            c_otp_btn, _ = st.columns([1, 1.5])
            with c_otp_btn: btn_send_otp = st.button("📩 Gửi mã xác thực OTP", use_container_width=True)
            
            if btn_send_otp:
                if not input_email or not re.match(GMAIL_REGEX, input_email):
                    st.error("❌ Vui lòng nhập địa chỉ Gmail hợp lệ (@gmail.com)!")
                else:
                    otp_random = str(random.randint(100000, 999999))
                    with st.spinner("Đang gửi mã xác thực về hộp thư của bạn..."):
                        if send_otp_email(input_email, otp_random):
                            st.session_state["generated_otp"] = otp_random
                            st.session_state["otp_target_email"] = input_email
                            st.success(f"✅ Đã gửi mã OTP đến {input_email}. Vui lòng mở hộp thư kiểm tra!")

            input_otp = st.text_input("Mã OTP (6 chữ số từ Gmail):", placeholder="VD: 123456", key="login_otp_input").strip()
            input_pass = st.text_input("Mã truy cập nội bộ:", type="password", key="login_pass_input")
            btn_login = st.button("Đăng nhập & Ghi nhớ máy này", type="primary", use_container_width=True)
            
            if btn_login:
                mat_khau_chuan = str(st.secrets.get("APP_PASSWORD", "123456")).strip()
                sent_otp = st.session_state.get("generated_otp")
                verified_email = st.session_state.get("otp_target_email")
                
                if not input_email or input_email != verified_email: st.error("❌ Email này chưa nhận mã OTP hoặc bị sửa đổi!")
                elif not input_otp or input_otp != sent_otp: st.error("❌ Mã OTP không chính xác!")
                elif input_pass != mat_khau_chuan: st.error("❌ Mã truy cập nội bộ không chính xác!")
                else:
                    with st.spinner("Đang lưu trạng thái và ghi nhớ thiết bị..."):
                        send_login_notification(input_email)
                        token = generate_auth_token(input_email)
                        auth_payload = json.dumps({"email": input_email, "token": token}, ensure_ascii=False)
                        local_storage.setItem(AUTH_STORAGE_KEY, auth_payload)
                        time.sleep(1.2)
                        st.session_state["password_correct"] = True
                        st.session_state["logged_in_user"] = input_email
                        if not st.session_state.get("sinh_vien"): st.session_state["sinh_vien"] = input_email.split("@")[0]
                        st.session_state.pop("generated_otp", None)
                        st.session_state.pop("otp_target_email", None)
                    st.toast("✅ Xác thực thành công!", icon="🎉")
                    st.rerun()
    return False

if not check_password(): st.stop()

# --- HÀM NÉN VÀ TỐI ƯU ẢNH PHIẾU XÉT NGHIỆM TRƯỚC KHI OCR ---
def optimize_lab_image(photo_file, max_dimension=1600, quality=85):
    try:
        photo_file.seek(0)
        img = Image.open(photo_file)
        if img.mode in ("RGBA", "P"): img = img.convert("RGB")
        width, height = img.size
        if max(width, height) > max_dimension:
            if width > height:
                new_width = max_dimension
                new_height = int(height * (max_dimension / width))
            else:
                new_height = max_dimension
                new_width = int(width * (max_dimension / height))
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=quality, optimize=True)
        buffer.seek(0)
        return Image.open(buffer)
    except Exception:
        photo_file.seek(0)
        return Image.open(photo_file)

# ==============================================================================
# DANH MỤC TRƯỜNG DỮ LIỆU & NẠP BẢN NHÁP TỰ ĐỘNG
# ==============================================================================
STORAGE_KEY = "clinical_report_draft"

FIELDS_TO_SAVE = [
    "loai_benh_an",  # Phân loại bệnh án
    "ho_ten", "tuoi", "gioi_tinh", "dan_tok", "nghe_nghiep", "khoa_phong", "dia_chi", "ngay_vao_vien", "sinh_vien",
    "ly_do_vao_vien", "benh_su", 
    "bs_truoc_mo", "bs_trong_mo", "bs_sau_mo", 
    "ts_san_khoa", "ts_phu_khoa", "ts_noi_ngoai_khoa",
    "ts_benh_ly", "ts_dinh_duong", "ts_san_khoa_nhi", "ts_tiem_chung",
    "ts_phat_trien", "ts_dich_te", "ts_di_ung",
    "ts_noi_khoa", "ts_ngoai_khoa", "ts_loi_song", "ts_gia_dinh",
    "kham_vao_vien", "kham_toan_than", "sh_mach", "sh_nhiet_do", "sh_ha", 
    "sh_nhip_tho", "sh_can_nang", "sh_chieu_cao", "sh_bmi", "sh_bmi_eval",
    "ngay_hau_phau", "kham_vet_mo", "kham_dan_luu", 
    "uu_tien_co_quan", "kham_tuan_hoan", "kham_ho_hap", "kham_tieu_hoa", 
    "kham_than_kinh", "kham_tiet_nieu", "kham_co_xuong_khop", "kham_co_quan_khac", "kham_san_phu_khoa",
    "tom_tat", "chan_doan_so_bo", "chan_doan_phan_biet", "bien_luan",
    "cls_dx_xac_dinh", "cls_dx_dieu_tri", "cls_dx_khac",
    "chan_doan_xac_dinh", "bien_luan_xac_dinh",
    "dt_muc_tieu", "dt_cu_the", "dt_theo_doi",
    "tien_luong", "tu_van", "so_hang_cls"
]

def load_draft_to_session(loaded_ls):
    if "so_hang_cls" in loaded_ls: st.session_state["so_hang_cls"] = int(loaded_ls["so_hang_cls"])
    for k in FIELDS_TO_SAVE:
        if k in loaded_ls: st.session_state[k] = loaded_ls[k]
        
    # Đảm bảo ô bs_trong_mo luôn có mẫu nếu bản nháp lưu chuỗi rỗng
    mau_5_dong = (
        "- Hình thức mổ: Mổ phiên / Mổ cấp cứu\n"
        "- Phương pháp mổ: \n"
        "- Phương pháp gây mê: \n"
        "- Quá trình mổ: Không có biến chứng\n"
        "- Chẩn đoán sau mổ: "
    )
    if not str(st.session_state.get("bs_trong_mo", "")).strip():
        st.session_state["bs_trong_mo"] = mau_5_dong
    try: st.session_state["tuoi"] = int(loaded_ls.get("tuoi", 45))
    except (ValueError, TypeError): st.session_state["tuoi"] = 45
    try: st.session_state["sh_can_nang"] = float(loaded_ls.get("sh_can_nang") or 0.0)
    except (ValueError, TypeError): st.session_state["sh_can_nang"] = 0.0
    try: st.session_state["sh_chieu_cao"] = float(loaded_ls.get("sh_chieu_cao") or 0.0)
    except (ValueError, TypeError): st.session_state["sh_chieu_cao"] = 0.0
    for i in range(st.session_state.get("so_hang_cls", 3)):
        if f"cls_kq_{i}" in loaded_ls: st.session_state[f"cls_kq_{i}"] = loaded_ls[f"cls_kq_{i}"]
        if f"cls_pg_{i}" in loaded_ls: st.session_state[f"cls_pg_{i}"] = loaded_ls[f"cls_pg_{i}"]

if "da_khoi_phuc_tu_dong" not in st.session_state:
    try:
        draft_raw = local_storage.getItem(STORAGE_KEY)
        if draft_raw:
            loaded_ls = json.loads(draft_raw) if isinstance(draft_raw, str) else draft_raw
            load_draft_to_session(loaded_ls)
    except Exception: pass
    st.session_state["da_khoi_phuc_tu_dong"] = True

# Khởi tạo giá trị mặc định
if "so_hang_cls" not in st.session_state: st.session_state["so_hang_cls"] = 1
for field in FIELDS_TO_SAVE:
    if field not in st.session_state:
        if field == "loai_benh_an": st.session_state[field] = "Nội khoa / Tiền phẫu"
        elif field == "bs_trong_mo":
            st.session_state[field] = (
                "- Hình thức mổ: Mổ phiên / Mổ cấp cứu\n"
                "- Phương pháp mổ: \n"
                "- Phương pháp gây mê: \n"
                "- Quá trình mổ: Không có biến chứng\n"
                "- Chẩn đoán sau mổ: "
            )
        elif field == "tuoi": st.session_state[field] = 45
        elif field == "gioi_tinh": st.session_state[field] = "Nam"
        elif field == "dan_tok": st.session_state[field] = "Kinh"
        elif field == "ngay_vao_vien": st.session_state[field] = datetime.now().strftime("%d/%m/%Y %H:%M")
        elif field in ["sh_can_nang", "sh_chieu_cao"]: st.session_state[field] = 0.0
        elif field == "uu_tien_co_quan": st.session_state[field] = "Không ưu tiên (Thứ tự mặc định)"
        else: st.session_state[field] = ""

for i in range(st.session_state["so_hang_cls"]):
    if f"cls_kq_{i}" not in st.session_state: st.session_state[f"cls_kq_{i}"] = ""
    if f"cls_pg_{i}" not in st.session_state: st.session_state[f"cls_pg_{i}"] = ""


# --- CSS TÙY BIẾN AN TOÀN THEO THEME GỐC CỦA STREAMLIT ---
st.markdown("""
<style>
    /* ==============================================================================
       RETRO OS DESKTOP Y2K (WINDOWS 2000 / MAC OS 9 CLASSIC SYSTEM AESTHETIC)
       Màu xám công nghiệp dịu mắt, hiệu ứng vát viền Bevel 3D kinh điển
       ============================================================================== */

    html {
        scroll-behavior: smooth;
    }

    /* 1. KHỐI CỬA SỔ HỆ THỐNG (WINDOW BOX & DIALOG) */
    div[data-testid="stExpander"] {
        border-top: 2px solid #ffffff !important;
        border-left: 2px solid #ffffff !important;
        border-right: 2px solid #404040 !important;
        border-bottom: 2px solid #404040 !important;
        border-radius: 0px !important;
        margin-bottom: 14px !important;
        background-color: #ece9d8 !important;
        box-shadow: 1px 1px 0px #000000 !important;
        overflow: hidden !important;
    }

    /* Thanh tiêu đề cửa sổ màu xanh Classic Navy đặc trưng */
    div[data-testid="stExpander"] > details > summary {
        background: linear-gradient(90deg, #0a246a 0%, #a6caf0 100%) !important;
        border-left: none !important;
        border-radius: 0px !important;
        padding: 6px 12px 6px 14px !important;
        font-family: 'Tahoma', 'Segoe UI', sans-serif !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
        letter-spacing: 0.5px !important;
        color: #ffffff !important;
        text-shadow: 1px 1px 1px #000000 !important;
        box-shadow: none !important;
    }

    div[data-testid="stExpander"] > details > summary:hover {
        background: linear-gradient(90deg, #1842a8 0%, #b8d5f5 100%) !important;
        color: #ffffff !important;
    }

    div[data-testid="stExpander"] > details > summary svg {
        fill: #ffffff !important;
    }

    /* 2. THANH TIÊU ĐỀ PHÂN MỤC (MENUBAR / STATUS STRIP) */
    .sidebar-header-amboss {
        background: #d4d0c8 !important;
        border-top: 2px solid #ffffff !important;
        border-left: 2px solid #ffffff !important;
        border-right: 2px solid #808080 !important;
        border-bottom: 2px solid #808080 !important;
        box-shadow: 1px 1px 0px #000000 !important;
        padding: 5px 12px !important;
        border-radius: 0px !important;
        font-family: 'Tahoma', sans-serif !important;
        font-size: 0.92rem !important;
        font-weight: 700 !important;
        color: #000000 !important;
        margin-bottom: 10px !important;
    }

    .sub-section-header {
        background: #e4e2d5 !important;
        border-top: 1px solid #ffffff !important;
        border-left: 4px solid #0a246a !important;
        border-right: 1px solid #808080 !important;
        border-bottom: 1px solid #808080 !important;
        padding: 4px 10px !important;
        border-radius: 0px !important;
        margin-top: 10px !important;
        margin-bottom: 8px !important;
        font-family: 'Tahoma', sans-serif !important;
        font-size: 0.88rem !important;
        font-weight: 700 !important;
        color: #102542 !important;
    }

    /* BỘ CHỌN MẪU BỆNH ÁN DẠNG NÚT BEVEL ĐỒNG BỘ THEME */
    div[data-testid="stRadio"] {
        background: #d4d0c8 !important;
        border-top: 2px solid #ffffff !important;
        border-left: 2px solid #ffffff !important;
        border-right: 2px solid #404040 !important;
        border-bottom: 2px solid #404040 !important;
        box-shadow: 1px 1px 0 #000000 !important;
        padding: 7px 9px 8px !important;
        margin: 8px 0 14px !important;
    }

    div[data-testid="stRadio"] > label {
        color: #000000 !important;
        font-family: 'Tahoma', 'Segoe UI', sans-serif !important;
        font-size: 0.9rem !important;
        font-weight: 700 !important;
        margin-bottom: 6px !important;
    }

    div[data-testid="stRadio"] [role="radiogroup"] {
        display: flex !important;
        flex-wrap: wrap !important;
        gap: 5px !important;
    }

    div[data-testid="stRadio"] [role="radio"] {
        min-height: 34px !important;
        padding: 5px 10px !important;
        background: #ece9d8 !important;
        border-top: 2px solid #ffffff !important;
        border-left: 2px solid #ffffff !important;
        border-right: 2px solid #404040 !important;
        border-bottom: 2px solid #404040 !important;
        border-radius: 0 !important;
        color: #000000 !important;
        font-family: 'Tahoma', 'Segoe UI', sans-serif !important;
        font-size: 0.84rem !important;
        font-weight: 700 !important;
        line-height: 1.25 !important;
        cursor: pointer !important;
    }

    div[data-testid="stRadio"] [role="radio"]:hover {
        background: #f5f4ea !important;
        color: #0a246a !important;
    }

    div[data-testid="stRadio"] [role="radio"][aria-checked="true"] {
        background: #0a246a !important;
        border-top-color: #404040 !important;
        border-left-color: #404040 !important;
        border-right-color: #ffffff !important;
        border-bottom-color: #ffffff !important;
        box-shadow: inset 1px 1px 0 #000000 !important;
        color: #ffffff !important;
        text-shadow: 1px 1px 1px #000000 !important;
    }

    div[data-testid="stRadio"] [role="radio"] p,
    div[data-testid="stRadio"] [role="radio"] span {
        color: inherit !important;
        font-family: inherit !important;
        font-size: inherit !important;
        font-weight: inherit !important;
    }

    @media (max-width: 720px) {
        div[data-testid="stRadio"] [role="radio"] {
            flex: 1 1 100% !important;
        }
    }

    .highlight-dx {
        background-color: #fffde6 !important;
        border: 1px solid #b8860b !important;
        color: #8b0000 !important;
        padding: 6px 12px !important;
        border-radius: 0px !important;
        font-weight: 700 !important;
        font-family: 'Tahoma', sans-serif !important;
    }

    .overview-panel {
        background: #ece9d8;
        border-top: 2px solid #ffffff;
        border-left: 2px solid #ffffff;
        border-right: 2px solid #404040;
        border-bottom: 2px solid #404040;
        box-shadow: 1px 1px 0 #000000;
        margin: 0 0 14px 0;
        padding: 0;
        color: #000000;
        font-family: 'Tahoma', 'Segoe UI', sans-serif;
    }

    .overview-panel-title {
        background: linear-gradient(90deg, #0a246a 0%, #a6caf0 100%);
        color: #ffffff;
        font-size: 0.95rem;
        font-weight: 700;
        letter-spacing: 0.5px;
        padding: 7px 12px;
        text-shadow: 1px 1px 1px #000000;
    }

    .overview-panel-body {
        padding: 10px 12px 8px 12px;
        background: #f5f4ea;
    }

    .overview-row {
        border-bottom: 1px solid #c8c5b8;
        line-height: 1.45;
        padding: 6px 4px;
    }

    .overview-row:last-child {
        border-bottom: none;
    }

    .overview-label {
        color: #0a246a;
        font-weight: 700;
    }

    .overview-diagnosis {
        background: #fffde6;
        border: 1px solid #b8860b;
        color: #8b0000;
        font-weight: 700;
        margin-top: 8px;
        padding: 7px 9px;
    }

    .overview-empty {
        color: #404040;
        font-style: italic;
        padding: 10px 4px;
    }

    /* ==============================================================================
       3. Ô NHẬP LIỆU CHÌM (CHUYỂN VIỀN ĐỎ BASEWEB THÀNH ĐEN HOÀN TOÀN)
       ============================================================================== */

    /* 1. KHUNG CHÌM 3D WIN2K NGOÀI CÙNG */
    div[data-testid="stTextInput"] > div,
    div[data-testid="stTextArea"] > div,
    div[data-testid="stNumberInput"] > div,
    div[data-testid="stSelectbox"] > div > div {
        border-top: 2px solid #404040 !important;
        border-left: 2px solid #404040 !important;
        border-right: 2px solid #ffffff !important;
        border-bottom: 2px solid #ffffff !important;
        background-color: #ffffff !important;
        border-radius: 0px !important;
        box-shadow: inset 1px 1px 0px #808080 !important;
        overflow: hidden !important;
    }

    /* 2. CHẶN VÀ ĐỔI TOÀN BỘ MÀU VIỀN ĐỎ NỘI BỘ SANG TRONG SUỐT / ĐEN */
    div[data-baseweb="base-input"],
    div[data-baseweb="input"],
    div[data-baseweb="textarea"] {
        border-color: transparent !important;
        border-radius: 0px !important;
        background-color: transparent !important;
        box-shadow: none !important;
    }

    /* Đổi màu đỏ khi focus sang đen hoặc triệt tiêu */
    div[data-baseweb="base-input"]:focus-within,
    div[data-baseweb="input"]:focus-within,
    div[data-baseweb="textarea"]:focus-within,
    div[data-baseweb="base-input"] > div:focus-within,
    div[data-baseweb="input"] > div:focus-within {
        border-color: transparent !important;
        box-shadow: none !important;
    }

    /* 3. KHI FOCUS: KHUNG NGOÀI CHUYỂN SANG ĐEN ĐẬM ĐỒNG BỘ */
    div[data-testid="stTextInput"] > div:focus-within,
    div[data-testid="stTextArea"] > div:focus-within,
    div[data-testid="stNumberInput"] > div:focus-within,
    div[data-testid="stSelectbox"] > div > div:focus-within {
        border-top: 2px solid #000000 !important;
        border-left: 2px solid #000000 !important;
        border-right: 2px solid #ffffff !important;
        border-bottom: 2px solid #ffffff !important;
        background-color: #ffffff !important;
        box-shadow: inset 1px 1px 0px #000000 !important;
    }

    /* 4. TẤT CẢ CÁC THẺ CON VÀ INPUT THUẦN */
    div[data-testid="stTextInput"] *,
    div[data-testid="stTextArea"] *,
    div[data-testid="stNumberInput"] * {
        border-radius: 0px !important;
    }

    div[data-testid="stTextInput"] input,
    div[data-testid="stTextArea"] textarea,
    div[data-testid="stNumberInput"] input {
        border: none !important;
        border-color: transparent !important;
        box-shadow: none !important;
        outline: none !important;
        color: #000000 !important;
        font-family: 'Tahoma', 'Segoe UI', sans-serif !important;
        font-size: 0.92rem !important;
        padding: 5px 8px !important;
        background-color: transparent !important;
    }

    /* 5. NÚT TĂNG GIẢM (+ / -) CỦA NUMBER INPUT */
    div[data-testid="stNumberInput"] button {
        border-radius: 0px !important;
        border-top: 1px solid #ffffff !important;
        border-left: 1px solid #ffffff !important;
        border-right: 1px solid #404040 !important;
        border-bottom: 1px solid #404040 !important;
        background-color: #ece9d8 !important;
        color: #000000 !important;
    }

    div[data-testid="stNumberInput"] button:active {
        border-top: 1px solid #404040 !important;
        border-left: 1px solid #404040 !important;
        border-right: 1px solid #ffffff !important;
        border-bottom: 1px solid #ffffff !important;
    }

    *:focus, *:focus-visible {
        outline: none !important;
    }

    /* 4. NÚT BẤM CÔNG HỘP CHỮ NHẬT VÁT NỔI 3D (CHUNKY BUTTONS) */
    button[kind="primary"] {
        background: #ece9d8 !important;
        border-top: 2px solid #ffffff !important;
        border-left: 2px solid #ffffff !important;
        border-right: 2px solid #404040 !important;
        border-bottom: 2px solid #404040 !important;
        box-shadow: 1px 1px 0px #000000 !important;
        border-radius: 0px !important;
        color: #000000 !important;
        font-family: 'Tahoma', sans-serif !important;
        font-weight: 700 !important;
        font-size: 0.88rem !important;
        padding: 4px 14px !important;
    }

    /* Hiệu ứng lún nút khi nhấn xuống (Pressed) */
    button[kind="primary"]:active {
        border-top: 2px solid #404040 !important;
        border-left: 2px solid #404040 !important;
        border-right: 2px solid #ffffff !important;
        border-bottom: 2px solid #ffffff !important;
        box-shadow: none !important;
        transform: translate(1px, 1px) !important;
    }

    button[kind="secondary"] {
        background: #ece9d8 !important;
        border-top: 2px solid #ffffff !important;
        border-left: 2px solid #ffffff !important;
        border-right: 2px solid #404040 !important;
        border-bottom: 2px solid #404040 !important;
        border-radius: 0px !important;
        font-family: 'Tahoma', sans-serif !important;
        font-weight: 600 !important;
        color: #000000 !important;
    }

    /* 5. THANH MỤC LỤC BÊN PHẢI DẠNG HỘP CÔNG CỤ SYSTEM TOOLBOX */
    .right-toc-container {
        position: fixed !important;
        top: 75px !important;
        right: 18px !important;
        z-index: 999999 !important;
    }

    .right-toc-trigger {
        background: #ece9d8 !important;
        border-top: 2px solid #ffffff !important;
        border-left: 2px solid #ffffff !important;
        border-right: 2px solid #404040 !important;
        border-bottom: 2px solid #404040 !important;
        box-shadow: 1px 1px 0px #000000 !important;
        color: #000000 !important;
        padding: 5px 12px !important;
        border-radius: 0px !important;
        font-family: 'Tahoma', sans-serif !important;
        font-size: 0.82rem !important;
        font-weight: 700 !important;
        cursor: pointer !important;
        list-style: none !important;
        user-select: none !important;
    }

    .right-toc-trigger:hover {
        background: #f5f4ea !important;
    }

    .right-toc-trigger:active {
        border-top: 2px solid #404040 !important;
        border-left: 2px solid #404040 !important;
        border-right: 2px solid #ffffff !important;
        border-bottom: 2px solid #ffffff !important;
        box-shadow: none !important;
    }

    .right-toc-menu {
        position: absolute !important;
        top: 30px !important;
        right: 0 !important;
        width: 250px !important;
        background: #ece9d8 !important;
        border-top: 2px solid #ffffff !important;
        border-left: 2px solid #ffffff !important;
        border-right: 2px solid #404040 !important;
        border-bottom: 2px solid #404040 !important;
        box-shadow: 3px 3px 0px #000000 !important;
        padding: 4px !important;
        display: flex !important;
        flex-direction: column !important;
        max-height: 75vh !important;
        overflow-y: auto !important;
    }

    .right-toc-header {
        font-family: 'Tahoma', sans-serif !important;
        font-size: 0.72rem !important;
        font-weight: 700 !important;
        background: #0a246a !important;
        color: #ffffff !important;
        padding: 3px 8px !important;
        letter-spacing: 0.5px !important;
        margin-bottom: 4px !important;
    }

    .toc-item {
        color: #000000 !important;
        font-family: 'Tahoma', sans-serif !important;
        font-size: 0.82rem !important;
        padding: 5px 8px !important;
        text-decoration: none !important;
        display: block !important;
        border: 1px solid transparent !important;
    }

    .toc-item:hover {
        background-color: #0a246a !important;
        color: #ffffff !important;
    }

    /* ==============================================================================
       6. KHUNG TẢI FILE (FILE UPLOADER SUNKEN BOX WIN2K)
       ============================================================================== */
    /* Triệt tiêu khung viền ngoài và các góc bo cong mặc định */
    div[data-testid="stFileUploader"],
    div[data-testid="stFileUploader"] * {
        border-radius: 0px !important;
    }

    /* Khung thả file (Dropzone): Nền trắng phẳng, rãnh chìm 3D */
    div[data-testid="stFileUploader"] section,
    div[data-testid="stFileUploader"] [data-testid="stFileUploaderDropzone"] {
        border-top: 2px solid #404040 !important;
        border-left: 2px solid #404040 !important;
        border-right: 2px solid #ffffff !important;
        border-bottom: 2px solid #ffffff !important;
        background-color: #ffffff !important;
        box-shadow: inset 1px 1px 0px #808080 !important;
        padding: 10px 14px !important;
    }

    /* Ép tất cả các thẻ nền bên trong vùng uploader sang trong suốt */
    div[data-testid="stFileUploader"] section > div {
        background-color: transparent !important;
    }

    /* Nút bấm Browse Files / Upload: Hộp nổi 3D vuông vức */
    div[data-testid="stFileUploader"] button {
        border-top: 2px solid #ffffff !important;
        border-left: 2px solid #ffffff !important;
        border-right: 2px solid #404040 !important;
        border-bottom: 2px solid #404040 !important;
        background-color: #ece9d8 !important;
        color: #000000 !important;
        border-radius: 0px !important;
        box-shadow: 1px 1px 0px #000000 !important;
        padding: 3px 12px !important;
        font-family: 'Tahoma', sans-serif !important;
        font-weight: 700 !important;
        font-size: 0.85rem !important;
    }

    /* Hiệu ứng lún nút Upload khi bấm */
    div[data-testid="stFileUploader"] button:active {
        border-top: 2px solid #404040 !important;
        border-left: 2px solid #404040 !important;
        border-right: 2px solid #ffffff !important;
        border-bottom: 2px solid #ffffff !important;
        box-shadow: none !important;
        transform: translate(1px, 1px) !important;
    }
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# HÀM HỖ TRỢ XUẤT FILE & AI CONTEXT
# ==============================================================================
def get_benh_su_text_for_ai():
    if is_postop_mode(st.session_state.get("loai_benh_an", "")):
        return f"- Trước mổ: {st.session_state.get('bs_truoc_mo')}\n- Trong mổ: {st.session_state.get('bs_trong_mo')}\n- Sau mổ: {st.session_state.get('bs_sau_mo')}"
    return st.session_state.get("benh_su")

def is_pediatric_mode(mode):
    return mode == "Nhi khoa"

def is_san_phu_khoa_mode(mode):
    return mode in ["Sản phụ khoa / Tiền phẫu", "Sản phụ khoa / Hậu phẫu"]

def is_postop_mode(mode):
    return mode in ["Hậu phẫu", "Sản phụ khoa / Hậu phẫu"]

def clinical_title(mode):
    if is_pediatric_mode(mode):
        return "BỆNH ÁN NHI KHOA"
    if mode == "Hậu phẫu" or mode == "Sản phụ khoa / Hậu phẫu":
        return "BỆNH ÁN HẬU PHẪU" if mode == "Hậu phẫu" else "BỆNH ÁN SẢN PHỤ KHOA (HẬU PHẪU)"
    if mode == "Sản phụ khoa / Tiền phẫu":
        return "BỆNH ÁN SẢN PHỤ KHOA (TIỀN PHẪU)"
    return "BỆNH ÁN LÂM SÀNG"

def pediatric_history_context():
    history_labels = [
        ("Tiền sử bệnh lý", "ts_benh_ly"),
        ("Tiền sử dinh dưỡng", "ts_dinh_duong"),
        ("Tiền sử sản khoa", "ts_san_khoa_nhi"),
        ("Tiền sử tiêm chủng", "ts_tiem_chung"),
        ("Tiền sử phát triển tâm thần vận động", "ts_phat_trien"),
        ("Tiền sử dịch tễ", "ts_dich_te"),
        ("Tiền sử dị ứng", "ts_di_ung"),
        ("Tiền sử gia đình", "ts_gia_dinh"),
    ]
    return "\n".join(f"{label}: {st.session_state.get(key, '')}" for label, key in history_labels)

def clinical_history_context(mode):
    if is_pediatric_mode(mode):
        return pediatric_history_context()
    return (
        f"Tiền sử nội khoa: {st.session_state.get('ts_noi_khoa')}\n"
        f"Tiền sử ngoại khoa và dị ứng: {st.session_state.get('ts_ngoai_khoa')}\n"
        f"Lối sống và thói quen: {st.session_state.get('ts_loi_song')}\n"
        f"Tiền sử gia đình: {st.session_state.get('ts_gia_dinh')}"
    )

def pediatric_age_group(age):
    try:
        age = int(age or 0)
    except (TypeError, ValueError):
        age = 0
    if age <= 0:
        return "sơ sinh"
    if age <= 5:
        return "trẻ nhỏ 1-5 tuổi"
    if age <= 10:
        return "trẻ học đường 6-10 tuổi"
    if age <= 15:
        return "trẻ vị thành niên 11-15 tuổi"
    return "vị thành niên 16-18 tuổi"

def pediatric_normal_history(age):
    age_group = pediatric_age_group(age)
    nutrition_by_age = {
        "sơ sinh": "Bú mẹ, phản xạ bú tốt, chưa ghi nhận khó khăn nuôi dưỡng hoặc nôn trớ bất thường.",
        "trẻ nhỏ 1-5 tuổi": "Ăn uống phù hợp lứa tuổi, ăn đa dạng, không biếng ăn kéo dài, không nôn hoặc tiêu chảy mạn tính.",
        "trẻ học đường 6-10 tuổi": "Chế độ ăn đa dạng, phù hợp lứa tuổi, phát triển thể chất phù hợp, không ghi nhận rối loạn dinh dưỡng.",
        "trẻ vị thành niên 11-15 tuổi": "Ăn uống đa dạng, phù hợp giai đoạn dậy thì, chưa ghi nhận rối loạn dinh dưỡng hoặc hành vi ăn uống bất thường.",
        "vị thành niên 16-18 tuổi": "Ăn uống đa dạng, phù hợp tuổi và mức độ hoạt động, chưa ghi nhận rối loạn dinh dưỡng hoặc hành vi ăn uống bất thường.",
    }
    development_by_age = {
        "sơ sinh": "Trẻ đáp ứng phù hợp, bú tốt, phản xạ sơ sinh phù hợp tuổi thai, chưa ghi nhận bất thường phát triển.",
        "trẻ nhỏ 1-5 tuổi": "Các mốc vận động, ngôn ngữ, nhận thức và giao tiếp xã hội phù hợp lứa tuổi; chưa ghi nhận thoái lui phát triển.",
        "trẻ học đường 6-10 tuổi": "Học tập, giao tiếp, vận động và tự chăm sóc phù hợp lứa tuổi; chưa ghi nhận khó khăn phát triển.",
        "trẻ vị thành niên 11-15 tuổi": "Phát triển thể chất, tâm lý, học tập và giao tiếp xã hội phù hợp lứa tuổi; chưa ghi nhận bất thường dậy thì.",
        "vị thành niên 16-18 tuổi": "Phát triển thể chất và tâm lý phù hợp lứa tuổi, học tập và giao tiếp xã hội ổn định.",
    }
    birth_history = "Sinh đủ tháng, đẻ thường, cân nặng sơ sinh phù hợp, không ghi nhận ngạt hoặc biến cố chu sinh." if age_group == "sơ sinh" else "Sinh đủ tháng, cân nặng sơ sinh phù hợp, không ghi nhận biến cố sản khoa hoặc bệnh lý chu sinh đáng chú ý."
    return {
        "ts_benh_ly": "Chưa ghi nhận bệnh lý nội khoa, ngoại khoa hoặc nhập viện trước đây đáng chú ý.",
        "ts_dinh_duong": nutrition_by_age[age_group],
        "ts_san_khoa_nhi": birth_history,
        "ts_tiem_chung": "Đã tiêm chủng đầy đủ theo lịch tiêm chủng của lứa tuổi, chưa ghi nhận phản ứng nặng sau tiêm.",
        "ts_phat_trien": development_by_age[age_group],
        "ts_dich_te": "Chưa ghi nhận tiếp xúc nguồn bệnh, ổ dịch, vật nuôi bất thường hoặc yếu tố dịch tễ đặc biệt.",
        "ts_di_ung": "Chưa ghi nhận dị ứng thuốc, thức ăn hoặc các dị nguyên khác.",
        "ts_gia_dinh": "Chưa ghi nhận bệnh di truyền, dị tật bẩm sinh, bệnh mạn tính hoặc bệnh truyền nhiễm đặc biệt trong gia đình.",
    }

def fill_pediatric_normal_history():
    filled_count = 0
    for field_key, normal_text in pediatric_normal_history(st.session_state.get("tuoi", 0)).items():
        if not str(st.session_state.get(field_key, "") or "").strip():
            st.session_state[field_key] = normal_text
            filled_count += 1
    st.session_state["_pediatric_history_fill_count"] = filled_count
def add_symptom_to_field(field_key, symptom_text):
    """Hàm chèn an toàn triệu chứng vào ô text_area mà không gây lỗi session_state"""
    val = str(st.session_state.get(field_key, "")).strip()
    lines = [l.strip() for l in val.split("\n") if l.strip()]
    formatted_sym = f"- {symptom_text}"
    if formatted_sym not in lines:
        lines.append(formatted_sym)
        updated_value = "\n".join(lines)
        st.session_state[field_key] = updated_value
        widget_key = f"_postop_{field_key}"
        if widget_key in st.session_state:
            st.session_state[widget_key] = updated_value

POSTOP_FIELDS = [
    "bs_truoc_mo", "bs_trong_mo", "bs_sau_mo",
    "ngay_hau_phau", "kham_vet_mo", "kham_dan_luu",
]

def initialize_postop_widgets(current_mode):
    mode_changed = st.session_state.get("_postop_last_mode") != current_mode
    for field_key in POSTOP_FIELDS:
        widget_key = f"_postop_{field_key}"
        if mode_changed or widget_key not in st.session_state:
            st.session_state[widget_key] = st.session_state.get(field_key, "")
    st.session_state["_postop_last_mode"] = current_mode

def sync_postop_field(field_key):
    widget_key = f"_postop_{field_key}"
    st.session_state[field_key] = st.session_state.get(widget_key, "")
def format_bullet_points(text):
    if not text or not str(text).strip(): return "Chưa ghi nhận thông tin."
    lines = str(text).strip().split("\n")
    formatted_lines = []
    for line in lines:
        cleaned = line.strip()
        if cleaned:
            if not cleaned.startswith("-") and not cleaned.startswith("*"): formatted_lines.append(f"- {cleaned}")
            else: formatted_lines.append(cleaned)
    return "\n".join(formatted_lines)

AI_PLAIN_LINE_FORMAT = """
QUY TẮC ĐỊNH DẠNG BẮT BUỘC: Mỗi ý trả lời phải là một câu hoàn chỉnh và nằm trên một dòng riêng. Không thêm gạch đầu dòng, số thứ tự, ký hiệu đầu dòng hoặc ký hiệu trang trí trước câu trả lời. Chỉ giữ lại các nhãn cấu trúc được yêu cầu.
"""

def clean_ai_lines(text):
    cleaned_lines = []
    for line in str(text).strip().splitlines():
        cleaned = re.sub(r"^\s*(?:[-*•]|\d+[.)])\s*", "", line).strip()
        if cleaned:
            cleaned_lines.append(cleaned)
    return "\n".join(cleaned_lines)

def format_history(text):
    if not text or not str(text).strip(): return "Chưa ghi nhận bất thường"
    return format_bullet_points(text)

def to_roman(number):
    roman_values = ((10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"))
    result = ""
    for value, symbol in roman_values:
        result += symbol * (number // value)
        number %= value
    return result

def get_section_numbers(data):
    section_numbers = {}
    next_number = 6

    def add_section(key, include=True):
        nonlocal next_number
        if include:
            section_numbers[key] = to_roman(next_number)
            next_number += 1

    add_section("tom_tat")
    add_section("chan_doan_so_bo")
    add_section("chan_doan_phan_biet", bool(str(data.get("chan_doan_phan_biet", "")).strip()))
    add_section("bien_luan", bool(str(data.get("bien_luan", "")).strip()))
    add_section("de_xuat_cls")
    add_section("cls_da_co")
    add_section("chan_doan_xac_dinh")
    add_section("bien_luan_xac_dinh", bool(str(data.get("bien_luan_xac_dinh", "")).strip()))
    add_section("dieu_tri")
    add_section("tien_luong", bool(str(data.get("tien_luong", "")).strip()))
    add_section("tu_van", bool(str(data.get("tu_van", "")).strip()))
    return section_numbers

NORMAL_ORGAN_FINDINGS = {
    "kham_san_phu_khoa": "- Âm hộ sạch, không tổn thương bất thường.\n- Âm đạo không ra máu hay khí hư bất thường.\n- Cổ tử cung không loét, không chảy máu khi chạm.\n- Tử cung không to bất thường, không đau khi di động.\n- Hai phần phụ không sờ thấy khối bất thường, không đau.",
    "kham_tuan_hoan": "- Lồng ngực cân đối, không ổ đập bất thường, không sẹo mổ cũ.\n- Mỏm tim đập ở khoang liên sườn V đường giữa đòn trái, diện đập 1-2 cm.\n- Dấu hiệu Hartzer (-), không có rung miêu.\n- Nhịp tim đều, tần số trùng nhịp mạch.\n- T1, T2 rõ, không nghe thấy tiếng tim bệnh lý (T3, T4, tiếng cọ màng ngoài tim).\n- Không có tiếng thổi bệnh lý ở các ổ van tim.\n- Mạch ngoại vi bắt rõ, đều hai bên.",
    "kham_ho_hap": "- Lồng ngực hai bên cân đối, di động đều theo nhịp thở, không co kéo cơ hô hấp phụ.\n- Khoang liên sườn không giãn rộng, không có tuần hoàn bàng hệ.\n- Rung thanh đều hai bên phế trường.\n- Gõ trong hai bên phổi.\n- Rì rào phế nang êm dịu hai phế trường.\n- Không nghe thấy rale ẩm, rale nổ, rale rít hay rale ngáy.",
    "kham_tieu_hoa": "- Bụng thon đều hai bên, di động theo nhịp thở, không chướng, không tuần hoàn bàng hệ, không sẹo mổ cũ.\n- Bụng mềm, không có điểm đau khu trú, không có phản ứng thành bụng hay cảm ứng phúc mạc.\n- Gan, lách không sờ thấy dưới bờ sườn, chiều cao gan trong giới hạn bình thường.\n- Các điểm đau ngoại khoa (Ruột thừa, Murphy, túi mật) âm tính.\n- Gõ trong toàn bụng, không có diện đục vùng thấp.\n- Tiếng nhu động ruột bình thường, không có tiếng thổi mạch máu bụng.",
    "kham_than_kinh": "- Bệnh nhân tỉnh táo, tiếp xúc tốt, Glasgow 15 điểm.\n- Không có dấu hiệu thần kinh khu trú.\n- Khám 12 đôi dây thần kinh sọ chưa phát hiện bệnh lý.\n- Trương lực cơ bình thường, cơ lực hai bên đều nhau (5/5).\n- Phản xạ gân xương tứ chi bình thường, đối xứng hai bên.\n- Dấu hiệu gáy mềm, Kernig (-), Brudzinski (-), Babinski (-) hai bên.\n- Cảm giác nông và sâu bình thường.",
    "kham_tiet_nieu": "- Hố thắt lưng hai bên cân đối, không sưng đỏ, không gồ cao.\n- Chạm thận (-), Bập bềnh thận (-).\n- Rung thận (-) hai bên.\n- Ấn các điểm niệu quản trên và giữa không đau.\n- Cầu bàng quang (-).",
    "kham_co_xuong_khop": "- Các khớp không sưng, nóng, đỏ, không biến dạng hay lệch trục.\n- Tầm vận động chủ động và thụ động các khớp trong giới hạn bình thường.\n- Không teo cơ, không cứng khớp buổi sáng.\n- Cột sống không gù vẹo, không có điểm đau chói dọc gai sống.",
    "kham_co_quan_khac": "- Răng - Hàm - Mặt, Tai - Mũi - Họng: Chưa phát hiện bất thường.\n- Nội tiết: Tuyến giáp không to, không có dấu hiệu suy hay cường giáp trên lâm sàng."
}
DETAILED_ORGAN_TEMPLATES = {
    "kham_san_phu_khoa": (
        "- NHÌN:\n"
        "  + Bụng dưới cân đối, không chướng, không sẹo mổ cũ bất thường.\n"
        "  + Âm hộ và tầng sinh môn sạch, không loét, không sưng nề hay tổn thương.\n"
        "  + Âm đạo không ra máu, không khí hư bất thường.\n"
        "- SỜ:\n"
        "  + Bụng mềm, không phản ứng thành bụng, không điểm đau khu trú.\n"
        "  + Tử cung không to bất thường, không đau khi di động.\n"
        "  + Hai phần phụ không sờ thấy khối bất thường, không đau.\n"
        "- KHÁM MỎ VỊT / KHÁM ÂM ĐẠO - CỔ TỬ CUNG KHI CÓ CHỈ ĐỊNH:\n"
        "  + Niêm mạc âm đạo hồng, không tổn thương; cổ tử cung không loét, không chảy máu khi chạm.\n"
        "  + Không ghi nhận dịch bất thường hoặc khối bất thường qua thăm khám."
    ),
    "kham_tuan_hoan": (
        "- NHÌN:\n"
        "  + Lồng ngực cân đối, không biến dạng, không sẹo mổ cũ, không tuần hoàn bàng hệ.\n"
        "  + Mỏm tim đập ở khoang liên sườn V đường giữa đòn trái, diện đập 1-2 cm.\n"
        "  + Không có ổ đập bất thường vùng trước tim, mũi ức hay hõm trên ức.\n"
        "- SỜ:\n"
        "  + Mỏm tim đập rõ, không có dấu hiệu nảy thất trái (Apex heave).\n"
        "  + Rung miêu (Thrills) (-), dấu hiệu Hartzer (-) ở mũi ức.\n"
        "  + Phản hồi gan - tĩnh mạch cổ (Hepatojugular reflux) (-).\n"
        "- NGHE:\n"
        "  + Nhịp tim đều, tần số ... chu kỳ/phút, trùng nhịp mạch quay.\n"
        "  + Tiếng T1, T2 nghe rõ, tách đôi sinh lý (nếu có), không nghe tiếng T3, T4, tiếng clack mở van hay click tống máu.\n"
        "  + Không có tiếng thổi tâm thu, tâm trương tại các ổ van (ĐMC, ĐMP, 2 lá, 3 lá).\n"
        "  + Tiếng cọ màng ngoài tim (-).\n"
        "- MẠCH MÁU NGOẠI VI:\n"
        "  + Mạch quay, mạch cảnh, mạch bẹn, khoeo, chày sau và mu chân bắt rõ, đều hai bên.\n"
        "  + Không có tiếng thổi động mạch cảnh, động mạch thận hay động mạch đùi."
    ),
    "kham_ho_hap": (
        "- NHÌN:\n"
        "  + Lồng ngực hai bên cân đối, di động nhịp nhàng theo nhịp thở.\n"
        "  + Không co kéo cơ hô hấp phụ (cơ ức đòn chũm, cơ liên sườn), không rút lõm hõm ức hay hõm trên đòn.\n"
        "  + Khoang liên sườn không giãn rộng, không có tuần hoàn bàng hệ hay sẹo mổ cũ.\n"
        "- SỜ:\n"
        "  + Khí quản nằm ở đường giữa, không lệch trục.\n"
        "  + Độ giãn nở lồng ngực (Chest expansion) đều hai bên.\n"
        "  + Rung thanh (Tactile fremitus) đều khắp hai phế trường, không tăng, không giảm.\n"
        "- GÕ:\n"
        "  + Gõ trong đều khắp hai phế trường từ đỉnh phổi xuống đáy phổi.\n"
        "  + Ranh giới gan - phổi và đáy phổi di động theo nhịp thở bình thường.\n"
        "- NGHE:\n"
        "  + Rì rào phế nang (Vesicular breath sounds) êm dịu hai bên phế trường.\n"
        "  + Tiếng thở thanh - khí quản bình thường, không có tiếng rít thanh quản (Stridor).\n"
        "  + Không có rale bệnh lý (Rale ẩm to/nhỏ hạt, rale nổ, rale rít, rale ngáy, tiếng cọ màng phổi)."
    ),
    "kham_tieu_hoa": (
        "- NHÌN:\n"
        "  + Bụng thon đều hai bên, di động nhịp nhàng theo nhịp thở, không chướng bè, không lõm lòng thuyền.\n"
        "  + Rốn lõm, không lồi, không chảy dịch; không sẹo mổ cũ, không quai ruột nổi hay dấu hiệu rắn bò.\n"
        "  + Không có tuần hoàn bàng hệ (kiểu cửa - chủ hoặc chủ - chủ).\n"
        "- NGHE:\n"
        "  + Nhu động ruột (Bowel sounds) nghe rõ, tần số khoảng 6-10 lần/phút, không nghe âm sắc kim loại/tiếng óc ách.\n"
        "  + Không có tiếng thổi mạch máu ổ bụng (động mạch chủ bụng, động mạch thận hai bên).\n"
        "- GÕ:\n"
        "  + Gõ trong toàn bụng, khoang Traube gõ vang.\n"
        "  + Chiều cao gan trên đường giữa đòn phải khoảng 9-11 cm, ranh giới rõ; lách không to qua gõ đục.\n"
        "  + Gõ đục vùng thấp (Shifting dullness) (-).\n"
        "- SỜ:\n"
        "  + Bụng mềm, ấn không đau, không có điểm đau khu trú.\n"
        "  + Phản ứng thành bụng (Guarding) (-), Cảm ứng phúc mạc (Peritoneal signs) (-), Dấu Blumberg (-).\n"
        "  + Gan không sờ thấy dưới bờ sườn, bờ gan mềm mại, bề mặt nhẵn.\n"
        "  + Lách không sờ thấy dưới bờ sườn trái (Phân độ lách độ 0).\n"
        "  + Các điểm đau ngoại khoa: Điểm MacBurney (-), Dấu hiệu Murphy (-), Điểm Mayo-Robson (-).\n"
        "  + Lỗ bẹn nông, vòng đùi hai bên bình thường, không thấy khối thoát vị."
    ),
    "kham_than_kinh": (
        "- TRI GIÁC & TÂM THẦN:\n"
        "  + Bệnh nhân tỉnh táo, tiếp xúc tốt, định hướng không gian - thời gian - bản thân chính xác, Glasgow 15 điểm (E4V5M6).\n"
        "  + Trí nhớ tức thì, gần và xa bình thường; ngôn ngữ lưu loát, không thất ngôn (Aphasia).\n"
        "- DẤU MÀNG NÃO & DÂY THẦN KINH SỌ (12 ĐÔI):\n"
        "  + Dấu hiệu gáy mềm, Kernig (-), Brudzinski (-).\n"
        "  + Dây I - XII: Thị lực và thị trường sơ bộ tốt, đồng tử hai bên đều (2.5mm), phản xạ ánh sáng (+).\n"
        "  + Vận nhãn bình thường (không lác, không sụp mi); cơ nhai khỏe, cảm giác mặt đối xứng.\n"
        "  + Mặt cân đối, không liệt mặt trung ương hay ngoại biên; thính lực hai bên đều.\n"
        "  + Màn hầu nâng đều, phản xạ nuốt tốt, vận động cơ ức đòn chũm và lưỡi bình thường.\n"
        "- VẬN ĐỘNG & TRƯƠNG LỰC CƠ:\n"
        "  + Cơ lực hai tay và hai chân đối xứng: 5/5 điểm toàn bộ.\n"
        "  + Trương lực cơ (độ co duỗi, độ chắc, độ ve vẩy) bình thường, không có co cứng tháp hay ngoại tháp.\n"
        "  + Dấu hiệu Babinski (-) hai bên, không có giật cơ (Clonus).\n"
        "- PHẢN XẠ GÂN XƯƠNG (DTR):\n"
        "  + Nhị đầu, tam đầu, gân gối, gân gót đều (2+) ở cả hai bên.\n"
        "- CẢM GIÁC & TIỀU NÃO:\n"
        "  + Cảm giác nông (đau, nhiệt, chạm nhẹ) và cảm giác sâu (vị thế khớp, rung âm thoa) bình thường.\n"
        "  + Nghiệm pháp ngón tay - chỉ mũi, gót - đầu gối chính xác; dấu Romberg (-)."
    ),
    "kham_tiet_nieu": (
        "- NHÌN:\n"
        "  + Hố thắt lưng hai bên phẳng, cân đối, không sưng nề, không bầm tím hay sẹo mổ cũ.\n"
        "  + Vùng hạ vị phẳng, không gồ cao, không thấy khối u hay cầu bàng quang nổi.\n"
        "- SỜ & GÕ:\n"
        "  + Chạm thận (-) hai bên, Bập bềnh thận (-) hai bên.\n"
        "  + Nghiệm pháp rung thận (Giordano) (-) hai bên.\n"
        "  + Ấn các điểm niệu quản trên (cạnh rốn) và điểm niệu quản giữa (đường nối 2 gai chậu trước trên) không đau.\n"
        "  + Điểm sườn - lưng, sườn - cột sống không có điểm đau chói.\n"
        "  + Cầu bàng quang (-), ấn vùng hạ vị không tức, gõ không đục.\n"
        "- NGHE:\n"
        "  + Không có tiếng thổi tâm thu động mạch thận hai bên ở thành bụng trước và sau lưng."
    ),
    "kham_co_xuong_khop": (
        "- NHÌN & TƯ THẾ:\n"
        "  + Dáng đi tự nhiên, trục chi thẳng, không khập khiễng, không biến dạng lệch trục chi.\n"
        "  + Các khớp ngoại vi (vai, khuỷu, cổ tay, bàn ngón tay, háng, gối, cổ chân) không sưng, không nóng, đỏ, không biến dạng hay teo cơ lân cận.\n"
        "  + Cột sống trục thẳng, còn đường cong sinh lý, không gù vẹo.\n"
        "- SỜ & VẬN ĐỘNG:\n"
        "  + Nhiệt độ da quanh khớp bình thường, ấn không có điểm đau chói quanh khớp hay dọc gai sống.\n"
        "  + Tầm vận động chủ động và thụ động (Active & Passive ROM) của tất cả các khớp trong giới hạn bình thường.\n"
        "  + Dấu hiệu bập bềnh xương bánh chè (-), dấu chạm xương bánh chè (-).\n"
        "  + Nghiệm pháp Lasegue (-) hai bên; không có dấu hiệu cứng khớp buổi sáng."
    ),
    "kham_co_quan_khac": (
        "- TAI - MŨI - HỌNG:\n"
        "  + Màng nhĩ hai bên sáng bóng, nón sáng rõ; họng sạch, niêm mạc hồng, Amidan không sưng đỏ, không hốc mủ.\n"
        "- RĂNG - HÀM - MẶT:\n"
        "  + Khớp cắn đúng, không lệch trục; niêm mạc miệng, nướu răng không viêm loét, không lung lay răng.\n"
        "- NỘI TIẾT:\n"
        "  + Tuyến giáp không to (độ 0), mật độ mềm đều, sờ không rung miêu, nghe không có tiếng thổi.\n"
        "  + Không có biểu hiện lồi mắt, run đầu chi hay các triệu chứng chuyển hóa lâm sàng."
    )
}

PEDIATRIC_NORMAL_ORGAN_FINDINGS = {
    "kham_tuan_hoan": "- Lồng ngực cân đối, không biến dạng, không co kéo, không có ổ đập bất thường.\n- Mỏm tim ở vị trí phù hợp theo tuổi, không có rung miêu.\n- Nhịp tim đều, tần số phù hợp tuổi; T1, T2 rõ, không nghe tiếng thổi bệnh lý.\n- Mạch ngoại vi rõ, đều hai bên, đầu chi ấm, thời gian làm đầy mao mạch dưới 2 giây.",
    "kham_ho_hap": "- Trẻ tỉnh, tự thở, lồng ngực cân đối, di động đều, không rút lõm lồng ngực hay phập phồng cánh mũi.\n- Không tím tái, không thở rên hoặc thở khò khè.\n- Rì rào phế nang rõ hai bên, không nghe ran bệnh lý.",
    "kham_tieu_hoa": "- Bụng mềm, không chướng, rốn sạch, không quai ruột nổi.\n- Gan có thể sờ dưới bờ sườn phải tùy tuổi; lách không to bất thường.\n- Không phản ứng thành bụng, nhu động ruột trong giới hạn bình thường.",
    "kham_than_kinh": "- Trẻ tỉnh, đáp ứng phù hợp lứa tuổi; đánh giá Glasgow hoặc AVPU khi cần.\n- Không dấu màng não, không yếu liệt khu trú, trương lực và phản xạ phù hợp tuổi.\n- Với trẻ nhỏ: thóp trước phẳng, vòng đầu và tương tác phù hợp lứa tuổi nếu có chỉ định.",
    "kham_tiet_nieu": "- Vùng hông lưng không sưng nề, không đau khi khám.\n- Không sờ thấy thận to bất thường, không cầu bàng quang.\n- Lượng nước tiểu và tình trạng tiểu tiện phù hợp tuổi.",
    "kham_co_xuong_khop": "- Tư thế và vận động phù hợp lứa tuổi, không biến dạng chi hay sưng nóng đỏ khớp.\n- Trương lực, cơ lực và tầm vận động phù hợp tuổi; không đau dọc cột sống.",
    "kham_co_quan_khac": "- Tai - mũi - họng: niêm mạc hồng, không xuất tiết bất thường.\n- Răng miệng và da niêm mạc chưa phát hiện bất thường.\n- Không ghi nhận hạch ngoại vi hoặc dấu hiệu nội tiết bất thường.",
}

PEDIATRIC_DETAILED_ORGAN_TEMPLATES = {
    "kham_tuan_hoan": "- NHÌN: Lồng ngực, sắc da, đầu chi, dấu suy tim và tím tái.\n- SỜ: Mỏm tim theo tuổi, rung miêu, mạch ngoại vi, thời gian làm đầy mao mạch.\n- NGHE: Tần số và nhịp tim theo tuổi, T1/T2, tiếng thổi và tiếng tim bất thường.",
    "kham_ho_hap": "- NHÌN: Tần số thở theo tuổi, kiểu thở, co kéo cơ hô hấp phụ, phập phồng cánh mũi, tím tái.\n- SỜ/GÕ: Độ giãn nở lồng ngực, rung thanh và vùng gõ bất thường khi phù hợp lứa tuổi.\n- NGHE: Rì rào phế nang, ran ẩm, ran rít, ran ngáy, tiếng thở rít.",
    "kham_tieu_hoa": "- NHÌN: Bụng chướng, rốn, tuần hoàn bàng hệ, quai ruột nổi, tình trạng dinh dưỡng.\n- SỜ/GÕ: Độ mềm bụng, phản ứng thành bụng, gan lách theo tuổi, khối bất thường.\n- NGHE: Nhu động ruột và tiếng thổi mạch máu khi có chỉ định.",
    "kham_than_kinh": "- TRI GIÁC: Đáp ứng với người chăm sóc, AVPU/Glasgow theo tuổi.\n- TRẺ NHỎ: Thóp, vòng đầu, giao tiếp mắt, trương lực và phản xạ nguyên thủy theo tuổi.\n- TRẺ LỚN: Dấu màng não, dây thần kinh sọ, cơ lực, phản xạ, cảm giác và phối hợp động tác.",
    "kham_tiet_nieu": "- NHÌN: Phù, màu da, vùng hông lưng và hạ vị.\n- SỜ/GÕ: Đau vùng thận, thận to, cầu bàng quang, điểm đau niệu quản khi phù hợp tuổi.\n- GHI NHẬN: Lượng nước tiểu, số lần tiểu và bất thường đường tiểu.",
    "kham_co_xuong_khop": "- ĐÁNH GIÁ: Tư thế, dáng đi và vận động theo lứa tuổi.\n- KHÁM: Sưng, nóng, đỏ, đau khớp; tầm vận động; cơ lực, trương lực và dấu hiệu viêm cơ.\n- TRẺ NHỎ: Khả năng lẫy, ngồi, đứng, đi và vận động đối xứng.",
    "kham_co_quan_khac": "- TAI - MŨI - HỌNG: Tai, mũi, họng, amidan và hạch cổ.\n- RĂNG MIỆNG: Niêm mạc, răng, lợi và tổn thương miệng.\n- DA - NIÊM MẠC: Ban, xuất huyết, vàng da, xanh tái, dấu mất nước và hạch ngoại vi.",
}

def organ_findings_for_mode(mode):
    return PEDIATRIC_NORMAL_ORGAN_FINDINGS if is_pediatric_mode(mode) else NORMAL_ORGAN_FINDINGS

def detailed_organ_templates_for_mode(mode):
    return PEDIATRIC_DETAILED_ORGAN_TEMPLATES if is_pediatric_mode(mode) else DETAILED_ORGAN_TEMPLATES

class BenhAnPDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.add_font("Roboto", "", "Roboto-Regular.ttf")
        self.add_font("Roboto-Bold", "", "Roboto-Bold.ttf")
        self.loai_ba = "Nội khoa / Tiền phẫu"

    def header(self):
        if self.page_no() == 1:
            self.set_font("Roboto-Bold", "", 15)
            title = clinical_title(self.loai_ba)
            self.cell(0, 8, title, align="C", new_x="LMARGIN", new_y="NEXT")
            self.set_font("Roboto", "", 8)
            self.cell(0, 4, f"Thời gian làm bệnh án: {datetime.now().strftime('%d/%m/%Y %H:%M')}", align="C", new_x="LMARGIN", new_y="NEXT")
            self.ln(3)

    def footer(self):
        self.set_y(-12)
        self.set_font("Roboto", "", 8)
        self.cell(0, 10, f"Trang {self.page_no()}/{{nb}}", align="C")

    def add_section_header(self, title):
        self.set_font("Roboto-Bold", "", 11)
        self.set_fill_color(225, 235, 245)
        self.cell(0, 7, title, fill=True, new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def add_subsection_header(self, title):
        self.set_font("Roboto-Bold", "", 10)
        self.cell(0, 6, title, new_x="LMARGIN", new_y="NEXT")

    def add_body_text(self, text):
        self.set_font("Roboto", "", 9.5)
        self.set_text_color(0, 0, 0)
        self.multi_cell(0, 5, str(text).strip() if str(text).strip() else "Chưa ghi nhận thông tin.")
        self.ln(2)

    def add_highlight_text(self, text):
        self.set_font("Roboto-Bold", "", 10.5)
        self.set_text_color(180, 0, 0)
        self.multi_cell(0, 5.5, str(text).strip() if str(text).strip() else "Chưa ghi nhận thông tin.")
        self.set_text_color(0, 0, 0)
        self.ln(2)

    def render_tom_tat_pdf(self, text):
        if not text or not str(text).strip():
            self.add_body_text("Chưa ghi nhận thông tin.")
            return
        lines = [line.strip() for line in str(text).strip().split("\n") if line.strip()]
        if not lines: return
        self.set_font("Roboto", "", 9.5)
        self.set_text_color(0, 0, 0)
        self.multi_cell(0, 5, lines[0])
        self.ln(1)
        orig_l_margin = self.l_margin
        self.set_left_margin(orig_l_margin + 8.0)
        for line in lines[1:]:
            bullet_line = line if (line.startswith("-") or line.startswith("*")) else f"- {line}"
            self.multi_cell(0, 5, bullet_line)
            self.ln(1)
        self.set_left_margin(orig_l_margin)
        self.ln(2)

    def render_table_cls(self, cls_rows):
        col_w = 95
        line_h = 5.0
        self.set_font("Roboto-Bold", "", 9.5)
        self.set_fill_color(230, 235, 245)
        if self.get_y() > 260: self.add_page()
        self.cell(col_w, 7, "KẾT QUẢ CẬN LÂM SÀNG", border=1, align="C", fill=True)
        self.cell(col_w, 7, "PHIÊN GIẢI / BIỆN GIẢI", border=1, align="C", fill=True, new_x="LMARGIN", new_y="NEXT")
        
        temp_files = []
        try:
            self.set_font("Roboto", "", 9)
            for kq, pg, img in cls_rows:
                txt_kq = format_bullet_points(kq) if kq else ""
                txt_pg = format_bullet_points(pg) if pg else "-"
                temp_img_path = None
                img_h = 0
                if img:
                    suffix = os.path.splitext(img.name)[1]
                    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as t_img:
                        t_img.write(img.getbuffer())
                        temp_img_path = t_img.name
                        temp_files.append(temp_img_path)
                    img_h = 45
                nb_lines_left = len(self.multi_cell(col_w - 4, line_h, txt_kq, dry_run=True, output="LINES"))
                nb_lines_right = len(self.multi_cell(col_w - 4, line_h, txt_pg, dry_run=True, output="LINES"))
                
                h_left = nb_lines_left * line_h + 4 + (img_h + 4 if img else 0)
                h_right = nb_lines_right * line_h + 4
                row_h = max(h_left, h_right, 8)
                
                if self.get_y() + row_h > 275:
                    self.add_page()
                    self.set_font("Roboto-Bold", "", 9.5)
                    self.set_fill_color(230, 235, 245)
                    self.cell(col_w, 7, "KẾT QUẢ CẬN LÂM SÀNG", border=1, align="C", fill=True)
                    self.cell(col_w, 7, "PHIÊN GIẢI / BIỆN GIẢI", border=1, align="C", fill=True, new_x="LMARGIN", new_y="NEXT")
                    self.set_font("Roboto", "", 9)

                curr_x = self.get_x()
                curr_y = self.get_y()
                self.rect(curr_x, curr_y, col_w, row_h)
                self.rect(curr_x + col_w, curr_y, col_w, row_h)
                
                self.set_xy(curr_x + 2, curr_y + 2)
                if txt_kq: self.multi_cell(col_w - 4, line_h, txt_kq)
                if temp_img_path:
                    y_img = self.get_y() + 1
                    img_w = min(col_w - 8, 70)
                    x_img = curr_x + (col_w - img_w) / 2
                    try: self.image(temp_img_path, x=x_img, y=y_img, w=img_w, h=img_h)
                    except: pass
                
                self.set_xy(curr_x + col_w + 2, curr_y + 2)
                self.multi_cell(col_w - 4, line_h, txt_pg)
                self.set_xy(curr_x, curr_y + row_h)
            self.ln(3)
        finally:
            for p in temp_files:
                if os.path.exists(p):
                    try: os.remove(p)
                    except: pass

def export_pdf(data):
    pdf = BenhAnPDF()
    pdf.loai_ba = data.get("loai_benh_an", "Nội khoa / Tiền phẫu")
    pdf.alias_nb_pages()
    pdf.add_page()
    
    # I. HÀNH CHÍNH
    pdf.add_section_header("I. PHẦN HÀNH CHÍNH")
    hc_text = (
        f"- Họ và tên: {str(data['ho_ten']).upper()}    |    Tuổi: {data['tuoi']}    |    Giới tính: {data['gioi_tinh']}\n"
        f"- Dân tộc: {data['dan_tok']}    |    Nghề nghiệp: {data['nghe_nghiep']}\n"
        f"- Khoa phòng: {data['khoa_phong']}\n"
        f"- Địa chỉ: {data['dia_chi']}\n"
        f"- Ngày giờ vào viện: {data['ngay_vao_vien']}\n"
        f"- Sinh viên thực hiện: {data['sinh_vien']}"
    )
    pdf.add_body_text(hc_text)

    # II. LÝ DO VÀO VIỆN
    pdf.add_section_header("II. LÝ DO VÀO VIỆN")
    pdf.add_body_text(data['ly_do_vao_vien'])

    if is_san_phu_khoa_mode(pdf.loai_ba):
        pdf.add_section_header("III. TIỀN SỬ")
        pdf.add_subsection_header("1. Tiền sử sản khoa:")
        pdf.add_body_text(format_history(data.get('ts_san_khoa', '')))
        pdf.add_subsection_header("2. Tiền sử phụ khoa:")
        pdf.add_body_text(format_history(data.get('ts_phu_khoa', '')))
        pdf.add_subsection_header("3. Tiền sử nội - ngoại khoa:")
        pdf.add_body_text(format_history(data.get('ts_noi_ngoai_khoa', '')))
        pdf.add_subsection_header("4. Tiền sử gia đình:")
        pdf.add_body_text(format_history(data.get('ts_gia_dinh', '')))

    # III hoặc IV. BỆNH SỬ
    pdf.add_section_header("IV. BỆNH SỬ" if is_san_phu_khoa_mode(pdf.loai_ba) else "III. BỆNH SỬ")
    if is_postop_mode(pdf.loai_ba):
        pdf.add_subsection_header("1. Tình trạng trước mổ:")
        pdf.add_body_text(format_bullet_points(data.get('bs_truoc_mo', '')))
        pdf.add_subsection_header("2. Tình trạng trong mổ:")
        pdf.add_body_text(format_bullet_points(data.get('bs_trong_mo', '')))
        pdf.add_subsection_header("3. Quá trình sau mổ:")
        pdf.add_body_text(format_bullet_points(data.get('bs_sau_mo', '')))
    else:
        pdf.add_body_text(data.get('benh_su', ''))

    if is_pediatric_mode(pdf.loai_ba):
        pdf.add_section_header("IV. TIỀN SỬ")
        pediatric_history = [
            ("1. Tiền sử bệnh lý:", "ts_benh_ly"),
            ("2. Tiền sử dinh dưỡng:", "ts_dinh_duong"),
            ("3. Tiền sử sản khoa:", "ts_san_khoa_nhi"),
            ("4. Tiền sử tiêm chủng:", "ts_tiem_chung"),
            ("5. Tiền sử phát triển tâm thần vận động:", "ts_phat_trien"),
            ("6. Tiền sử dịch tễ:", "ts_dich_te"),
            ("7. Tiền sử dị ứng:", "ts_di_ung"),
            ("8. Tiền sử gia đình:", "ts_gia_dinh"),
        ]
        for title, key in pediatric_history:
            pdf.add_subsection_header(title)
            pdf.add_body_text(format_history(data.get(key, '')))
    elif not is_san_phu_khoa_mode(pdf.loai_ba):
        pdf.add_section_header("IV. TIỀN SỬ")
        pdf.add_subsection_header("1. Tiền sử nội khoa:")
        pdf.add_body_text(format_history(data.get('ts_noi_khoa', '')))
        pdf.add_subsection_header("2. Tiền sử ngoại khoa và dị ứng:")
        pdf.add_body_text(format_history(data.get('ts_ngoai_khoa', '')))
        pdf.add_subsection_header("3. Lối sống và thói quen:")
        pdf.add_body_text(format_history(data.get('ts_loi_song', '')))
        pdf.add_subsection_header("4. Tiền sử gia đình:")
        pdf.add_body_text(format_history(data.get('ts_gia_dinh', '')))

    # V. THĂM KHÁM LÂM SÀNG
    pdf.add_section_header("V. THĂM KHÁM LÂM SÀNG")
    
    if is_postop_mode(pdf.loai_ba):
        pdf.add_subsection_header("1. Thăm khám hiện tại:")
        pdf.add_highlight_text(f"Hậu phẫu: {data.get('ngay_hau_phau', '...')}")
        pdf.add_body_text("a. Toàn thân:")
    else:
        pdf.add_subsection_header("1. Thăm khám lúc vào viện:")
        pdf.add_body_text(format_bullet_points(data.get('kham_vao_vien', '')))
        pdf.add_subsection_header("2. Thăm khám hiện tại:")
        pdf.add_body_text("a. Toàn thân:")

    pdf.add_body_text(format_bullet_points(data.get('kham_toan_than', '')))
    
    # BẢNG SINH HIỆU
    mach_val = data.get('sh_mach') or "--"
    nhiet_val = data.get('sh_nhiet_do') or "--"
    ha_val = data.get('sh_ha') or "--"
    nt_val = data.get('sh_nhip_tho') or "--"
    cn_val = data.get('sh_can_nang') if float(data.get('sh_can_nang', 0)) > 0 else "--"
    cc_val = data.get('sh_chieu_cao') if float(data.get('sh_chieu_cao', 0)) > 0 else "--"
    bmi_num = data.get('sh_bmi', '')
    bmi_txt = data.get('sh_bmi_eval', '')

    pdf.ln(1)
    pdf.set_draw_color(180, 180, 180)
    pdf.set_fill_color(245, 247, 250)
    col_w4 = (pdf.w - pdf.l_margin - pdf.r_margin) / 4.0
    pdf.set_font("Roboto-Bold", size=8.5)
    pdf.cell(col_w4, 5.5, f"Mạch: {mach_val} ck/phút", border=1, fill=True)
    pdf.cell(col_w4, 5.5, f"Nhiệt độ: {nhiet_val} °C", border=1, fill=True)
    pdf.cell(col_w4, 5.5, f"Huyết áp: {ha_val} mmHg", border=1, fill=True)
    pdf.cell(col_w4, 5.5, f"Nhịp thở: {nt_val} l/phút", border=1, fill=True, ln=True)
    
    col_w3_1 = col_w4
    col_w3_2 = col_w4
    col_w3_3 = col_w4 * 2.0
    bmi_display = f"BMI: {bmi_num} kg/m² ({bmi_txt})" if bmi_num else "BMI: --"
    pdf.cell(col_w3_1, 5.5, f"Chiều cao: {cc_val} cm", border=1)
    pdf.cell(col_w3_2, 5.5, f"Cân nặng: {cn_val} kg", border=1)
    pdf.cell(col_w3_3, 5.5, bmi_display, border=1, ln=True)
    pdf.ln(2)
    
    if is_postop_mode(pdf.loai_ba):
        pdf.add_body_text("b. Tình trạng vết mổ và dẫn lưu:")
        pdf.add_subsection_header("Vết mổ:")
        pdf.add_body_text(format_bullet_points(data.get('kham_vet_mo', '')))
        pdf.add_subsection_header("Ống dẫn lưu:")
        pdf.add_body_text(format_bullet_points(data.get('kham_dan_luu', '')))
        pdf.add_body_text("c. Các cơ quan:")
    else:
        pdf.add_body_text("b. Các cơ quan:")
        
    organ_list = ([{"key": "kham_san_phu_khoa", "name": "Sản phụ khoa"}] if is_san_phu_khoa_mode(pdf.loai_ba) else []) + [
        {"key": "kham_tuan_hoan", "name": "Tuần hoàn"},
        {"key": "kham_ho_hap", "name": "Hô hấp"},
        {"key": "kham_tieu_hoa", "name": "Tiêu hóa"},
        {"key": "kham_than_kinh", "name": "Thần kinh"},
        {"key": "kham_tiet_nieu", "name": "Thận - Tiết niệu"},
        {"key": "kham_co_xuong_khop", "name": "Cơ xương khớp"},
        {"key": "kham_co_quan_khac", "name": "Các cơ quan khác"},
    ]
    selected_organ = data.get("uu_tien_co_quan", "Không ưu tiên (Thứ tự mặc định)")
    if selected_organ != "Không ưu tiên (Thứ tự mặc định)":
        fav = next((it for it in organ_list if it["name"] == selected_organ), None)
        others = [it for it in organ_list if it["name"] != selected_organ]
        render_list = ([fav] + others) if fav else organ_list
    else:
        render_list = organ_list

    for org in render_list:
        content = format_bullet_points(data.get(org["key"], ""))
        title_text = f"{org['name']}:"
        pdf.add_subsection_header(title_text)
        text_w = pdf.get_string_width(title_text)
        x = pdf.get_x()
        y = pdf.get_y() - 1
        pdf.set_draw_color(50, 50, 50)
        pdf.set_line_width(0.3)
        pdf.line(x, y, x + text_w, y)
        pdf.add_body_text(content)

    section_numbers = get_section_numbers(data)
    num_tt = section_numbers["tom_tat"]
    num_cdsb = section_numbers["chan_doan_so_bo"]
    num_dxcls = section_numbers["de_xuat_cls"]
    num_cls = section_numbers["cls_da_co"]
    num_cdxd = section_numbers["chan_doan_xac_dinh"]

    def pdf_tt():
        pdf.add_section_header(f"{num_tt}. TÓM TẮT BỆNH ÁN")
        pdf.render_tom_tat_pdf(data.get('tom_tat', ''))

    def pdf_cdsb():
        pdf.add_section_header(f"{num_cdsb}. CHẨN ĐOÁN SƠ BỘ")
        pdf.add_body_text(data.get('chan_doan_so_bo', ''))
        if "chan_doan_phan_biet" in section_numbers:
            pdf.add_section_header(f"{section_numbers['chan_doan_phan_biet']}. CHẨN ĐOÁN PHÂN BIỆT")
            pdf.add_body_text(data.get('chan_doan_phan_biet', ''))
        bl = str(data.get('bien_luan', '')).strip()
        if bl:
            pdf.add_section_header(f"{section_numbers['bien_luan']}. BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ")
            pdf.add_body_text(bl)

    def pdf_cls():
        pdf.add_section_header(f"{num_dxcls}. ĐỀ XUẤT CẬN LÂM SÀNG")
        nhan_cls1 = "1. Phát hiện biến chứng / Đánh giá sau mổ:" if is_postop_mode(pdf.loai_ba) else "1. Phục vụ chẩn đoán xác định:"
        pdf.add_subsection_header(nhan_cls1)
        pdf.add_body_text(format_bullet_points(data.get('cls_dx_xac_dinh', '')))
        nhan_cls2 = "2. Theo dõi hồi phục & Điều trị:" if is_postop_mode(pdf.loai_ba) else "2. Phục vụ điều trị:"
        pdf.add_subsection_header(nhan_cls2)
        pdf.add_body_text(format_bullet_points(data.get('cls_dx_dieu_tri', '')))
        pdf.add_subsection_header("3. Cận lâm sàng khác:")
        pdf.add_body_text(format_bullet_points(data.get('cls_dx_khac', '')))

        pdf.add_section_header(f"{num_cls}. CẬN LÂM SÀNG ĐÃ CÓ")
        cls_rows = []
        so_hang = data.get("so_hang_cls", 3)
        for i in range(so_hang):
            kq = data.get(f"cls_kq_{i}", "").strip()
            pg = data.get(f"cls_pg_{i}", "").strip()
            img = data.get(f"cls_img_{i}", None)
            if kq or pg or img: cls_rows.append((kq, pg, img))
        if not cls_rows: pdf.add_body_text("Chưa ghi nhận kết quả cận lâm sàng.")
        else: pdf.render_table_cls(cls_rows)

    def pdf_cdxd():
        pdf.add_section_header(f"{num_cdxd}. CHẨN ĐOÁN XÁC ĐỊNH")
        pdf.add_highlight_text(format_bullet_points(data.get('chan_doan_xac_dinh', '')))
        noi_dung_bl_xd = str(data.get('bien_luan_xac_dinh', '')).strip()
        if noi_dung_bl_xd:
            pdf.add_section_header(f"{section_numbers['bien_luan_xac_dinh']}. BIỆN LUẬN CHẨN ĐOÁN XÁC ĐỊNH")
            pdf.add_body_text(format_bullet_points(noi_dung_bl_xd))

    # THỨ TỰ THỐNG NHẤT CHO CẢ TIỀN PHẪU VÀ HẬU PHẪU
    pdf_tt()
    pdf_cdsb()
    pdf_cls()
    pdf_cdxd()

    pdf.add_section_header(f"{section_numbers['dieu_tri']}. ĐIỀU TRỊ")
    pdf.add_subsection_header("1. Mục tiêu điều trị:")
    pdf.add_body_text(format_bullet_points(data.get('dt_muc_tieu', '')))
    pdf.add_subsection_header("2. Điều trị cụ thể:")
    pdf.add_body_text(format_bullet_points(data.get('dt_cu_the', '')))
    pdf.add_subsection_header("3. Theo dõi sau điều trị:")
    pdf.add_body_text(format_bullet_points(data.get('dt_theo_doi', '')))

    noi_dung_tien_luong = str(data.get("tien_luong", "")).strip()
    if noi_dung_tien_luong:
        pdf.add_section_header(f"{section_numbers['tien_luong']}. TIÊN LƯỢNG")
        pdf.add_body_text(format_bullet_points(noi_dung_tien_luong))

    noi_dung_tu_van = str(data.get("tu_van", "")).strip()
    if noi_dung_tu_van:
        ten_de_muc_tu_van = f"{section_numbers['tu_van']}. TƯ VẤN"
        pdf.add_section_header(ten_de_muc_tu_van)
        pdf.add_body_text(format_bullet_points(noi_dung_tu_van))

    return bytes(pdf.output())

def set_cell_background(cell, fill_hex):
    """Tô màu nền cho ô trong bảng docx."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Căn lề đệm bên trong ô."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def export_docx(data):
    doc = Document()
    
    # Thiết lập lề trang 2cm tiêu chuẩn văn bản y khoa
    for sec in doc.sections:
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
        sec.left_margin = Inches(0.8)
        sec.right_margin = Inches(0.8)

    loai_ba = data.get("loai_benh_an", "Nội khoa / Tiền phẫu")

    # Hàm trợ giúp thêm đoạn văn bản có định dạng
    def add_sec_title(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.bold = True
        run.font.size = Pt(11.5)
        run.font.color.rgb = RGBColor(10, 36, 106)  # Classic Navy Blue

    def add_subsec_title(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.bold = True
        run.font.size = Pt(10.5)
        run.font.color.rgb = RGBColor(30, 41, 59)

    def add_bullet_list(text):
        if not text or not str(text).strip():
            p = doc.add_paragraph("Chưa ghi nhận thông tin.")
            p.runs[0].font.size = Pt(10)
            p.runs[0].font.italic = True
            p.paragraph_format.space_after = Pt(3)
            return
        lines = [l.strip() for l in str(text).strip().split("\n") if l.strip()]
        for line in lines:
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(2)
            clean_text = line.lstrip("-*• ")
            run = p.add_run(clean_text)
            run.font.size = Pt(10)

    def add_normal_text(text, bold=False, red=False):
        if not text or not str(text).strip():
            p = doc.add_paragraph("Chưa ghi nhận thông tin.")
            p.runs[0].font.size = Pt(10)
            p.runs[0].font.italic = True
            p.paragraph_format.space_after = Pt(3)
            return
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(str(text).strip())
        run.font.size = Pt(10)
        run.bold = bold
        if red:
            run.font.color.rgb = RGBColor(180, 0, 0)

    # --- TIÊU ĐỀ TRANG ---
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    run_title = p_title.add_run(clinical_title(loai_ba))
    run_title.bold = True
    run_title.font.size = Pt(16)
    run_title.font.color.rgb = RGBColor(10, 36, 106)

    p_time = doc.add_paragraph()
    p_time.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_time.paragraph_format.space_after = Pt(10)
    run_time = p_time.add_run(f"Thời gian lập hồ sơ: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    run_time.font.size = Pt(9)
    run_time.font.italic = True

    # I. HÀNH CHÍNH
    add_sec_title("I. PHẦN HÀNH CHÍNH")
    hc_lines = [
        f"Họ và tên: {str(data.get('ho_ten', '')).upper()}    |    Tuổi: {data.get('tuoi', '')}    |    Giới tính: {data.get('gioi_tinh', '')}",
        f"Dân tộc: {data.get('dan_tok', '')}    |    Nghề nghiệp: {data.get('nghe_nghiep', '')}",
        f"Khoa / Phòng: {data.get('khoa_phong', '')}",
        f"Địa chỉ: {data.get('dia_chi', '')}",
        f"Ngày giờ vào viện: {data.get('ngay_vao_vien', '')}",
        f"Bác sĩ / Sinh viên thực hiện: {data.get('sinh_vien', '')}"
    ]
    for line in hc_lines:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(line)
        run.font.size = Pt(10)

    # II. LÝ DO VÀO VIỆN
    add_sec_title("II. LÝ DO VÀO VIỆN")
    add_normal_text(data.get('ly_do_vao_vien', ''))

    if is_san_phu_khoa_mode(loai_ba):
        add_sec_title("III. TIỀN SỬ")
        add_subsec_title("1. Tiền sử sản khoa:")
        add_bullet_list(format_history(data.get('ts_san_khoa', '')))
        add_subsec_title("2. Tiền sử phụ khoa:")
        add_bullet_list(format_history(data.get('ts_phu_khoa', '')))
        add_subsec_title("3. Tiền sử nội - ngoại khoa:")
        add_bullet_list(format_history(data.get('ts_noi_ngoai_khoa', '')))
        add_subsec_title("4. Tiền sử gia đình:")
        add_bullet_list(format_history(data.get('ts_gia_dinh', '')))

    # III hoặc IV. BỆNH SỬ
    add_sec_title("IV. BỆNH SỬ" if is_san_phu_khoa_mode(loai_ba) else "III. BỆNH SỬ")
    if is_postop_mode(loai_ba):
        add_subsec_title("1. Tình trạng trước mổ:")
        add_bullet_list(data.get('bs_truoc_mo', ''))
        add_subsec_title("2. Tình trạng trong mổ:")
        add_bullet_list(data.get('bs_trong_mo', ''))
        add_subsec_title("3. Quá trình sau mổ:")
        add_bullet_list(data.get('bs_sau_mo', ''))
    else:
        add_normal_text(data.get('benh_su', ''))

    if is_pediatric_mode(loai_ba):
        add_sec_title("IV. TIỀN SỬ")
        pediatric_history = [
            ("1. Tiền sử bệnh lý:", "ts_benh_ly"),
            ("2. Tiền sử dinh dưỡng:", "ts_dinh_duong"),
            ("3. Tiền sử sản khoa:", "ts_san_khoa_nhi"),
            ("4. Tiền sử tiêm chủng:", "ts_tiem_chung"),
            ("5. Tiền sử phát triển tâm thần vận động:", "ts_phat_trien"),
            ("6. Tiền sử dịch tễ:", "ts_dich_te"),
            ("7. Tiền sử dị ứng:", "ts_di_ung"),
            ("8. Tiền sử gia đình:", "ts_gia_dinh"),
        ]
        for title, key in pediatric_history:
            add_subsec_title(title)
            add_bullet_list(format_history(data.get(key, '')))
    elif not is_san_phu_khoa_mode(loai_ba):
        add_sec_title("IV. TIỀN SỬ")
        add_subsec_title("1. Tiền sử nội khoa:")
        add_bullet_list(format_history(data.get('ts_noi_khoa', '')))
        add_subsec_title("2. Tiền sử ngoại khoa & dị ứng:")
        add_bullet_list(format_history(data.get('ts_ngoai_khoa', '')))
        add_subsec_title("3. Lối sống & thói quen:")
        add_bullet_list(format_history(data.get('ts_loi_song', '')))
        add_subsec_title("4. Tiền sử gia đình:")
        add_bullet_list(format_history(data.get('ts_gia_dinh', '')))

    # V. THĂM KHÁM LÂM SÀNG
    add_sec_title("V. THĂM KHÁM LÂM SÀNG")
    if is_postop_mode(loai_ba):
        add_subsec_title("1. Thăm khám hiện tại:")
        p_hp = doc.add_paragraph()
        run_hp = p_hp.add_run(f"Hậu phẫu: {data.get('ngay_hau_phau', '...')}")
        run_hp.bold = True
        run_hp.font.size = Pt(10)
        run_hp.font.color.rgb = RGBColor(180, 0, 0)
        add_subsec_title("a. Toàn thân:")
    else:
        add_subsec_title("1. Thăm khám lúc vào viện:")
        add_bullet_list(data.get('kham_vao_vien', ''))
        add_subsec_title("2. Thăm khám hiện tại:")
        add_subsec_title("a. Toàn thân:")

    add_bullet_list(data.get('kham_toan_than', ''))

    # Bảng sinh hiệu (Vital signs)
    mach_val = data.get('sh_mach') or "--"
    nhiet_val = data.get('sh_nhiet_do') or "--"
    ha_val = data.get('sh_ha') or "--"
    nt_val = data.get('sh_nhip_tho') or "--"
    cn_val = data.get('sh_can_nang') if float(data.get('sh_can_nang', 0)) > 0 else "--"
    cc_val = data.get('sh_chieu_cao') if float(data.get('sh_chieu_cao', 0)) > 0 else "--"
    bmi_num = data.get('sh_bmi', '')
    bmi_txt = data.get('sh_bmi_eval', '')
    bmi_display = f"BMI: {bmi_num} kg/m² ({bmi_txt})" if bmi_num else "BMI: --"

    tbl_sh = doc.add_table(rows=2, cols=4)
    tbl_sh.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_sh.autofit = False

    sh_data = [
        [f"Mạch: {mach_val} ck/p", f"Nhiệt độ: {nhiet_val} °C", f"Huyết áp: {ha_val} mmHg", f"Nhịp thở: {nt_val} l/p"],
        [f"Chiều cao: {cc_val} cm", f"Cân nặng: {cn_val} kg", bmi_display, ""]
    ]

    for r_idx, row in enumerate(tbl_sh.rows):
        for c_idx, cell in enumerate(row.cells):
            cell.text = sh_data[r_idx][c_idx]
            p = cell.paragraphs[0]
            p.runs[0].font.size = Pt(9)
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
            if r_idx == 0:
                set_cell_background(cell, "EAECEF")
                p.runs[0].bold = True

    # Hợp nhất 2 ô cuối của hàng 2 cho phần hiển thị BMI
    tbl_sh.cell(1, 2).merge(tbl_sh.cell(1, 3))

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    if is_postop_mode(loai_ba):
        add_subsec_title("b. Vết mổ & Dẫn lưu:")
        p_vm = doc.add_paragraph(style='List Bullet')
        p_vm.add_run("Vết mổ: ").bold = True
        p_vm.add_run(str(data.get('kham_vet_mo', 'Chưa ghi nhận.')))
        p_vm.runs[0].font.size = Pt(10)
        p_vm.runs[1].font.size = Pt(10)

        p_dl = doc.add_paragraph(style='List Bullet')
        p_dl.add_run("Ống dẫn lưu: ").bold = True
        p_dl.add_run(str(data.get('kham_dan_luu', 'Chưa ghi nhận.')))
        p_dl.runs[0].font.size = Pt(10)
        p_dl.runs[1].font.size = Pt(10)
        add_subsec_title("c. Khám các cơ quan:")
    else:
        add_subsec_title("b. Khám các cơ quan:")

    organ_list = ([{"key": "kham_san_phu_khoa", "name": "Sản phụ khoa"}] if is_san_phu_khoa_mode(loai_ba) else []) + [
        {"key": "kham_tuan_hoan", "name": "Tuần hoàn"},
        {"key": "kham_ho_hap", "name": "Hô hấp"},
        {"key": "kham_tieu_hoa", "name": "Tiêu hóa"},
        {"key": "kham_than_kinh", "name": "Thần kinh"},
        {"key": "kham_tiet_nieu", "name": "Thận - Tiết niệu"},
        {"key": "kham_co_xuong_khop", "name": "Cơ xương khớp"},
        {"key": "kham_co_quan_khac", "name": "Các cơ quan khác"},
    ]
    selected_organ = data.get("uu_tien_co_quan", "Không ưu tiên (Thứ tự mặc định)")
    if selected_organ != "Không ưu tiên (Thứ tự mặc định)":
        fav = next((it for it in organ_list if it["name"] == selected_organ), None)
        others = [it for it in organ_list if it["name"] != selected_organ]
        render_list = ([fav] + others) if fav else organ_list
    else:
        render_list = organ_list

    for org in render_list:
        add_subsec_title(f"{org['name']}:")
        add_bullet_list(data.get(org["key"], ""))

    # VI. TÓM TẮT BỆNH ÁN
    add_sec_title("VI. TÓM TẮT BỆNH ÁN")
    lines_tt = [l.strip() for l in str(data.get("tom_tat", "")).split("\n") if l.strip()]
    if lines_tt:
        p_first = doc.add_paragraph(lines_tt[0])
        p_first.runs[0].font.size = Pt(10)
        p_first.paragraph_format.space_after = Pt(2)
        for line in lines_tt[1:]:
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(line.lstrip("-*• "))
            run.font.size = Pt(10)
    else:
        add_normal_text("Chưa ghi nhận thông tin.")

    section_numbers = get_section_numbers(data)

    # VII & VIII & IX. CHẨN ĐOÁN SƠ BỘ, PHÂN BIỆT & BIỆN LUẬN
    add_sec_title(f"{section_numbers['chan_doan_so_bo']}. CHẨN ĐOÁN SƠ BỘ")
    add_normal_text(data.get('chan_doan_so_bo', ''))

    if "chan_doan_phan_biet" in section_numbers:
        add_sec_title(f"{section_numbers['chan_doan_phan_biet']}. CHẨN ĐOÁN PHÂN BIỆT")
        add_normal_text(data.get('chan_doan_phan_biet', ''))

    bl_sb = str(data.get('bien_luan', '')).strip()
    if bl_sb:
        add_sec_title(f"{section_numbers['bien_luan']}. BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ")
        add_bullet_list(bl_sb)

    # X. ĐỀ XUẤT CẬN LÂM SÀNG
    add_sec_title(f"{section_numbers['de_xuat_cls']}. ĐỀ XUẤT CẬN LÂM SÀNG")
    nhan_cls1 = "1. Phát hiện biến chứng / Đánh giá sau mổ:" if is_postop_mode(loai_ba) else "1. Phục vụ chẩn đoán xác định:"
    add_subsec_title(nhan_cls1)
    add_bullet_list(data.get('cls_dx_xac_dinh', ''))
    nhan_cls2 = "2. Theo dõi hồi phục & Điều trị:" if is_postop_mode(loai_ba) else "2. Phục vụ điều trị:"
    add_subsec_title(nhan_cls2)
    add_bullet_list(data.get('cls_dx_dieu_tri', ''))
    add_subsec_title("3. Cận lâm sàng khác:")
    add_bullet_list(data.get('cls_dx_khac', ''))

    # XI. CẬN LÂM SÀNG ĐÃ CÓ (BẢNG WORD KÈM ẢNH NẾU CÓ)
    add_sec_title(f"{section_numbers['cls_da_co']}. CẬN LÂM SÀNG ĐÃ CÓ")
    cls_rows = []
    so_hang = data.get("so_hang_cls", 3)
    for i in range(so_hang):
        kq = data.get(f"cls_kq_{i}", "").strip()
        pg = data.get(f"cls_pg_{i}", "").strip()
        img = data.get(f"cls_img_{i}", None)
        if kq or pg or img:
            cls_rows.append((kq, pg, img))

    if not cls_rows:
        add_normal_text("Chưa ghi nhận kết quả cận lâm sàng.")
    else:
        table_cls = doc.add_table(rows=1, cols=2)
        table_cls.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_cells = table_cls.rows[0].cells
        hdr_cells[0].text = "KẾT QUẢ CẬN LÂM SÀNG"
        hdr_cells[1].text = "PHIÊN GIẢI / BIỆN GIẢI"
        for c in hdr_cells:
            set_cell_background(c, "E1EBF5")
            p = c.paragraphs[0]
            p.runs[0].bold = True
            p.runs[0].font.size = Pt(10)
            set_cell_margins(c, top=100, bottom=100, left=150, right=150)

        temp_docx_imgs = []
        try:
            for kq, pg, img in cls_rows:
                row_cells = table_cls.add_row().cells
                set_cell_margins(row_cells[0], top=80, bottom=80, left=120, right=120)
                set_cell_margins(row_cells[1], top=80, bottom=80, left=120, right=120)
                
                # Cột Kết quả
                p_kq = row_cells[0].paragraphs[0]
                p_kq.text = kq if kq else "-"
                p_kq.runs[0].font.size = Pt(9.5)
                
                if img:
                    suffix = os.path.splitext(img.name)[1]
                    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as t_f:
                        t_f.write(img.getbuffer())
                        t_path = t_f.name
                        temp_docx_imgs.append(t_path)
                    p_img = row_cells[0].add_paragraph()
                    p_img.add_run().add_picture(t_path, width=Inches(2.5))
                
                # Cột Biện giải
                p_pg = row_cells[1].paragraphs[0]
                p_pg.text = pg if pg else "-"
                p_pg.runs[0].font.size = Pt(9.5)
        finally:
            for p_path in temp_docx_imgs:
                if os.path.exists(p_path):
                    try:
                        os.remove(p_path)
                    except Exception:
                        pass

    # XII. CHẨN ĐOÁN XÁC ĐỊNH
    add_sec_title(f"{section_numbers['chan_doan_xac_dinh']}. CHẨN ĐOÁN XÁC ĐỊNH")
    add_normal_text(data.get('chan_doan_xac_dinh', ''), bold=True, red=True)

    # XIII. BIỆN LUẬN CHẨN ĐOÁN XÁC ĐỊNH
    noi_dung_bl_xd = str(data.get('bien_luan_xac_dinh', '')).strip()
    if noi_dung_bl_xd:
        add_sec_title(f"{section_numbers['bien_luan_xac_dinh']}. BIỆN LUẬN CHẨN ĐOÁN XÁC ĐỊNH")
        add_bullet_list(noi_dung_bl_xd)

    # XIV. ĐIỀU TRỊ
    add_sec_title(f"{section_numbers['dieu_tri']}. ĐIỀU TRỊ")
    add_subsec_title("1. Mục tiêu điều trị:")
    add_bullet_list(data.get('dt_muc_tieu', ''))
    add_subsec_title("2. Điều trị cụ thể:")
    add_bullet_list(data.get('dt_cu_the', ''))
    add_subsec_title("3. Theo dõi sau điều trị:")
    add_bullet_list(data.get('dt_theo_doi', ''))

    # XV. TIÊN LƯỢNG & XVI. TƯ VẤN
    tl_str = str(data.get("tien_luong", "")).strip()
    if tl_str:
        add_sec_title(f"{section_numbers['tien_luong']}. TIÊN LƯỢNG")
        add_bullet_list(tl_str)

    tv_str = str(data.get("tu_van", "")).strip()
    if tv_str:
        ten_de_muc_tv = f"{section_numbers['tu_van']}. TƯ VẤN"
        add_sec_title(ten_de_muc_tv)
        add_bullet_list(tv_str)

    docx_io = io.BytesIO()
    doc.save(docx_io)
    docx_io.seek(0)
    return docx_io.getvalue()

# ==============================================================================
# SIDEBAR: QUẢN LÝ BẢN NHÁP & ĐĂNG XUẤT
# ==============================================================================
with st.sidebar:
    st.markdown("<div class='sidebar-header-amboss'>Quản lý bản nháp</div>", unsafe_allow_html=True)
    st.caption("🟢 **Tự động lưu:** Dữ liệu được ghi nhớ tự động vào trình duyệt mỗi khi nhập liệu.")
    
    if st.button("🔄 Nạp lại bản nháp từ trình duyệt", type="primary", use_container_width=True):
        saved_raw = local_storage.getItem(STORAGE_KEY)
        if saved_raw:
            try:
                loaded_ls = json.loads(saved_raw) if isinstance(saved_raw, str) else saved_raw
                load_draft_to_session(loaded_ls)
                st.toast("Đã khôi phục bệnh án thành công!", icon="✅")
                st.rerun()
            except Exception as e: st.error(f"Lỗi khi đọc bản nháp: {e}")
        else: st.warning("Không tìm thấy dữ liệu nháp nào.")

    if st.button("🗑️ Xóa bản nháp (Làm bệnh án mới)", use_container_width=True):
        # 1. Xóa an toàn chống KeyError từ streamlit_local_storage
        try:
            local_storage.deleteItem(STORAGE_KEY)
        except Exception:
            pass

        # 2. Xóa snapshot đệm để tránh ghi đè dữ liệu rác
        st.session_state["last_saved_snapshot"] = ""

        # Mẫu khung sườn cố định cho phần Trong mổ
        mau_trong_mo = (
            "- Hình thức mổ: Mổ phiên / Mổ cấp cứu\n"
            "- Phương pháp mổ: \n"
            "- Phương pháp gây mê: \n"
            "- Quá trình mổ: Không có biến chứng\n"
            "- Chẩn đoán sau mổ: "
        )

        # 3. Đặt lại tất cả các trường dữ liệu
        for k in FIELDS_TO_SAVE:
            if k == "tuoi":
                st.session_state[k] = 45
            elif k in ["sh_can_nang", "sh_chieu_cao"]:
                st.session_state[k] = 0.0
            elif k == "gioi_tinh":
                st.session_state[k] = "Nam"
            elif k == "dan_tok":
                st.session_state[k] = "Kinh"
            elif k == "loai_benh_an":
                st.session_state[k] = "Nội khoa / Tiền phẫu"
            elif k == "bs_trong_mo":
                st.session_state[k] = mau_trong_mo
            elif k == "uu_tien_co_quan":
                st.session_state[k] = "Không ưu tiên (Thứ tự mặc định)"
            elif k == "ngay_vao_vien":
                st.session_state[k] = datetime.now().strftime("%d/%m/%Y %H:%M")
            else:
                st.session_state[k] = ""

        # 4. Đặt lại số hàng cận lâm sàng
        st.session_state["so_hang_cls"] = 1
        for i in range(15):
            st.session_state[f"cls_kq_{i}"] = ""
            st.session_state[f"cls_pg_{i}"] = ""

        st.toast("Đã xóa sạch bản nháp và làm mới form!", icon="🗑️")
        st.rerun()

    st.markdown("---")
    st.caption("Hoặc lưu trữ dạng tập tin JSON tải về máy:")
    current_data = {k: st.session_state.get(k, "") for k in FIELDS_TO_SAVE}
    for i in range(current_data.get("so_hang_cls", 3)):
        current_data[f"cls_kq_{i}"] = st.session_state.get(f"cls_kq_{i}", "")
        current_data[f"cls_pg_{i}"] = st.session_state.get(f"cls_pg_{i}", "")
        
    json_string = json.dumps(current_data, ensure_ascii=False, indent=2)
    ten_benh_nhan = str(st.session_state.get("ho_ten", "chua_dat_ten")).strip().replace(" ", "_")
    if not ten_benh_nhan: ten_benh_nhan = "chua_dat_ten"
        
    st.download_button("📥 Lưu bản nháp về máy (.json)", data=json_string, file_name=f"Ban_nhap_{ten_benh_nhan}_{datetime.now().strftime('%Y%m%d_%H%M')}.json", mime="application/json", use_container_width=True)
    st.markdown("---")
    st.markdown("**Khôi phục dữ liệu từ bản nháp:**")
    file_nhap = st.file_uploader("Chọn tập tin .json đã lưu:", type=["json"], key="uploader_nhap_json")
    if file_nhap is not None:
        if st.button("🔄 Nhấn vào đây để nạp dữ liệu", type="primary", use_container_width=True):
            try:
                loaded_data = json.load(file_nhap)
                load_draft_to_session(loaded_data)
                st.success("Đã nạp bản nháp thành công!")
                st.rerun()
            except Exception as e: st.error(f"Không thể đọc file: {e}")

    if st.session_state.get("password_correct") and not st.session_state.get("is_admin"):
        st.markdown("---")
        if st.button("🚪 Đăng xuất khỏi thiết bị này"):
            local_storage.deleteItem(AUTH_STORAGE_KEY)
            st.session_state["password_correct"] = False
            st.session_state.pop("logged_in_user", None)
            st.rerun()

# ==============================================================================
# GIAO DIỆN CHÍNH (3 TABS)
# ==============================================================================
st.title("Bệnh Án Lâm Sàng")
st.caption("Cấu trúc bệnh án trình bày ca bệnh và thi lâm sàng (Hỗ trợ Nội khoa, Ngoại khoa, Hậu phẫu, Sản phụ khoa, Nhi khoa).")

loai_benh_an = st.selectbox(
    "**LỰA CHỌN MẪU BỆNH ÁN:**",
    ["Nội khoa / Tiền phẫu", "Nhi khoa", "Hậu phẫu", "Sản phụ khoa / Tiền phẫu", "Sản phụ khoa / Hậu phẫu"],
    key="loai_benh_an",
)
initialize_postop_widgets(loai_benh_an)

# Khai báo Dictionary lưu trữ ảnh toàn cục
uploaded_imgs = {}


# CÁC HÀM UI RỜI RẠC DÙNG CHUNG (Để hoán đổi vị trí)
# --- HÀM TỰ ĐỘNG TẠO CÂU DẪN TÓM TẮT BỆNH ÁN NỘI KHOA (KHÔNG DÙNG AI) ---
def generate_intro_tom_tat_noi_khoa():
    gioi_tinh = st.session_state.get("gioi_tinh", "Nam")
    tuoi = st.session_state.get("tuoi", "")
    tuoi_str = f"{tuoi} tuổi" if tuoi else ""
    
    ts_list = []
    history_keys = ["ts_benh_ly", "ts_dinh_duong", "ts_san_khoa_nhi", "ts_tiem_chung", "ts_phat_trien", "ts_dich_te", "ts_di_ung", "ts_gia_dinh"] if is_pediatric_mode(loai_benh_an) else ["ts_noi_khoa", "ts_ngoai_khoa"]
    for k in history_keys:
        val = str(st.session_state.get(k, "")).strip()
        if val:
            lines = [l.strip().lstrip("-*• ") for l in val.split("\n") if l.strip()]
            if lines:
                ts_list.append(", ".join(lines))
    tien_su_str = "; ".join(ts_list) if ts_list else "chưa ghi nhận bất thường"

    ly_do = str(st.session_state.get("ly_do_vao_vien", "")).strip() or "..."

    ngay_vv_raw = str(st.session_state.get("ngay_vao_vien", "")).strip()
    so_ngay_nam_vien = 0
    if ngay_vv_raw:
        try:
            date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", ngay_vv_raw)
            if date_match:
                d_vv = datetime.strptime(date_match.group(1), "%d/%m/%Y").date()
                d_hientai = datetime.now().date()
                so_ngay_nam_vien = max(0, (d_hientai - d_vv).days)
        except Exception:
            so_ngay_nam_vien = 0

    benh_su = str(st.session_state.get("benh_su", "")).lower()
    so_ngay_truoc_vv = 0
    match_ngay = re.search(r"cách (?:vào viện|nhập viện)\s*(?:khoảng|tầm)?\s*(\d+)\s*ngày", benh_su)
    match_gio = re.search(r"cách (?:vào viện|nhập viện)\s*(?:khoảng|tầm)?\s*(\d+)\s*giờ", benh_su)
    match_tuan = re.search(r"cách (?:vào viện|nhập viện)\s*(?:khoảng|tầm)?\s*(\d+)\s*tuần", benh_su)

    if match_ngay:
        so_ngay_truoc_vv = int(match_ngay.group(1))
    elif match_gio:
        so_ngay_truoc_vv = max(1, round(int(match_gio.group(1)) / 24))
    elif match_tuan:
        so_ngay_truoc_vv = int(match_tuan.group(1)) * 7

    tong_ngay = so_ngay_truoc_vv + so_ngay_nam_vien
    if tong_ngay > 0:
        dien_bien_str = f"bệnh diễn biến {tong_ngay} ngày nay"
    else:
        dien_bien_str = "bệnh diễn biến cấp tính"

    return (
        f"Bệnh nhân {gioi_tinh.lower()} {tuoi_str}, tiền sử {tien_su_str} vào viện vì {ly_do}, "
        f"{dien_bien_str}. Qua thăm khám và hỏi bệnh phát hiện các hội chứng và triệu chứng sau:"
    )

# --- HÀM TỰ ĐỘNG TẠO CÂU DẪN TÓM TẮT BỆNH ÁN HẬU PHẪU (KHÔNG DÙNG AI) ---
# --- HÀM TỰ ĐỘNG TẠO CÂU DẪN TÓM TẮT BỆNH ÁN HẬU PHẪU (KHÔNG DÙNG AI) ---
def generate_intro_tom_tat_hau_phau():
    gioi_tinh = st.session_state.get("gioi_tinh", "Nam")
    tuoi = st.session_state.get("tuoi", "")
    tuoi_str = f"{tuoi} tuổi" if tuoi else "..."

    # 1. Trích xuất tiền sử (Nội khoa & Ngoại khoa)
    ts_list = []
    for k in ["ts_noi_khoa", "ts_ngoai_khoa"]:
        val = str(st.session_state.get(k, "")).strip()
        if val:
            lines = [l.strip().lstrip("-*• ") for l in val.split("\n") if l.strip()]
            if lines:
                ts_list.append(", ".join(lines))
    tien_su_str = "; ".join(ts_list) if ts_list else "chưa ghi nhận bất thường"

    # 2. Lý do vào viện
    ly_do = str(st.session_state.get("ly_do_vao_vien", "")).strip() or "..."

    bs_truoc_mo = str(st.session_state.get("bs_truoc_mo", "")).strip()
    bs_trong_mo = str(st.session_state.get("bs_trong_mo", "")).strip()
    cd_so_bo = str(st.session_state.get("chan_doan_so_bo", "")).strip()

    # 3. Trích xuất Chẩn đoán trước mổ
    cd_truoc_mo = "..."
    m_cdtm = re.search(r"(?:chẩn đoán trước mổ|cđ trước mổ)(?:\s*là)?[:\s\-]+([^.\n;]+)", bs_truoc_mo, re.IGNORECASE)
    if m_cdtm and m_cdtm.group(1).strip():
        cd_truoc_mo = m_cdtm.group(1).strip()
    elif bs_truoc_mo:
        cd_truoc_mo = bs_truoc_mo.split("\n")[-1].strip().lstrip("-*• ")

    # 4. Trích xuất các trường từ form mẫu 5 dòng trong mổ
    pp_mo = "..."
    loai_mo = "cấp cứu/phiên"
    cd_sau_mo = "..."
    dien_bien_phau_thuat = "Quá trình mổ không có biến chứng."

    if bs_trong_mo:
        # Hình thức mổ: Mổ phiên hay Mổ cấp cứu
        m_ht = re.search(r"(?:hình thức mổ|mổ phiên/mổ cấp cứu|mổ phiên hay mổ cấp cứu)[:\s\-]+([^\n]+)", bs_trong_mo, re.IGNORECASE)
        if m_ht:
            txt_ht = m_ht.group(1).lower().strip()
            if "cấp cứu" in txt_ht and "phiên" not in txt_ht:
                loai_mo = "cấp cứu"
            elif "phiên" in txt_ht or "chương trình" in txt_ht:
                loai_mo = "chương trình"

        # Phương pháp mổ
        m_pp = re.search(r"(?:phương pháp mổ|phương pháp phẫu thuật|pt)[:\s\-]+([^\n]+)", bs_trong_mo, re.IGNORECASE)
        if m_pp and m_pp.group(1).strip():
            pp_mo = m_pp.group(1).strip()

        # Quá trình mổ
        m_qtm = re.search(r"(?:quá trình mổ|diễn biến mổ)[:\s\-]+([^\n]+)", bs_trong_mo, re.IGNORECASE)
        if m_qtm and m_qtm.group(1).strip():
            val_qtm = m_qtm.group(1).strip().rstrip(".")
            dien_bien_phau_thuat = f"Quá trình mổ {val_qtm[0].lower() + val_qtm[1:] if val_qtm.lower().startswith('không') or val_qtm.lower().startswith('có') or val_qtm.lower().startswith('thuận') else val_qtm}."

        # Chẩn đoán sau mổ
        m_sm = re.search(r"(?:chẩn đoán sau mổ|cđ sau mổ)[:\s\-]+([^\n]+)", bs_trong_mo, re.IGNORECASE)
        if m_sm and m_sm.group(1).strip():
            cd_sau_mo = m_sm.group(1).strip()

    # Dự phòng chẩn đoán sau mổ từ chẩn đoán sơ bộ
    if cd_sau_mo == "..." and cd_so_bo:
        cd_clean = re.sub(r"^hậu phẫu ngày[^-\:]*[-\:]\s*", "", cd_so_bo, flags=re.IGNORECASE)
        cd_sau_mo = cd_clean.split("-")[0].strip() or cd_so_bo

    # 5. Trích xuất Ngày hậu phẫu
    ngay_hp_val = str(st.session_state.get("ngay_hau_phau", "")).strip()
    m_hp = re.search(r"(?:ngày\s*(?:thứ)?\s*)(\d+)", ngay_hp_val, re.IGNORECASE)
    if m_hp:
        ngay_hp_str = f"ngày thứ {m_hp.group(1)}"
    elif ngay_hp_val:
        ngay_hp_str = ngay_hp_val
    else:
        ngay_hp_str = "ngày thứ ..."

    # Đoạn dẫn chính
    cau_dan_chinh = (
        f"Bệnh nhân {gioi_tinh.lower()} {tuoi_str}, tiền sử {tien_su_str} vào viện vì {ly_do}, "
        f"chẩn đoán trước mổ là {cd_truoc_mo}, được mổ bằng phương pháp {pp_mo}, "
        f"mổ {loai_mo}, chẩn đoán sau mổ là {cd_sau_mo}. "
        f"{dien_bien_phau_thuat} Hiện tại hậu phẫu {ngay_hp_str}. "
        f"Qua thăm khám và hỏi bệnh phát hiện các hội chứng và triệu chứng sau:"
    )

    # 6. Trích xuất Sinh hiệu để tạo dòng 1
    sh_items = []
    if str(st.session_state.get("sh_mach", "")).strip():
        sh_items.append(f"Mạch {st.session_state.get('sh_mach')} lần/phút")
    if str(st.session_state.get("sh_ha", "")).strip():
        sh_items.append(f"HA {st.session_state.get('sh_ha')} mmHg")
    if str(st.session_state.get("sh_nhiet_do", "")).strip():
        sh_items.append(f"Nhiệt độ {st.session_state.get('sh_nhiet_do')} °C")
    if str(st.session_state.get("sh_nhip_tho", "")).strip():
        sh_items.append(f"Nhịp thở {st.session_state.get('sh_nhip_tho')} lần/phút")
    sh_str = ", ".join(sh_items) if sh_items else "Mạch, HA, Nhiệt độ trong giới hạn bình thường"
    line_sh = f"- Tỉnh, tiếp xúc tốt, Sinh hiệu: {sh_str}"

    # 7. Trích xuất Khám vết mổ để tạo dòng 2
    raw_vm = str(st.session_state.get("kham_vet_mo", "")).strip()
    if raw_vm:
        vm_clean = "; ".join([l.strip().lstrip("-*• ") for l in raw_vm.split("\n") if l.strip()])
    else:
        vm_clean = "khô, không sưng đỏ, chân chỉ sạch"
    line_vm = f"- Vết mổ: {vm_clean}"

    # 8. Trích xuất Khám dẫn lưu để tạo dòng 3
    raw_dl = str(st.session_state.get("kham_dan_luu", "")).strip()
    if raw_dl:
        dl_clean = "; ".join([l.strip().lstrip("-*• ") for l in raw_dl.split("\n") if l.strip()])
    else:
        dl_clean = "chân dẫn lưu sạch, không rỉ dịch bất thường"
    line_dl = f"- Dẫn lưu: {dl_clean}"

    # Ghép câu dẫn hoàn chỉnh kèm 3 dòng xuống hàng
    return f"{cau_dan_chinh}\n{line_sh}\n{line_vm}\n{line_dl}"

def ui_tom_tat(num):
    col_tt_title, col_tt_btn = st.columns([1.5, 0.5])
    with col_tt_title:
        st.markdown(f"**{num}. Tóm tắt bệnh án:**")
    with col_tt_btn:
        # Tự động chọn câu dẫn phù hợp với loại bệnh án
        help_text = "Tự động trích xuất thông tin phẫu thuật và điền câu dẫn" if is_postop_mode(loai_benh_an) else "Tự động tính ngày và điền câu dẫn mở đầu"
        if st.button("⚡ Tạo câu dẫn", key="btn_auto_cau_dan_tt", help=help_text, use_container_width=True):
            if is_postop_mode(loai_benh_an):
                cau_dan_moi = generate_intro_tom_tat_hau_phau()
            else:
                cau_dan_moi = generate_intro_tom_tat_noi_khoa()

            current_tt = str(st.session_state.get("tom_tat", "")).strip()
            if current_tt:
                lines = current_tt.split("\n")
                if len(lines) > 1 and (lines[1].strip().startswith("-") or lines[1].strip().startswith("*")):
                    st.session_state["tom_tat"] = cau_dan_moi + "\n" + "\n".join(lines[1:])
                else:
                    st.session_state["tom_tat"] = cau_dan_moi + "\n" + current_tt
            else:
                st.session_state["tom_tat"] = cau_dan_moi

    st.text_area(
        f"{num}. Tóm tắt bệnh án:", 
        key="tom_tat", 
        height=130, 
        label_visibility="collapsed",
        placeholder="Dòng đầu tiên là câu dẫn tóm tắt. Các dòng tiếp theo ghi các hội chứng và triệu chứng có giá trị..."
    )

def ui_cdsb(num_sb, num_pb, num_bl):
    # Hàng 1: Tiêu đề + Nút bấm căn ngang hàng
    col_h_left, col_h_right_title, col_h_right_btn = st.columns([1, 0.65, 0.35])
    with col_h_left:
        st.markdown(f"**{num_sb}. Chẩn đoán sơ bộ:**")
    with col_h_right_title:
        st.markdown(f"**{num_pb}. Chẩn đoán phân biệt:**")
    with col_h_right_btn:
        btn_ai_cdpb = st.button("🪄 Làm phép", key="btn_ai_cdpb", type="primary", use_container_width=True)

    # Xử lý logic AI khi bấm nút
    if btn_ai_cdpb:
        if "GEMINI_API_KEY" not in st.secrets:
            st.error("⚠️ Chưa cài đặt API Key bí mật!")
        elif not str(st.session_state.get("chan_doan_so_bo", "")).strip():
            st.warning("⚠️ Vui lòng nhập Chẩn đoán sơ bộ!")
        else:
            with st.spinner("AI đang phân tích lập luận lâm sàng để tạo Chẩn đoán phân biệt & Biện luận..."):
                try:
                    benh_su_str = get_benh_su_text_for_ai()
                    context_cdpb = (
                        f"Loại bệnh án: {loai_benh_an}\n"
                        f"Bệnh nhân: {st.session_state.get('tuoi')} tuổi, Giới tính: {st.session_state.get('gioi_tinh')}\n"
                        f"Lý do vào viện: {st.session_state.get('ly_do_vao_vien')}\n"
                        f"Bệnh sử: {benh_su_str}\n"
                        f"{clinical_history_context(loai_benh_an)}\n"
                        f"Sinh hiệu: Mạch {st.session_state.get('sh_mach')}, HA {st.session_state.get('sh_ha')}, Nhiệt độ {st.session_state.get('sh_nhiet_do')}\n"
                        f"Khám toàn thân: {st.session_state.get('kham_toan_than')}\n"
                    )
                    if is_postop_mode(loai_benh_an):
                        context_cdpb += (
                            f"Ngày hậu phẫu: {st.session_state.get('ngay_hau_phau')}\n"
                            f"Khám vết mổ: {st.session_state.get('kham_vet_mo')}\n"
                            f"Khám dẫn lưu: {st.session_state.get('kham_dan_luu')}\n"
                        )
                    context_cdpb += f"CHẨN ĐOÁN SƠ BỘ: {st.session_state.get('chan_doan_so_bo')}"

                    model = get_feature_model("KEY_AI", "gemini-3.1-flash-lite")
                    prompt_cdpb = f"""
                    Bạn là {'bác sĩ nhi khoa' if is_pediatric_mode(loai_benh_an) else 'bác sĩ lâm sàng'} thực thụ và giàu kinh nghiệm. Hãy nhìn vào toàn thể ca bệnh dưới đây, phân tích logic giữa bệnh cảnh, triệu chứng cơ năng, thực thể và chẩn đoán sơ bộ để đưa ra:
                    {'Khi phân tích bệnh nhi, phải đối chiếu tuổi, mốc phát triển, dinh dưỡng, tiêm chủng và bệnh thường gặp theo lứa tuổi; không dùng ngưỡng người lớn.' if is_pediatric_mode(loai_benh_an) else ''}
                    1. Danh sách CHẨN ĐOÁN PHÂN BIỆT (Differential Diagnosis): sắp xếp thứ tự từ khả năng cao nhất đến thấp hơn, từ bệnh lý cấp cứu nguy hiểm đến ít cấp cứu hơn.
                    2. BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ: Lập luận chặt chẽ vì sao nghĩ đến chẩn đoán sơ bộ và vì sao cần phân biệt với các bệnh lý nêu trên.

                    Dữ kiện ca bệnh:
                    {context_cdpb}

                    YÊU CẦU ĐẦU RA (Xuất ra đúng 2 khối nhãn sau, không viết thêm lời dẫn chào hỏi):
                    {AI_PLAIN_LINE_FORMAT}
                    [CHAN_DOAN_PHAN_BIET]
                    Tên bệnh A
                    Tên bệnh B
                    Tên bệnh C

                    [BIEN_LUAN_SO_BO]
                    (Nội dung đoạn văn biện luận logic, súc tích).
                    """
                    res_text = model.generate_content(prompt_cdpb).text

                    if "[CHAN_DOAN_PHAN_BIET]" in res_text and "[BIEN_LUAN_SO_BO]" in res_text:
                        parts = res_text.split("[BIEN_LUAN_SO_BO]")
                        st.session_state["chan_doan_phan_biet"] = parts[0].replace("[CHAN_DOAN_PHAN_BIET]", "").strip()
                        st.session_state["bien_luan"] = parts[1].strip()
                        st.toast("✨ Đã tạo gợi ý Chẩn đoán phân biệt & Biện luận!", icon="🪄")
                    else:
                        st.session_state["chan_doan_phan_biet"] = res_text.strip()
                except Exception as e:
                    st.error(f"Lỗi AI: {e}")

    # Hàng 2: Hai ô nhập liệu ngang hàng nhau, cùng chiều cao
    c_cd1, c_cd2 = st.columns(2)
    with c_cd1:
        placeholder_cd = "Hậu phẫu ngày thứ [X]... mổ phiên/cấp cứu do [Bệnh lý]..." if is_postop_mode(loai_benh_an) else "Chẩn đoán sơ bộ..."
        st.text_area(f"{num_sb}. Chẩn đoán sơ bộ:", key="chan_doan_so_bo", height=100, placeholder=placeholder_cd, label_visibility="collapsed")
    with c_cd2:
        st.text_area(f"{num_pb}. Chẩn đoán phân biệt:", key="chan_doan_phan_biet", height=100, label_visibility="collapsed")

    # Hàng 3: Biện luận chẩn đoán sơ bộ
    st.markdown(f"**{num_bl}. Biện luận chẩn đoán sơ bộ:**")
    st.text_area(f"{num_bl}. Biện luận chẩn đoán sơ bộ:", key="bien_luan", height=130, label_visibility="collapsed")

def ui_cls(num_dx, num_kq):
    st.markdown(f"<div class='sub-section-header'>{num_dx}. Đề xuất cận lâm sàng</div>", unsafe_allow_html=True)
    if st.button("🪄 Làm phép", type="primary", key="btn_ai_cls"):
        if "GEMINI_API_KEY" not in st.secrets:
            st.error("⚠️ Chưa cài đặt API Key!")
        else:
            with st.spinner("AI đang phân tích chỉ định cận lâm sàng tối ưu..."):
                try:
                    benh_su_str = get_benh_su_text_for_ai()
                    model = get_feature_model("KEY_AI", "gemini-3.1-flash-lite")
                    
                    if is_postop_mode(loai_benh_an):
                        # Prompt chuyên biệt hóa tuyệt đối cho Hậu phẫu
                        context_cls = (
                            f"LOẠI BỆNH ÁN: HẬU PHẪU\n"
                            f"Bệnh nhân: {st.session_state.get('tuoi')} tuổi, Giới tính: {st.session_state.get('gioi_tinh')}\n"
                            f"Diễn biến trước/trong/sau mổ:\n{benh_su_str}\n"
                            f"Khám toàn thân & Sinh hiệu: Mạch {st.session_state.get('sh_mach')}, HA {st.session_state.get('sh_ha')}, Nhiệt độ {st.session_state.get('sh_nhiet_do')}, {st.session_state.get('kham_toan_than')}\n"
                            f"Khám ngày hậu phẫu: {st.session_state.get('ngay_hau_phau')}\n"
                            f"Tình trạng vết mổ: {st.session_state.get('kham_vet_mo')}\n"
                            f"Tình trạng ống dẫn lưu: {st.session_state.get('kham_dan_luu')}\n"
                            f"Chẩn đoán sơ bộ hậu phẫu: {st.session_state.get('chan_doan_so_bo')}\n"
                            f"Chẩn đoán phân biệt / Biến chứng nghi ngờ: {st.session_state.get('chan_doan_phan_biet')}\n"
                        )
                        prompt_cls = f"""
                        Bạn là một phẫu thuật viên / bác sĩ ngoại khoa giàu kinh nghiệm. 
                        Đối với ca bệnh HẬU PHẪU dưới đây, chẩn đoán bệnh nguyên đã rõ ràng qua phẫu thuật. 
                        QUY TẮC CỐT LÕI: TUYỆT ĐỐI KHÔNG đề xuất lại các xét nghiệm chẩn đoán bệnh ban đầu (như siêu âm tìm sỏi, nội soi chẩn đoán u...). 
                        CHỈ ĐỀ XUẤT các cận lâm sàng để theo dõi và phát hiện CÁC VẤN ĐỀ SAU MỔ, bao gồm:
                        - Tầm soát và đánh giá biến chứng ngoại khoa khi cần thiết và nghi ngờ như: Chảy máu sau mổ (Hemoglobin/Hct tụt), tụ dịch/áp xe tồn dư, rò miệng nối/xì rò tiêu hóa, bục vết mổ, xẹp phổi/viêm phổi hậu phẫu, tắc ruột sau mổ.
                        - Đánh giá hồi phục chức năng và chuyển hóa: Điện giải đồ (đặc biệt K+ trong hồi phục nhu động ruột), bilan viêm/nhiễm trùng (CTM, CRP/PCT), chức năng thận (Ure, Creatinine), vi sinh cấy dịch vết mổ/dẫn lưu nếu nghi nhiễm trùng.

                        Dữ kiện ca bệnh:
                        {context_cls}

                        YÊU CẦU ĐẦU RA (Xuất đúng 3 nhãn sau, mỗi xét nghiệm xuống 1 dòng, không dùng gạch đầu dòng, không giải thích dài dòng):
                        {AI_PLAIN_LINE_FORMAT}
                        [CLS_XAC_DINH]
                        (Các CLS để phát hiện/loại trừ biến chứng sau mổ đang theo dõi: VD Siêu âm ổ bụng kiểm tra dịch tồn dư, X-quang ngực thẳng, X-quang bụng không chuẩn bị...)
                        [CLS_DIEU_TRI]
                        (Các CLS theo dõi hồi phục và định hướng điều trị/chăm sóc: VD Tổng phân tích tế bào máu, Điện giải đồ, CRP, Ure, Creatinine, đường huyết...)
                        [CLS_KHAC]
                        (Cấy vi sinh dịch dẫn lưu/mủ vết mổ làm kháng sinh đồ nếu có chỉ định, khí máu động mạch, đông máu toàn bộ...)
                        """
                    elif is_pediatric_mode(loai_benh_an):
                        context_cls = (
                            f"LOẠI BỆNH ÁN: NHI KHOA\n"
                            f"Bệnh nhi: {st.session_state.get('tuoi')} tuổi, Giới tính: {st.session_state.get('gioi_tinh')}\n"
                            f"Bệnh sử: {benh_su_str}\n"
                            f"{clinical_history_context(loai_benh_an)}\n"
                            f"Khám toàn thân & Sinh hiệu: Mạch {st.session_state.get('sh_mach')}, HA {st.session_state.get('sh_ha')}, Nhiệt độ {st.session_state.get('sh_nhiet_do')}, Nhịp thở {st.session_state.get('sh_nhip_tho')}, Cân nặng {st.session_state.get('sh_can_nang')}, Chiều cao {st.session_state.get('sh_chieu_cao')}\n"
                            f"Khám: {st.session_state.get('kham_toan_than')}\n"
                            f"Chẩn đoán sơ bộ: {st.session_state.get('chan_doan_so_bo')}\n"
                            f"Chẩn đoán phân biệt: {st.session_state.get('chan_doan_phan_biet')}"
                        )
                        prompt_cls = f"""
                        Bạn là bác sĩ nhi khoa. Dựa vào ca bệnh ({context_cls}), hãy chỉ định cận lâm sàng phù hợp với tuổi và cân nặng của bệnh nhi, cân nhắc nguy cơ bức xạ và tránh lạm dụng xét nghiệm.
                        Ưu tiên xét nghiệm cần thiết để chẩn đoán, đánh giá mức độ nặng, dinh dưỡng, mất nước và nhiễm trùng theo bệnh cảnh; nêu rõ khi nào cần siêu âm hoặc xét nghiệm chuyên sâu.
                        Trả về đúng 3 nhãn: [CLS_XAC_DINH], [CLS_DIEU_TRI], [CLS_KHAC] dưới dạng danh sách xuống dòng, không dùng gạch đầu dòng, không giải thích thừa.
                        {AI_PLAIN_LINE_FORMAT}
                        """
                    else:
                        # Prompt chuẩn cho Nội khoa / Tiền phẫu
                        context_cls = (
                            f"Loại bệnh án: {loai_benh_an}\n"
                            f"Bệnh nhân: {st.session_state.get('tuoi')} tuổi, Giới tính: {st.session_state.get('gioi_tinh')}\n"
                            f"Bệnh sử: {benh_su_str}\n"
                            f"Khám: {st.session_state.get('kham_toan_than')}\n"
                            f"Chẩn đoán sơ bộ: {st.session_state.get('chan_doan_so_bo')}\n"
                            f"Chẩn đoán phân biệt: {st.session_state.get('chan_doan_phan_biet')}"
                        )
                        prompt_cls = f"""
                        Bạn là bác sĩ lâm sàng. Dựa vào ca bệnh ({context_cls}), hãy chỉ định CẬN LÂM SÀNG cần thiết, hợp lý, tránh lạm dụng xét nghiệm:
                        Trả về đúng 3 nhãn: [CLS_XAC_DINH], [CLS_DIEU_TRI], [CLS_KHAC] dưới dạng danh sách xuống dòng, không dùng gạch đầu dòng, không giải thích thừa.
                        {AI_PLAIN_LINE_FORMAT}
                        """

                    res_cls_text = model.generate_content(prompt_cls).text

                    if "[CLS_XAC_DINH]" in res_cls_text and "[CLS_DIEU_TRI]" in res_cls_text:
                        p1 = res_cls_text.split("[CLS_DIEU_TRI]")
                        part_xd = p1[0].replace("[CLS_XAC_DINH]", "").strip()
                        if "[CLS_KHAC]" in p1[1]:
                            p2 = p1[1].split("[CLS_KHAC]")
                            part_dt, part_khac = p2[0].strip(), p2[1].strip()
                        else:
                            part_dt, part_khac = p1[1].strip(), ""
                        st.session_state["cls_dx_xac_dinh"] = part_xd
                        st.session_state["cls_dx_dieu_tri"] = part_dt
                        st.session_state["cls_dx_khac"] = part_khac
                        st.success("✨ Đã gợi ý danh mục CLS theo dõi sau mổ thành công!")
                    else:
                        st.error("AI trả về sai định dạng cấu trúc nhãn.")
                except Exception as e:
                    st.error(f"Lỗi AI: {e}")

    c_cls1, c_cls2, c_cls3 = st.columns(3)
    with c_cls1:
        nhan_cls1 = "1. Phát hiện biến chứng / Đánh giá sau mổ:" if is_postop_mode(loai_benh_an) else "1. Phục vụ chẩn đoán xác định:"
        st.text_area(nhan_cls1, key="cls_dx_xac_dinh", height=130)
    with c_cls2:
        nhan_cls2 = "2. Theo dõi hồi phục & Điều trị:" if is_postop_mode(loai_benh_an) else "2. Phục vụ điều trị:"
        st.text_area(nhan_cls2, key="cls_dx_dieu_tri", height=130)
    with c_cls3:
        st.text_area("3. Cận lâm sàng khác:", key="cls_dx_khac", height=130)
        
    st.markdown(f"<div class='sub-section-header'>{num_kq}. Cận lâm sàng đã có ({st.session_state['so_hang_cls']} xét nghiệm)</div>", unsafe_allow_html=True)
    with st.container():
        col_ocr_file, col_ocr_act = st.columns([2.5, 1])
        with col_ocr_file: lab_photos = st.file_uploader("📷 Tải lên ảnh phiếu xét nghiệm (cho phép nhiều ảnh):", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="uploader_ocr_lab_multi")
        with col_ocr_act: 
            st.write(""); st.write("")
            btn_ocr = st.button("⚡ Phân tích tất cả ảnh", type="primary", use_container_width=True, key="btn_ocr_lab_batch")

        if btn_ocr and lab_photos:
            vision_model = get_feature_model("KEY_OCR", "gemini-3.1-flash-lite")
            if not vision_model: st.error("⚠️ Hệ thống chưa được cấu hình API Key!")
            else:
                progress_bar = st.progress(0, text="Bắt đầu phân tích...")
                ocr_prompt = "Bạn là bác sĩ xét nghiệm. Đọc phiếu này và trả về JSON có 2 khóa: 'ket_qua' (liệt kê chỉ số dạng \\n- ) và 'phien_giai' (biện luận chỉ số bất thường)."
                so_hang_cls = int(st.session_state.get("so_hang_cls", 3))
                last_used_idx = -1
                for r in range(so_hang_cls):
                    val_k, val_p = str(st.session_state.get(f"cls_kq_{r}", "")).strip(), str(st.session_state.get(f"cls_pg_{r}", "")).strip()
                    if (val_k and val_k not in ["None", "-"]) or (val_p and val_p not in ["None", "-"]): last_used_idx = r

                start_row = last_used_idx + 1
                empty_rows = [start_row + i for i in range(len(lab_photos))]
                if start_row + len(lab_photos) > so_hang_cls: st.session_state["so_hang_cls"] = start_row + len(lab_photos)

                thanh_cong = 0
                for idx, photo in enumerate(lab_photos):
                    target_row = empty_rows[idx]
                    progress_bar.progress(int((idx + 1) / len(lab_photos) * 100), text=f"Xử lý ảnh {idx + 1}/{len(lab_photos)}...")
                    try:
                        img_input = optimize_lab_image(photo)
                        resp = vision_model.generate_content([ocr_prompt, img_input])
                        raw_text = resp.text.strip()
                        if raw_text.startswith("```json"): raw_text = raw_text[7:]
                        elif raw_text.startswith("```"): raw_text = raw_text[3:]
                        if raw_text.endswith("```"): raw_text = raw_text[:-3]

                        lab_data = json.loads(raw_text.strip())
                        if isinstance(lab_data, dict):
                            st.session_state[f"cls_kq_{target_row}"] = lab_data.get("ket_qua", "")
                            st.session_state[f"cls_pg_{target_row}"] = lab_data.get("phien_giai", "")
                            thanh_cong += 1
                    except Exception as e: st.warning(f"Không thể phân tích ảnh {photo.name}: {e}")

                progress_bar.empty()
                if thanh_cong > 0:
                    st.toast(f"✅ Đã phân tích xong {thanh_cong} ảnh!", icon="🧪")
        st.divider()

    for i in range(st.session_state["so_hang_cls"]):
        st.markdown(f"**Hàng {i + 1}:**")
        col_left, col_right = st.columns([1, 1])
        with col_left:
            st.text_area(f"Kết quả cận lâm sàng {i + 1}:", key=f"cls_kq_{i}", height=85)
            img = st.file_uploader(f"Đính kèm ảnh cho hàng {i + 1}:", type=["png", "jpg", "jpeg"], key=f"uploader_cls_img_{i}")
            if img:
                uploaded_imgs[f"cls_img_{i}"] = img
                st.image(img, width=160, caption=f"Ảnh hàng {i + 1}")
        with col_right:
            st.text_area(f"Biện giải cận lâm sàng {i + 1}:", key=f"cls_pg_{i}", height=85)
        st.divider()

    col_btn_them, col_btn_bot, _ = st.columns([2, 2, 6])
    with col_btn_them:
        if st.button("➕ Thêm kết quả cận lâm sàng"): st.session_state["so_hang_cls"] += 1; st.rerun()
    with col_btn_bot:
        if st.session_state["so_hang_cls"] > 1:
            if st.button("➖ Bớt hàng cuối"): st.session_state["so_hang_cls"] -= 1; st.rerun()

def ui_cdxd(num_xd, num_blxd):
    placeholder_xd = "Phẫu thuật [Tên PT] mổ [phiên/cấp cứu] ngày thứ [X] do [Bệnh lý] hiện tại [ổn định/biến chứng...]" if is_postop_mode(loai_benh_an) else "Chẩn đoán xác định..."
    st.text_area(f"{num_xd}. Chẩn đoán xác định:", key="chan_doan_xac_dinh", height=90, placeholder=placeholder_xd)
    st.text_area(f"{num_blxd}. Biện luận chẩn đoán xác định:", key="bien_luan_xac_dinh", height=110)
def check_section_has_data(keys):
    """Kiểm tra xem ít nhất một trường trong danh sách có chứa dữ liệu hay không."""
    for k in keys:
        val = st.session_state.get(k, "")
        if isinstance(val, (int, float)) and val > 0:
            return True
        if isinstance(val, str) and val.strip() and val.strip() not in ["0.0", "0", "None", "-"]:
            return True
    return False
tab1, tab2, tab3 = st.tabs(["Nhập liệu hồ sơ", "Xuất tập tin", "Phản biện lâm sàng"])

with tab1:
    # ==============================================================================
    # THANH MỤC LỤC NHANH BÊN PHẢI (RIGHT FLOATING TABLE OF CONTENTS)
    # ==============================================================================
    toc_items = [
        ("#sec-hanh-chinh", "I. Hành chính"),
        ("#sec-ly-do-benh-su", "II & III. Lý do & Bệnh sử" + (" sản phụ khoa" if is_san_phu_khoa_mode(loai_benh_an) else " hậu phẫu" if is_postop_mode(loai_benh_an) else "")),
        ("#sec-tien-su", "IV. Tiền sử"),
        ("#sec-kham-lam-sang", "V. Thăm khám lâm sàng"),
        ("#sec-tom-tat-so-bo", "VI - IX. Tóm tắt & CĐ sơ bộ"),
        ("#sec-can-lam-sang", "X & XI. Cận lâm sàng"),
        ("#sec-chan-doan-xac-dinh", "XII & XIII. CĐ xác định & Biện luận"),
        ("#sec-dieu-tri", "XIV. Điều trị"),
        ("#sec-tien-luong-tu-van", "XV & XVI. Tiên lượng & Tư vấn"),
    ]

    toc_links_html = "".join([f"<a href='{href}' class='toc-item'>{title}</a>" for href, title in toc_items])

    st.markdown(f"""
    <div class="right-toc-container">
        <details class="right-toc-details">
            <summary class="right-toc-trigger">📑 Mục lục bệnh án</summary>
            <div class="right-toc-menu">
                <div class="right-toc-header">ĐIỀU HƯỚNG NHANH</div>
                {toc_links_html}
            </div>
        </details>
    </div>
    """, unsafe_allow_html=True)
    
    
    # -------------------------------------------------------------------------
    # I. HÀNH CHÍNH (Mở nếu có dữ liệu hoặc mặc định luôn mở)
    # -------------------------------------------------------------------------
    has_hc = check_section_has_data(["ho_ten", "dan_tok", "nghe_nghiep", "khoa_phong", "dia_chi"])
    st.markdown("<div id='sec-hanh-chinh'></div>", unsafe_allow_html=True)
    with st.expander("I. PHẦN HÀNH CHÍNH", expanded=True):
        c_hc1, c_hc2, c_hc3 = st.columns(3)
        with c_hc1:
            st.text_input("Họ và tên người bệnh", key="ho_ten", placeholder="Nguyễn Văn A")
            st.text_input("Dân tộc", key="dan_tok", placeholder="Kinh, Tày, Nùng...")
        with c_hc2:
            st.number_input("Tuổi", min_value=0, max_value=120, key="tuoi")
            st.text_input("Nghề nghiệp", key="nghe_nghiep", placeholder="Kỹ sư, Hưu trí, Nông dân...")
        with c_hc3:
            st.selectbox("Giới tính", ["Nam", "Nữ", "Khác"], key="gioi_tinh")
            st.text_input("Khoa / Phòng điều trị", key="khoa_phong", placeholder="Khoa Ngoại Tiêu Hóa, Nội Tim mạch...")
        
        c_hc4, c_hc5, c_hc6 = st.columns(3)
        with c_hc4: st.text_input("Địa chỉ", key="dia_chi", placeholder="Quận Đống Đa, TP. Hà Nội")
        with c_hc5: st.text_input("Bác sĩ hoặc Sinh viên phụ trách", key="sinh_vien", placeholder="Bác sĩ nội trú, Sinh viên Y...")
        with c_hc6: st.text_input("Ngày giờ vào viện", key="ngay_vao_vien")

    # -------------------------------------------------------------------------
    # II & III. LÝ DO VÀ BỆNH SỬ (Tự mở khi có dữ liệu)
    # -------------------------------------------------------------------------
    mau_5_dong = (
        "- Hình thức mổ: Mổ phiên / Mổ cấp cứu\n"
        "- Phương pháp mổ: \n"
        "- Phương pháp gây mê: \n"
        "- Quá trình mổ: Không có biến chứng\n"
        "- Chẩn đoán sau mổ: "
    )

    keys_bs = ["ly_do_vao_vien", "benh_su", "bs_truoc_mo", "bs_sau_mo", "ts_san_khoa", "ts_phu_khoa", "ts_noi_ngoai_khoa", "ts_benh_ly", "ts_dinh_duong", "ts_san_khoa_nhi", "ts_tiem_chung", "ts_phat_trien", "ts_dich_te", "ts_di_ung", "ts_gia_dinh"]
    if is_postop_mode(loai_benh_an) or is_san_phu_khoa_mode(loai_benh_an):
        val_tm = str(st.session_state.get("bs_trong_mo", "")).strip()
        # Mở nếu có thông tin khác với mẫu 5 dòng trống hoặc các trường bệnh sử khác có chữ
        if val_tm and val_tm != mau_5_dong.strip():
            has_bs = True
        else:
            has_bs = check_section_has_data(keys_bs)
    else:
        has_bs = check_section_has_data(keys_bs)

    st.markdown("<div id='sec-ly-do-benh-su'></div>", unsafe_allow_html=True)
    tieu_de_ly_do = "II VÀ III. LÝ DO, TIỀN SỬ VÀ BỆNH SỬ" if is_san_phu_khoa_mode(loai_benh_an) else "II VÀ III. LÝ DO VÀO VIỆN VÀ BỆNH SỬ"
    with st.expander(tieu_de_ly_do, expanded=has_bs):
        st.text_area("Lý do vào viện:", key="ly_do_vao_vien", placeholder="Ví dụ: Giống bệnh án tiền phẫu", height=65)
        
        if is_san_phu_khoa_mode(loai_benh_an):
            st.markdown("**TIỀN SỬ SẢN PHỤ KHOA:**")
            c_spk1, c_spk2 = st.columns(2)
            with c_spk1:
                st.markdown("**1. Tiền sử sản khoa**")
                st.text_area("Nội dung tiền sử sản khoa:", key="ts_san_khoa", height=90, label_visibility="collapsed")
                st.markdown("**2. Tiền sử phụ khoa**")
                st.text_area("Nội dung tiền sử phụ khoa:", key="ts_phu_khoa", height=90, label_visibility="collapsed")
            with c_spk2:
                st.markdown("**3. Tiền sử nội - ngoại khoa**")
                st.text_area("Nội dung tiền sử nội - ngoại khoa:", key="ts_noi_ngoai_khoa", height=90, label_visibility="collapsed")
                st.markdown("**4. Tiền sử gia đình**")
                st.text_area("Nội dung tiền sử gia đình:", key="ts_gia_dinh", height=90, label_visibility="collapsed")
            st.markdown("**BỆNH SỬ SẢN PHỤ KHOA:**")

        if is_postop_mode(loai_benh_an):
            st.markdown("**BỆNH SỬ HẬU PHẪU:**")
            st.text_area(
                "1. Tình trạng trước mổ:",
                key="_postop_bs_truoc_mo",
                height=90,
                placeholder="Chỉ nêu các triệu chứng chính và Chẩn đoán trước mổ...",
                on_change=sync_postop_field,
                args=("bs_truoc_mo",),
            )

            val_trong_mo = st.session_state.get("_postop_bs_trong_mo", "")
            if not str(val_trong_mo).strip():
                val_trong_mo = mau_5_dong
                st.session_state["_postop_bs_trong_mo"] = mau_5_dong
                st.session_state["bs_trong_mo"] = mau_5_dong

            st.text_area(
                "2. Tình trạng trong mổ:",
                key="_postop_bs_trong_mo",
                height=130,
                on_change=sync_postop_field,
                args=("bs_trong_mo",),
            )
            st.text_area(
                "3. Quá trình sau mổ:",
                key="_postop_bs_sau_mo",
                height=90,
                placeholder="Từ lúc rời phòng hồi tỉnh đến nay: Tri giác, đau, trung tiện, tiểu tiện, tình trạng dẫn lưu, ăn uống...",
                on_change=sync_postop_field,
                args=("bs_sau_mo",),
            )
        else:
            st.text_area("Bệnh sử:", key="benh_su", placeholder="Mô tả hoàn cảnh khởi phát, triệu chứng cơ năng điển hình...", height=130)

    # -------------------------------------------------------------------------
    # IV. TIỀN SỬ (Tự mở khi có dữ liệu)
    # -------------------------------------------------------------------------
    if not is_san_phu_khoa_mode(loai_benh_an):
        pediatric_history = ["ts_benh_ly", "ts_dinh_duong", "ts_san_khoa_nhi", "ts_tiem_chung", "ts_phat_trien", "ts_dich_te", "ts_di_ung", "ts_gia_dinh"]
        standard_history = ["ts_noi_khoa", "ts_ngoai_khoa", "ts_loi_song", "ts_gia_dinh"]
        history_keys = pediatric_history if is_pediatric_mode(loai_benh_an) else standard_history
        has_ts = check_section_has_data(history_keys)
        st.markdown("<div id='sec-tien-su'></div>", unsafe_allow_html=True)
        with st.expander("IV. TIỀN SỬ", expanded=has_ts):
            if is_pediatric_mode(loai_benh_an):
                st.markdown("**MẪU TIỀN SỬ BÌNH THƯỜNG THEO TUỔI:**")
                st.caption(f"Nhóm tuổi hiện tại: {pediatric_age_group(st.session_state.get('tuoi', 0))}. Chỉ các ô đang trống mới được điền.")
                col_fill_history, col_reset_history = st.columns([2, 1])
                with col_fill_history:
                    st.button(
                        "⚡ Điền tiền sử bình thường",
                        key="btn_fill_pediatric_history",
                        on_click=fill_pediatric_normal_history,
                        use_container_width=True,
                        help="Tạo mẫu tiền sử bình thường theo tuổi đã nhập, không sử dụng AI và không ghi đè nội dung đã có.",
                    )
                with col_reset_history:
                    st.button(
                        "🔄 Xóa tiền sử",
                        key="btn_clear_pediatric_history",
                        on_click=lambda: [st.session_state.__setitem__(field_key, "") for field_key in pediatric_normal_history(0)],
                        use_container_width=True,
                    )
                if "_pediatric_history_fill_count" in st.session_state:
                    filled_count = st.session_state.pop("_pediatric_history_fill_count")
                    if filled_count:
                        st.toast(f"Đã điền mẫu bình thường cho {filled_count} mục tiền sử theo tuổi.", icon="📋")
                    else:
                        st.info("Các mục tiền sử nhi khoa đều đã có nội dung.")

                pediatric_labels = [
                    ("1. Tiền sử bệnh lý", "ts_benh_ly"),
                    ("2. Tiền sử dinh dưỡng", "ts_dinh_duong"),
                    ("3. Tiền sử sản khoa", "ts_san_khoa_nhi"),
                    ("4. Tiền sử tiêm chủng", "ts_tiem_chung"),
                    ("5. Tiền sử phát triển tâm thần vận động", "ts_phat_trien"),
                    ("6. Tiền sử dịch tễ", "ts_dich_te"),
                    ("7. Tiền sử dị ứng", "ts_di_ung"),
                    ("8. Tiền sử gia đình", "ts_gia_dinh"),
                ]
                c_ts1, c_ts2 = st.columns(2)
                for index, (label, key) in enumerate(pediatric_labels):
                    with c_ts1 if index % 2 == 0 else c_ts2:
                        st.markdown(f"<div class='sub-section-header'>{label}</div>", unsafe_allow_html=True)
                        st.text_area(f"Nội dung {label.lower()}:", key=key, height=90, label_visibility="collapsed")
            else:
                c_ts1, c_ts2 = st.columns(2)
                with c_ts1:
                    st.markdown("<div class='sub-section-header'>1. Tiền sử nội khoa</div>", unsafe_allow_html=True)
                    st.text_area("Nội dung tiền sử nội khoa:", key="ts_noi_khoa", height=90, label_visibility="collapsed")
                    st.markdown("<div class='sub-section-header'>2. Tiền sử ngoại khoa và dị ứng</div>", unsafe_allow_html=True)
                    st.text_area("Nội dung tiền sử ngoại khoa và dị ứng:", key="ts_ngoai_khoa", height=90, label_visibility="collapsed")
                with c_ts2:
                    st.markdown("<div class='sub-section-header'>3. Lối sống và thói quen</div>", unsafe_allow_html=True)
                    st.text_area("Nội dung lối sống và thói quen:", key="ts_loi_song", height=90, label_visibility="collapsed")
                    st.markdown("<div class='sub-section-header'>4. Tiền sử gia đình</div>", unsafe_allow_html=True)
                    st.text_area("Nội dung tiền sử gia đình:", key="ts_gia_dinh", height=90, label_visibility="collapsed")

    # -------------------------------------------------------------------------
    # V. THĂM KHÁM LÂM SÀNG (Tự mở khi có dữ liệu hoặc khi bấm nút điền mẫu)
    # -------------------------------------------------------------------------
    has_kham = check_section_has_data([
        "kham_vao_vien", "kham_toan_than", "sh_mach", "sh_nhiet_do", "sh_ha", "sh_nhip_tho",
        "sh_can_nang", "sh_chieu_cao", "ngay_hau_phau", "kham_vet_mo", "kham_dan_luu",
        "kham_tuan_hoan", "kham_ho_hap", "kham_tieu_hoa", "kham_than_kinh", "kham_tiet_nieu",
        "kham_co_xuong_khop", "kham_co_quan_khac", "kham_san_phu_khoa"
    ])
    st.markdown("<div id='sec-kham-lam-sang'></div>", unsafe_allow_html=True)
    with st.expander("V. THĂM KHÁM LÂM SÀNG", expanded=has_kham):
        if is_postop_mode(loai_benh_an):
            st.markdown("<div class='sub-section-header'>1. Thăm khám hiện tại - Toàn thân & Sinh hiệu</div>", unsafe_allow_html=True)
            st.text_input(
                "Khám hậu phẫu ngày thứ mấy? Giờ thứ mấy?",
                key="_postop_ngay_hau_phau",
                placeholder="VD: Ngày thứ 3 sau mổ (Giờ thứ 72)...",
                on_change=sync_postop_field,
                args=("ngay_hau_phau",),
            )
        else:
            st.markdown("<div class='sub-section-header'>1. Thăm khám lúc vào viện</div>", unsafe_allow_html=True)
            st.text_area("Nội dung khám lúc vào viện:", key="kham_vao_vien", height=80, label_visibility="collapsed")
            st.markdown("<div class='sub-section-header'>2. Thăm khám hiện tại - Toàn thân & Sinh hiệu</div>", unsafe_allow_html=True)

        LIST_TOAN_THAN = [
            ("🟢 Tỉnh táo, tiếp xúc tốt, GCS 15 điểm", "Bệnh nhân tỉnh táo, tiếp xúc tốt, Glasgow 15 điểm"),
            ("🟢 Da niêm mạc hồng hào", "Da niêm mạc hồng hào"),
            ("🟢 Không phù, không xuất huyết dưới da", "Không phù, không xuất huyết dưới da"),
            ("🟢 Tuyến giáp không to, hạch ngoại vi không sờ thấy", "Tuyến giáp không to, hạch ngoại vi không sờ thấy"),
            ("🟢 Thể trạng trung bình", "Thể trạng trung bình"),
            ("🟠 Da niêm mạc nhợt / thiếu máu", "Da niêm mạc nhợt, nghi ngờ thiếu máu"),
            ("🟠 Sốt nhẹ / gai rét", "Bệnh nhân có sốt nhẹ, gai rét"),
            ("🟠 Vã mồ hôi, đầu chi lạnh", "Vã mồ hôi, đầu chi lạnh ẩm"),
            ("🔴 Li bì, tiếp xúc chậm, Glasgow < 15đ", "Bệnh nhân li bì, tiếp xúc chậm"),
            ("🔴 Phù mềm hai chi dưới", "Phù mềm, ấn lõm hai chi dưới")
        ]

        LIST_VET_MO = [
            ("🟢 Vết mổ khô, sạch, chân chỉ không nề đỏ", "Vết mổ khô, sạch, chân chỉ không nề đỏ"),
            ("🟢 Mép mổ phẳng, liền tốt, không đau tức", "Mép mổ phẳng, liền tốt, ấn không đau tức"),
            ("🟠 Rỉ ít dịch hồng thấm băng", "Vết mổ rỉ ít dịch hồng thấm băng"),
            ("🟠 Tụ máu / bầm tím nhẹ quanh mép mổ", "Có vết bầm tím quanh mép mổ, không sưng phồng"),
            ("🔴 Chân chỉ sưng nề, tấy đỏ, đau tức nhiều", "Chân chỉ sưng nề, tấy đỏ, ấn đau tức nhiều"),
            ("🔴 Rỉ dịch mủ / dịch đục có mùi hôi", "Vết mổ rỉ dịch mủ vàng đục, có mùi hôi"),
            ("🔴 Hở mép mổ / bục chỉ vết mổ", "Hở mép mổ, bục chỉ một phần")
        ]

        LIST_DAN_LUU = [
            ("🟢 Bệnh nhân không mang ống dẫn lưu", "Bệnh nhân không đặt ống dẫn lưu"),
            ("🟢 Dẫn lưu ra lượng ít dịch hồng nhạt / thanh dịch (< 30ml/24h)", "Chân dẫn lưu sạch, ra lượng ít dịch hồng nhạt (< 30ml/24h)"),
            ("🟢 Hệ thống áp lực âm hoạt động tốt, không rỉ chân", "Hệ thống dẫn lưu hoạt động tốt, chân dẫn lưu khô"),
            ("🟠 Dẫn lưu ra ít dịch vàng chanh / cặn nhẹ", "Dẫn lưu ra ít dịch thanh vàng chanh"),
            ("🔴 Ra máu đỏ tươi liên tục nghi chảy máu trong", "Dẫn lưu ra máu đỏ tươi liên tục (> 50ml/h), nghi ngờ chảy máu trong"),
            ("🔴 Ra dịch mủ đục / dịch tiêu hóa / dịch mật nghi rò", "Dẫn lưu ra dịch mủ đục / dịch tiêu hóa nghi rò"),
            ("🔴 Tắc ống dẫn lưu / ngừng ra dịch bất thường", "Ống dẫn lưu bị tắc, không ra thêm dịch")
        ]

        col_tt_mo_ta, col_tt_sh = st.columns([1.25, 1])
        with col_tt_mo_ta:
            c_title_tt, c_pop_tt = st.columns([2, 1.2])
            with c_title_tt: st.markdown("**Mô tả khám toàn thân:**")
            with c_pop_tt:
                with st.popover("⚡ Chọn nhanh", use_container_width=True):
                    st.caption("Danh mục triệu chứng (Ưu tiên bình thường):")
                    for label, full_text in LIST_TOAN_THAN:
                        c_txt, c_btn = st.columns([3.5, 1.2])
                        with c_txt: st.markdown(f"<span style='font-size: 0.88rem;'>{label}</span>", unsafe_allow_html=True)
                        with c_btn: st.button("➕", key=f"add_tt_{label}", on_click=add_symptom_to_field, args=("kham_toan_than", full_text), use_container_width=True)

            st.text_area("Nội dung khám toàn thân:", key="kham_toan_than", height=155, label_visibility="collapsed")

        with col_tt_sh:
            st.markdown("**Dấu hiệu sinh tồn (Vital Signs):**")
            c_sh1, c_sh2 = st.columns(2)
            with c_sh1:
                st.text_input("Mạch (lần/phút):", key="sh_mach", placeholder="VD: 80")
                st.text_input("Huyết áp (mmHg):", key="sh_ha", placeholder="VD: 120/80")
                st.number_input("Cân nặng (kg):", key="sh_can_nang", min_value=0.0, max_value=250.0, value=float(st.session_state.get("sh_can_nang", 0.0)), step=0.5)
            with c_sh2:
                st.text_input("Nhiệt độ (°C):", key="sh_nhiet_do", placeholder="VD: 37.0")
                st.text_input("Nhịp thở (lần/phút):", key="sh_nhip_tho", placeholder="VD: 18")
                st.number_input("Chiều cao (cm):", key="sh_chieu_cao", min_value=0.0, max_value=230.0, value=float(st.session_state.get("sh_chieu_cao", 0.0)), step=1.0)
            
            w, h = st.session_state.get("sh_can_nang", 0.0), st.session_state.get("sh_chieu_cao", 0.0)
            if w > 0 and h > 0:
                bmi_val = round(w / ((h / 100.0) ** 2), 1)
                bmi_eval = "Gầy / Thiếu cân" if bmi_val < 18.5 else "Bình thường" if bmi_val <= 22.9 else "Thừa cân" if bmi_val <= 24.9 else "Béo phì"
                st.session_state["sh_bmi"], st.session_state["sh_bmi_eval"] = str(bmi_val), bmi_eval
                st.caption(f"📊 **BMI:** `{bmi_val} kg/m²` — **Đánh giá:** *{bmi_eval}*")
            else:
                st.session_state["sh_bmi"], st.session_state["sh_bmi_eval"] = "", ""

        if is_postop_mode(loai_benh_an):
            st.markdown("<div class='sub-section-header'>2. Thăm khám Vết mổ & Dẫn lưu</div>", unsafe_allow_html=True)
            c_vm, c_dl = st.columns(2)
            with c_vm:
                c_title_vm, c_pop_vm = st.columns([2, 1.2])
                with c_title_vm: st.markdown("**Tình trạng vết mổ:**")
                with c_pop_vm:
                    with st.popover("⚡ Chọn nhanh", use_container_width=True):
                        st.caption("Dấu hiệu vết mổ:")
                        for label, full_text in LIST_VET_MO:
                            c_txt, c_btn = st.columns([3.5, 1.2])
                            with c_txt: st.markdown(f"<span style='font-size: 0.88rem;'>{label}</span>", unsafe_allow_html=True)
                            with c_btn: st.button("➕", key=f"add_vm_{label}", on_click=add_symptom_to_field, args=("kham_vet_mo", full_text), use_container_width=True)
                st.text_area(
                    "Tình trạng vết mổ:",
                    key="_postop_kham_vet_mo",
                    height=90,
                    label_visibility="collapsed",
                    on_change=sync_postop_field,
                    args=("kham_vet_mo",),
                )

            with c_dl:
                c_title_dl, c_pop_dl = st.columns([2, 1.2])
                with c_title_dl: st.markdown("**Tình trạng ống dẫn lưu:**")
                with c_pop_dl:
                    with st.popover("⚡ Chọn nhanh", use_container_width=True):
                        st.caption("Dấu hiệu dẫn lưu:")
                        for label, full_text in LIST_DAN_LUU:
                            c_txt, c_btn = st.columns([3.5, 1.2])
                            with c_txt: st.markdown(f"<span style='font-size: 0.88rem;'>{label}</span>", unsafe_allow_html=True)
                            with c_btn: st.button("➕", key=f"add_dl_{label}", on_click=add_symptom_to_field, args=("kham_dan_luu", full_text), use_container_width=True)
                st.text_area(
                    "Tình trạng ống dẫn lưu:",
                    key="_postop_kham_dan_luu",
                    height=90,
                    label_visibility="collapsed",
                    on_change=sync_postop_field,
                    args=("kham_dan_luu",),
                )

            st.markdown("<div class='sub-section-header'>3. Thăm khám hiện tại - Các cơ quan</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='sub-section-header'>3. Thăm khám hiện tại - Các cơ quan</div>", unsafe_allow_html=True)

        active_normal_findings = organ_findings_for_mode(loai_benh_an)
        active_detailed_templates = detailed_organ_templates_for_mode(loai_benh_an)

        def xu_ly_dien_kham_binh_thuong():
            dem = 0
            for k_cq, norm_val in active_normal_findings.items():
                noi_dung = str(st.session_state.get(k_cq, "") or "").strip()
                if not noi_dung:
                    st.session_state[k_cq] = norm_val
                    dem += 1
            st.session_state["_msg_dien_cq"] = dem

        def xu_ly_xoa_cac_co_quan():
            for k_cq in active_normal_findings.keys():
                st.session_state[k_cq] = ""
            st.session_state["_msg_xoa_cq"] = True

        col_btn_fill, col_clear_cq = st.columns([2, 1])
        with col_btn_fill: st.button("⚡ Điền khám bình thường cho các cơ quan để trống", on_click=xu_ly_dien_kham_binh_thuong, use_container_width=True)
        with col_clear_cq: st.button("🔄 Đặt lại các cơ quan", on_click=xu_ly_xoa_cac_co_quan, use_container_width=True)

        if "_msg_dien_cq" in st.session_state:
            d = st.session_state.pop("_msg_dien_cq")
            if d > 0: st.toast(f"Đã điền mẫu cho {d} cơ quan còn lại!", icon="✨")
            else: st.info("Tất cả các cơ quan đều đã có dữ liệu.")

        if st.session_state.pop("_msg_xoa_cq", False):
            st.toast("Đã làm trống các ô khám cơ quan!", icon="🧹")

        st.markdown("---")

        ORGAN_DEF = [
            {"key": "kham_san_phu_khoa", "name": "Sản phụ khoa"},
            {"key": "kham_tuan_hoan", "name": "Tuần hoàn"},
            {"key": "kham_ho_hap", "name": "Hô hấp"},
            {"key": "kham_tieu_hoa", "name": "Tiêu hóa"},
            {"key": "kham_than_kinh", "name": "Thần kinh"},
            {"key": "kham_tiet_nieu", "name": "Thận - Tiết niệu"},
            {"key": "kham_co_xuong_khop", "name": "Cơ xương khớp"},
            {"key": "kham_co_quan_khac", "name": "Các cơ quan khác"}
        ]
        
        organ_options = ORGAN_DEF if is_san_phu_khoa_mode(loai_benh_an) else ORGAN_DEF[1:]
        if not is_san_phu_khoa_mode(loai_benh_an) and st.session_state.get("uu_tien_co_quan") == "Sản phụ khoa":
            st.session_state["uu_tien_co_quan"] = "Không ưu tiên (Thứ tự mặc định)"
        selected_organ_name = st.selectbox(
            "Chọn cơ quan chuyên khoa ưu tiên:", 
            ["Không ưu tiên (Thứ tự mặc định)"] + [item["name"] for item in organ_options], 
            index=0, 
            key="uu_tien_co_quan"
        )

        if selected_organ_name != "Không ưu tiên (Thứ tự mặc định)":
            fav = next(item for item in organ_options if item["name"] == selected_organ_name)
            others = [item for item in organ_options if item["name"] != selected_organ_name]
            
            # Khối cơ quan chuyên khoa trọng điểm
            c_fav_title, c_fav_btn, c_fav_ai = st.columns([2.1, 1.2, 1.6])
            with c_fav_title:
                st.markdown(f"⭐ **{fav['name'].upper()} (CƠ QUAN CHUYÊN KHOA TRỌNG ĐIỂM):**")
            with c_fav_btn:
                # Nút cho phép nạp mẫu chuyên sâu hoặc chèn mẫu nếu ô đang trống
                if st.button(f"⚡ Mẫu khám sâu {fav['name']}", key=f"btn_fill_deep_{fav['key']}", use_container_width=True):
                    st.session_state[fav["key"]] = active_detailed_templates.get(fav["key"], "")
                    st.toast(f"Đã nạp khung khám chuyên sâu cho cơ quan {fav['name']}!", icon="🩺")
                    st.rerun()
            with c_fav_ai:
                btn_ai_organ = st.button(
                    f"Gợi ý khám {fav['name']}",
                    key=f"btn_ai_organ_{fav['key']}",
                    type="primary",
                    use_container_width=True,
                )

            if btn_ai_organ:
                if "GEMINI_API_KEY" not in st.secrets:
                    st.error("⚠️ Chưa cài đặt API Key!")
                else:
                    with st.spinner(f"AI đang phân tích nội dung cần khám {fav['name']}..."):
                        try:
                            organ_context = f"""
                            Cơ quan được ưu tiên khám: {fav['name']}
                            Đối tượng: {'nhi khoa, cần diễn giải theo tuổi' if is_pediatric_mode(loai_benh_an) else 'người bệnh thông thường'}
                            Lý do vào viện: {st.session_state.get('ly_do_vao_vien')}
                            Bệnh sử: {get_benh_su_text_for_ai()}
                            {clinical_history_context(loai_benh_an)}
                            """
                            prompt_organ = f"""
                            Bạn là {'bác sĩ nhi khoa' if is_pediatric_mode(loai_benh_an) else 'bác sĩ lâm sàng'} giàu kinh nghiệm. Dựa duy nhất vào lý do vào viện, bệnh sử và tiền sử dưới đây, hãy gợi ý cho người dùng những nội dung quan trọng cần hỏi và thăm khám đối với cơ quan {fav['name']}.
                            Với bệnh nhi, luôn điều chỉnh nội dung theo tuổi, mốc phát triển và đặc điểm khám trẻ em; không áp dụng máy móc tiêu chuẩn người lớn.
                            Hãy tạo một khung khám để điền trực tiếp vào ô bệnh án. Mỗi dòng là một nội dung cần kiểm tra, kết thúc bằng dấu hai chấm để người dùng tự ghi kết quả sau khi khám, ví dụ: "Mỏm tim: ".
                            Chỉ nêu các điểm cần quan sát, sờ, gõ, nghe và nghiệm pháp cần cân nhắc nếu thực sự liên quan đến cơ quan này và bệnh cảnh.
                            Không được tự suy đoán hoặc điền kết quả bình thường/bất thường, không đưa ra kết luận chẩn đoán.

                            Dữ kiện ca bệnh:
                            {organ_context}

                            Chỉ trả về khung các nội dung cần hỏi và thăm khám, không thêm lời mở đầu, nhận xét ngoài lề hay nhãn bao quanh.
                            {AI_PLAIN_LINE_FORMAT}
                            """
                            model = get_feature_model("KEY_AI", "gemini-3.1-flash-lite")
                            suggestion = clean_ai_lines(model.generate_content(prompt_organ).text)
                            if suggestion:
                                current_exam = str(st.session_state.get(fav["key"], "")).strip()
                                st.session_state[fav["key"]] = f"{current_exam}\n{suggestion}".strip() if current_exam else suggestion
                                st.toast(f"Đã điền khung khám {fav['name']}. Bạn hãy bổ sung kết quả thực tế.", icon="🩺")
                            else:
                                st.error("AI không trả về nội dung khám.")
                        except Exception as e:
                            st.error(f"Lỗi AI: {e}")

            st.text_area(
                f"Khám chi tiết {fav['name']}:", 
                key=fav["key"], 
                height=220, 
                help="Mẫu khám chuyên khoa đầy đủ trình tự Nhìn - Sờ - Gõ - Nghe và các nghiệm pháp đặc hiệu."
            )
            
            st.markdown("---")
            st.markdown("**Các cơ quan khác (Khám định kỳ/toàn diện):**")
            c_cq1, c_cq2 = st.columns(2)
            half = len(others) // 2 + len(others) % 2
            with c_cq1:
                for org in others[:half]: 
                    st.text_area(f"{org['name']}:", key=org["key"], height=85)
            with c_cq2:
                for org in others[half:]: 
                    st.text_area(f"{org['name']}:", key=org["key"], height=85)
        else:
            # Khi không chọn ưu tiên: hiển thị 2 cột mặc định
            if is_san_phu_khoa_mode(loai_benh_an):
                st.text_area("Sản phụ khoa:", key="kham_san_phu_khoa", height=110)
            c_cq1, c_cq2 = st.columns(2)
            with c_cq1:
                st.text_area("Tuần hoàn:", key="kham_tuan_hoan", height=85)
                st.text_area("Hô hấp:", key="kham_ho_hap", height=85)
                st.text_area("Tiêu hóa:", key="kham_tieu_hoa", height=85)
                st.text_area("Thần kinh:", key="kham_than_kinh", height=85)
            with c_cq2:
                st.text_area("Thận - Tiết niệu:", key="kham_tiet_nieu", height=85)
                st.text_area("Cơ xương khớp:", key="kham_co_xuong_khop", height=85)
                st.text_area("Các cơ quan khác:", key="kham_co_quan_khac", height=85)

    # -------------------------------------------------------------------------
    # VI ĐẾN IX. TÓM TẮT & BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ
    # -------------------------------------------------------------------------
    has_tt_sobo = check_section_has_data(["tom_tat", "chan_doan_so_bo", "chan_doan_phan_biet", "bien_luan"])
    st.markdown("<div id='sec-tom-tat-so-bo'></div>", unsafe_allow_html=True)
    with st.expander("VI ĐẾN IX. TÓM TẮT VÀ BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ", expanded=has_tt_sobo):
        ui_tom_tat("VI")
        ui_cdsb("VII", "VIII", "IX")

    # -------------------------------------------------------------------------
    # X VÀ XI. CẬN LÂM SÀNG
    # -------------------------------------------------------------------------
    cls_keys = ["cls_dx_xac_dinh", "cls_dx_dieu_tri", "cls_dx_khac"]
    for i in range(st.session_state.get("so_hang_cls", )):
        cls_keys.extend([f"cls_kq_{i}", f"cls_pg_{i}"])
    has_cls = check_section_has_data(cls_keys)

    st.markdown("<div id='sec-can-lam-sang'></div>", unsafe_allow_html=True)
    with st.expander("X VÀ XI. CẬN LÂM SÀNG", expanded=has_cls):
        ui_cls("X", "XI")

    # -------------------------------------------------------------------------
    # XII VÀ XIII. CHẨN ĐOÁN XÁC ĐỊNH VÀ BIỆN LUẬN
    # -------------------------------------------------------------------------
    has_cdxd = check_section_has_data(["chan_doan_xac_dinh", "bien_luan_xac_dinh"])
    st.markdown("<div id='sec-chan-doan-xac-dinh'></div>", unsafe_allow_html=True)
    with st.expander("XII VÀ XIII. CHẨN ĐOÁN XÁC ĐỊNH VÀ BIỆN LUẬN", expanded=has_cdxd):
        ui_cdxd("XII", "XIII")

    # -------------------------------------------------------------------------
    # XIV. HƯỚNG DẪN VÀ KẾ HOẠCH ĐIỀU TRỊ
    # -------------------------------------------------------------------------
    has_dt = check_section_has_data(["dt_muc_tieu", "dt_cu_the", "dt_theo_doi"])
    st.markdown("<div id='sec-dieu-tri'></div>", unsafe_allow_html=True)
    with st.expander("XIV. HƯỚNG DẪN VÀ KẾ HOẠCH ĐIỀU TRỊ", expanded=has_dt):
        if st.button("🪄 Làm phép", key="btn_ai_dt", type="primary"):
            if "GEMINI_API_KEY" not in st.secrets: st.error("⚠️ Chưa cài đặt API Key!")
            else:
                with st.spinner("AI đang phân tích phác đồ điều trị..."):
                    try:
                        cls_da_co_str = "".join([f"+ {st.session_state.get(f'cls_kq_{i}', '')} -> {st.session_state.get(f'cls_pg_{i}', '')}\n" for i in range(st.session_state.get("so_hang_cls", 3)) if st.session_state.get(f'cls_kq_{i}', '').strip()])
                        context_dt = f"Loại: {loai_benh_an}\nBệnh nhân: {st.session_state.get('tuoi')} tuổi, {st.session_state.get('gioi_tinh')}\n{clinical_history_context(loai_benh_an)}\nChẩn đoán: {st.session_state.get('chan_doan_xac_dinh')}\nCLS quan trọng:\n{cls_da_co_str}"
                        model = get_feature_model("KEY_AI", "gemini-3.1-flash-lite")
                        prompt_dt = f"Bạn là {'bác sĩ nhi khoa' if is_pediatric_mode(loai_benh_an) else 'bác sĩ điều trị'}. Xây dựng phác đồ cho ca bệnh ({context_dt}). {'Tính liều theo cân nặng/tuổi, ghi rõ chống chỉ định và tư vấn người chăm sóc; không dùng liều người lớn.' if is_pediatric_mode(loai_benh_an) else ''} Yêu cầu trả về đúng 3 tag: [MUC_TIEU], [DIEU_TRI_CU_THE] (ghi rõ thuốc/chăm sóc vết mổ nếu hậu phẫu), [THEO_DOI]. {AI_PLAIN_LINE_FORMAT}"
                        txt = model.generate_content(prompt_dt).text

                        if "[MUC_TIEU]" in txt and "[DIEU_TRI_CU_THE]" in txt and "[THEO_DOI]" in txt:
                            p1 = txt.split("[DIEU_TRI_CU_THE]")
                            st.session_state["dt_muc_tieu"] = p1[0].replace("[MUC_TIEU]", "").strip()
                            p2 = p1[1].split("[THEO_DOI]")
                            st.session_state["dt_cu_the"] = p2[0].strip()
                            st.session_state["dt_theo_doi"] = p2[1].strip()
                            st.success("✨ Đã lên phác đồ điều trị thành công!")
                        else: st.error("AI phản hồi sai cấu trúc.")
                    except Exception as e: st.error(f"Lỗi AI: {e}")

        c_mt, c_ct, c_td = st.columns(3)
        with c_mt: st.text_area("1. Mục tiêu điều trị:", key="dt_muc_tieu", height=220)
        with c_ct: st.text_area("2. Điều trị cụ thể:", key="dt_cu_the", height=220)
        with c_td: st.text_area("3. Theo dõi:", key="dt_theo_doi", height=220)

    # -------------------------------------------------------------------------
    # XV VÀ XVI. TIÊN LƯỢNG VÀ TƯ VẤN
    # -------------------------------------------------------------------------
    has_tltv = check_section_has_data(["tien_luong", "tu_van"])
    st.markdown("<div id='sec-tien-luong-tu-van'></div>", unsafe_allow_html=True)
    with st.expander("XV VÀ XVI. TIÊN LƯỢNG VÀ TƯ VẤN", expanded=has_tltv):
        if st.button("🪄 Làm phép", type="primary", key="btn_ai_tienluong"):
            if "GEMINI_API_KEY" not in st.secrets: st.error("⚠️ Chưa cài đặt API Key!")
            else:
                with st.spinner("AI đang phân tích logic lâm sàng..."):
                    try:
                        context = f"Loại: {loai_benh_an}\nTuổi: {st.session_state.get('tuoi')}, Giới tính: {st.session_state.get('gioi_tinh')}\n{clinical_history_context(loai_benh_an)}\nChẩn đoán: {st.session_state.get('chan_doan_xac_dinh')}\nĐiều trị: {st.session_state.get('dt_cu_the')}"
                        model = get_feature_model("KEY_AI", "gemini-3.1-flash-lite")
                        prompt = f"Bạn là {'bác sĩ nhi khoa' if is_pediatric_mode(loai_benh_an) else 'bác sĩ lâm sàng'}. Đưa ra TIÊN LƯỢNG và TƯ VẤN cho ca bệnh ({context}). {'Tư vấn phải hướng tới cha mẹ/người chăm sóc và có dấu hiệu cảnh báo cần đưa trẻ đi khám ngay.' if is_pediatric_mode(loai_benh_an) else ''} Yêu cầu trả về đúng 2 tag: [TIEN_LUONG] và [TU_VAN]. {AI_PLAIN_LINE_FORMAT}"
                        res_text = model.generate_content(prompt).text

                        if "[TIEN_LUONG]" in res_text and "[TU_VAN]" in res_text:
                            parts = res_text.split("[TU_VAN]")
                            st.session_state["tien_luong"] = parts[0].replace("[TIEN_LUONG]", "").strip()
                            st.session_state["tu_van"] = parts[1].strip()
                            st.success("✨ Đã tạo gợi ý thành công!")
                        else: st.error("AI trả về sai định dạng.")
                    except Exception as e: st.error(f"Lỗi AI: {e}")

        c_pl, c_tv = st.columns(2)
        with c_pl: st.text_area("XV. Tiên lượng:", key="tien_luong", height=250)
        with c_tv: st.text_area("XVI. Tư vấn:", key="tu_van", height=250)
# Gom dữ liệu để xuất file
data_benh_an = {k: st.session_state.get(k, "") for k in FIELDS_TO_SAVE}
data_benh_an["loai_benh_an"] = loai_benh_an
data_benh_an["sh_mach"] = str(st.session_state.get("sh_mach", "")).strip()
data_benh_an["sh_nhiet_do"] = str(st.session_state.get("sh_nhiet_do", "")).strip()
data_benh_an["sh_ha"] = str(st.session_state.get("sh_ha", "")).strip()
data_benh_an["sh_nhip_tho"] = str(st.session_state.get("sh_nhip_tho", "")).strip()
data_benh_an["sh_can_nang"] = str(st.session_state.get("sh_can_nang", 0.0))
data_benh_an["sh_chieu_cao"] = str(st.session_state.get("sh_chieu_cao", 0.0))
data_benh_an["sh_bmi"] = str(st.session_state.get("sh_bmi", ""))
data_benh_an["sh_bmi_eval"] = str(st.session_state.get("sh_bmi_eval", ""))
data_benh_an["so_hang_cls"] = st.session_state.get("so_hang_cls", )
for i in range(data_benh_an["so_hang_cls"]):
    data_benh_an[f"cls_kq_{i}"] = st.session_state.get(f"cls_kq_{i}", "")
    data_benh_an[f"cls_pg_{i}"] = st.session_state.get(f"cls_pg_{i}", "")
if 'uploaded_imgs' in locals(): data_benh_an.update(uploaded_imgs)

# --- TAB 2: XEM TRƯỚC VÀ XUẤT TẬP TIN ---
with tab2:
    ho_ten_val = str(st.session_state.get("ho_ten", "")).strip()
    if ho_ten_val:
        so_hang = st.session_state.get("so_hang_cls", 3)
        dem_cls = sum(1 for i in range(so_hang) if str(st.session_state.get(f"cls_kq_{i}", "")).strip() or (locals().get('uploaded_imgs') and uploaded_imgs.get(f"cls_img_{i}")))
        overview_rows = [
            f"<div class='overview-row'><span class='overview-label'>Bệnh nhân:</span> {html.escape(ho_ten_val.upper())} | {html.escape(str(st.session_state.get('tuoi')))} tuổi | Giới tính: {html.escape(str(st.session_state.get('gioi_tinh')))} | Loại bệnh án: {html.escape(loai_benh_an)}</div>",
            f"<div class='overview-row'><span class='overview-label'>Khoa phòng:</span> {html.escape(str(st.session_state.get('khoa_phong') or 'Chưa điền'))} | <span class='overview-label'>Lý do vào viện:</span> {html.escape(str(st.session_state.get('ly_do_vao_vien') or 'Chưa điền'))}</div>",
            f"<div class='overview-row'><span class='overview-label'>Cận lâm sàng đã nhập:</span> {dem_cls}/{so_hang} hàng</div>",
        ]
        chan_doan_so_bo = str(st.session_state.get("chan_doan_so_bo", "")).strip()
        if chan_doan_so_bo:
            overview_rows.append(f"<div class='overview-row'><span class='overview-label'>Chẩn đoán sơ bộ:</span> {html.escape(chan_doan_so_bo)}</div>")
        chan_doan_xac_dinh = str(st.session_state.get("chan_doan_xac_dinh", "")).strip()
        if chan_doan_xac_dinh:
            overview_rows.append(f"<div class='overview-diagnosis'>Chẩn đoán xác định: {html.escape(chan_doan_xac_dinh)}</div>")
        st.markdown(
            "<div class='overview-panel'>"
            "<div class='overview-panel-title'>XEM TRƯỚC THÔNG TIN TỔNG QUAN</div>"
            f"<div class='overview-panel-body'>{''.join(overview_rows)}</div>"
            "</div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            "<div class='overview-panel'>"
            "<div class='overview-panel-title'>XEM TRƯỚC THÔNG TIN TỔNG QUAN</div>"
            "<div class='overview-panel-body'><div class='overview-empty'>Vui lòng điền thông tin bên tab Nhập liệu hồ sơ.</div></div>"
            "</div>",
            unsafe_allow_html=True,
        )

    st.markdown("---")
    col_dl_pdf, col_dl_docx = st.columns(2)
    
    with col_dl_pdf:
        if st.button("📄 Tạo & Xem trước tập tin PDF", type="primary", use_container_width=True):
            if not ho_ten_val: 
                st.error("Vui lòng điền tối thiểu Họ và tên người bệnh trước khi xuất tập tin!")
            elif not os.path.exists("Roboto-Regular.ttf") or not os.path.exists("Roboto-Bold.ttf"): 
                st.error("Chưa tìm thấy tập tin font 'Roboto-Regular.ttf' và 'Roboto-Bold.ttf' trong cùng thư mục với app.py!")
            else:
                with st.spinner("Đang kết xuất văn bản PDF..."):
                    pdf_bytes = export_pdf(data_benh_an)
                    st.session_state["pdf_bytes_preview"] = pdf_bytes
                    ten_mau_file = {
                        "Nhi khoa": "Nhi_khoa_",
                        "Hậu phẫu": "Hau_phau_",
                        "Sản phụ khoa / Tiền phẫu": "San_phu_khoa_Tien_phau_",
                        "Sản phụ khoa / Hậu phẫu": "San_phu_khoa_Hau_phau_",
                    }.get(loai_benh_an, "")
                    st.session_state["ten_file_pdf"] = f"Benh_an_{ten_mau_file}{ho_ten_val.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf"
                    st.session_state["active_preview"] = "pdf"

        if st.session_state.get("pdf_bytes_preview"):
            st.download_button("📥 Tải PDF về máy", data=st.session_state["pdf_bytes_preview"], file_name=st.session_state.get("ten_file_pdf", "benh_an.pdf"), mime="application/pdf", use_container_width=True)

    with col_dl_docx:
        if st.button("📝 Tạo & Xem trước tập tin Word (.docx)", type="secondary", use_container_width=True):
            if not ho_ten_val:
                st.error("Vui lòng điền tối thiểu Họ và tên người bệnh!")
            else:
                with st.spinner("Đang kết xuất và chuyển đổi tài liệu Word..."):
                    docx_bytes = export_docx(data_benh_an)
                    ten_mau_file = {
                        "Nhi khoa": "Nhi_khoa_",
                        "Hậu phẫu": "Hau_phau_",
                        "Sản phụ khoa / Tiền phẫu": "San_phu_khoa_Tien_phau_",
                        "Sản phụ khoa / Hậu phẫu": "San_phu_khoa_Hau_phau_",
                    }.get(loai_benh_an, "")
                    ten_file_docx = f"Benh_an_{ten_mau_file}{ho_ten_val.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.docx"
                    st.session_state["docx_bytes_data"] = docx_bytes
                    st.session_state["ten_file_docx"] = ten_file_docx
                    
                    # Chuyển đổi DOCX sang HTML để hiển thị xem trước
                    try:
                        res_html = mammoth.convert_to_html(io.BytesIO(docx_bytes))
                        st.session_state["docx_html_preview"] = res_html.value
                    except Exception as err:
                        st.session_state["docx_html_preview"] = f"<p style='color:red;'>Lỗi hiển thị bản xem trước: {err}</p>"
                        
                    st.session_state["active_preview"] = "docx"
                    st.success("Tạo tài liệu Word thành công!")

        if st.session_state.get("docx_bytes_data"):
            st.download_button(
                "📥 Tải Word (.docx) về máy",
                data=st.session_state["docx_bytes_data"],
                file_name=st.session_state.get("ten_file_docx", "benh_an.docx"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    # --- KHU VỰC HIỂN THỊ XEM TRƯỚC (PREVIEW) ---
    if st.session_state.get("active_preview") == "docx" and st.session_state.get("docx_html_preview"):
        st.markdown("---")
        st.markdown("#### 📝 Bản xem trước tài liệu Word trực tiếp:")
        
        # Bọc mã HTML trong container mô phỏng trang tài liệu A4
        styled_word_preview = f"""
        <div style="
            background-color: #525659;
            padding: 25px 15px;
            display: flex;
            justify-content: center;
            border-radius: 4px;
        ">
            <div style="
                background: #ffffff;
                color: #24292e;
                width: 100%;
                max-width: 800px;
                min-height: 900px;
                padding: 40px 50px;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.35);
                font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
                font-size: 14px;
                line-height: 1.6;
            ">
                <style>
                    table {{
                        width: 100% !important;
                        border-collapse: collapse !important;
                        margin: 12px 0 !important;
                    }}
                    th, td {{
                        border: 1px solid #c8d1dc !important;
                        padding: 6px 10px !important;
                        vertical-align: top !important;
                    }}
                    th {{
                        background-color: #e1ebf5 !important;
                        font-weight: bold !important;
                    }}
                    ul, ol {{
                        margin-top: 4px !important;
                        margin-bottom: 6px !important;
                        padding-left: 24px !important;
                    }}
                    p {{
                        margin-top: 3px !important;
                        margin-bottom: 5px !important;
                    }}
                    img {{
                        max-width: 100% !important;
                        height: auto !important;
                        margin: 8px 0 !important;
                        border: 1px solid #ddd !important;
                    }}
                </style>
                {st.session_state['docx_html_preview']}
            </div>
        </div>
        """
        st.markdown(styled_word_preview, unsafe_allow_html=True)

    elif st.session_state.get("active_preview") == "pdf" and st.session_state.get("pdf_bytes_preview"):
        st.markdown("---")
        st.markdown("#### 📄 Bản xem trước PDF trực tiếp:")
        pdf_viewer(input=st.session_state["pdf_bytes_preview"], width=750, height=850)
# --- TAB 3: PHẢN BIỆN BỆNH ÁN ---
with tab3:
    st.markdown("### Giảng viên lâm sàng phản biện ca bệnh")
    st.caption("Giảng viên lâm sàng giàu kinh nghiệm, rà soát tính logic toàn diện của ca bệnh và đặt câu hỏi.")

    benh_su_str_pb = get_benh_su_text_for_ai()
    cls_da_co_str_pb = "\n".join(
        f"Hàng {row_index + 1}: Kết quả: {st.session_state.get(f'cls_kq_{row_index}', '')} | Biện giải: {st.session_state.get(f'cls_pg_{row_index}', '')}"
        for row_index in range(st.session_state.get("so_hang_cls", 1))
        if str(st.session_state.get(f"cls_kq_{row_index}", "")).strip()
        or str(st.session_state.get(f"cls_pg_{row_index}", "")).strip()
    ) or "Chưa ghi nhận kết quả cận lâm sàng"
    ca_benh_summary = f"""
    Loại bệnh án: {loai_benh_an}
    Hành chính: Họ tên {st.session_state.get('ho_ten')}, tuổi {st.session_state.get('tuoi')}, giới tính {st.session_state.get('gioi_tinh')}, dân tộc {st.session_state.get('dan_tok')}, nghề nghiệp {st.session_state.get('nghe_nghiep')}, khoa/phòng {st.session_state.get('khoa_phong')}
    Lý do vào viện: {st.session_state.get('ly_do_vao_vien')}
    Bệnh sử: {benh_su_str_pb}
    {clinical_history_context(loai_benh_an)}
    Khám toàn thân: {st.session_state.get('kham_toan_than')}
    Sinh hiệu: Mạch {st.session_state.get('sh_mach')}, nhiệt độ {st.session_state.get('sh_nhiet_do')}, huyết áp {st.session_state.get('sh_ha')}, nhịp thở {st.session_state.get('sh_nhip_tho')}, cân nặng {st.session_state.get('sh_can_nang')}, chiều cao {st.session_state.get('sh_chieu_cao')}, BMI {st.session_state.get('sh_bmi')} {st.session_state.get('sh_bmi_eval')}
    Khám lúc vào viện: {st.session_state.get('kham_vao_vien')}
    Khám tuần hoàn: {st.session_state.get('kham_tuan_hoan')}
    Khám hô hấp: {st.session_state.get('kham_ho_hap')}
    Khám tiêu hóa: {st.session_state.get('kham_tieu_hoa')}
    Khám thần kinh: {st.session_state.get('kham_than_kinh')}
    Khám thận - tiết niệu: {st.session_state.get('kham_tiet_nieu')}
    Khám cơ xương khớp: {st.session_state.get('kham_co_xuong_khop')}
    Khám các cơ quan khác: {st.session_state.get('kham_co_quan_khac')}
    Ngày hậu phẫu: {st.session_state.get('ngay_hau_phau', 'Không có')}
    Tình trạng vết mổ: {st.session_state.get('kham_vet_mo')}
    Tình trạng ống dẫn lưu: {st.session_state.get('kham_dan_luu')}
    Tóm tắt bệnh án: {st.session_state.get('tom_tat')}
    Chẩn đoán sơ bộ: {st.session_state.get('chan_doan_so_bo')}
    Chẩn đoán phân biệt: {st.session_state.get('chan_doan_phan_biet')}
    Biện luận chẩn đoán sơ bộ: {st.session_state.get('bien_luan')}
    Đề xuất cận lâm sàng phục vụ chẩn đoán: {st.session_state.get('cls_dx_xac_dinh')}
    Đề xuất cận lâm sàng phục vụ điều trị: {st.session_state.get('cls_dx_dieu_tri')}
    Đề xuất cận lâm sàng khác: {st.session_state.get('cls_dx_khac')}
    Kết quả cận lâm sàng đã có:
    {cls_da_co_str_pb}
    Chẩn đoán xác định: {st.session_state.get('chan_doan_xac_dinh')}
    Biện luận chẩn đoán xác định: {st.session_state.get('bien_luan_xac_dinh')}
    Mục tiêu điều trị: {st.session_state.get('dt_muc_tieu')}
    Điều trị cụ thể: {st.session_state.get('dt_cu_the')}
    Theo dõi sau điều trị: {st.session_state.get('dt_theo_doi')}
    Tiên lượng: {st.session_state.get('tien_luong')}
    Tư vấn: {st.session_state.get('tu_van')}
    """

    col_btn_pb, col_mode = st.columns([1, 1.5])
    with col_mode: phong_cach = st.selectbox("Phong cách chất vấn:", ["Học thuật & Hướng dẫn", "Nghiêm khắc & Thách thức", "Thực chiến giao ban"])
    with col_btn_pb: st.write(""); btn_phan_bien = st.button("Giảng viên phản biện & Đặt câu hỏi", type="primary")

    if btn_phan_bien:
        model = get_feature_model("KEY_ATTENDING", "gemini-3.6-flash")
        if not model: st.error("⚠️ Hệ thống chưa cấu hình API Key!")
        elif not st.session_state.get("chan_doan_so_bo"): st.warning("⚠️ Vui lòng nhập tối thiểu Chẩn đoán sơ bộ ở Tab 1!")
        else:
            with st.spinner("Thầy/Cô đang đọc kỹ bệnh án..."):
                try:
                    prompt_phan_bien = f"""Bạn là {'Giảng viên Nhi khoa' if is_pediatric_mode(loai_benh_an) else 'Giảng viên lâm sàng'}. Hãy nhận xét ca bệnh ({ca_benh_summary}) theo phong cách {phong_cach}.
                    {'Tập trung kiểm tra đánh giá theo tuổi, dinh dưỡng, tiêm chủng, phát triển tâm thần vận động, liều thuốc theo cân nặng và tư vấn người chăm sóc.' if is_pediatric_mode(loai_benh_an) else ''}
                    Trả về đúng định dạng JSON: {{"nhan_xet_tong_the": "...", "danh_sach_cau_hoi": [{{"chu_de": "...", "cau_hoi": "...", "goi_y_tra_loi": "..."}}]}}"""
                    res_pb = model.generate_content(prompt_phan_bien)
                    res_raw = res_pb.text.strip()
                    if res_raw.startswith("```json"): res_raw = res_raw[7:]
                    elif res_raw.startswith("```"): res_raw = res_raw[3:]
                    if res_raw.endswith("```"): res_raw = res_raw[:-3]
                    st.session_state["data_phan_bien_json"] = json.loads(res_raw.strip())
                except Exception as e: st.error(f"Lỗi AI: {e}")

    if st.session_state.get("data_phan_bien_json"):
        data_pb = st.session_state["data_phan_bien_json"]
        st.divider()
        st.markdown("#### Nhận xét tổng thể & Điểm cần lưu ý:")
        st.info(data_pb.get("nhan_xet_tong_the", ""))
        st.markdown("#### ❓ Câu hỏi vấn đáp (Bấm vào từng câu để xem gợi ý đáp án):")
        for idx, item in enumerate(data_pb.get("danh_sach_cau_hoi", [])):
            with st.expander(f"**Câu {idx + 1} ({item.get('chu_de')}):** {item.get('cau_hoi')}", expanded=False):
                st.markdown("**Gợi ý hướng trả lời (Teaching Points):**")
                st.markdown(item.get("goi_y_tra_loi"))

# ==============================================================================
# CƠ CHẾ TỰ ĐỘNG LƯU NHÁP VÀO LOCALSTORAGE TRÌNH DUYỆT
# ==============================================================================
co_du_lieu = any(bool(str(st.session_state.get(k, "")).strip()) for k in [
    "ho_ten", "benh_su", "bs_truoc_mo", "ly_do_vao_vien", "kham_toan_than", "chan_doan_so_bo",
    "ts_benh_ly", "ts_dinh_duong", "ts_san_khoa_nhi", "ts_tiem_chung", "ts_phat_trien", "ts_dich_te", "ts_di_ung", "ts_gia_dinh"
])
if co_du_lieu:
    current_snapshot = {k: st.session_state.get(k, "") for k in FIELDS_TO_SAVE}
    for i in range(st.session_state.get("so_hang_cls", 3)):
        current_snapshot[f"cls_kq_{i}"] = st.session_state.get(f"cls_kq_{i}", "")
        current_snapshot[f"cls_pg_{i}"] = st.session_state.get(f"cls_pg_{i}", "")
    snapshot_json = json.dumps(current_snapshot, ensure_ascii=False)
    if st.session_state.get("last_saved_snapshot") != snapshot_json:
        local_storage.setItem(STORAGE_KEY, snapshot_json)
        st.session_state["last_saved_snapshot"] = snapshot_json