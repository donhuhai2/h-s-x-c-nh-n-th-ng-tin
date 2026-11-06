# -*- coding: utf-8 -*-
"""
Web App Nhập liệu thẩm định LIS & Xuất báo cáo Word / Excel.
Đã tích hợp:
 1. Đăng nhập / Đăng xuất, mã hóa mật khẩu SHA-256.
 2. Phân quyền Admin / User (chỉ Admin mới có quyền xóa dữ liệu và tạo người dùng).
 3. Đổi mật khẩu cá nhân an toàn.
 4. Cơ sở dữ liệu đám mây Supabase cho cả hồ sơ xét nghiệm và tài khoản.
"""

import os
import io
import re
import json
import zipfile
import subprocess
import tempfile
import urllib.request
import time
import hashlib
from datetime import date, datetime
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from PIL import Image

try:
    from google import genai
    from google.genai import types
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False

try:
    from supabase import create_client, Client
    HAS_SUPABASE = True
except ImportError:
    HAS_SUPABASE = False

EXCEL_FILE = "danh_sach_phieu.xlsx"
TEMPLATE_FILE = "template.docx"
API_KEY_FILE = "api_key_config.txt"

st.set_page_config(
    page_title="Hệ Thống Thẩm Định PXN - LIS",
    page_icon="🧪",
    layout="wide"
)

# ================= QUẢN LÝ BẢO MẬT & BĂM MẬT KHẨU =================
def hash_password(password):
    """Mã hóa mật khẩu bằng SHA-256."""
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

# ================= KẾT NỐI SUPABASE =================
def get_supabase_client():
    if not HAS_SUPABASE:
        return None
    url = os.environ.get("SUPABASE_URL") or st.secrets.get("SUPABASE_URL", "")
    key = os.environ.get("SUPABASE_KEY") or st.secrets.get("SUPABASE_KEY", "")
    if url and key:
        try:
            return create_client(url, key)
        except Exception:
            return None
    return None

supabase: Client = get_supabase_client()

# ================= TÀI KHOẢN & XÁC THỰC =================
def get_user(username):
    """Tìm thông tin tài khoản."""
    u_clean = username.strip().lower()
    if supabase:
        try:
            res = supabase.table("users").select("*").eq("username", u_clean).execute()
            if res.data:
                return res.data[0]
        except Exception:
            pass
    # Tài khoản dự phòng nội bộ nếu chưa kết nối bảng users
    if u_clean == "admin":
        return {
            "username": "admin",
            "password_hash": hash_password("admin"),
            "role": "Admin",
            "full_name": "Quản trị viên"
        }
    return None

def change_user_password(username, new_password):
    """Đổi mật khẩu tài khoản."""
    new_hash = hash_password(new_password)
    if supabase:
        try:
            supabase.table("users").update({"password_hash": new_hash}).eq("username", username.strip().lower()).execute()
            return True, "Đổi mật khẩu thành công!"
        except Exception as e:
            return False, f"Lỗi đổi mật khẩu: {e}"
    return True, "Đã đổi mật khẩu tạm thời."

def create_new_user(username, password, role, full_name):
    """Tạo người dùng mới (chỉ Admin)."""
    u_clean = username.strip().lower()
    p_hash = hash_password(password)
    if supabase:
        try:
            check = supabase.table("users").select("id").eq("username", u_clean).execute()
            if check.data:
                return False, "Tên đăng nhập đã tồn tại!"
            supabase.table("users").insert({
                "username": u_clean,
                "password_hash": p_hash,
                "role": role,
                "full_name": full_name
            }).execute()
            return True, f"Tạo tài khoản {username} thành công!"
        except Exception as e:
            return False, f"Lỗi tạo tài khoản: {e}"
    return False, "Cần kết nối Supabase để lưu trữ tài khoản!"

def fetch_all_users():
    """Lấy danh sách tài khoản."""
    if supabase:
        try:
            res = supabase.table("users").select("id, username, role, full_name").execute()
            return pd.DataFrame(res.data)
        except Exception:
            pass
    return pd.DataFrame()

# Khởi tạo trạng thái đăng nhập
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "user_info" not in st.session_state:
    st.session_state["user_info"] = None

# ================= GIAO DIỆN ĐĂNG NHẬP =================
if not st.session_state["logged_in"]:
    st.title("🔐 Đăng Nhập Hệ Thống PXN")
    st.markdown("Vui lòng đăng nhập để sử dụng phần mềm quản lý và thẩm định kết quả.")

    col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
    with col_l2:
        with st.form("login_form"):
            login_user = st.text_input("Tên đăng nhập:", placeholder="Nhập username...")
            login_pass = st.text_input("Mật khẩu:", type="password", placeholder="Nhập password...")
            btn_login = st.form_submit_button("Đăng nhập", use_container_width=True)

        if btn_login:
            user_data = get_user(login_user)
            if user_data and user_data["password_hash"] == hash_password(login_pass):
                st.session_state["logged_in"] = True
                st.session_state["user_info"] = user_data
                st.success(f"Chào mừng {user_data.get('full_name', login_user)}!")
                st.rerun()
            else:
                st.error("Tên đăng nhập hoặc mật khẩu không chính xác!")
        
        st.info("💡 Tài khoản mặc định ban đầu: **admin** / mật khẩu: **admin**")
    st.stop()

# ================= XỬ LÝ DỮ LIỆU XÉT NGHIỆM =================
def clean_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", str(filename)).strip()

def load_saved_api_key():
    if os.path.exists(API_KEY_FILE):
        try:
            with open(API_KEY_FILE, "r", encoding="utf-8") as f:
                saved = f.read().strip()
                if saved:
                    return saved
        except Exception:
            pass
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass
    return ""

def save_api_key_to_file(key_val):
    try:
        with open(API_KEY_FILE, "w", encoding="utf-8") as f:
            f.write(str(key_val).strip())
        return True
    except Exception:
        return False

def fetch_all_records_df():
    if supabase:
        try:
            res = supabase.table("lab_records").select("id, data").order("id", desc=False).execute()
            rows = [item["data"] for item in res.data]
            if rows:
                return pd.DataFrame(rows).fillna("")
            return pd.DataFrame()
        except Exception as e:
            st.error(f"Lỗi truy vấn Supabase: {e}")
    if os.path.exists(EXCEL_FILE):
        try:
            return pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def save_single_record(row_dict):
    clean_dict = {k: ("" if v is None else str(v)) for k, v in row_dict.items()}
    if supabase:
        try:
            supabase.table("lab_records").insert({"data": clean_dict}).execute()
        except Exception as e:
            st.warning(f"Lỗi lưu mây: {e}")
    new_df = pd.DataFrame([clean_dict])
    if os.path.exists(EXCEL_FILE):
        try:
            old_df = pd.read_excel(EXCEL_FILE, dtype=str)
            combined = pd.concat([old_df, new_df], ignore_index=True)
        except Exception:
            combined = new_df
    else:
        combined = new_df
    combined.to_excel(EXCEL_FILE, index=False)

def delete_records_by_date(target_date):
    deleted_count = 0
    t_date = str(target_date).strip()
    if supabase:
        try:
            res = supabase.table("lab_records").select("id, data").execute()
            ids_to_del = []
            for item in res.data:
                d1 = str(item["data"].get("Ngay", "")).strip()
                d2 = str(item["data"].get("Ngay_xac_nhan", "")).strip()
                if d1 == t_date or d2 == t_date:
                    ids_to_del.append(item["id"])
            if ids_to_del:
                supabase.table("lab_records").delete().in_("id", ids_to_del).execute()
                deleted_count = len(ids_to_del)
        except Exception as e:
            st.error(f"Lỗi khi xóa trên Supabase: {e}")
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
            mask1 = (df.get("Ngay", "").str.strip() == t_date)
            mask2 = (df.get("Ngay_xac_nhan", "").str.strip() == t_date)
            df_rem = df[~(mask1 | mask2)]
            df_rem.to_excel(EXCEL_FILE, index=False)
            if not supabase:
                deleted_count = len(df) - len(df_rem)
        except Exception:
            pass
    return deleted_count

def delete_all_records():
    if supabase:
        try:
            supabase.table("lab_records").delete().neq("id", 0).execute()
        except Exception as e:
            st.error(f"Lỗi khi xóa trên Supabase: {e}")
    if os.path.exists(EXCEL_FILE):
        try:
            os.remove(EXCEL_FILE)
        except Exception:
            pass

def export_single_word(context, template_path=TEMPLATE_FILE):
    doc = DocxTemplate(template_path)
    clean_ctx = {k: ("" if v is None else str(v)) for k, v in context.items()}
    doc.render(clean_ctx)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def convert_docx_to_pdf_bytes(docx_bytes):
    with tempfile.TemporaryDirectory() as tmp_dir:
        input_docx = os.path.join(tmp_dir, "input.docx")
        with open(input_docx, "wb") as f:
            f.write(docx_bytes)
        cmd = None
        for exec_name in ["libreoffice", "soffice"]:
            try:
                subprocess.run([exec_name, "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
                cmd = exec_name
                break
            except Exception:
                continue
        if cmd:
            try:
                subprocess.run(
                    [cmd, "--headless", "--convert-to", "pdf", input_docx, "--outdir", tmp_dir],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    check=True
                )
                output_pdf = os.path.join(tmp_dir, "input.pdf")
                if os.path.exists(output_pdf):
                    with open(output_pdf, "rb") as f_pdf:
                        return f_pdf.read()
            except Exception:
                pass
    return None

def get_bytes_from_gdrive_url(url):
    match = re.search(r'(?:/file/d/|id=)([a-zA-Z0-9_-]+)', url)
    if not match:
        raise ValueError("Đường dẫn Google Drive không hợp lệ! Vui lòng đảm bảo link có dạng: https://drive.google.com/file/d/.../view")
    file_id = match.group(1)
    direct_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    req = urllib.request.Request(direct_url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as resp:
        return resp.read()

def extract_lab_data_from_image(image_bytes, api_key, mode="patient"):
    client = genai.Client(api_key=api_key)
    if mode == "recommendation":
        prompt = """
        Bạn là chuyên gia thẩm định hồ sơ kết quả xét nghiệm (LIS/PXN).
        Hãy đọc kỹ ảnh phiếu xét nghiệm và trích xuất các thông tin phục vụ mục "XÁC NHẬN XÉT NGHIỆM ĐƯỢC CHỈNH SỬA, KHUYẾN CÁO":
        1. SID_KQCS: Mã bệnh nhân / Barcode / SID.
        2. Ngay_KQCS: Ngày lấy mẫu hoặc ngày xét nghiệm (dd/mm/yyyy).
        3. Thiet_bi_KQCS: Tên máy/thiết bị phân tích (nếu có, không có để trống).
        4. XN_KQCS: Tên chỉ số xét nghiệm liên quan (ví dụ: Glucose, Urea, AST...).
        5. Ghi_chu_Khuyen_cao: Trích xuất toàn bộ nội dung "Ghi chú", "Khuyến cáo", "Lời dặn", hoặc thông tin trạng thái mẫu.

        Trả về DUY NHẤT một chuỗi JSON thuần túy:
        {
            "SID_KQCS": "...",
            "Ngay_KQCS": "...",
            "Thiet_bi_KQCS": "...",
            "XN_KQCS": "...",
            "Ghi_chu_Khuyen_cao": "..."
        }
        """
    else:
        prompt = """
        Bạn là chuyên gia thẩm định hồ sơ xét nghiệm LIS/PXN.
        Hãy đọc ảnh phiếu kết quả xét nghiệm và trích xuất dữ liệu theo đúng quy tắc sau:

        1. THÔNG TIN CHUNG:
           - SID: Mã bệnh nhân / Barcode / SID.
           - Ngay: Ngày xét nghiệm (dd/mm/yyyy).
           - Gio_nhan_mau: Giờ nhận mẫu (hh:mm).
           - May: Tên thiết bị phân tích (nếu có).

        2. THẨM ĐỊNH KẾT QUẢ THÔNG THƯỜNG:
           - Chọn BẤT KỲ 1 chỉ số xét nghiệm thuộc chuyên khoa SINH HÓA hoặc MIỄN DỊCH có kết quả NẰM TRONG KHOẢNG THAM CHIẾU SINH HỌC.
           - Trích xuất: XN, KQXN, ĐVKQXN, Tham_chieu_KQXN, KET_LUAN: "Đạt".

        3. THẨM ĐỊNH KẾT QUẢ BẤT THƯỜNG:
           - Tìm 1 chỉ số có KẾT QUẢ NẰM NGOÀI KHOẢNG THAM CHIẾU (in đậm, cờ H/L cảnh báo).
           - Trích xuất: XN_BAT_THUONG, XN_BAT_THUONG_CO_KHONG ("Có"), KET_LUAN_XN_BAT_THUONG ("Đạt").
           - Nếu toàn bộ phiếu bình thường: XN_BAT_THUONG: "", XN_BAT_THUONG_CO_KHONG: "Không", KET_LUAN_XN_BAT_THUONG: "Không áp dụng".

        Trả về DUY NHẤT một chuỗi JSON thuần túy:
        {
            "SID": "...",
            "Ngay": "...",
            "Gio_nhan_mau": "...",
            "May": "...",
            "XN": "...",
            "KQXN": "...",
            "ĐVKQXN": "...",
            "Tham_chieu_KQXN": "...",
            "KET_LUAN": "Đạt",
            "XN_BAT_THUONG": "...",
            "XN_BAT_THUONG_CO_KHONG": "Có",
            "KET_LUAN_XN_BAT_THUONG": "Đạt"
        }
        """

    candidate_models = ["gemini-3.8-flash", "gemini-3.1-pro-preview"]
    last_error = None
    for model_name in candidate_models:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=[
                        types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
                        prompt
                    ]
                )
                raw_text = response.text.strip()
                raw_text = re.sub(r"^```json\s*", "", raw_text)
                raw_text = re.sub(r"^```\s*", "", raw_text)
                raw_text = re.sub(r"```$", "", raw_text).strip()
                return json.loads(raw_text)
            except Exception as e:
                last_error = e
                err_str = str(e)
                if "503" in err_str or "UNAVAILABLE" in err_str:
                    time.sleep(2)
                    continue
                else:
                    break
    raise last_error

# Khởi tạo session_state cho 5 bệnh nhân
for i in range(1, 6):
    if f"sid_{i}" not in st.session_state:
        st.session_state[f"sid_{i}"] = f"BN00{i}"
    if f"ngay_{i}" not in st.session_state:
        st.session_state[f"ngay_{i}"] = date.today().strftime("%d/%m/%Y")
    if f"gio_{i}" not in st.session_state:
        st.session_state[f"gio_{i}"] = f"0{7+i}:30"
    if f"may_{i}" not in st.session_state:
        st.session_state[f"may_{i}"] = "AU 5822_1"
    if f"xn_{i}" not in st.session_state:
        st.session_state[f"xn_{i}"] = "Glucose"
    if f"kq_{i}" not in st.session_state:
        st.session_state[f"kq_{i}"] = "5.6"
    if f"dv_{i}" not in st.session_state:
        st.session_state[f"dv_{i}"] = "mmol/L"
    if f"tc_{i}" not in st.session_state:
        st.session_state[f"tc_{i}"] = "3.9 - 6.4"
    if f"kl_{i}" not in st.session_state:
        st.session_state[f"kl_{i}"] = "Đạt"
    if f"xnbt_{i}" not in st.session_state:
        st.session_state[f"xnbt_{i}"] = "Ure / BUN"
    if f"btck_{i}" not in st.session_state:
        st.session_state[f"btck_{i}"] = "Có"
    if f"btkl_{i}" not in st.session_state:
        st.session_state[f"btkl_{i}"] = "Đạt"

if "sid_kqcs" not in st.session_state:
    st.session_state["sid_kqcs"] = "BN_TEST_01"
if "ngay_kqcs" not in st.session_state:
    st.session_state["ngay_kqcs"] = date.today().strftime("%d/%m/%Y")
if "thiet_bi_kqcs" not in st.session_state:
    st.session_state["thiet_bi_kqcs"] = "AU 5822_1"
if "xn_kqcs" not in st.session_state:
    st.session_state["xn_kqcs"] = "Urea"
if "ghi_chu_khuyen_cao" not in st.session_state:
    st.session_state["ghi_chu_khuyen_cao"] = "Mẫu tán huyết nhẹ, đã kiểm tra đối chứng."

# ================= GIAO DIỆN CHÍNH =================
user_info = st.session_state["user_info"]
is_admin = (user_info.get("role") == "Admin")

st.title("🧪 Phần mềm Thẩm định PXN & Nhận diện Ảnh AI")

# Sidebar
with st.sidebar:
    st.markdown(f"👤 **Xin chào:** {user_info.get('full_name', user_info['username'])}")
    st.markdown(f"🔰 **Vai trò:** `{user_info.get('role', 'User')}`")
    
    if st.button("🚪 Đăng xuất", use_container_width=True):
        st.session_state["logged_in"] = False
        st.session_state["user_info"] = None
        st.rerun()

    st.markdown("---")
    with st.expander("🔑 Đổi Mật Khẩu Cá Nhân"):
        old_pwd = st.text_input("Mật khẩu cũ:", type="password", key="chg_old")
        new_pwd = st.text_input("Mật khẩu mới:", type="password", key="chg_new")
        confirm_pwd = st.text_input("Nhập lại mật khẩu mới:", type="password", key="chg_conf")
        if st.button("Cập nhật mật khẩu", use_container_width=True):
            if hash_password(old_pwd) != user_info["password_hash"]:
                st.error("Mật khẩu cũ không chính xác!")
            elif not new_pwd or len(new_pwd) < 4:
                st.warning("Mật khẩu mới phải có ít nhất 4 ký tự!")
            elif new_pwd != confirm_pwd:
                st.error("Xác nhận mật khẩu mới không khớp!")
            else:
                ok, msg = change_user_password(user_info["username"], new_pwd)
                if ok:
                    st.success(msg)
                    st.session_state["user_info"]["password_hash"] = hash_password(new_pwd)
                else:
                    st.error(msg)

    st.markdown("---")
    st.subheader("🔑 Cấu hình Gemini AI")
    saved_key = load_saved_api_key()
    api_key_input = st.text_input(
        "Nhập Gemini API Key:",
        value=saved_key,
        type="password",
        help="Lấy API Key tại: https://aistudio.google.com/app/apikey"
    )
    col_save_k1, col_save_k2 = st.columns([1, 1])
    with col_save_k1:
        if st.button("💾 Lưu Key", use_container_width=True):
            if api_key_input:
                save_api_key_to_file(api_key_input)
                st.success("✅ Đã lưu API Key!")
            else:
                st.warning("Vui lòng nhập Key trước khi lưu.")
    with col_save_k2:
        if st.button("🗑️ Xóa Key", use_container_width=True):
            if os.path.exists(API_KEY_FILE):
                os.remove(API_KEY_FILE)
            st.info("Đã xóa Key lưu!")
            st.rerun()

    st.markdown("---")
    st.subheader("☁️ Google Drive Lưu Trữ")
    gdrive_folder_url = st.text_input(
        "Đường dẫn Thư mục Google Drive:",
        value="https://drive.google.com",
        help="Dán link thư mục Google Drive để mở nhanh"
    )
    if gdrive_folder_url:
        st.markdown(f"[📂 Mở Thư mục Google Drive]({gdrive_folder_url})", unsafe_allow_html=True)

# Danh sách tab (Nếu là Admin thì có thêm tab Quản lý tài khoản)
tab_titles = ["📸 Quét Ảnh Tự Động (AI)", "📝 Form Nhập Liệu", "📊 Quản Lý Dữ Liệu Đám Mây", "⚡ Xuất Báo Cáo"]
if is_admin:
    tab_titles.append("👥 Quản Trị Người Dùng")

tabs = st.tabs(tab_titles)

# ================= TAB 0: QUÉT ẢNH TỰ ĐỘNG =================
with tabs[0]:
    st.subheader("📷 Nạp ảnh kết quả xét nghiệm để AI tự động trích xuất")
    col_up1, col_up2 = st.columns([1, 1])
    with col_up1:
        target_choice = st.selectbox(
            "Chọn vị trí bạn muốn AI điền dữ liệu vào:",
            options=["bn_1", "bn_2", "bn_3", "bn_4", "bn_5", "khuyen_cao"],
            format_func=lambda x: {
                "bn_1": "Bệnh nhân 1 (Kết quả thông thường & Bất thường)",
                "bn_2": "Bệnh nhân 2 (Kết quả thông thường & Bất thường)",
                "bn_3": "Bệnh nhân 3 (Kết quả thông thường & Bất thường)",
                "bn_4": "Bệnh nhân 4 (Kết quả thông thường & Bất thường)",
                "bn_5": "Bệnh nhân 5 (Kết quả thông thường & Bất thường)",
                "khuyen_cao": "📌 Mục Xác nhận xét nghiệm được chỉnh sửa, khuyến cáo"
            }[x]
        )
        input_source = st.radio(
            "Phương thức nạp ảnh:",
            options=["📤 Tải file từ máy / điện thoại", "📷 Chụp trực tiếp bằng Camera", "🔗 Nhập link Google Drive"],
            horizontal=True
        )
        image_bytes = None
        if input_source == "📤 Tải file từ máy / điện thoại":
            uploaded_file = st.file_uploader("Chọn ảnh phiếu kết quả xét nghiệm (JPG, PNG, JPEG):", type=["jpg", "png", "jpeg"])
            if uploaded_file is not None:
                image_bytes = uploaded_file.getvalue()
                st.image(image_bytes, caption="Ảnh đã tải lên", use_container_width=True)
        elif input_source == "📷 Chụp trực tiếp bằng Camera":
            camera_pic = st.camera_input("Chụp ảnh phiếu xét nghiệm qua camera điện thoại / máy tính:")
            if camera_pic is not None:
                image_bytes = camera_pic.getvalue()
                st.image(image_bytes, caption="Ảnh vừa chụp", use_container_width=True)
        elif input_source == "🔗 Nhập link Google Drive":
            gdrive_img_url = st.text_input(
                "Dán đường link chia sẻ ảnh từ Google Drive:",
                placeholder="https://drive.google.com/file/d/.../view?usp=sharing"
            )
            if gdrive_img_url:
                try:
                    with st.spinner("Đang tải ảnh từ Google Drive..."):
                        image_bytes = get_bytes_from_gdrive_url(gdrive_img_url)
                        st.image(image_bytes, caption="Ảnh tải từ Google Drive thành công", use_container_width=True)
                except Exception as ex_drive:
                    st.error(f"Lỗi đọc ảnh từ Google Drive: {ex_drive}")

        if image_bytes is not None:
            st.success("✅ Ảnh đã sẵn sàng trong hệ thống.")
            st.markdown(f"👉 [Nhấp vào đây để mở Thư mục Google Drive và lưu trữ ảnh]({gdrive_folder_url})")

    with col_up2:
        st.markdown("**Kết quả AI nhận diện và trích xuất:**")
        if image_bytes is not None and st.button("🔍 Bắt đầu phân tích ảnh", use_container_width=True):
            active_key = api_key_input or saved_key
            if not active_key:
                st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh bên trái trước khi quét ảnh!")
            elif not HAS_GENAI:
                st.error("⚠️ Cần cài đặt thư viện google-genai!")
            else:
                is_recommendation = (target_choice == "khuyen_cao")
                spinner_msg = "AI đang trích xuất Mã BN, Ngày XN và Khuyến cáo..." if is_recommendation else "AI đang phân tích chỉ số thường và chỉ số bất thường..."
                with st.spinner(spinner_msg):
                    try:
                        mode = "recommendation" if is_recommendation else "patient"
                        extracted = extract_lab_data_from_image(image_bytes, active_key, mode=mode)
                        if is_recommendation:
                            if extracted.get("SID_KQCS"):
                                st.session_state["sid_kqcs"] = extracted.get("SID_KQCS")
                            if extracted.get("Ngay_KQCS"):
                                st.session_state["ngay_kqcs"] = extracted.get("Ngay_KQCS")
                            if extracted.get("Thiet_bi_KQCS"):
                                st.session_state["thiet_bi_kqcs"] = extracted.get("Thiet_bi_KQCS")
                            if extracted.get("XN_KQCS"):
                                st.session_state["xn_kqcs"] = extracted.get("XN_KQCS")
                            if extracted.get("Ghi_chu_Khuyen_cao"):
                                st.session_state["ghi_chu_khuyen_cao"] = extracted.get("Ghi_chu_Khuyen_cao")
                            st.success("✅ Đã cập nhật vào mục 'Xác nhận xét nghiệm được chỉnh sửa, khuyến cáo'!")
                        else:
                            p_idx = int(target_choice.split("_")[1])
                            st.session_state[f"sid_{p_idx}"] = extracted.get("SID", "")
                            st.session_state[f"ngay_{p_idx}"] = extracted.get("Ngay", "")
                            st.session_state[f"gio_{p_idx}"] = extracted.get("Gio_nhan_mau", "")
                            if extracted.get("May"):
                                st.session_state[f"may_{p_idx}"] = extracted.get("May", "")
                            st.session_state[f"xn_{p_idx}"] = extracted.get("XN", "")
                            st.session_state[f"kq_{p_idx}"] = extracted.get("KQXN", "")
                            st.session_state[f"dv_{p_idx}"] = extracted.get("ĐVKQXN", "")
                            st.session_state[f"tc_{p_idx}"] = extracted.get("Tham_chieu_KQXN", "")
                            st.session_state[f"kl_{p_idx}"] = extracted.get("KET_LUAN", "Đạt")
                            st.session_state[f"xnbt_{p_idx}"] = extracted.get("XN_BAT_THUONG", "")
                            st.session_state[f"btck_{p_idx}"] = extracted.get("XN_BAT_THUONG_CO_KHONG", "Có")
                            st.session_state[f"btkl_{p_idx}"] = extracted.get("KET_LUAN_XN_BAT_THUONG", "Đạt")
                            st.success(f"✅ Đã phân tích thành công cho Bệnh nhân {p_idx}!")
                        st.json(extracted)
                        st.info("👉 Hãy chuyển sang Tab '📝 Form Nhập Liệu' để kiểm tra lại và bấm Lưu.")
                    except Exception as err:
                        st.error(f"Lỗi phân tích ảnh: {err}")

# ================= TAB 1: FORM NHẬP LIỆU =================
with tabs[1]:
    with st.form("form_nhap_lieu"):
        st.subheader("1. Thông tin chung & Phần mềm LIS")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            ngay_danh_gia = st.date_input("Ngày đánh giá", value=date.today())
            ten_phan_mem = st.text_input("Tên phần mềm LIS", value="VNPT-LIS")
        with col2:
            hang_phan_mem = st.text_input("Hãng sản xuất", value="VNPT IT")
            phien_ban_phan_mem = st.text_input("Phiên bản", value="v3.2.1")
        with col3:
            so_mau = st.text_input("Cỡ mẫu kiểm tra", value="5")
            muc_dich_khac = st.text_input("Mục đích khác (nếu có)", value="Thẩm định định kỳ")
        with col4:
            nguoi_xac_nhan = st.text_input("Người xác nhận / Thẩm định", value=user_info.get("full_name", "Đỗ Như Hải"))
            ngay_xn_str = st.text_input("Ngày xác nhận (dd/mm/yyyy)", value=ngay_danh_gia.strftime("%d/%m/%Y"))

        st.markdown("---")
        st.subheader("2. Thẩm định kết quả thông thường (Chỉ số Sinh hóa / Miễn dịch bình thường)")
        p_cols = st.columns(5)
        patient_data = {}
        for i in range(1, 6):
            with p_cols[i-1]:
                st.markdown(f"**Bệnh nhân {i}**")
                patient_data[f"SID_{i}"] = st.text_input(f"Mã SID {i}", value=st.session_state[f"sid_{i}"])
                patient_data[f"Ngay_{i}"] = st.text_input(f"Ngày XN {i}", value=st.session_state[f"ngay_{i}"])
                patient_data[f"Gio_nhan_mau_{i}"] = st.text_input(f"Giờ nhận {i}", value=st.session_state[f"gio_{i}"])
                patient_data[f"May_{i}"] = st.text_input(f"Thiết bị {i}", value=st.session_state[f"may_{i}"])
                patient_data[f"Co_Khong_{i}"] = st.selectbox(f"LIS chính xác?", ["Có", "Không"], key=f"form_ck_{i}")
                patient_data[f"XN_{i}"] = st.text_input(f"Tên XN (SH/MD) {i}", value=st.session_state[f"xn_{i}"])
                patient_data[f"KQXN_{i}"] = st.text_input(f"Kết quả {i}", value=st.session_state[f"kq_{i}"])
                patient_data[f"ĐVKQXN_{i}"] = st.text_input(f"Đơn vị {i}", value=st.session_state[f"dv_{i}"])
                patient_data[f"Tham_chieu_KQXN_{i}"] = st.text_input(f"Tham chiếu {i}", value=st.session_state[f"tc_{i}"])
                default_kl_idx = 0 if st.session_state[f"kl_{i}"] == "Đạt" else 1
                patient_data[f"KET_LUAN_XN_{i}"] = st.selectbox(f"Kết luận {i}", ["Đạt", "Không đạt"], index=default_kl_idx, key=f"form_kl_{i}")

        st.markdown("---")
        st.subheader("3. Thẩm định xét nghiệm BẤT THƯỜNG / BÁO ĐỘNG (Ngoài khoảng tham chiếu, in đậm)")
        bt_cols = st.columns(5)
        bt_data = {}
        for i in range(1, 6):
            with bt_cols[i-1]:
                st.markdown(f"**Bệnh nhân {i} (Báo động)**")
                bt_data[f"XN_BAT_THUONG_{i}"] = st.text_input(f"XN Báo động (in đậm) {i}", value=st.session_state[f"xnbt_{i}"])
                default_btck_idx = 0 if st.session_state[f"btck_{i}"] == "Có" else 1
                bt_data[f"XN_BAT_THUONG_CO_KHONG_{i}"] = st.selectbox(f"Đánh dấu nhận biết?", ["Có", "Không"], index=default_btck_idx, key=f"form_btck_{i}")
                default_btkl_idx = 0 if st.session_state[f"btkl_{i}"] == "Đạt" else 1
                bt_data[f"KET_LUAN_XN_BAT_THUONG_{i}"] = st.selectbox(f"Kết luận BĐ {i}", ["Đạt", "Không đạt"], index=default_btkl_idx, key=f"form_btkl_{i}")

        st.markdown("---")
        st.subheader("4. Thẩm định xét nghiệm CHỈNH SỬA & KHUYẾN CÁO")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            sid_kqcs = st.text_input("Mã BN (giả định)", value=st.session_state["sid_kqcs"])
            ngay_kqcs = st.text_input("Ngày XN (giả định)", value=st.session_state["ngay_kqcs"])
            thiet_bi_kqcs = st.text_input("Tên thiết bị XN", value=st.session_state["thiet_bi_kqcs"])
            xn_kqcs = st.text_input("Loại xét nghiệm", value=st.session_state["xn_kqcs"])
        with c2:
            gia_tri_dau = st.text_input("Giá trị ban đầu", value="15.2")
            gia_tri_cs1 = st.text_input("Giá trị sửa lần 1", value="15.8")
            gia_tri_cs2 = st.text_input("Giá trị sửa lần 2", value="15.5")
        with c3:
            ghi_chu_khuyen_cao = st.text_area("Thông tin khuyến cáo (hiển thị trong ghi chú)", value=st.session_state["ghi_chu_khuyen_cao"])
            lis_co_khong = st.selectbox("LIS thể hiện đủ thời gian/người sửa?", ["Có", "Không"])
        with c4:
            kl_chinh_sua = st.selectbox("Kết luận chỉnh sửa", ["Đạt", "Không đạt"])

        st.markdown("---")
        st.subheader("5. Đánh giá tỷ lệ sai sót & Tình huống")
        s1, s2 = st.columns(2)
        with s1:
            danh_gia_1 = st.selectbox("1. Lấy thông tin & ghi nhận mẫu", ["Đạt", "Không đạt"])
            danh_gia_2 = st.selectbox("2. Đổ thông tin LIS lên máy", ["Đạt", "Không đạt"])
            danh_gia_3 = st.selectbox("3. Máy nhận đủ thông tin để chạy", ["Đạt", "Không đạt"])
            danh_gia_4 = st.selectbox("4. Xuất KQ từ máy vào LIS", ["Đạt", "Không đạt"])
            danh_gia_5 = st.selectbox("5. In kết quả ra phiếu", ["Đạt", "Không đạt"])
        with s2:
            tinh_huong_1 = st.text_input("Tình huống 1", value="Mất kết nối mạng LAN đột ngột")
            tinh_huong_2 = st.text_input("Tình huống 2", value="Trùng lặp mã Barcode tiếp nhận")
            tinh_huong_3 = st.text_input("Tình huống 3", value="Sửa kết quả sau khi duyệt sơ bộ")
            tinh_huong_4 = st.text_input("Tình huống 4", value="In lại phiếu kết quả sau 24h")

        st.markdown("---")
        st.subheader("6. Nội dung thẩm định & Kết luận chung")
        k1, k2 = st.columns(2)
        with k1:
            kq_thong_thuong = st.selectbox("Xác nhận KQ thông thường", ["Đạt", "Không đạt"])
            kq_bao_dong = st.selectbox("Xác nhận KQ báo động", ["Đạt", "Không đạt"])
            tt_sua_doi = st.selectbox("Xác nhận thông tin sửa đổi", ["Đạt", "Không đạt"])
            tinh_toan = st.selectbox("Xác nhận tính toán/chuyển đổi", ["Đạt", "Không đạt"])
        with k2:
            de_xuat = st.text_area("Đề xuất / nhận xét", value="Hệ thống đáp ứng tốt các yêu cầu vận hành, kết nối 2 chiều ổn định.")
            ket_luan_lis = st.selectbox("Kết luận: Sử dụng hệ thống LIS", ["Đạt", "Không đạt"])

        submitted = st.form_submit_button("💾 LƯU DỮ LIỆU LÊN ĐÁM MÂY", use_container_width=True)

    if submitted:
        row_payload = {
            "Day": str(ngay_danh_gia.day),
            "Month": str(ngay_danh_gia.month),
            "year": str(ngay_danh_gia.year),
            "Ngay": ngay_danh_gia.strftime("%d/%m/%Y"),
            "So_mau": so_mau,
            "Nguoi_Xac_Nhan": nguoi_xac_nhan,
            "Ngay_xac_nhan": ngay_xn_str,
            "Ten_phan_mem": ten_phan_mem,
            "Hang_phan_mem": hang_phan_mem,
            "Phien_ban_phan_mem": phien_ban_phan_mem,
            "muc_dich_khac": muc_dich_khac,
            **patient_data,
            **bt_data,
            "SID_KQCS": sid_kqcs,
            "Ngay_KQCS": ngay_kqcs,
            "Thiet_bi_KQCS": thiet_bi_kqcs,
            "XN_KQCS": xn_kqcs,
            "Gia_tri_dau_tien_KQCS": gia_tri_dau,
            "Gia_tri_CS_lan_1": gia_tri_cs1,
            "Gia_tri_CS_lan_2": gia_tri_cs2,
            "Ghi_chu_Khuyen_cao": ghi_chu_khuyen_cao,
            "Lis_co_khong": lis_co_khong,
            "Ket_luan_chinh_sưa_dat_khong": kl_chinh_sua,
            "xac_nhan_KQ_thong_thuong": kq_thong_thuong,
            "xac_nhan_KQ_bao_dong": kq_bao_dong,
            "xac_nhan_TT_sua_doi": tt_sua_doi,
            "xac_nhan_tinh_toan": tinh_toan,
            "De_xuat": de_xuat,
            "Ket_luan_sư_dung_LIS_dat": ket_luan_lis,
            "Danh_gia_1": danh_gia_1,
            "Danh_gia_2": danh_gia_2,
            "Danh_gia_3": danh_gia_3,
            "Danh_gia_4": danh_gia_4,
            "Danh_gia_5": danh_gia_5,
            "Tinh_huong_1": tinh_huong_1,
            "Tinh_huong_2": tinh_huong_2,
            "Tinh_huong_3": tinh_huong_3,
            "Tinh_huong_4": tinh_huong_4,
        }
        save_single_record(row_payload)
        st.success("✅ Đã lưu phiếu thành công lên Đám mây!")
        if os.path.exists(TEMPLATE_FILE):
            try:
                word_buffer = export_single_word(row_payload)
                st.download_button(
                    label="📥 Tải ngay Phiếu Word (.docx) của bản ghi này",
                    data=word_buffer,
                    file_name=f"Ket_qua_{clean_filename(ten_phan_mem)}_{clean_filename(ngay_xn_str)}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Lỗi xuất Word: {e}")

# ================= TAB 2: XEM & QUẢN LÝ DỮ LIỆU ĐÁM MÂY =================
with tabs[2]:
    st.subheader("📊 Quản lý Dữ liệu Đám mây & Khôi phục")
    df_current = fetch_all_records_df()
    
    if is_admin:
        with st.expander("📤 NẠP / KHÔI PHỤC DỮ LIỆU TỪ FILE EXCEL (Chỉ Admin)", expanded=False):
            uploaded_excel = st.file_uploader(
                "Chọn file Excel sao lưu (.xlsx hoặc .xls):",
                type=["xlsx", "xls"],
                key="excel_uploader"
            )
            if uploaded_excel is not None:
                try:
                    df_imported = pd.read_excel(uploaded_excel, dtype=str).fillna("")
                    st.write(f"🔍 Đã phát hiện **{len(df_imported)}** dòng dữ liệu trong file:")
                    st.dataframe(df_imported.head(5), use_container_width=True)
                    upload_action = st.radio(
                        "Chọn phương thức nạp dữ liệu:",
                        options=["➕ Ghi thêm (Nối tiếp dữ liệu mới)", "🔄 Ghi đè (Xóa sạch dữ liệu cũ và thay thế)"]
                    )
                    if st.button("🚀 Xác nhận Nạp vào Hệ thống", use_container_width=True):
                        if "Ghi đè" in upload_action:
                            delete_all_records()
                        for _, r in df_imported.iterrows():
                            save_single_record(r.to_dict())
                        st.success("✅ Đã cập nhật thành công dữ liệu!")
                        st.rerun()
                except Exception as e:
                    st.error(f"Lỗi đọc file Excel: {e}")

    st.markdown("### Dữ liệu hiện có trong Hệ thống")
    if not df_current.empty:
        st.write(f"Tổng số bản ghi hiện có: **{len(df_current)}**")
        st.dataframe(df_current, use_container_width=True)
        
        excel_buffer = io.BytesIO()
        df_current.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)
        st.download_button(
            label="📥 Tải toàn bộ file Excel về máy tính (.xlsx)",
            data=excel_buffer,
            file_name=f"danh_sach_phieu_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # Chỉ Admin mới thấy khu vực xóa dữ liệu
        if is_admin:
            st.markdown("---")
            st.markdown("### 🔒 Khu Vực Xóa Dữ Liệu (Dành cho Quản trị viên)")
            unique_dates_del = set()
            for _, r in df_current.iterrows():
                d1 = str(r.get("Ngay", "")).strip()
                d2 = str(r.get("Ngay_xac_nhan", "")).strip()
                if d1:
                    unique_dates_del.add(d1)
                if d2:
                    unique_dates_del.add(d2)
            list_dates_del = sorted(list(unique_dates_del))

            col_del1, col_del2 = st.columns([1, 1])
            with col_del1:
                st.markdown("**1. Xóa dữ liệu của 1 ngày bất kỳ**")
                if list_dates_del:
                    del_date_chosen = st.selectbox("Chọn ngày muốn xóa:", options=list_dates_del, key="select_del_date")
                    if st.button(f"❌ Xóa toàn bộ phiếu ngày {del_date_chosen}", use_container_width=True):
                        deleted_count = delete_records_by_date(del_date_chosen)
                        st.success(f"✅ Đã xóa {deleted_count} bản ghi của ngày {del_date_chosen}!")
                        st.rerun()
                else:
                    st.info("Không có ngày để xóa.")

            with col_del2:
                st.markdown("**2. Xóa toàn bộ dữ liệu (Làm mới hoàn toàn)**")
                if st.button("⚠️ XÓA SẠCH TẤT CẢ DỮ LIỆU", use_container_width=True):
                    delete_all_records()
                    st.warning("✅ Đã xóa sạch toàn bộ cơ sở dữ liệu!")
                    st.rerun()
        else:
            st.caption("ℹ️ *Bạn đang đăng nhập bằng tài khoản Người dùng (User). Chỉ Admin mới có quyền xóa dữ liệu.*")
    else:
        st.info("Chưa có dữ liệu nào được lưu trong hệ thống.")

# ================= TAB 3: XUẤT BÁO CÁO (WORD / PDF / ZIP) =================
with tabs[3]:
    st.subheader("⚡ Tùy chọn Xuất Báo Cáo")
    df_all = fetch_all_records_df()
    if df_all.empty:
        st.warning("⚠️ Chưa có dữ liệu trong hệ thống! Vui lòng nhập dữ liệu trước.")
    elif not os.path.exists(TEMPLATE_FILE):
        st.error(f"⚠️ Không tìm thấy file mẫu '{TEMPLATE_FILE}'!")
    else:
        unique_dates = set()
        for _, r in df_all.iterrows():
            d1 = str(r.get("Ngay", "")).strip()
            d2 = str(r.get("Ngay_xac_nhan", "")).strip()
            if d1:
                unique_dates.add(d1)
            if d2:
                unique_dates.add(d2)
        date_list = sorted(list(unique_dates))
        
        scope_choice = st.radio(
            "Phạm vi xuất phiếu:",
            options=["Toàn bộ tất cả phiếu", "Lọc theo ngày"],
            horizontal=True
        )
        selected_dates = []
        is_single_date = False
        if scope_choice == "Lọc theo ngày":
            if date_list:
                selected_dates = st.multiselect(
                    "Chọn ngày cần xuất (có thể chọn 1 ngày hoặc nhiều ngày):",
                    options=date_list,
                    default=[date_list[-1]]
                )
            else:
                input_d = st.text_input("Nhập ngày cần xuất (dd/mm/yyyy):", value=date.today().strftime("%d/%m/%Y"))
                selected_dates = [input_d] if input_d else []
            is_single_date = (len(selected_dates) == 1)
        else:
            is_single_date = False

        st.markdown("---")
        if is_single_date:
            st.info(f"📌 Đang chọn đúng **1 ngày** ({selected_dates[0]}): Hệ thống sẽ **xuất bản riêng trực tiếp (không đóng gói ZIP)**.")
        else:
            if scope_choice == "Toàn bộ tất cả phiếu":
                st.info("📦 Đang chọn **Toàn bộ tất cả phiếu**: Hệ thống sẽ **đóng gói file ZIP**.")
            else:
                st.info(f"📦 Đang chọn **{len(selected_dates)} ngày**: Hệ thống sẽ **đóng gói file ZIP**.")

        file_format = st.radio(
            "Chọn định dạng xuất:",
            options=["Word (.docx)", "PDF (.pdf)"],
            horizontal=True
        )

        if st.button("🚀 Bắt đầu Tạo & Tải File", use_container_width=True):
            if scope_choice == "Lọc theo ngày":
                if not selected_dates:
                    st.warning("Vui lòng chọn ít nhất 1 ngày!")
                    st.stop()
                clean_selected = [str(d).strip() for d in selected_dates]
                mask1 = df_all.get("Ngay", "").str.strip().isin(clean_selected)
                mask2 = df_all.get("Ngay_xac_nhan", "").str.strip().isin(clean_selected)
                filtered_df = df_all[mask1 | mask2]
            else:
                filtered_df = df_all

            total_records = len(filtered_df)
            if total_records == 0:
                st.warning("Không tìm thấy phiếu nào phù hợp với điều kiện đã chọn!")
                st.stop()

            with st.spinner("Đang kết xuất dữ liệu và tạo tệp..."):
                if is_single_date and total_records == 1:
                    row = filtered_df.iloc[0]
                    ctx = row.to_dict()
                    pm = clean_filename(row.get("Ten_phan_mem", "Phieu"))
                    ng = clean_filename(row.get("Ngay", clean_filename(row.get("Ngay_xac_nhan", ""))))
                    base_name = f"Ket_qua_{pm}_{ng}"
                    word_bytes = export_single_word(ctx, TEMPLATE_FILE).getvalue()

                    if file_format == "Word (.docx)":
                        st.success(f"✅ Đã tạo thành công phiếu Word ngày {selected_dates[0]}!")
                        st.download_button(
                            label=f"📥 TẢI XUỐNG FILE WORD: {base_name}.docx",
                            data=word_bytes,
                            file_name=f"{base_name}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    else:
                        pdf_bytes = convert_docx_to_pdf_bytes(word_bytes)
                        if pdf_bytes:
                            st.success(f"✅ Đã tạo thành công phiếu PDF ngày {selected_dates[0]}!")
                            st.download_button(
                                label=f"📥 TẢI XUỐNG FILE PDF: {base_name}.pdf",
                                data=pdf_bytes,
                                file_name=f"{base_name}.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )
                        else:
                            st.warning("⚠️ Môi trường chưa cài LibreOffice. Cung cấp file Word:")
                            st.download_button(
                                label=f"📥 TẢI XUỐNG FILE WORD: {base_name}.docx",
                                data=word_bytes,
                                file_name=f"{base_name}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                else:
                    zip_buffer = io.BytesIO()
                    count = 0
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for idx, r in filtered_df.iterrows():
                            ctx = r.to_dict()
                            pm = clean_filename(r.get("Ten_phan_mem", f"Phieu_{idx+1}"))
                            ng = clean_filename(r.get("Ngay", clean_filename(r.get("Ngay_xac_nhan", ""))))
                            file_code = f"Ket_qua_{pm}_{ng}_{idx+1}"
                            word_bytes = export_single_word(ctx, TEMPLATE_FILE).getvalue()

                            if file_format == "Word (.docx)":
                                zip_file.writestr(f"{file_code}.docx", word_bytes)
                                count += 1
                            else:
                                pdf_bytes = convert_docx_to_pdf_bytes(word_bytes)
                                if pdf_bytes:
                                    zip_file.writestr(f"{file_code}.pdf", pdf_bytes)
                                else:
                                    zip_file.writestr(f"{file_code}.docx", word_bytes)
                                count += 1

                    zip_buffer.seek(0)
                    tag_name = "Toan_Bo" if scope_choice == "Toàn bộ tất cả phiếu" else f"{len(selected_dates)}_Ngay"
                    fmt_tag = "PDF" if file_format == "PDF (.pdf)" else "Word"
                    zip_filename = f"Ket_qua_PXN_{tag_name}_{fmt_tag}.zip"

                    st.success(f"🎉 Đã đóng gói thành công {count} phiếu vào file ZIP ({file_format})!")
                    st.download_button(
                        label=f"📥 TẢI VỀ GÓI NÉN ZIP: {zip_filename}",
                        data=zip_buffer.getvalue(),
                        file_name=zip_filename,
                        mime="application/zip",
                        use_container_width=True
                    )

# ================= TAB 4: QUẢN TRỊ NGƯỜI DÙNG (CHỈ ADMIN) =================
if is_admin:
    with tabs[4]:
        st.subheader("👥 Quản Trị Tài Khoản & Phân Quyền")
        
        col_u1, col_u2 = st.columns([1, 1.2])
        with col_u1:
            st.markdown("#### Tạo Tài Khoản Mới")
            with st.form("create_user_form"):
                new_u = st.text_input("Tên đăng nhập mới (username):")
                new_p = st.text_input("Mật khẩu ban đầu:", type="password")
                new_fn = st.text_input("Họ và tên nhân viên:")
                new_role = st.selectbox("Phân quyền sử dụng:", ["User", "Admin"])
                btn_create = st.form_submit_button("➕ Tạo người dùng", use_container_width=True)
                
            if btn_create:
                if not new_u or not new_p:
                    st.warning("Vui lòng điền đầy đủ tên đăng nhập và mật khẩu!")
                else:
                    ok_u, msg_u = create_new_user(new_u, new_p, new_role, new_fn)
                    if ok_u:
                        st.success(msg_u)
                        st.rerun()
                    else:
                        st.error(msg_u)

        with col_u2:
            st.markdown("#### Danh Sách Người Dùng Hiện Có")
            df_users = fetch_all_users()
            if not df_users.empty:
                st.dataframe(df_users, use_container_width=True)
            else:
                st.info("Chưa có danh sách người dùng trên Supabase.")