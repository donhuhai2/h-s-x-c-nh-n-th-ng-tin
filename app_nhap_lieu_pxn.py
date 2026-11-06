# -*- coding: utf-8 -*-
"""
Web App Nhập liệu thẩm định LIS & Xuất báo cáo Word / Excel.
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
from datetime import date
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

EXCEL_FILE = "danh_sach_phieu.xlsx"
TEMPLATE_FILE = "template.docx"
API_KEY_FILE = "api_key_config.txt"
ADMIN_SECURITY_PIN = "175175"

st.set_page_config(
    page_title="Nhập liệu & Thẩm định PXN - AI Vision",
    page_icon="🧪",
    layout="wide"
)

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

def append_to_excel(row_dict, file_path=EXCEL_FILE):
    new_df = pd.DataFrame([row_dict])
    if os.path.exists(file_path):
        try:
            old_df = pd.read_excel(file_path, dtype=str)
            combined = pd.concat([old_df, new_df], ignore_index=True)
        except Exception:
            combined = new_df
    else:
        combined = new_df
    combined.to_excel(file_path, index=False)

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
        raise ValueError("Đường dẫn Google Drive không hợp lệ! Vui lòng đảm bảo link có dạng: [https://drive.google.com/file/d/.../view](https://drive.google.com/file/d/.../view)")
    
    file_id = match.group(1)
    direct_url = f"[https://drive.google.com/uc?export=download&id=](https://drive.google.com/uc?export=download&id=){file_id}"
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

st.title("🧪 Phần mềm Thẩm định PXN & Nhận diện Ảnh AI")
st.markdown("Hỗ trợ quét ảnh phiếu xét nghiệm (Upload / Camera / Google Drive), quản lý cơ sở dữ liệu **Excel** và xuất **Báo cáo Word / PDF**.")

with st.sidebar:
    st.header("🔑 Cấu hình Hệ thống")
    
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

    if saved_key:
        st.caption("🔒 *API Key đang được ghi nhớ tự động.*")

    st.markdown("---")
    st.subheader("☁️ Google Drive Lưu Trữ")
    gdrive_folder_url = st.text_input(
        "Đường dẫn Thư mục Google Drive lưu ảnh:",
        value="https://drive.google.com",
        help="Dán đường link thư mục Google Drive của bạn để mở nhanh và lưu trữ ảnh xét nghiệm"
    )
    if gdrive_folder_url:
        st.markdown(f"[📂 Mở Thư mục Google Drive của bạn]({gdrive_folder_url})", unsafe_allow_html=True)

tabs = st.tabs(["📸 Quét Ảnh Tự Động (AI)", "📝 Form Nhập Liệu", "📊 Xem & Quản Lý Excel", "⚡ Xuất Báo Cáo"])

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
                    st.error(f"Lỗi đọc ảnh từ Google Drive: {ex_drive}. Hãy chắc chắn quyền truy cập file trên Drive là 'Bất kỳ ai có đường liên kết'.")

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
                st.error("⚠️ Cần cài đặt thư viện google-genai: chạy lệnh `python -m pip install google-genai`")
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
            nguoi_xac_nhan = st.text_input("Người xác nhận / Thẩm định", value="Đỗ Như Hải")
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
            st.markdown("**Đánh giá hoạt động**")
            danh_gia_1 = st.selectbox("1. Lấy thông tin & ghi nhận mẫu", ["Đạt", "Không đạt"])
            danh_gia_2 = st.selectbox("2. Đổ thông tin LIS lên máy", ["Đạt", "Không đạt"])
            danh_gia_3 = st.selectbox("3. Máy nhận đủ thông tin để chạy", ["Đạt", "Không đạt"])
            danh_gia_4 = st.selectbox("4. Xuất KQ từ máy vào LIS", ["Đạt", "Không đạt"])
            danh_gia_5 = st.selectbox("5. In kết quả ra phiếu", ["Đạt", "Không đạt"])
        with s2:
            st.markdown("**Tình huống thẩm định**")
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

        submitted = st.form_submit_button("💾 LƯU DỮ LIỆU VÀO EXCEL", use_container_width=True)

    if submitted:
        row_payload = {
            "Day": str(ngay_danh_gia.day),
            "Month": str(ngay_danh_gia.month),
            "year": str(ngay_danh_gia.year),
            "Ngay": ngay_danh_gia.strftime("%d/%m/%Y"),
            "So_mau": so_mau,
            "Nguoi_Xac_Nhan": nguoi_xac_nhan,
            "Ngay_