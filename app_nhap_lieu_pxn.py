# -*- coding: utf-8 -*-
"""
Web App Nhập liệu thẩm định LIS & Xuất báo cáo Word / Excel.
Cập nhật logic Xuất báo cáo:
 - Xuất toàn bộ HOẶC chọn từ 2 ngày trở lên -> Đóng gói ZIP (Tùy chọn Word hoặc PDF).
 - Chọn đúng 1 ngày -> Xuất bản riêng trực tiếp (không nén ZIP), có nút tùy chọn tải Word hoặc PDF.
"""

import os
import io
import re
import json
import zipfile
import subprocess
import tempfile
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

st.set_page_config(
    page_title="Nhập liệu & Thẩm định PXN - AI Vision",
    page_icon="🧪",
    layout="wide"
)

def clean_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", str(filename)).strip()

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
    """
    Chuyển đổi DOCX sang PDF bằng công cụ headless (hỗ trợ LibreOffice trên Streamlit Cloud / Linux / Windows).
    Nếu môi trường chưa cài LibreOffice, hàm sẽ trả về None để xử lý fallback.
    """
    with tempfile.TemporaryDirectory() as tmp_dir:
        input_docx = os.path.join(tmp_dir, "input.docx")
        with open(input_docx, "wb") as f:
            f.write(docx_bytes)
        
        # Thử lệnh libreoffice hoặc soffice
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

    fallback_models = ["gemini-3.8-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
    last_error = None
    for model_name in fallback_models:
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
            continue

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
st.markdown("Hỗ trợ quét ảnh phiếu xét nghiệm, lưu trữ vào **Excel** và xuất **Báo cáo Word / PDF**.")

with st.sidebar:
    st.header("🔑 Cấu hình AI Vision")
    api_key_input = st.text_input(
        "Nhập Gemini API Key:",
        type="password",
        help="Lấy API Key tại: https://aistudio.google.com/app/apikey"
    )
    st.info("💡 Bạn có thể quét ảnh cho từng Bệnh nhân (1-5) hoặc quét riêng cho phiếu có Ghi chú/Khuyến cáo.")

tabs = st.tabs(["📸 Quét Ảnh Tự Động (AI)", "📝 Form Nhập Liệu", "📊 Xem & Tải Excel", "⚡ Xuất Báo Cáo"])

# ================= TAB 0: QUÉT ẢNH TỰ ĐỘNG =================
with tabs[0]:
    st.subheader("📷 Tải ảnh kết quả xét nghiệm để AI tự động trích xuất")
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
        uploaded_file = st.file_uploader("Chọn ảnh phiếu kết quả xét nghiệm (JPG, PNG, JPEG):", type=["jpg", "png", "jpeg"])
        if uploaded_file is not None:
            image = Image.open(uploaded_file)
            st.image(image, caption="Ảnh phiếu xét nghiệm đã chọn", use_container_width=True)

    with col_up2:
        st.markdown("**Kết quả AI nhận diện và trích xuất:**")
        if uploaded_file and st.button("🔍 Bắt đầu phân tích ảnh", use_container_width=True):
            if not api_key_input:
                st.error("⚠️ Vui lòng nhập Gemini API Key ở thanh bên trái trước khi quét ảnh!")
            elif not HAS_GENAI:
                st.error("⚠️ Cần cài đặt thư viện google-genai: chạy lệnh `python -m pip install google-genai`")
            else:
                is_recommendation = (target_choice == "khuyen_cao")
                spinner_msg = "AI đang trích xuất Mã BN, Ngày XN và Khuyến cáo..." if is_recommendation else "AI đang phân tích chỉ số thường và chỉ số bất thường..."
                
                with st.spinner(spinner_msg):
                    try:
                        bytes_data = uploaded_file.getvalue()
                        mode = "recommendation" if is_recommendation else "patient"
                        extracted = extract_lab_data_from_image(bytes_data, api_key_input, mode=mode)
                        
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
        
        append_to_excel(row_payload)
        st.success("✅ Đã lưu phiếu thành công vào file Excel!")
        
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
        else:
            st.warning(f"Chưa tìm thấy file '{TEMPLATE_FILE}' trong thư mục để tạo nút tải Word.")

# ================= TAB 2: XEM & QUẢN LÝ DỮ LIỆU EXCEL =================
with tabs[2]:
    st.subheader("Dữ liệu đã lưu trong file Excel")
    if os.path.exists(EXCEL_FILE):
        try:
            df_current = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
            st.write(f"Tổng số bản ghi hiện có: **{len(df_current)}**")
            st.dataframe(df_current, use_container_width=True)
            
            with open(EXCEL_FILE, "rb") as f:
                st.download_button(
                    label="📥 Tải toàn bộ file Excel về máy",
                    data=f,
                    file_name=EXCEL_FILE,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Lỗi khi đọc file Excel: {e}")
    else:
        st.info("Chưa có dữ liệu nào được lưu.")

# ================= TAB 3: XUẤT BÁO CÁO (WORD / PDF / ZIP) =================
with tabs[3]:
    st.subheader("⚡ Tùy chọn Xuất Báo Cáo")
    
    if not os.path.exists(EXCEL_FILE):
        st.warning("⚠️ Chưa có file Excel dữ liệu! Vui lòng nhập dữ liệu trước.")
    elif not os.path.exists(TEMPLATE_FILE):
        st.error(f"⚠️ Không tìm thấy file mẫu '{TEMPLATE_FILE}'!")
    else:
        df_all = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
        
        # Lấy danh sách ngày có trong Excel
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
        
        # Hiển thị thông báo logic tương ứng
        if is_single_date:
            st.info(f"📌 Đang chọn đúng **1 ngày** ({selected_dates[0]}): Hệ thống sẽ **xuất bản riêng trực tiếp (không đóng gói ZIP)**.")
        else:
            if scope_choice == "Toàn bộ tất cả phiếu":
                st.info("📦 Đang chọn **Toàn bộ tất cả phiếu**: Hệ thống sẽ **đóng gói file ZIP**.")
            else:
                st.info(f"📦 Đang chọn **{len(selected_dates)} ngày**: Hệ thống sẽ **đóng gói file ZIP**.")

        # Tùy chọn định dạng
        file_format = st.radio(
            "Chọn định dạng xuất:",
            options=["Word (.docx)", "PDF (.pdf)"],
            horizontal=True
        )

        btn_label = "🚀 Bắt đầu Tạo & Tải File"
        if st.button(btn_label, use_container_width=True):
            # Lọc dữ liệu theo phạm vi
            if scope_choice == "Lọc theo ngày":
                if not selected_dates:
                    st.warning("Vui lòng chọn ít nhất 1 ngày!")
                    st.stop()
                filter_mask = df_all.get("Ngay", "").str.strip().isin([d.strip() for d in selected_dates]) | \
                              df_all.get("Ngay_xac_nhan", "").str.strip().isin([d.strip() for d in selected_dates])
                filtered_df = df_all[filter_mask]
            else:
                filtered_df = df_all

            total_records = len(filtered_df)
            if total_records == 0:
                st.warning("Không tìm thấy phiếu nào phù hợp với điều kiện đã chọn!")
                st.stop()

            with st.spinner("Đang kết xuất dữ liệu và tạo tệp..."):
                # ================= TRƯỜNG HỢP 1: XUẤT 1 BẢN RIÊNG (CHỌN 1 NGÀY VÀ RA 1 BẢN) =================
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
                            st.warning("⚠️ Môi trường chưa cài LibreOffice để chuyển đổi tự động sang PDF. Đang cung cấp file Word:")
                            st.download_button(
                                label=f"📥 TẢI XUỐNG FILE WORD: {base_name}.docx",
                                data=word_bytes,
                                file_name=f"{base_name}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )

                # ================= TRƯỜNG HỢP 2: ĐÓNG GÓI ZIP (TOÀN BỘ HOẶC >= 2 NGÀY HOẶC NHIỀU BẢN GHI) =================
                else:
                    zip_buffer = io.BytesIO()
                    count = 0
                    pdf_failed_notice = False

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
                                    count += 1
                                else:
                                    # Fallback lưu word nếu không convert được pdf
                                    zip_file.writestr(f"{file_code}.docx", word_bytes)
                                    count += 1
                                    pdf_failed_notice = True

                    zip_buffer.seek(0)
                    tag_name = "Toan_Bo" if scope_choice == "Toàn bộ tất cả phiếu" else f"{len(selected_dates)}_Ngay"
                    fmt_tag = "PDF" if file_format == "PDF (.pdf)" else "Word"
                    zip_filename = f"Ket_qua_PXN_{tag_name}_{fmt_tag}.zip"

                    if pdf_failed_notice:
                        st.info("💡 Lưu ý: Do môi trường chưa có công cụ LibreOffice, các file trong gói ZIP được lưu tạm dưới dạng Word (.docx).")

                    st.success(f"🎉 Đã đóng gói thành công {count} phiếu vào file ZIP ({file_format})!")
                    st.download_button(
                        label=f"📥 TẢI VỀ GÓI NÉN ZIP: {zip_filename}",
                        data=zip_buffer.getvalue(),
                        file_name=zip_filename,
                        mime="application/zip",
                        use_container_width=True
                    )
