# -*- coding: utf-8 -*-
"""
Web App Nhập liệu thẩm định LIS & Xuất báo cáo Word / Excel.
Cập nhật:
 - Cho phép chọn vị trí điền dữ liệu từ ảnh:
   + Bệnh nhân 1 -> 5 (Thẩm định thông thường & bất thường)
   + "Xác nhận xét nghiệm được chỉnh sửa, khuyến cáo" (Mã BN, Ngày XN, Thiết bị, Tên XN, và Thông tin khuyến cáo/ghi chú)
"""

import os
import io
import re
import json
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

def extract_lab_data_from_image(image_bytes, api_key, mode="patient"):
    """
    Sử dụng Gemini Vision để phân tích phiếu xét nghiệm theo 2 chế độ:
    1. mode="patient": Bệnh nhân 1-5 (Chỉ số thường + Chỉ số bất thường in đậm ngoài khoảng tham chiếu).
    2. mode="recommendation": Mục Chỉnh sửa, khuyến cáo (Mã BN, Ngày XN, Thiết bị, Tên XN, Nội dung khuyến cáo/ghi chú).
    """
    client = genai.Client(api_key=api_key)
    
    if mode == "recommendation":
        prompt = """
        Bạn là chuyên gia thẩm định hồ sơ kết quả xét nghiệm (LIS/PXN).
        Hãy đọc kỹ ảnh phiếu xét nghiệm và trích xuất các thông tin phục vụ mục "XÁC NHẬN XÉT NGHIỆM ĐƯỢC CHỈNH SỬA, KHUYẾN CÁO":
        1. SID_KQCS: Mã bệnh nhân / Barcode / SID.
        2. Ngay_KQCS: Ngày lấy mẫu hoặc ngày xét nghiệm (dd/mm/yyyy).
        3. Thiet_bi_KQCS: Tên máy/thiết bị phân tích (nếu có, không có để trống).
        4. XN_KQCS: Tên chỉ số xét nghiệm liên quan (ví dụ: Glucose, Urea, AST...).
        5. Ghi_chu_Khuyen_cao: Trích xuất toàn bộ nội dung "Ghi chú", "Khuyến cáo", "Lời dặn", hoặc thông tin trạng thái mẫu (ví dụ: 'Mẫu tán huyết', 'Mẫu vỡ hồng cầu', 'Cần kiểm tra lại sau 24h', 'Khuyến cáo nhịn ăn...'). Nếu không có ghi chú rõ ràng, trích xuất ghi chú lâm sàng hoặc để trống.

        Trả về DUY NHẤT một chuỗi JSON thuần túy (không dùng code block markdown, không giải thích thêm):
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
           - Chọn BẤT KỲ 1 chỉ số xét nghiệm thuộc chuyên khoa SINH HÓA hoặc MIỄN DỊCH có kết quả NẰM TRONG KHOẢNG THAM CHIẾU SINH HỌC (bình thường, không in đậm, không báo động).
           - Trích xuất: XN (tên chỉ số), KQXN (kết quả), ĐVKQXN (đơn vị), Tham_chieu_KQXN (khoảng tham chiếu), KET_LUAN: "Đạt".

        3. THẨM ĐỊNH KẾT QUẢ BẤT THƯỜNG:
           - Tìm 1 chỉ số có KẾT QUẢ NẰM NGOÀI KHOẢNG THAM CHIẾU (in đậm ở cột kết quả, có cờ H/L hoặc cảnh báo).
           - Trích xuất: XN_BAT_THUONG, XN_BAT_THUONG_CO_KHONG ("Có" nếu được đánh dấu nhận biết), KET_LUAN_XN_BAT_THUONG ("Đạt").
           - Nếu toàn bộ phiếu bình thường: Để XN_BAT_THUONG là "", XN_BAT_THUONG_CO_KHONG là "Không", KET_LUAN_XN_BAT_THUONG là "Không áp dụng".

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

    response = client.models.generate_content(
        model="gemini-3.8-flash",
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

# Khởi tạo session_state cho 5 bệnh nhân
for i in range(1, 6):
    if f"sid_{i}" not in st.session_state:
        st.session_state[f"sid_{i}"] = f"BN00{i}"
    if f"ngay_{i}" not in st.session_state:
        st.session_state[f"ngay_{i}"] = date.today().strftime("%d/%m/%Y")
    if f"gio_{i}" not in st.session_state:
        st.session_state[f"gio_{i}"] = f"0{7+i}:30"
    if f"may_{i}" not in st.session_state:
        st.session_state[f"may_{i}"] = "Cobas c501"
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
        st.session_state[f"xnbt_{i}"] = "Troponin T"
    if f"btck_{i}" not in st.session_state:
        st.session_state[f"btck_{i}"] = "Có"
    if f"btkl_{i}" not in st.session_state:
        st.session_state[f"btkl_{i}"] = "Đạt"

# Session state cho mục Chỉnh sửa & Khuyến cáo
if "sid_kqcs" not in st.session_state:
    st.session_state["sid_kqcs"] = "BN_TEST_01"
if "ngay_kqcs" not in st.session_state:
    st.session_state["ngay_kqcs"] = date.today().strftime("%d/%m/%Y")
if "thiet_bi_kqcs" not in st.session_state:
    st.session_state["thiet_bi_kqcs"] = "Cobas c501"
if "xn_kqcs" not in st.session_state:
    st.session_state["xn_kqcs"] = "Urea"
if "ghi_chu_khuyen_cao" not in st.session_state:
    st.session_state["ghi_chu_khuyen_cao"] = "Mẫu tán huyết nhẹ, đã chạy lại kiểm chứng."

st.title("🧪 Phần mềm Thẩm định PXN & Nhận diện Ảnh AI")
st.markdown("Hỗ trợ quét ảnh phiếu xét nghiệm tự động điền vào: **Bệnh nhân 1-5** hoặc **Mục Chỉnh sửa, khuyến cáo**.")

with st.sidebar:
    st.header("🔑 Cấu hình AI Vision")
    api_key_input = st.text_input(
        "Nhập Gemini API Key:",
        type="password",
        help="Lấy API Key miễn phí tại: https://aistudio.google.com/app/apikey"
    )
    st.info("💡 Bạn có thể quét ảnh cho từng Bệnh nhân (1-5) hoặc quét riêng cho phiếu có Ghi chú/Khuyến cáo.")

tabs = st.tabs(["📸 Quét Ảnh Tự Động (AI)", "📝 Form Nhập Liệu", "📊 Xem & Tải Excel", "⚡ Xuất Hàng Loạt"])

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
                spinner_msg = "AI đang trích xuất Mã BN, Ngày XN và Nội dung khuyến cáo / ghi chú..." if is_recommendation else "AI đang phân tích chỉ số thường và chỉ số bất thường..."
                
                with st.spinner(spinner_msg):
                    try:
                        bytes_data = uploaded_file.getvalue()
                        mode = "recommendation" if is_recommendation else "patient"
                        extracted = extract_lab_data_from_image(bytes_data, api_key_input, mode=mode)
                        
                        if is_recommendation:
                            # Cập nhật mục Chỉnh sửa, khuyến cáo
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
                            st.success("✅ Đã trích xuất và cập nhật thành công vào mục 'Xác nhận xét nghiệm được chỉnh sửa, khuyến cáo'!")
                        else:
                            # Cập nhật Bệnh nhân 1-5
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
            ngay_xn_str = st.text_input("Ngày xác nhận (text)", value=ngay_danh_gia.strftime("%d/%m/%Y"))

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
            st.write(f"Tổng số bản ghi: **{len(df_current)}**")
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

# ================= TAB 3: XUẤT HÀNG LOẠT =================
with tabs[3]:
    st.subheader("Xuất hàng loạt tất cả phiếu trong Excel ra thư mục")
    if st.button("🚀 Chạy xuất tất cả file Word"):
        if not os.path.exists(EXCEL_FILE):
            st.error("Chưa có file Excel dữ liệu!")
        elif not os.path.exists(TEMPLATE_FILE):
            st.error(f"Không tìm thấy file mẫu '{TEMPLATE_FILE}'!")
        else:
            out_folder = "Ket_Qua_Xuat"
            os.makedirs(out_folder, exist_ok=True)
            df_all = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
            count = 0
            for idx, r in df_all.iterrows():
                ctx = r.to_dict()
                pm = clean_filename(r.get("Ten_phan_mem", f"Phieu_{idx+1}"))
                ng = clean_filename(r.get("Ngay", ""))
                fname = f"Ket_qua_{pm}_{ng}_{idx+1}.docx"
                doc = DocxTemplate(TEMPLATE_FILE)
                doc.render(ctx)
                doc.save(os.path.join(out_folder, fname))
                count += 1
            st.success(f"🎉 Đã xuất thành công {count} file vào thư mục '{out_folder}/'!")
