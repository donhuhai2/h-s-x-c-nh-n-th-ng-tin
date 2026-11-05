# -*- coding: utf-8 -*-
"""
Web App Nhập liệu thẩm định LIS & Xuất báo cáo Word / Excel.
Cập nhật mới nhất:
 1. Tự động lưu và nhớ Gemini API Key vào file cấu hình (không phải nhập lại mỗi khi mở lại trang).
 2. Ẩn hoàn toàn ô PIN khỏi Sidebar (mã PIN được khóa bí mật ngầm trong hệ thống).
 3. Bảo vệ thao tác xóa Excel bằng MÃ PIN bảo mật.
 4. Đa dạng nguồn ảnh (Tải lên, Chụp trực tiếp bằng Camera, Link Google Drive) và liên kết Google Drive.
 5. Xuất báo cáo linh hoạt: 1 ngày xuất riêng lẻ (Word/PDF), >= 2 ngày hoặc toàn bộ đóng gói ZIP (Word/PDF).
"""

import os
import io
import re
import json
import zipfile
import subprocess
import tempfile
import urllib.request
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
ADMIN_SECURITY_PIN = "175175"  # Mã PIN quản trị xóa dữ liệu

st.set_page_config(
    page_title="Nhập liệu & Thẩm định PXN - AI Vision",
    page_icon="🧪",
    layout="wide"
)

def clean_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", str(filename)).strip()

def load_saved_api_key():
    """Đọc API Key đã lưu từ file cấu hình hoặc Streamlit secrets."""
    if os.path.exists(API_KEY_FILE):
        try:
            with open(API_KEY_FILE, "r", encoding="utf-8") as f:
                saved = f.read().strip()
                if saved:
                    return saved
        except Exception:
            pass
    if "GEMINI_API_KEY" in st.secrets:
        return st.secrets["GEMINI_API_KEY"]
    return ""

def save_api_key_to_file(key_val):
    """Lưu API Key vào file cấu hình."""
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
        raise ValueError("Đường dẫn Google Drive không hợp lệ! Vui lòng đảm bảo link có dạng: https://drive.google.com/file/d/.../view")
    
    file_id = match.group(1)
    direct_url = f"