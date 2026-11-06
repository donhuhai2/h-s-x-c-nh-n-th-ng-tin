# -*- coding: utf-8 -*-
"""
Web App Nhập liệu thẩm định LIS & Xuất báo cáo Word / Excel.
Bản hoàn chỉnh tích hợp toàn diện:
 1. Tự động dò tìm Model Gemini khả dụng (client.models.list + v1beta) -> Khắc phục triệt để lỗi 404 NOT_FOUND.
 2. Hỗ trợ xử lý ảnh xoay ngang/dọc tự động qua PIL.Image.
 3. Đăng nhập, Phân quyền (Admin/User), Đổi mật khẩu & Đặt lại mật khẩu chính xác 100%.
 4. Đa tầng cơ sở dữ liệu: PostgreSQL trực tiếp (DATABASE_URL) -> Supabase API -> Lưu trữ JSON/Excel bền vững.
 5. Tự động chuẩn hóa cột Ngày (ghép Day/Month/year nếu trống) giúp lọc và xuất báo cáo Word/PDF/ZIP không lỗi.
 6. Hỗ trợ cấu hình đường dẫn lưu file tự động và Sao lưu/Khôi phục toàn hệ thống (.json / .xlsx).
"""

import os
import io
import re
import json
import base64
import shutil
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

try:
    from werkzeug.security import check_password_hash
    HAS_WERKZEUG = True
except ImportError:
    HAS_WERKZEUG = False

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    HAS_PSYCOPG2 = True
except ImportError:
    HAS_PSYCOPG2 = False

EXCEL_FILE = "danh_sach_phieu.xlsx"
USERS_LOCAL_FILE = "users_data.json"
RECORDS_LOCAL_FILE = "records_data.json"
TEMPLATE_FILE = "template.docx"
API_KEY_FILE = "api_key_config.txt"
SUPABASE_CFG_FILE = "supabase_config.json"

st.set_page_config(
    page_title="Hệ Thống Thẩm Định PXN - LIS",
    page_icon="🧪",
    layout="wide"
)

# ================= QUẢN LÝ BẢO MẬT & MẬT KHẨU =================
def hash_password(password):
    """Mã hóa mật khẩu bằng SHA-256."""
    return hashlib.sha256(str(password).strip().encode("utf-8")).hexdigest()

def verify_password(plain_password, stored_hash, session_plain=None):
    """Kiểm tra mật khẩu tương thích cả phiên hiện tại, scrypt/pbkdf2 và SHA-256."""
    plain_str = str(plain_password).strip()
    if session_plain is not None and plain_str == str(session_plain).strip():
        return True
    if not stored_hash:
        return False
    stored_str = str(stored_hash).strip()
    if stored_str.startswith("scrypt:") or stored_str.startswith("pbkdf2:"):
        if HAS_WERKZEUG:
            try:
                if check_password_hash(stored_str, plain_str):
                    return True
            except Exception:
                pass
    if hash_password(plain_str) == stored_str:
        return True
    if plain_str == stored_str:
        return True
    return False

# ================= CẤU HÌNH ĐƯỜNG DẪN LƯU FILE & ĐÁM MÂY =================
def load_cloud_config():
    url = os.environ.get("SUPABASE_URL", "").strip()
    key = os.environ.get("SUPABASE_KEY", "").strip()
    db_url = os.environ.get("DATABASE_URL", "").strip()
    local_dir = os.environ.get("BACKUP_LOCAL_DIR", "").strip()
    gdrive_webhook = os.environ.get("GDRIVE_WEBHOOK_URL", "").strip()

    if os.path.exists(SUPABASE_CFG_FILE):
        try:
            with open(SUPABASE_CFG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                if cfg.get("url"):
                    url = cfg["url"].strip()
                if cfg.get("key"):
                    key = cfg["key"].strip()
                if cfg.get("db_url"):
                    db_url = cfg["db_url"].strip()
                if cfg.get("local_dir"):
                    local_dir = cfg["local_dir"].strip()
                if cfg.get("gdrive_webhook"):
                    gdrive_webhook = cfg["gdrive_webhook"].strip()
        except Exception:
            pass
    return url, key, db_url, local_dir, gdrive_webhook

def save_cloud_config(url, key, db_url="", local_dir="", gdrive_webhook=""):
    try:
        with open(SUPABASE_CFG_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "url": url.strip(),
                "key": key.strip(),
                "db_url": db_url.strip(),
                "local_dir": local_dir.strip(),
                "gdrive_webhook": gdrive_webhook.strip()
            }, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False

def auto_sync_to_custom_paths():
    """Tự động đồng bộ file Excel và JSON vào thư mục cấu hình & Google Drive Webhook."""
    _, _, _, local_dir, gdrive_webhook = load_cloud_config()
    msg_list = []

    if local_dir:
        try:
            os.makedirs(local_dir, exist_ok=True)
            if os.path.exists(EXCEL_FILE):
                shutil.copy2(EXCEL_FILE, os.path.join(local_dir, "danh_sach_phieu.xlsx"))
            if os.path.exists(RECORDS_LOCAL_FILE):
                shutil.copy2(RECORDS_LOCAL_FILE, os.path.join(local_dir, "backup_phieu_pxn.json"))
            msg_list.append(f"Đã lưu vào thư mục: {local_dir}")
        except Exception as e:
            msg_list.append(f"Lưu thư mục cục bộ lỗi: {e}")

    if gdrive_webhook and os.path.exists(EXCEL_FILE):
        try:
            with open(EXCEL_FILE, "rb") as f:
                b64_xlsx = base64.b64encode(f.read()).decode("utf-8")
            payload = json.dumps({
                "filename": "danh_sach_phieu.xlsx",
                "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "filedata": b64_xlsx
            }).encode("utf-8")
            req = urllib.request.Request(
                gdrive_webhook,
                data=payload,
                headers={"Content-Type": "application/json"}
            )
            with urllib.request.urlopen(req, timeout=10):
                pass
            msg_list.append("Đã đồng bộ tự động lên Google Drive!")
        except Exception as e:
            msg_list.append(f"Lỗi đồng bộ Google Drive Webhook: {e}")

    return msg_list

def get_supabase_client():
    if not HAS_SUPABASE:
        return None
    url, key, _, _, _ = load_cloud_config()
    if url and key:
        try:
            return create_client(url, key)
        except Exception:
            return None
    return None

supabase: Client = get_supabase_client()

def get_pg_conn():
    if not HAS_PSYCOPG2:
        return None
    _, _, db_url, _, _ = load_cloud_config()
    if db_url:
        try:
            return psycopg2.connect(db_url, connect_timeout=5)
        except Exception:
            return None
    return None

def init_pg_tables_and_grants():
    conn = get_pg_conn()
    if not conn:
        return False
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS public.users (
                    id SERIAL PRIMARY KEY,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    role TEXT DEFAULT 'User',
                    full_name TEXT DEFAULT ''
                );
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS public.lab_records (
                    id BIGSERIAL PRIMARY KEY,
                    created_at TIMESTAMPTZ DEFAULT NOW(),
                    data JSONB NOT NULL
                );
            """)
            cur.execute("ALTER TABLE public.users DISABLE ROW LEVEL SECURITY;")
            cur.execute("ALTER TABLE public.lab_records DISABLE ROW LEVEL SECURITY;")
            cur.execute("GRANT ALL ON TABLE public.users TO anon, authenticated, service_role;")
            cur.execute("GRANT ALL ON TABLE public.lab_records TO anon, authenticated, service_role;")
            cur.execute("GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO anon, authenticated, service_role;")
        conn.commit()
        conn.close()
        return True
    except Exception:
        try:
            conn.close()
        except Exception:
            pass
        return False

init_pg_tables_and_grants()

# ================= QUẢN LÝ NGƯỜI DÙNG =================
def load_local_users():
    default_users = [
        {
            "id": 1,
            "username": "admin",
            "password_hash": hash_password("admin"),
            "role": "Admin",
            "full_name": "Quản trị viên PXN"
        },
        {
            "id": 2,
            "username": "nhuhai",
            "password_hash": hash_password("123456"),
            "role": "Admin",
            "full_name": "CN. Đỗ Như Hải"
        }
    ]
    if os.path.exists(USERS_LOCAL_FILE):
        try:
            with open(USERS_LOCAL_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list) and len(data) > 0:
                    return data
        except Exception:
            pass
    save_local_users(default_users)
    return default_users

def save_local_users(users_list):
    try:
        with open(USERS_LOCAL_FILE, "w", encoding="utf-8") as f:
            json.dump(users_list, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def fetch_all_users_list():
    conn = get_pg_conn()
    if conn:
        try:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("SELECT id, username, password_hash, role, full_name FROM public.users ORDER BY id ASC;")
                rows = [dict(r) for r in cur.fetchall()]
            conn.close()
            if rows:
                save_local_users(rows)
                return rows, "PostgreSQL Cloud"
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

    if supabase:
        try:
            res = supabase.table("users").select("*").order("id", desc=False).execute()
            if res.data:
                save_local_users(res.data)
                return res.data, "Supabase Cloud"
        except Exception:
            pass

    return load_local_users(), "Hệ thống Lưu trữ Nội bộ"

def get_user(username):
    u_clean = str(username).strip().lower()
    users_list, _ = fetch_all_users_list()
    for u in users_list:
        if str(u.get("username", "")).strip().lower() == u_clean:
            return u
    return None

def change_user_password(username, new_password, full_name="", role="Admin"):
    u_clean = str(username).strip().lower()
    new_hash = hash_password(new_password)
    cloud_updated = False

    conn = get_pg_conn()
    if conn:
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT id FROM public.users WHERE lower(username) = %s;", (u_clean,))
                if cur.fetchone():
                    cur.execute("UPDATE public.users SET password_hash = %s WHERE lower(username) = %s;", (new_hash, u_clean))
                else:
                    cur.execute(
                        "INSERT INTO public.users (username, password_hash, role, full_name) VALUES (%s, %s, %s, %s);",
                        (u_clean, new_hash, role, full_name)
                    )
            conn.commit()
            conn.close()
            cloud_updated = True
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

    if supabase and not cloud_updated:
        try:
            check = supabase.table("users").select("id").eq("username", u_clean).execute()
            if check.data:
                supabase.table("users").update({"password_hash": new_hash}).eq("username", u_clean).execute()
            else:
                supabase.table("users").insert({
                    "username": u_clean,
                    "password_hash": new_hash,
                    "role": role,
                    "full_name": full_name
                }).execute()
        except Exception:
            pass

    users_list = load_local_users()
    found = False
    for u in users_list:
        if str(u.get("username", "")).strip().lower() == u_clean:
            u["password_hash"] = new_hash
            if full_name:
                u["full_name"] = full_name
            if role:
                u["role"] = role
            found = True
            break
    if not found:
        new_id = max([int(x.get("id", 0)) for x in users_list] + [0]) + 1
        users_list.append({
            "id": new_id,
            "username": u_clean,
            "password_hash": new_hash,
            "role": role,
            "full_name": full_name or u_clean
        })
    save_local_users(users_list)
    return True, f"✅ Đã đổi mật khẩu cho tài khoản '{u_clean}' thành công!"

def create_new_user(username, password, role, full_name):
    u_clean = str(username).strip().lower()
    p_hash = hash_password(password)
    users_list, _ = fetch_all_users_list()

    for u in users_list:
        if str(u.get("username", "")).strip().lower() == u_clean:
            return False, f"⛔ Tên đăng nhập '{u_clean}' đã tồn tại!"

    conn = get_pg_conn()
    if conn:
        try:
            with conn.cursor() as cur:
                cur.execute(
                    "INSERT INTO public.users (username, password_hash, role, full_name) VALUES (%s, %s, %s, %s);",
                    (u_clean, p_hash, role, full_name)
                )
            conn.commit()
            conn.close()
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

    if supabase:
        try:
            supabase.table("users").insert({
                "username": u_clean,
                "password_hash": p_hash,
                "role": role,
                "full_name": full_name
            }).execute()
        except Exception:
            pass

    local_users = load_local_users()
    new_id = max([int(x.get("id", 0)) for x in local_users] + [0]) + 1
    local_users.append({
        "id": new_id,
        "username": u_clean,
        "password_hash": p_hash,
        "role": role,
        "full_name": full_name
    })
    save_local_users(local_users)
    return True, f"✅ Đã tạo tài khoản '{u_clean}' ({full_name}) thành công!"

def delete_user_by_username(username):
    u_clean = str(username).strip().lower()
    if u_clean == "admin":
        return False, "Không thể xóa tài khoản admin gốc!"
    conn = get_pg_conn()
    if conn:
        try:
            with conn.cursor() as cur:
                cur.execute("DELETE FROM public.users WHERE lower(username) = %s;", (u_clean,))
            conn.commit()
            conn.close()
        except Exception:
            pass
    if supabase:
        try:
            supabase.table("users").delete().eq("username", u_clean).execute()
        except Exception:
            pass
    local_users = [u for u in load_local_users() if str(u.get("username", "")).strip().lower() != u_clean]
    save_local_users(local_users)
    return True, f"✅ Đã xóa tài khoản '{u_clean}'!"

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "user_info" not in st.session_state:
    st.session_state["user_info"] = None
if "login_plain_pwd" not in st.session_state:
    st.session_state["login_plain_pwd"] = None

# ================= GIAO DIỆN ĐĂNG NHẬP =================
if not st.session_state["logged_in"]:
    st.title("🔐 Đăng Nhập Hệ Thống Thẩm Định PXN")
    col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
    with col_l2:
        with st.form("login_form"):
            login_user = st.text_input("Tên đăng nhập:", placeholder="Nhập username...")
            login_pass = st.text_input("Mật khẩu:", type="password", placeholder="Nhập mật khẩu...")
            btn_login = st.form_submit_button("🚀 Đăng nhập", use_container_width=True)

        if btn_login:
            u_clean = login_user.strip().lower()
            p_clean = login_pass.strip()
            user_data = get_user(u_clean)
            if user_data and verify_password(p_clean, user_data.get("password_hash")):
                st.session_state["logged_in"] = True
                st.session_state["user_info"] = user_data
                st.session_state["login_plain_pwd"] = p_clean
                st.rerun()
            else:
                st.error("⛔ Tên đăng nhập hoặc mật khẩu không chính xác!")
        st.info("💡 Tài khoản mặc định: **admin** (mật khẩu: **admin**) hoặc **nhuhai** (mật khẩu: **123456**)")
    st.stop()

# ================= XỬ LÝ DỮ LIỆU PHIẾU XÉT NGHIỆM & AI =================
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
    return os.environ.get("GEMINI_API_KEY", "").strip()

def save_api_key_to_file(key_val):
    try:
        with open(API_KEY_FILE, "w", encoding="utf-8") as f:
            f.write(str(key_val).strip())
        return True
    except Exception:
        return False

def normalize_records_df(df):
    """Tự động điền cột Ngay và Ngay_xac_nhan từ Day/Month/year nếu bị trống."""
    if df.empty:
        return df
    df = df.fillna("")
    for idx, row in df.iterrows():
        ngay_val = str(row.get("Ngay", "")).strip()
        ngay_xn_val = str(row.get("Ngay_xac_nhan", "")).strip()
        if not ngay_val:
            d = str(row.get("Day", "")).strip()
            m = str(row.get("Month", "")).strip()
            y = str(row.get("year", "")).strip()
            if d and m and y:
                ngay_val = f"{d.zfill(2)}/{m.zfill(2)}/{y}"
                df.at[idx, "Ngay"] = ngay_val
        if not ngay_xn_val and ngay_val:
            df.at[idx, "Ngay_xac_nhan"] = ngay_val
    return df

def fetch_all_records_df():
    conn = get_pg_conn()
    if conn:
        try:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("SELECT id, data FROM public.lab_records ORDER BY id ASC;")
                rows = [r["data"] for r in cur.fetchall() if isinstance(r.get("data"), dict)]
            conn.close()
            if rows:
                df = normalize_records_df(pd.DataFrame(rows))
                df.to_excel(EXCEL_FILE, index=False)
                return df
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

    if supabase:
        try:
            res = supabase.table("lab_records").select("id, data").order("id", desc=False).execute()
            rows = [item["data"] for item in res.data if isinstance(item.get("data"), dict)]
            if rows:
                df = normalize_records_df(pd.DataFrame(rows))
                df.to_excel(EXCEL_FILE, index=False)
                return df
        except Exception:
            pass

    if os.path.exists(EXCEL_FILE):
        try:
            return normalize_records_df(pd.read_excel(EXCEL_FILE, dtype=str))
        except Exception:
            pass
    if os.path.exists(RECORDS_LOCAL_FILE):
        try:
            with open(RECORDS_LOCAL_FILE, "r", encoding="utf-8") as f:
                rows = json.load(f)
                if rows:
                    return normalize_records_df(pd.DataFrame(rows))
        except Exception:
            pass
    return pd.DataFrame()

def save_single_record(row_dict):
    clean_dict = {k: ("" if v is None else str(v)) for k, v in row_dict.items()}
    conn = get_pg_conn()
    if conn:
        try:
            with conn.cursor() as cur:
                cur.execute("INSERT INTO public.lab_records (data) VALUES (%s);", (json.dumps(clean_dict),))
            conn.commit()
            conn.close()
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

    if supabase:
        try:
            supabase.table("lab_records").insert({"data": clean_dict}).execute()
        except Exception:
            pass

    new_df = pd.DataFrame([clean_dict])
    if os.path.exists(EXCEL_FILE):
        try:
            old_df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
            combined = pd.concat([old_df, new_df], ignore_index=True)
        except Exception:
            combined = new_df
    else:
        combined = new_df
    combined = normalize_records_df(combined)
    combined.to_excel(EXCEL_FILE, index=False)
    try:
        with open(RECORDS_LOCAL_FILE, "w", encoding="utf-8") as f:
            json.dump(combined.to_dict(orient="records"), f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    auto_sync_to_custom_paths()

def delete_records_by_date(target_date):
    deleted_count = 0
    t_date = str(target_date).strip()
    conn = get_pg_conn()
    if conn:
        try:
            with conn.cursor() as cur:
                cur.execute(
                    "DELETE FROM public.lab_records WHERE (data->>'Ngay') = %s OR (data->>'Ngay_xac_nhan') = %s;",
                    (t_date, t_date)
                )
                deleted_count = cur.rowcount
            conn.commit()
            conn.close()
        except Exception:
            pass

    if os.path.exists(EXCEL_FILE):
        try:
            df = normalize_records_df(pd.read_excel(EXCEL_FILE, dtype=str))
            mask1 = (df.get("Ngay", "").str.strip() == t_date)
            mask2 = (df.get("Ngay_xac_nhan", "").str.strip() == t_date)
            df_rem = df[~(mask1 | mask2)]
            deleted_count = max(deleted_count, len(df) - len(df_rem))
            df_rem.to_excel(EXCEL_FILE, index=False)
            with open(RECORDS_LOCAL_FILE, "w", encoding="utf-8") as f:
                json.dump(df_rem.to_dict(orient="records"), f, ensure_ascii=False, indent=2)
            auto_sync_to_custom_paths()
        except Exception:
            pass
    return deleted_count

def delete_all_records():
    conn = get_pg_conn()
    if conn:
        try:
            with conn.cursor() as cur:
                cur.execute("TRUNCATE TABLE public.lab_records;")
            conn.commit()
            conn.close()
        except Exception:
            pass
    if supabase:
        try:
            supabase.table("lab_records").delete().neq("id", 0).execute()
        except Exception:
            pass
    if os.path.exists(EXCEL_FILE):
        try:
            os.remove(EXCEL_FILE)
        except Exception:
            pass
    if os.path.exists(RECORDS_LOCAL_FILE):
        try:
            os.remove(RECORDS_LOCAL_FILE)
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
        raise ValueError("Đường dẫn Google Drive không hợp lệ!")
    file_id = match.group(1)
    direct_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    req = urllib.request.Request(direct_url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as resp:
        return resp.read()

# ================= NHẬN DIỆN ẢNH AI (TỰ ĐỘNG DÒ MODEL KHẢ DỤNG) =================
def get_available_vision_models(client):
    """Tự động lấy danh sách các model thực tế đang mở cho API Key này từ máy chủ Google."""
    available = []
    try:
        for m in client.models.list():
            name = getattr(m, "name", "") or ""
            clean_name = name.replace("models/", "")
            methods = getattr(m, "supported_actions", None) or getattr(m, "supported_generation_methods", []) or []
            if "gemini" in clean_name.lower() and not any(
                x in clean_name.lower() for x in ["embedding", "tts", "audio", "imagen", "veo", "aqa", "thinking"]
            ):
                if not methods or "generateContent" in str(methods):
                    available.append(clean_name)
    except Exception:
        pass

    preferred_order = [
        "gemini-2.5-flash",
        "gemini-2.5-pro",
        "gemini-2.0-flash",
        "gemini-2.0-flash-001",
        "gemini-2.0-flash-lite",
        "gemini-flash-latest",
        "gemini-pro-latest"
    ]

    final_list = []
    for pref in preferred_order:
        for avail in available:
            if avail == pref and avail not in final_list:
                final_list.append(avail)

    for avail in available:
        if avail not in final_list:
            final_list.append(avail)

    if not final_list:
        final_list = preferred_order

    return final_list

def extract_lab_data_from_image(image_bytes, api_key, mode="patient"):
    client = genai.Client(api_key=api_key.strip(), http_options={"api_version": "v1beta"})

    pil_img = Image.open(io.BytesIO(image_bytes))
    if pil_img.mode not in ("RGB", "L"):
        pil_img = pil_img.convert("RGB")

    if mode == "recommendation":
        prompt = """
        Bạn là chuyên gia thẩm định hồ sơ kết quả xét nghiệm (LIS/PXN).
        Lưu ý: Ảnh có thể bị xoay ngang 90 độ hoặc ngược, hãy tự động đọc theo chiều chữ đúng.
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
        Lưu ý: Ảnh có thể bị xoay ngang 90 độ hoặc ngược, hãy tự động đọc theo chiều chữ đúng.
        Hãy đọc ảnh phiếu kết quả xét nghiệm và trích xuất dữ liệu theo đúng quy tắc sau:

        1. THÔNG TIN CHUNG:
           - SID: Mã bệnh nhân / Barcode / SID.
           - Ngay: Ngày xét nghiệm (dd/mm/yyyy).
           - Gio_nhan_mau: Giờ nhận mẫu (hh:mm).
           - May: Tên thiết bị phân tích (nếu có, ví dụ AU 5822_1, AU 5822_2, DxI800_3...).

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

    candidate_models = get_available_vision_models(client)
    errors_log = []

    for model_name in candidate_models:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[pil_img, prompt]
            )
            raw_text = (response.text or "").strip()
            raw_text = re.sub(r"^```json\s*", "", raw_text)
            raw_text = re.sub(r"^```\s*", "", raw_text)
            raw_text = re.sub(r"```$", "", raw_text).strip()
            start_idx = raw_text.find("{")
            end_idx = raw_text.rfind("}")
            if start_idx != -1 and end_idx != -1:
                raw_text = raw_text[start_idx:end_idx + 1]
            return json.loads(raw_text)
        except Exception as e:
            errors_log.append(f"{model_name}: {e}")
            continue

    raise RuntimeError("Không thể phân tích ảnh với các model hiện có. Chi tiết: " + " | ".join(errors_log[:3]))

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
user_role_str = str(user_info.get("role", "")).strip().lower()
is_admin = (user_role_str in ["admin", "quản trị viên", "quantri"]) or (str(user_info.get("username", "")).lower() in ["admin", "nhuhai"])

st.title("🧪 Phần mềm Thẩm định PXN & Nhận diện Ảnh AI")

with st.sidebar:
    st.markdown(f"👤 **Xin chào:** {user_info.get('full_name', user_info.get('username'))}")
    st.markdown(f"🔰 **Vai trò:** `{'Admin' if is_admin else 'User'}`")

    if st.button("🚪 Đăng xuất", use_container_width=True):
        st.session_state["logged_in"] = False
        st.session_state["user_info"] = None
        st.session_state["login_plain_pwd"] = None
        st.rerun()

    st.markdown("---")
    with st.expander("🔑 Đổi Mật Khẩu Cá Nhân", expanded=False):
        old_pwd = st.text_input("Mật khẩu hiện tại:", type="password", key="chg_old")
        new_pwd = st.text_input("Mật khẩu mới:", type="password", key="chg_new")
        confirm_pwd = st.text_input("Nhập lại mật khẩu mới:", type="password", key="chg_conf")
        if st.button("💾 Cập nhật mật khẩu", use_container_width=True):
            latest_user = get_user(user_info.get("username", "admin")) or user_info
            is_valid_old = verify_password(
                old_pwd,
                latest_user.get("password_hash"),
                session_plain=st.session_state.get("login_plain_pwd")
            )
            if not is_valid_old:
                st.error("⛔ Mật khẩu hiện tại không chính xác!")
            elif not new_pwd or len(new_pwd.strip()) < 4:
                st.warning("⚠️ Mật khẩu mới phải có ít nhất 4 ký tự!")
            elif new_pwd.strip() != confirm_pwd.strip():
                st.error("⛔ Xác nhận mật khẩu mới không khớp!")
            else:
                ok, msg = change_user_password(
                    user_info["username"],
                    new_pwd.strip(),
                    user_info.get("full_name", ""),
                    user_info.get("role", "Admin" if is_admin else "User")
                )
                if ok:
                    st.session_state["user_info"]["password_hash"] = hash_password(new_pwd.strip())
                    st.session_state["login_plain_pwd"] = new_pwd.strip()
                    st.success(msg)
                else:
                    st.error(msg)

    if is_admin:
        with st.expander("☁️ Cấu hình Cloud / Database", expanded=False):
            cur_url, cur_key, cur_db_url, cur_local_dir, cur_webhook = load_cloud_config()
            inp_db_url = st.text_input(
                "1. PostgreSQL URI (DATABASE_URL - Khuyên dùng):",
                value=cur_db_url,
                type="password",
                help="Dán chuỗi postgresql://postgres.gjgvfhiktredztgpcrfy:... từ mục Shared pooler trên Supabase"
            )
            inp_local_dir = st.text_input(
                "2. Đường dẫn thư mục lưu tự động trên máy (nếu chạy Local):",
                value=cur_local_dir,
                placeholder="D:/Backup_PXN"
            )
            inp_webhook = st.text_input(
                "3. Link Google Drive Webhook (nếu dùng):",
                value=cur_webhook
            )
            if st.button("🔌 Lưu Cấu Hình Cloud", use_container_width=True):
                save_cloud_config(cur_url, cur_key, inp_db_url, inp_local_dir, inp_webhook)
                init_pg_tables_and_grants()
                auto_sync_to_custom_paths()
                st.success("✅ Đã lưu cấu hình thành công!")
                time.sleep(0.5)
                st.rerun()

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
        value="https://drive.google.com"
    )
    if gdrive_folder_url:
        st.markdown(f"[📂 Mở Thư mục Google Drive]({gdrive_folder_url})", unsafe_allow_html=True)

tab_titles = ["📸 Quét Ảnh Tự Động (AI)", "📝 Form Nhập Liệu", "📊 Quản Lý Dữ Liệu & Sao Lưu", "⚡ Xuất Báo Cáo"]
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
                placeholder="[https://drive.google.com/file/d/.../view?usp=sharing](https://drive.google.com/file/d/.../view?usp=sharing)"
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
                            for k_js, k_st in [
                                ("SID_KQCS", "sid_kqcs"),
                                ("Ngay_KQCS", "ngay_kqcs"),
                                ("Thiet_bi_KQCS", "thiet_bi_kqcs"),
                                ("XN_KQCS", "xn_kqcs"),
                                ("Ghi_chu_Khuyen_cao", "ghi_chu_khuyen_cao")
                            ]:
                                if extracted.get(k_js):
                                    st.session_state[k_st] = extracted.get(k_js)
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
            nguoi_xac_nhan = st.text_input("Người xác nhận / Thẩm định", value=user_info.get("full_name", "CN. Đỗ Như Hải"))
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

        submitted = st.form_submit_button("💾 LƯU DỮ LIỆU HỆ THỐNG", use_container_width=True)

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
        st.success("✅ Đã lưu phiếu xét nghiệm thành công!")
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

# ================= TAB 2: QUẢN LÝ DỮ LIỆU & SAO LƯU =================
with tabs[2]:
    st.subheader("📊 Quản lý Dữ liệu & Khôi phục Hệ thống")
    df_current = fetch_all_records_df()

    if is_admin:
        with st.expander("📤 NẠP / KHÔI PHỤC DỮ LIỆU TỪ FILE EXCEL HOẶC GÓI BACKUP (.xlsx / .json)", expanded=False):
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                st.markdown("**1. Nạp từ file Excel (.xlsx)**")
                uploaded_excel = st.file_uploader("Chọn file Excel sao lưu:", type=["xlsx", "xls"], key="excel_uploader")
                if uploaded_excel is not None:
                    try:
                        df_imported = normalize_records_df(pd.read_excel(uploaded_excel, dtype=str))
                        st.write(f"🔍 Phát hiện **{len(df_imported)}** phiếu:")
                        upload_action = st.radio("Phương thức nạp:", options=["➕ Ghi thêm", "🔄 Ghi đè toàn bộ"])
                        if st.button("🚀 Nạp File Excel", use_container_width=True):
                            if "Ghi đè" in upload_action:
                                delete_all_records()
                            for _, r in df_imported.iterrows():
                                save_single_record(r.to_dict())
                            st.success("✅ Đã khôi phục dữ liệu Excel thành công!")
                            st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi đọc file Excel: {e}")

            with col_b2:
                st.markdown("**2. Khôi phục Gói Toàn Hệ Thống (.json)**")
                uploaded_json = st.file_uploader("Chọn file Backup Hệ thống (.json):", type=["json"], key="json_uploader")
                if uploaded_json is not None:
                    if st.button("🔄 Khôi phục Toàn Hệ Thống", use_container_width=True):
                        try:
                            pkg = json.load(uploaded_json)
                            if "users" in pkg and isinstance(pkg["users"], list):
                                save_local_users(pkg["users"])
                            if "records" in pkg and isinstance(pkg["records"], list):
                                delete_all_records()
                                for rec in pkg["records"]:
                                    save_single_record(rec)
                            st.success("✅ Đã khôi phục toàn bộ Tài khoản và Phiếu xét nghiệm!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Lỗi file JSON: {e}")

    st.markdown("### Dữ liệu Phiếu Xét Nghiệm Hiện Có")
    if not df_current.empty:
        st.write(f"Tổng số bản ghi hiện có: **{len(df_current)}**")
        st.dataframe(df_current, use_container_width=True)

        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            excel_buffer = io.BytesIO()
            df_current.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)
            st.download_button(
                label="📥 Tải toàn bộ file Excel (.xlsx)",
                data=excel_buffer,
                file_name=f"danh_sach_phieu_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with c_dl2:
            full_pkg = {
                "users": load_local_users(),
                "records": df_current.to_dict(orient="records")
            }
            json_bytes = json.dumps(full_pkg, ensure_ascii=False, indent=2).encode("utf-8")
            st.download_button(
                label="📦 Tải Gói Backup Toàn Hệ Thống (Tài khoản + Phiếu .json)",
                data=json_bytes,
                file_name=f"backup_he_thong_pxn_{datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json",
                use_container_width=True
            )

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
            with col_del2:
                st.markdown("**2. Xóa toàn bộ dữ liệu (Làm mới hoàn toàn)**")
                if st.button("⚠️ XÓA SẠCH TẤT CẢ DỮ LIỆU", use_container_width=True):
                    delete_all_records()
                    st.warning("✅ Đã xóa sạch toàn bộ cơ sở dữ liệu!")
                    st.rerun()
    else:
        st.info("Chưa có phiếu xét nghiệm nào trong hệ thống.")

# ================= TAB 3: XUẤT BÁO CÁO (WORD / PDF / ZIP) =================
with tabs[3]:
    st.subheader("⚡ Tùy chọn Xuất Báo Cáo")
    df_all = fetch_all_records_df()
    if df_all.empty:
        st.warning("⚠️ Chưa có dữ liệu trong hệ thống!")
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

        scope_choice = st.radio("Phạm vi xuất phiếu:", options=["Toàn bộ tất cả phiếu", "Lọc theo ngày"], horizontal=True)
        selected_dates = []
        if scope_choice == "Lọc theo ngày":
            if date_list:
                selected_dates = st.multiselect("Chọn ngày cần xuất:", options=date_list, default=[date_list[-1]])
            else:
                input_d = st.text_input("Nhập ngày cần xuất (dd/mm/yyyy):", value=date.today().strftime("%d/%m/%Y"))
                selected_dates = [input_d] if input_d else []

        file_format = st.radio("Chọn định dạng xuất:", options=["Word (.docx)", "PDF (.pdf)"], horizontal=True)

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
                if total_records == 1:
                    row = filtered_df.iloc[0]
                    ctx = row.to_dict()
                    pm = clean_filename(row.get("Ten_phan_mem", "Phieu"))
                    ng = clean_filename(row.get("Ngay", clean_filename(row.get("Ngay_xac_nhan", ""))))
                    base_name = f"Ket_qua_{pm}_{ng}"
                    word_bytes = export_single_word(ctx, TEMPLATE_FILE).getvalue()

                    if file_format == "Word (.docx)":
                        st.success(f"✅ Đã tạo thành công phiếu Word ({ng})!")
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
                    zip_filename = f"Ket_qua_PXN_{count}_Phieu.zip"
                    st.success(f"🎉 Đã đóng gói thành công {count} phiếu vào file ZIP!")
                    st.download_button(
                        label=f"📥 TẢI VỀ GÓI NÉN ZIP: {zip_filename}",
                        data=zip_buffer.getvalue(),
                        file_name=zip_filename,
                        mime="application/zip",
                        use_container_width=True
                    )

# ================= TAB 4: QUẢN TRỊ NGƯỜI DÙNG =================
if is_admin:
    with tabs[4]:
        st.subheader("👥 Quản Trị Tài Khoản & Phân Quyền")
        col_u1, col_u2 = st.columns([1, 1.3])
        with col_u1:
            st.markdown("#### 1. Tạo Tài Khoản Mới")
            with st.form("create_user_form"):
                new_u = st.text_input("Tên đăng nhập mới (username):")
                new_p = st.text_input("Mật khẩu ban đầu:", type="password")
                new_fn = st.text_input("Họ và tên nhân viên:")
                new_role = st.selectbox("Phân quyền sử dụng:", ["Admin", "User"])
                btn_create = st.form_submit_button("➕ Tạo người dùng", use_container_width=True)

            if btn_create:
                if not new_u or not new_p:
                    st.warning("Vui lòng điền đầy đủ tên đăng nhập và mật khẩu!")
                else:
                    ok_u, msg_u = create_new_user(new_u, new_p, new_role, new_fn)
                    if ok_u:
                        st.success(msg_u)
                        time.sleep(0.5)
                        st.rerun()
                    else:
                        st.error(msg_u)

            st.markdown("---")
            st.markdown("#### 2. Đặt Lại Mật Khẩu Cho Tài Khoản Bất Kỳ")
            with st.form("reset_pwd_form"):
                target_u = st.text_input("Tên đăng nhập cần đặt lại mật khẩu:")
                reset_p = st.text_input("Mật khẩu mới:", type="password")
                btn_reset = st.form_submit_button("🔄 Đặt lại mật khẩu", use_container_width=True)
            if btn_reset:
                if not target_u or not reset_p:
                    st.warning("Vui lòng nhập tên tài khoản và mật khẩu mới!")
                else:
                    ok_r, msg_r = change_user_password(target_u, reset_p)
                    if ok_r:
                        if target_u.strip().lower() == str(user_info.get("username", "")).lower():
                            st.session_state["user_info"]["password_hash"] = hash_password(reset_p)
                            st.session_state["login_plain_pwd"] = reset_p.strip()
                        st.success(msg_r)
                    else:
                        st.error(msg_r)

        with col_u2:
            st.markdown("#### Danh Sách Người Dùng Hiện Có")
            users_data, source_label = fetch_all_users_list()
            st.caption(f"📡 *Nguồn dữ liệu đang hoạt động: {source_label}*")
            if users_data:
                df_u = pd.DataFrame(users_data)
                display_cols = [c for c in ["id", "username", "full_name", "role"] if c in df_u.columns]
                st.dataframe(df_u[display_cols], use_container_width=True)

                with st.expander("🗑️ Xóa tài khoản nhân viên"):
                    del_u_list = [u["username"] for u in users_data if u.get("username") != "admin"]
                    if del_u_list:
                        u_to_del = st.selectbox("Chọn tài khoản cần xóa:", options=del_u_list)
                        if st.button(f"❌ Xóa tài khoản {u_to_del}", use_container_width=True):
                            ok_d, msg_d = delete_user_by_username(u_to_del)
                            if ok_d:
                                st.success(msg_d)
                                time.sleep(0.5)
                                st.rerun()
                            else:
                                st.error(msg_d)
                    else:
                        st.info("Không có tài khoản phụ để xóa.")