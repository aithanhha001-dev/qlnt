import calendar
import hashlib
import io
import os
import re
import warnings
from contextlib import contextmanager
from datetime import date, datetime

import pandas as pd
import streamlit as st

# ==========================================
# 1. CẤU HÌNH TRANG & HÀM TIỆN ÍCH
# ==========================================
st.set_page_config(page_title="Quản lý nhân sự", layout="wide", page_icon="🏫")

# Ẩn thanh công cụ của Streamlit
st.markdown("""
<style>
[data-testid="stToolbar"], [data-testid="stMainMenu"], [data-testid="stAppDeployButton"],
[data-testid="stDeployButton"], .stDeployButton, #MainMenu, footer,
[data-testid="stDecoration"], [data-testid="stStatusWidget"] { display: none !important; visibility: hidden !important; }
</style>
""", unsafe_allow_html=True)

_ST_VER = tuple(int(x) for x in re.findall(r"\d+", st.__version__)[:2])
W = {"width": "stretch"} if _ST_VER >= (1, 50) else {"use_container_width": True}

for _k, _v in {"logged_in": False, "username": "", "role": "", "ma_vien_chuc": "", "flash": ""}.items():
    st.session_state.setdefault(_k, _v)

MIN_DATE = date(1940, 1, 1)
MAX_DATE = date(2040, 12, 31)
DATE_FORMAT = "DD/MM/YYYY"
YEARS = list(range(2020, 2041))
GIOI_TINH = ["Nam", "Nữ"]
TRINH_DO = ["Đại học", "Cao đẳng", "Trung cấp", "Thạc sĩ", "Tiến sĩ", "Khác"]
HANG_CDNN = ["Hạng I", "Hạng II", "Hạng III", "Hạng IV", "Khác"]
TRANG_THAI_CC = ["X", "P", "K", "VR", "CT", "O"]
LUONG_CO_SO = 2530000
CHUC_VU_LIST = ["Hiệu trưởng", "Phó hiệu trưởng", "Tổ trưởng CM", "Tổ phó CM", "Giáo viên",
                "Nhân viên y tế", "Nhân viên Hành chính", "Nhân viên khác"]
CHUC_VU_CU = {"Phó Hiệu trưởng": "Phó hiệu trưởng", "Tổ trưởng": "Tổ trưởng CM", "Nhân viên": "Nhân viên khác"}

def make_hashes(password): return hashlib.sha256(str.encode(password)).hexdigest()
def check_hashes(password, hashed_text): return make_hashes(password) == hashed_text

# ==========================================
# 2. MODULE KẾT NỐI CƠ SỞ DỮ LIỆU ĐÁM MÂY
# ==========================================
try:
    HAS_SECRETS = "DATABASE_URL" in st.secrets
except Exception:
    HAS_SECRETS = False
USE_POSTGRES = HAS_SECRETS or "DATABASE_URL" in os.environ

if USE_POSTGRES:
    import psycopg2
    from psycopg2 import IntegrityError
    DB_URL = st.secrets.get("DATABASE_URL", os.environ.get("DATABASE_URL", ""))
    
    @contextmanager
    def db():
        conn = psycopg2.connect(DB_URL)
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def q(sql, params=()):
        sql = sql.replace('?', '%s')
        sql = sql.replace("CAST(strftime('%d', ngay_cham_cong) AS INTEGER)", "CAST(EXTRACT(DAY FROM ngay_cham_cong) AS INTEGER)")
        with db() as conn:
            with warnings.catch_warnings():
                warnings.simplefilter('ignore', UserWarning)
                return pd.read_sql_query(sql, conn, params=params)

    def run(sql, params=()):
        sql = sql.replace('?', '%s')
        with db() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, params)

    def run_many(sql, records):
        sql = sql.replace('?', '%s')
        with db() as conn:
            with conn.cursor() as cur:
                cur.executemany(sql, records)
else:
    import sqlite3
    from sqlite3 import IntegrityError
    DB_NAME = "quanly_nhatruong.db"

    @contextmanager
    def db():
        conn = sqlite3.connect(DB_NAME)
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def q(sql, params=()):
        with db() as conn: return pd.read_sql_query(sql, conn, params=params)

    def run(sql, params=()):
        with db() as conn: conn.execute(sql, params)

    def run_many(sql, records):
        with db() as conn: conn.executemany(sql, records)

def flash_rerun(msg, *clear_keys):
    st.session_state["flash"] = msg
    for k in clear_keys: st.session_state.pop(k, None)
    st.rerun()

# ---------- Hàm Ngày tháng / định dạng ----------
def date_or_none(val):
    if val is None: return None
    try:
        if pd.isna(val): return None
    except (TypeError, ValueError): pass
    s = str(val).strip()
    if s in ("", "NaT", "None", "nan", "<NA>"): return None
    try:
        if isinstance(val, datetime): d = val.date()
        elif isinstance(val, date): d = val
        elif re.match(r"^\d{4}-\d{2}-\d{2}", s): d = pd.to_datetime(s).date()
        else: d = pd.to_datetime(s, dayfirst=True).date()
    except Exception: return None
    return min(max(d, MIN_DATE), MAX_DATE)

def date_or_today(val): return date_or_none(val) or date.today()
def sql_date(d): return d.strftime("%Y-%m-%d") if d else None
def format_date_vn(val):
    d = date_or_none(val)
    return d.strftime("%d/%m/%Y") if d else ""

def safe_str(val):
    try: return "" if val is None or pd.isna(val) else str(val)
    except (TypeError, ValueError): return str(val)

def pick(options, value, default=0): return options.index(value) if value in options else default

# ---------- Hàm Excel ----------
def _autofit(ws, df):
    from openpyxl.utils import get_column_letter
    for i, col in enumerate(df.columns, start=1):
        width = max([len(str(col))] + [len(str(v)) for v in df[col].head(500)]) + 2
        ws.column_dimensions[get_column_letter(i)].width = min(max(width, 8), 40)
    ws.freeze_panes = "A2"

def to_excel_sheets(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
            _autofit(writer.sheets[name[:31]], df)
    return output.getvalue()

def to_excel_bytes(df, sheet_name="DanhSach"): return to_excel_sheets({sheet_name: df})

def quarter_of(ts): return (ts.month - 1) // 3 + 1

def filter_quarter(df, col, year, quarter):
    mask = df[col].apply(lambda t: pd.notna(t) and t.year == year and quarter_of(t) == quarter)
    return df[mask].sort_values(col)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

def quarter_export_ui(df, date_col, cols, title, file_prefix, key):
    st.subheader(title)
    today = date.today()
    c1, c2 = st.columns(2)
    nam_x = c1.number_input("Chọn Năm", min_value=2000, max_value=2100, value=today.year, step=1, key=f"{key}_nam")
    quy_x = c2.selectbox("Chọn Quý", [1, 2, 3, 4], index=quarter_of(today) - 1,
                         format_func=lambda x: f"Quý {x} (tháng {3 * x - 2}-{3 * x})", key=f"{key}_quy")
    if df.empty:
        st.info("Chưa có dữ liệu.")
        return

    def prep(sub):
        out = sub[cols].copy()
        out.insert(0, "STT", range(1, len(out) + 1))
        return out

    df_q = prep(filter_quarter(df, date_col, int(nam_x), int(quy_x)))
    if df_q.empty:
        st.info(f"Không có cán bộ đến hạn trong Quý {quy_x}/{int(nam_x)}.")
    else:
        st.dataframe(df_q, hide_index=True, **W)
        st.download_button(f"📥 Xuất Excel Quý {quy_x}/{int(nam_x)}", data=to_excel_bytes(df_q, f"Quy{quy_x}_{int(nam_x)}"),
                           mime=XLSX_MIME, file_name=f"{file_prefix}_Q{quy_x}_{int(nam_x)}.xlsx", key=f"{key}_dl_q")
    all_sheets = {f"Quy{i}": prep(filter_quarter(df, date_col, int(nam_x), i)) for i in range(1, 5)}
    if any(not v.empty for v in all_sheets.values()):
        st.download_button(f"📥 Xuất cả năm {int(nam_x)} (4 sheet theo quý)", data=to_excel_sheets(all_sheets),
                           mime=XLSX_MIME, file_name=f"{file_prefix}_Nam{int(nam_x)}.xlsx", key=f"{key}_dl_year")

def generate_excel_template():
    cols = ["Mã VC", "Họ và tên", "Giới tính", "Ngày sinh", "Dân tộc", "Ngày tuyển dụng",
            "Ngày vào Đảng", "Trình độ", "Chuyên ngành", "Ghi chú"]
    return to_excel_bytes(pd.DataFrame(columns=cols))

def style_cham_cong(val):
    return {
        "P": "color: red; font-weight: bold;",
        "VR": "color: blue; font-weight: bold;",
        "K": "color: white; background-color: darkred; font-weight: bold;",
        "CT": "color: purple; font-weight: bold;",
        "O": "color: #ff9800; font-weight: bold;",
        "X": "color: #4caf50;",
    }.get(val, "")

def style_days(df, day_cols):
    s = df.style
    fn = s.map if hasattr(s, "map") else s.applymap
    return fn(style_cham_cong, subset=day_cols)

# ==========================================
# 3. KHỞI TẠO BẢNG TỰ ĐỘNG
# ==========================================
# 3. KHỞI TẠO BẢNG TỰ ĐỘNG (ĐÃ TỐI ƯU ĐÁM MÂY)
# ==========================================
VC_EXTRA_COLS = {
    "he_so_luong": "REAL", "phu_cap": "REAL", "ngay_huong_luong": "DATE", "chu_ky_nang_luong": "INTEGER",
    "muc_tham_nien": "REAL", "ngay_huong_tham_nien": "DATE", "chu_ky_nang_tham_nien": "INTEGER",
}

@st.cache_resource(show_spinner=False)
def init_db():
    try:
        with db() as conn:
            if USE_POSTGRES:
                with conn.cursor() as c:
                    c.execute("""CREATE TABLE IF NOT EXISTS users (
                                 id SERIAL PRIMARY KEY, username TEXT UNIQUE NOT NULL, password TEXT NOT NULL, 
                                 role TEXT NOT NULL, ma_vien_chuc TEXT)""")
                    c.execute("""CREATE TABLE IF NOT EXISTS chuc_vu (id SERIAL PRIMARY KEY, ten_chuc_vu TEXT NOT NULL)""")
                    c.execute("""CREATE TABLE IF NOT EXISTS vien_chuc (
                                    id SERIAL PRIMARY KEY, ma_vien_chuc TEXT UNIQUE NOT NULL, ho_ten TEXT NOT NULL,
                                    gioi_tinh TEXT, ngay_sinh DATE, dan_toc TEXT, chuc_vu_id INTEGER, ngay_tuyen_dung DATE,
                                    ngay_vao_dang DATE, hang_chuc_danh TEXT, trinh_do TEXT, ten_truong_dao_tao TEXT,
                                    chuyen_nganh TEXT, hinh_thuc_dao_tao TEXT, thac_si_chuyen_nganh TEXT, quan_ly_nha_nuoc TEXT,
                                    ly_luan_chinh_tri TEXT, tin_hoc TEXT, ngoai_ngu TEXT, chung_chi_khac TEXT, ghi_chu TEXT,
                                    he_so_luong REAL, phu_cap REAL, ngay_huong_luong DATE, chu_ky_nang_luong INTEGER,
                                    muc_tham_nien REAL, ngay_huong_tham_nien DATE, chu_ky_nang_tham_nien INTEGER,
                                    FOREIGN KEY(chuc_vu_id) REFERENCES chuc_vu(id))""")
                    c.execute("""CREATE TABLE IF NOT EXISTS cham_cong (
                                    id SERIAL PRIMARY KEY, vien_chuc_id INTEGER NOT NULL,
                                    ngay_cham_cong DATE NOT NULL, trang_thai TEXT DEFAULT 'X', ghi_chu TEXT,
                                    UNIQUE(vien_chuc_id, ngay_cham_cong))""")
                    c.execute("""CREATE TABLE IF NOT EXISTS khen_thuong (
                                    id SERIAL PRIMARY KEY, vien_chuc_id INTEGER NOT NULL,
                                    ngay_thang DATE, thanh_tich TEXT, danh_hieu TEXT, hinh_thuc TEXT,
                                    cap_khen TEXT, nam_khen INTEGER, so_quyet_dinh TEXT,
                                    FOREIGN KEY(vien_chuc_id) REFERENCES vien_chuc(id))""")
                    c.execute("""CREATE TABLE IF NOT EXISTS vu_viec (
                                    id SERIAL PRIMARY KEY, ngay_thang DATE, noi_dung TEXT, phuong_an TEXT)""")

                    def add_missing(table, cols):
                        c.execute("SELECT column_name FROM information_schema.columns WHERE table_name = %s", (table,))
                        existing = [r[0] for r in c.fetchall()]
                        for name, typ in cols.items():
                            if name not in existing:
                                c.execute(f"ALTER TABLE {table} ADD COLUMN {name} {typ}")
                    
                    add_missing("users", {"ma_vien_chuc": "TEXT"})
                    add_missing("vien_chuc", VC_EXTRA_COLS)

                    # Tối ưu chèn Admin: Dùng ON CONFLICT để bỏ qua nếu đã tồn tại, tránh lỗi UniqueViolation
                    c.execute("""INSERT INTO users (username, password, role, ma_vien_chuc) 
                                 VALUES (%s, %s, %s, '') 
                                 ON CONFLICT (username) DO NOTHING""",
                              ("admin", make_hashes("admin123"), "admin"))
                    
                    for old, new in CHUC_VU_CU.items():
                        c.execute("SELECT 1 FROM chuc_vu WHERE ten_chuc_vu=%s", (new,))
                        if not c.fetchone():
                            c.execute("UPDATE chuc_vu SET ten_chuc_vu=%s WHERE ten_chuc_vu=%s", (new, old))
                    for ten in CHUC_VU_LIST:
                        c.execute("SELECT 1 FROM chuc_vu WHERE ten_chuc_vu=%s", (ten,))
                        if not c.fetchone():
                            c.execute("INSERT INTO chuc_vu (ten_chuc_vu) VALUES (%s)", (ten,))
                    
                    placeholders = ','.join(['%s'] * len(CHUC_VU_LIST))
                    c.execute(f"""DELETE FROM chuc_vu WHERE ten_chuc_vu NOT IN ({placeholders})
                                  AND id NOT IN (SELECT chuc_vu_id FROM vien_chuc WHERE chuc_vu_id IS NOT NULL)""", CHUC_VU_LIST)
            else:
                c = conn.cursor()
                c.execute('''CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT,
                             username TEXT UNIQUE NOT NULL, password TEXT NOT NULL, role TEXT NOT NULL, ma_vien_chuc TEXT)''')
                c.execute('''CREATE TABLE IF NOT EXISTS chuc_vu (id INTEGER PRIMARY KEY AUTOINCREMENT, ten_chuc_vu TEXT NOT NULL)''')
                c.execute('''CREATE TABLE IF NOT EXISTS vien_chuc (
                                id INTEGER PRIMARY KEY AUTOINCREMENT, ma_vien_chuc TEXT UNIQUE NOT NULL, ho_ten TEXT NOT NULL,
                                gioi_tinh TEXT, ngay_sinh DATE, dan_toc TEXT, chuc_vu_id INTEGER, ngay_tuyen_dung DATE,
                                ngay_vao_dang DATE, hang_chuc_danh TEXT, trinh_do TEXT, ten_truong_dao_tao TEXT,
                                chuyen_nganh TEXT, hinh_thuc_dao_tao TEXT, thac_si_chuyen_nganh TEXT, quan_ly_nha_nuoc TEXT,
                                ly_luan_chinh_tri TEXT, tin_hoc TEXT, ngoai_ngu TEXT, chung_chi_khac TEXT, ghi_chu TEXT,
                                he_so_luong REAL, phu_cap REAL, ngay_huong_luong DATE, chu_ky_nang_luong INTEGER,
                                muc_tham_nien REAL, ngay_huong_tham_nien DATE, chu_ky_nang_tham_nien INTEGER,
                                FOREIGN KEY(chuc_vu_id) REFERENCES chuc_vu(id))''')
                c.execute('''CREATE TABLE IF NOT EXISTS cham_cong (
                                id INTEGER PRIMARY KEY AUTOINCREMENT, vien_chuc_id INTEGER NOT NULL,
                                ngay_cham_cong DATE NOT NULL, trang_thai TEXT DEFAULT 'X', ghi_chu TEXT,
                                UNIQUE(vien_chuc_id, ngay_cham_cong))''')
                c.execute('''CREATE TABLE IF NOT EXISTS khen_thuong (
                                id INTEGER PRIMARY KEY AUTOINCREMENT, vien_chuc_id INTEGER NOT NULL,
                                ngay_thang DATE, thanh_tich TEXT, danh_hieu TEXT, hinh_thuc TEXT,
                                cap_khen TEXT, nam_khen INTEGER, so_quyet_dinh TEXT,
                                FOREIGN KEY(vien_chuc_id) REFERENCES vien_chuc(id))''')
                c.execute('''CREATE TABLE IF NOT EXISTS vu_viec (
                                id INTEGER PRIMARY KEY AUTOINCREMENT, ngay_thang DATE, noi_dung TEXT, phuong_an TEXT)''')

                def add_missing(table, cols):
                    existing = [r[1] for r in c.execute(f"PRAGMA table_info({table})").fetchall()]
                    for name, typ in cols.items():
                        if name not in existing: c.execute(f"ALTER TABLE {table} ADD COLUMN {name} {typ}")

                add_missing("users", {"ma_vien_chuc": "TEXT"})
                add_missing("vien_chuc", VC_EXTRA_COLS)

                c.execute("INSERT OR IGNORE INTO users (username, password, role, ma_vien_chuc) VALUES (?, ?, ?, '')",
                          ("admin", make_hashes("admin123"), "admin"))
                for old, new in CHUC_VU_CU.items():
                    if not c.execute("SELECT 1 FROM chuc_vu WHERE ten_chuc_vu=?", (new,)).fetchone():
                        c.execute("UPDATE chuc_vu SET ten_chuc_vu=? WHERE ten_chuc_vu=?", (new, old))
                for ten in CHUC_VU_LIST:
                    if not c.execute("SELECT 1 FROM chuc_vu WHERE ten_chuc_vu=?", (ten,)).fetchone():
                        c.execute("INSERT INTO chuc_vu (ten_chuc_vu) VALUES (?)", (ten,))
                c.execute(f"""DELETE FROM chuc_vu WHERE ten_chuc_vu NOT IN ({','.join('?' * len(CHUC_VU_LIST))})
                              AND id NOT IN (SELECT chuc_vu_id FROM vien_chuc WHERE chuc_vu_id IS NOT NULL)""", CHUC_VU_LIST)
    except Exception as e:
        print(f"Lỗi khởi tạo DB: {e}")

init_db()

# ==========================================
# 4. MÀN HÌNH ĐĂNG NHẬP
# ==========================================
def logout():
    st.session_state.update({"logged_in": False, "username": "", "role": "", "ma_vien_chuc": "", "flash": ""})
    st.rerun()

if not st.session_state["logged_in"]:
    st.markdown("<h1 style='text-align: center; color: #1E88E5; margin-top: 50px;'>🏫 HỆ THỐNG QUẢN LÝ NHÂN SỰ</h1>", unsafe_allow_html=True)
    st.markdown("---")
    _, col2, _ = st.columns([1, 1, 1])
    with col2:
        with st.form("login_form"):
            st.subheader("🔒 Đăng nhập")
            username = st.text_input("Tài khoản").strip()
            password = st.text_input("Mật khẩu", type="password")
            submitted = st.form_submit_button("🚀 Đăng nhập", **W)
        if submitted:
            res = q("SELECT password, role, ma_vien_chuc FROM users WHERE username=?", (username,))
            if not res.empty and check_hashes(password, res.iloc[0]["password"]):
                st.session_state.update({"logged_in": True, "username": username, "role": res.iloc[0]["role"], "ma_vien_chuc": safe_str(res.iloc[0]["ma_vien_chuc"])})
                st.rerun()
            else:
                st.error("❌ Tài khoản hoặc mật khẩu không chính xác!")
    st.stop()

is_admin = st.session_state["role"] == "admin"
current_ma_vc = st.session_state["ma_vien_chuc"]

if not is_admin and not current_ma_vc:
    st.error("Tài khoản của bạn chưa được liên kết với hồ sơ Cán bộ nào. Vui lòng liên hệ Quản trị viên.")
    if st.button("Đăng xuất"): logout()
    st.stop()

def get_current_vc_id():
    if is_admin: return None
    res = q("SELECT id FROM vien_chuc WHERE ma_vien_chuc = ?", (current_ma_vc,))
    return int(res.iloc[0]["id"]) if not res.empty else None

def scope_where(alias="vc", prefix="WHERE"):
    if is_admin: return "", ()
    return f"{prefix} {alias}.ma_vien_chuc = ?", (current_ma_vc,)

def delete_vien_chuc(ids):
    ids = [int(i) for i in ids]
    if not ids: return
    ph = ",".join("?" * len(ids))
    with db() as conn:
        if USE_POSTGRES:
            cur = conn.cursor()
            cur.execute(f"SELECT ma_vien_chuc FROM vien_chuc WHERE id IN ({ph})".replace('?', '%s'), ids)
            codes = [r[0] for r in cur.fetchall()]
            cur.execute(f"DELETE FROM cham_cong WHERE vien_chuc_id IN ({ph})".replace('?', '%s'), ids)
            cur.execute(f"DELETE FROM khen_thuong WHERE vien_chuc_id IN ({ph})".replace('?', '%s'), ids)
            cur.execute(f"DELETE FROM vien_chuc WHERE id IN ({ph})".replace('?', '%s'), ids)
            if codes:
                cur.execute(f"UPDATE users SET ma_vien_chuc='' WHERE ma_vien_chuc IN ({','.join('?' * len(codes))})".replace('?', '%s'), codes)
            cur.close()
        else:
            codes = [r[0] for r in conn.execute(f"SELECT ma_vien_chuc FROM vien_chuc WHERE id IN ({ph})", ids)]
            conn.execute(f"DELETE FROM cham_cong WHERE vien_chuc_id IN ({ph})", ids)
            conn.execute(f"DELETE FROM khen_thuong WHERE vien_chuc_id IN ({ph})", ids)
            conn.execute(f"DELETE FROM vien_chuc WHERE id IN ({ph})", ids)
            if codes:
                conn.execute(f"UPDATE users SET ma_vien_chuc='' WHERE ma_vien_chuc IN ({','.join('?' * len(codes))})", codes)

df_chuc_vu = q("SELECT * FROM chuc_vu")
_order = lambda t: CHUC_VU_LIST.index(t) if t in CHUC_VU_LIST else len(CHUC_VU_LIST)
df_chuc_vu = df_chuc_vu.assign(_o=df_chuc_vu["ten_chuc_vu"].map(_order)).sort_values(["_o", "id"])
CV_NAMES = dict(zip(df_chuc_vu["id"].astype(int).tolist(), df_chuc_vu["ten_chuc_vu"].tolist()))

st.sidebar.title("🏫 Menu Chức Năng")
st.sidebar.info(f"👤 Cán bộ: **{st.session_state['username'].upper()}**\n\n🔑 Quyền: **{'Admin' if is_admin else 'Giáo viên'}**\n\n🌐 Mạng: **{'Postgres Đám mây' if USE_POSTGRES else 'SQLite Local'}**")

menu_options = ["📊 Tổng quan", "👥 Quản lý Nhân sự", "⏱️ Chấm công", "💰 Quản lý Lương", "⏳ Quản lý Thâm niên", "🏆 Thi đua - Khen thưởng", "🚨 Quản lý Vụ việc"]
if is_admin: menu_options.append("⚙️ Quản lý Tài khoản")
menu = st.sidebar.radio("Điều hướng", menu_options)
st.sidebar.markdown("---")
if st.sidebar.button("🚪 Đăng xuất", **W): logout()

if st.session_state.get("flash"): st.success(st.session_state.pop("flash"))

# ==========================================
# CÁC MODULE CHỨC NĂNG
# ==========================================
def vc_fields(info=None):
    edit = info is not None
    g = (lambda k: safe_str(info[k])) if edit else (lambda k: "")
    d = (lambda k: date_or_none(info[k])) if edit else (lambda k: None)

    st.markdown("**1. Thông tin cơ bản**")
    c1, c2, c3, c4 = st.columns(4)
    ma = c1.text_input("Mã viên chức (*)", value=g("ma_vien_chuc"), disabled=edit)
    ho_ten = c2.text_input("Họ và tên (*)", value=g("ho_ten"))
    gioi_tinh = c3.selectbox("Giới tính", GIOI_TINH, index=pick(GIOI_TINH, g("gioi_tinh")))
    ngay_sinh = c4.date_input("Ngày sinh", value=d("ngay_sinh"), min_value=MIN_DATE, max_value=MAX_DATE, format=DATE_FORMAT)

    c5, c6, c7, c8 = st.columns(4)
    dan_toc = c5.text_input("Dân tộc", value=g("dan_toc"))
    cv_ids = list(CV_NAMES)
    cur_cv = info["chuc_vu_id"] if edit else None
    cv_idx = cv_ids.index(int(cur_cv)) if edit and pd.notna(cur_cv) and int(cur_cv) in cv_ids else 0
    chuc_vu_id = c6.selectbox("Chức vụ", options=cv_ids, index=cv_idx, format_func=CV_NAMES.get)
    ngay_td = c7.date_input("Ngày tuyển dụng", value=d("ngay_tuyen_dung"), min_value=MIN_DATE, max_value=MAX_DATE, format=DATE_FORMAT)
    ngay_dang = c8.date_input("Ngày vào Đảng", value=d("ngay_vao_dang"), min_value=MIN_DATE, max_value=MAX_DATE, format=DATE_FORMAT)

    st.markdown("**2. Thông tin Trình độ & Đào tạo**")
    c9, c10, c11, c12 = st.columns(4)
    trinh_do = c9.selectbox("Trình độ", TRINH_DO, index=pick(TRINH_DO, g("trinh_do")))
    truong_dt = c10.text_input("Tên trường đào tạo", value=g("ten_truong_dao_tao"))
    chuyen_nganh = c11.text_input("Chuyên ngành", value=g("chuyen_nganh"))
    ht_dt = c12.text_input("Hình thức đào tạo", value=g("hinh_thuc_dao_tao"))

    st.markdown("**3. Hạng chức danh & Các chứng chỉ**")
    c13, c14, c15 = st.columns(3)
    hang = c13.selectbox("Hạng CDNN", HANG_CDNN, index=pick(HANG_CDNN, g("hang_chuc_danh")))
    thac_si = c14.text_input("Thạc sĩ chuyên ngành", value=g("thac_si_chuyen_nganh"))
    qlnn = c15.text_input("Quản lý nhà nước", value=g("quan_ly_nha_nuoc"))

    c16, c17, c18, c19 = st.columns(4)
    llct = c16.text_input("Lý luận chính trị", value=g("ly_luan_chinh_tri"))
    tin_hoc = c17.text_input("Tin học", value=g("tin_hoc"))
    ngoai_ngu = c18.text_input("Ngoại ngữ", value=g("ngoai_ngu"))
    cc_khac = c19.text_input("Chứng chỉ khác", value=g("chung_chi_khac"))
    ghi_chu = st.text_input("Ghi chú thêm", value=g("ghi_chu"))

    vals = {
        "ho_ten": ho_ten.strip(), "gioi_tinh": gioi_tinh, "ngay_sinh": sql_date(ngay_sinh), "dan_toc": dan_toc,
        "chuc_vu_id": int(chuc_vu_id), "ngay_tuyen_dung": sql_date(ngay_td), "ngay_vao_dang": sql_date(ngay_dang),
        "hang_chuc_danh": hang, "trinh_do": trinh_do, "ten_truong_dao_tao": truong_dt, "chuyen_nganh": chuyen_nganh,
        "hinh_thuc_dao_tao": ht_dt, "thac_si_chuyen_nganh": thac_si, "quan_ly_nha_nuoc": qlnn,
        "ly_luan_chinh_tri": llct, "tin_hoc": tin_hoc, "ngoai_ngu": ngoai_ngu, "chung_chi_khac": cc_khac, "ghi_chu": ghi_chu,
    }
    return ma.strip(), vals

if menu == "📊 Tổng quan":
    st.title("📊 Bảng điều khiển Tổng quan")
    row = q("""SELECT COUNT(*) AS tong,
                      COALESCE(SUM(CASE WHEN gioi_tinh='Nam' THEN 1 ELSE 0 END), 0) AS nam,
                      COALESCE(SUM(CASE WHEN gioi_tinh='Nữ' THEN 1 ELSE 0 END), 0) AS nu FROM vien_chuc""").iloc[0]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("👥 Tổng số CB/GV", f"{int(row['tong'])} người")
    c2.metric("👨 Nam", f"{int(row['nam'])} người")
    c3.metric("👩 Nữ", f"{int(row['nu'])} người")
    c4.metric("🏆 Đang công tác", f"{int(row['tong'])} người")

elif menu == "👥 Quản lý Nhân sự":
    st.title("👥 Quản lý Hồ sơ Cán bộ, Giáo viên")
    if is_admin:
        tabs = st.tabs(["📋 Danh sách Đơn vị", "➕ Thêm mới", "✏️ Cập nhật & Xóa", "📂 Nhập Excel"])
        tab_list, tab_add, tab_
