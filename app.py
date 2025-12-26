import streamlit as st
import pandas as pd
import numpy as np
import folium
from streamlit_folium import st_folium
from pyproj import Transformer
import geopandas as gpd
from shapely.geometry import Point
import os
import re
from datetime import datetime, date
import plotly.express as px
import io
import json
from sqlalchemy import create_engine, text

# =========================================================
# 0) 基本設定
# =========================================================
st.set_page_config(page_title="彰化農地定監管理系統", layout="wide")

EXCEL_PATH = "彰化農地管理資料庫.xlsx"
METALS = ['汞', '砷', '銅', '鉻', '鎘', '鉛', '鋅', '鎳']
transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)

# --- Secrets / Env ---
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "")
DATABASE_URL = st.secrets.get("DATABASE_URL", "") or os.environ.get("DATABASE_URL", "")
DB_ENABLED = bool(DATABASE_URL)

# =========================================================
# 1) CSS 美化（保留你原本風格）
# =========================================================
st.markdown("""
<style>
    th { color: #0a3d2a !important; }
    .grid-box { padding: 10px; border-radius: 10px; margin-bottom: 8px; font-weight: 700; text-align:center; }
    .bg-persistent { background: #ffe6e6; border: 1px solid #ffb3b3; }
    .bg-prolonged { background: #e6f0ff; border: 1px solid #b3d1ff; }
    .bg-exited { background: #e6ffe6; border: 1px solid #b3ffb3; }
    .small-note { color:#666; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

# =========================================================
# 2) 小工具函數
# =========================================================
def clean_id(val):
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

def get_minguo_date():
    now = datetime.now()
    return f"民國 {now.year - 1911} 年 {now.month} 月 {now.day} 日"

def to_float(x):
    try:
        if pd.isna(x): return None
        return float(x)
    except:
        return None

def dist_m(x1, y1, x2, y2):
    if None in [x1,y1,x2,y2]: return None
    return float(((x1-x2)**2 + (y1-y2)**2) ** 0.5)

def safe_str(x):
    if pd.isna(x): return ""
    return str(x).strip()

def ensure_session_state():
    if "admin_ok" not in st.session_state:
        st.session_state.admin_ok = False
    if "batch_new_samples" not in st.session_state:
        st.session_state.batch_new_samples = []  # list of dicts

ensure_session_state()

# =========================================================
# 3) DB 連線 / 初始化
# =========================================================
@st.cache_resource
def get_engine():
    if not DB_ENABLED:
        return None
    return create_engine(DATABASE_URL, pool_pre_ping=True)

engine = get_engine() if DB_ENABLED else None

def db_init():
    """建立最小可用表：lands / samples / standards / blocks"""
    if not DB_ENABLED:
        return
    with engine.begin() as conn:
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS lands (
            land_id SERIAL PRIMARY KEY,
            lot_no TEXT UNIQUE NOT NULL,
            sgm_id TEXT,
            land_no TEXT,
            grid_id TEXT,
            township TEXT,
            survey_type TEXT,          -- 系統型農地 / 個案型農地
            rep TEXT,                  -- 代表點 / 備用點 / ...
            irrigation_type TEXT,
            x DOUBLE PRECISION,
            y DOUBLE PRECISION,

            admin_status TEXT,         -- 農地監測狀態：監測/管制/建物/正常/難以採樣
            metal_result TEXT,         -- 重金屬判定結果：增量/延長/正常/管制
            freq TEXT,                 -- 監測頻率：持續/延長/退場/管制
            last_year INTEGER,

            extra JSONB DEFAULT '{}'::jsonb,
            updated_at TIMESTAMP DEFAULT NOW()
        );
        """))
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS samples (
            sample_id SERIAL PRIMARY KEY,
            lot_no TEXT NOT NULL,
            year INTEGER NOT NULL,
            sample_date DATE,

            coord_x DOUBLE PRECISION,
            coord_y DOUBLE PRECISION,

            photo_site TEXT,
            photo_sample TEXT,

            xrf JSONB,
            total JSONB,
            used_total BOOLEAN DEFAULT FALSE,

            admin_status TEXT,        -- 監測/管制/建物/正常/難以採樣（本次現勘）
            metal_result TEXT,        -- 增量/延長/正常/管制（本次判定）
            freq TEXT,                -- 持續/延長/退場/管制（策略頻率）

            da_pct JSONB,             -- 你的 DA% 定義
            er JSONB,                 -- ER=濃度/監測標準

            created_at TIMESTAMP DEFAULT NOW()
        );
        """))
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS standards (
            metal TEXT PRIMARY KEY,
            monitor_std DOUBLE PRECISION,
            control_std DOUBLE PRECISION,
            da_threshold DOUBLE PRECISION
        );
        """))
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS blocks (
            block_id TEXT NOT NULL,
            lot_no TEXT NOT NULL,
            is_rep BOOLEAN DEFAULT FALSE,
            PRIMARY KEY (block_id, lot_no)
        );
        """))

db_init()

# =========================================================
# 4) Excel 讀取（fallback）
# =========================================================
@st.cache_data
def load_all_data_from_excel():
    if not os.path.exists(EXCEL_PATH):
        return None, None, None, None
    try:
        xl = pd.ExcelFile(EXCEL_PATH)
        actual_sheets = xl.sheet_names

        def get_s(n):
            return next((s for s in actual_sheets if n == s.strip()), None)

        df_m = pd.read_excel(xl, sheet_name=get_s("農地現況主檔"))
        df_h = pd.read_excel(xl, sheet_name=get_s("歷年調查紀錄"))
        df_b = pd.read_excel(xl, sheet_name=get_s("同坵塊對照表"))
        df_s = pd.read_excel(xl, sheet_name=get_s("判定標準表"))

        # clean
        df_m.columns = df_m.columns.astype(str).str.strip()
        if "網格編號" in df_m.columns:
            df_m["網格編號"] = df_m["網格編號"].apply(clean_id)
        if "地段地號" in df_m.columns:
            df_m["地段地號"] = df_m["地段地號"].astype(str).str.strip()

        df_h.columns = df_h.columns.astype(str).str.strip()
        if "SGM編號" in df_h.columns:
            df_h["SGM編號"] = df_h["SGM編號"].apply(clean_id)

        df_b.columns = df_b.columns.astype(str).str.strip()
        df_s.columns = df_s.columns.astype(str).str.strip()

        return df_m, df_h, df_b, df_s
    except Exception as e:
        st.error(f"Excel 讀取失敗：{e}")
        return None, None, None, None

# =========================================================
# 5) DB 讀取（優先）
# =========================================================
def load_all_data_from_db():
    if not DB_ENABLED:
        return None, None, None, None
    try:
        df_m = pd.read_sql("SELECT * FROM lands", engine)
        df_h = pd.read_sql("SELECT * FROM samples", engine)
        df_b = pd.read_sql("SELECT * FROM blocks", engine)
        df_s = pd.read_sql("SELECT * FROM standards", engine)
        return df_m, df_h, df_b, df_s
    except Exception as e:
        st.warning(f"DB 讀取失敗，改用 Excel：{e}")
        return None, None, None, None

# 主資料載入（DB 優先，否則 Excel）
df_master, df_history, df_block, df_settings = load_all_data_from_db()
if df_master is None:
    df_master, df_history, df_block, df_settings = load_all_data_from_excel()

# =========================================================
# 6) 管理員登入（可選）
# =========================================================
def admin_login_box():
    with st.sidebar.expander("🔐 管理員登入", expanded=False):
        if st.session_state.admin_ok:
            st.success("已登入管理員")
            if st.button("登出"):
                st.session_state.admin_ok = False
        else:
            pwd = st.text_input("管理員密碼", type="password")
            if st.button("登入"):
                if ADMIN_PASSWORD and pwd == ADMIN_PASSWORD:
                    st.session_state.admin_ok = True
                    st.success("登入成功")
                else:
                    st.error("密碼錯誤或尚未設定 ADMIN_PASSWORD")

# =========================================================
# 7) 規則：重金屬判定 → 監測頻率
#    注意：這裡先給「可用的預設規則」，你之後要更精準我再幫你改成 rules table
# =========================================================
def standards_dict(df_std: pd.DataFrame):
    """
    回傳 dict: metal -> {monitor_std, control_std, da_threshold}
    """
    d = {}
    if df_std is None or len(df_std) == 0:
        return d
    # Excel/DB 欄位名可能不同，做兼容
    cols = [c.strip() for c in df_std.columns]
    # DB: metal/monitor_std/control_std/da_threshold
    if set(["metal", "monitor_std", "control_std", "da_threshold"]).issubset(set(cols)):
        for _, r in df_std.iterrows():
            m = safe_str(r["metal"])
            d[m] = {
                "monitor": to_float(r["monitor_std"]),
                "control": to_float(r["control_std"]),
                "da": to_float(r["da_threshold"])
            }
        return d

    # Excel: 項目名稱/監測標準/管制標準/上升標準 (DA門檻)
    if set(["項目名稱", "監測標準", "管制標準", "上升標準 (DA門檻)"]).issubset(set(cols)):
        for _, r in df_std.iterrows():
            m = safe_str(r["項目名稱"])
            d[m] = {
                "monitor": to_float(r["監測標準"]),
                "control": to_float(r["管制標準"]),
                "da": to_float(r["上升標準 (DA門檻)"])
            }
        return d

    # 容錯：嘗試猜欄位
    name_col = next((c for c in cols if "名稱" in c), None)
    mon_col = next((c for c in cols if "監測" in c), None)
    ctl_col = next((c for c in cols if "管制" in c), None)
    da_col = next((c for c in cols if "上升" in c or "DA" in c), None)
    if name_col and mon_col and ctl_col:
        for _, r in df_std.iterrows():
            m = safe_str(r[name_col])
            d[m] = {"monitor": to_float(r[mon_col]), "control": to_float(r[ctl_col]), "da": to_float(r[da_col]) if da_col else None}
    return d

STD = standards_dict(df_settings)

def choose_concentration(xrf_dict, total_dict):
    """
    若全量有填（任一金屬非空），就用全量覆蓋作為計算濃度；否則用 XRF
    """
    used_total = False
    conc = {}
    total_has_any = False
    if isinstance(total_dict, dict):
        total_has_any = any(to_float(total_dict.get(m)) is not None for m in METALS)
    if total_has_any:
        used_total = True
        for m in METALS:
            conc[m] = to_float(total_dict.get(m))
    else:
        for m in METALS:
            conc[m] = to_float((xrf_dict or {}).get(m))
    return conc, used_total

def compute_er(conc: dict):
    er = {}
    for m, v in conc.items():
        mon = (STD.get(m, {}) or {}).get("monitor")
        if v is None or mon in [None, 0]:
            er[m] = None
        else:
            er[m] = float(v / mon)
    return er

def get_baseline_value(lot_no: str, metal: str):
    """
    先用最接近/最新的一筆歷史紀錄當 baseline（先可用）
    之後你要改成「納入定監當年 baseline」也可以做（會新增 baseline_year/值）
    """
    if df_history is None or len(df_history) == 0:
        return None
    sub = df_history[df_history["lot_no"].astype(str) == lot_no].copy() if "lot_no" in df_history.columns else None
    if sub is None or len(sub) == 0:
        # Excel history 若是用 SGM，這裡就先回 None（你以 lot_no 為主）
        return None
    # 取最新 year
    if "year" in sub.columns:
        sub = sub.sort_values("year", ascending=False)
    # baseline 用 used_total 的濃度（若 total 存在）
    for _, r in sub.iterrows():
        used_total = bool(r.get("used_total", False))
        if used_total and isinstance(r.get("total"), dict):
            v = to_float(r["total"].get(metal))
        else:
            v = to_float((r.get("xrf") or {}).get(metal)) if isinstance(r.get("xrf"), dict) else None
        if v is not None:
            return v
    return None

def compute_da_pct(lot_no: str, conc: dict):
    """
    DA% = ((本次 - baseline) / baseline) * 100
    baseline 先用「上一筆有值」(可用版)；之後可改成「納入定監當年」固定 baseline
    """
    da = {}
    for m, v in conc.items():
        base = get_baseline_value(lot_no, m)
        if v is None or base is None:
            da[m] = None
            continue
        if base == 0:
            if v == 0:
                da[m] = 0.0
            else:
                da[m] = float("inf")  # 保守：視為極大上升
        else:
            da[m] = float(((v - base) / base) * 100.0)
    return da

def judge_metal_result(conc: dict, er: dict, da_pct: dict):
    """
    重金屬判定結果：管制/增量/延長/正常
    預設規則（可用版）：
      1) 任一金屬 >= 管制標準 → 管制
      2) 否則 若任一 ER>=1 或 任一 DA% >= 對應門檻 → 增量
      3) 否則 若任一 ER>=0.7 → 延長（你可調）
      4) 否則 → 正常
    """
    # 1) 管制
    for m, v in conc.items():
        ctl = (STD.get(m, {}) or {}).get("control")
        if v is not None and ctl is not None and v >= ctl:
            return "管制"

    # 2) 增量：ER>=1 或 DA%超門檻
    for m in METALS:
        er_v = er.get(m)
        if er_v is not None and er_v >= 1.0:
            return "增量"
        da_v = da_pct.get(m)
        da_th = (STD.get(m, {}) or {}).get("da")
        if da_v is not None:
            if da_v == float("inf"):
                return "增量"
            if da_th is not None and da_v >= da_th:
                return "增量"

    # 3) 延長：ER>=0.7
    for m in METALS:
        er_v = er.get(m)
        if er_v is not None and er_v >= 0.7:
            return "延長"

    return "正常"

def metal_result_to_freq(metal_result: str):
    """
    你指定的對應：
      增量 -> 持續
      延長 -> 延長
      正常 -> 退場
      管制 -> 管制
    """
    if metal_result == "增量":
        return "持續"
    if metal_result == "延長":
        return "延長"
    if metal_result == "正常":
        return "退場"
    if metal_result == "管制":
        return "管制"
    return None

def derive_admin_status(admin_input: str, metal_result: str):
    """
    若使用者有填 admin_status（建物/難以採樣/管制/正常/監測）就以輸入為主；
    若沒填，依 metal_result 推：
      管制->管制
      正常->正常
      增量/延長->監測
    """
    ai = safe_str(admin_input)
    if ai in ["監測", "管制", "建物", "正常", "難以採樣"]:
        return ai
    if metal_result == "管制":
        return "管制"
    if metal_result == "正常":
        return "正常"
    if metal_result in ["增量", "延長"]:
        return "監測"
    return None

# =========================================================
# 8) DB 寫入工具
# =========================================================
def db_upsert_land(land: dict):
    """
    land: {lot_no, sgm_id, land_no, grid_id, township, survey_type, rep, irrigation_type, x, y,
           admin_status, metal_result, freq, last_year}
    """
    if not DB_ENABLED:
        return
    with engine.begin() as conn:
        conn.execute(text("""
        INSERT INTO lands (lot_no, sgm_id, land_no, grid_id, township, survey_type, rep, irrigation_type,
                           x, y, admin_status, metal_result, freq, last_year, updated_at)
        VALUES (:lot_no, :sgm_id, :land_no, :grid_id, :township, :survey_type, :rep, :irrigation_type,
                :x, :y, :admin_status, :metal_result, :freq, :last_year, NOW())
        ON CONFLICT (lot_no) DO UPDATE SET
            sgm_id=EXCLUDED.sgm_id,
            land_no=EXCLUDED.land_no,
            grid_id=EXCLUDED.grid_id,
            township=EXCLUDED.township,
            survey_type=EXCLUDED.survey_type,
            rep=EXCLUDED.rep,
            irrigation_type=EXCLUDED.irrigation_type,
            x=EXCLUDED.x,
            y=EXCLUDED.y,
            admin_status=EXCLUDED.admin_status,
            metal_result=EXCLUDED.metal_result,
            freq=EXCLUDED.freq,
            last_year=EXCLUDED.last_year,
            updated_at=NOW();
        """), land)

def db_insert_sample(sample: dict):
    if not DB_ENABLED:
        return
    with engine.begin() as conn:
        conn.execute(text("""
        INSERT INTO samples (lot_no, year, sample_date, coord_x, coord_y, photo_site, photo_sample,
                             xrf, total, used_total, admin_status, metal_result, freq, da_pct, er)
        VALUES (:lot_no, :year, :sample_date, :coord_x, :coord_y, :photo_site, :photo_sample,
                :xrf::jsonb, :total::jsonb, :used_total, :admin_status, :metal_result, :freq,
                :da_pct::jsonb, :er::jsonb);
        """), {
            **sample,
            "xrf": json.dumps(sample.get("xrf") or {}, ensure_ascii=False),
            "total": json.dumps(sample.get("total") or {}, ensure_ascii=False),
            "da_pct": json.dumps(sample.get("da_pct") or {}, ensure_ascii=False),
            "er": json.dumps(sample.get("er") or {}, ensure_ascii=False),
        })

def refresh_data():
    global df_master, df_history, df_block, df_settings, STD
    if DB_ENABLED:
        df_master, df_history, df_block, df_settings = load_all_data_from_db()
    else:
        df_master, df_history, df_block, df_settings = load_all_data_from_excel()
    STD = standards_dict(df_settings)

# =========================================================
# 9) Sidebar 導覽
# =========================================================
st.sidebar.title("🌿 系統導覽")
admin_login_box()

menu = st.sidebar.radio(
    "功能導覽",
    ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"]
)

# =========================================================
# 10) 若無資料
# =========================================================
if df_master is None:
    st.error("❌ 讀取資料失敗：請確認 Excel 檔案在 repo 中，或 DB 已正確設定 DATABASE_URL")
    st.stop()

# =========================================================
# 11) 統計首頁（保留你原本重點）
# =========================================================
def compute_dashboard_counts(df_m: pd.DataFrame):
    # 兼容：DB欄位/Excel欄位
    if "lot_no" in df_m.columns:
        abs_total = len(df_m)
        rep_col = "rep" if "rep" in df_m.columns else None
        admin_col = "admin_status" if "admin_status" in df_m.columns else None
    else:
        abs_total = len(df_m)
        rep_col = "代表性" if "代表性" in df_m.columns else None
        admin_col = "農地監測狀態" if "農地監測狀態" in df_m.columns else None

    sampling_pts = 0
    if rep_col:
        sampling_pts = len(df_m[df_m[rep_col].astype(str).isin(["代表點", "備用點"])])

    def count_contains(col, kw):
        if not col or col not in df_m.columns:
            return 0
        return int(df_m[col].astype(str).str.contains(kw, na=False).sum())

    control_count = count_contains(admin_col, "管制")
    build_count = count_contains(admin_col, "建物")
    hard_count = count_contains(admin_col, "難以採樣")
    normal_count = count_contains(admin_col, "正常")

    return abs_total, sampling_pts, control_count, build_count, hard_count, normal_count

# =========================================================
# 12) Excel→DB 一鍵匯入（管理員用）
# =========================================================
def admin_import_excel_to_db():
    st.subheader("🛠️ 管理員工具：Excel → DB 匯入")
    if not DB_ENABLED:
        st.warning("尚未設定 DATABASE_URL（Streamlit Secrets），目前不能匯入 DB。")
        return
    if not st.session_state.admin_ok:
        st.info("請先在側邊欄登入管理員")
        return

    st.caption("此工具會把 repo 裡的 Excel 主檔、標準表、同坵塊表匯入 DB。")
    if st.button("🚀 一鍵匯入（lands / standards / blocks）"):
        df_mx, df_hx, df_bx, df_sx = load_all_data_from_excel()
        if df_mx is None:
            st.error("Excel 讀不到，無法匯入。")
            return

        # 1) standards
        std_d = standards_dict(df_sx)
        with engine.begin() as conn:
            for m in METALS:
                v = std_d.get(m, {"monitor": None, "control": None, "da": None})
                conn.execute(text("""
                INSERT INTO standards (metal, monitor_std, control_std, da_threshold)
                VALUES (:metal, :monitor_std, :control_std, :da_threshold)
                ON CONFLICT (metal) DO UPDATE SET
                  monitor_std=EXCLUDED.monitor_std,
                  control_std=EXCLUDED.control_std,
                  da_threshold=EXCLUDED.da_threshold;
                """), {
                    "metal": m,
                    "monitor_std": v.get("monitor"),
                    "control_std": v.get("control"),
                    "da_threshold": v.get("da"),
                })

        # 2) blocks
        if df_bx is not None and len(df_bx) > 0:
            bcols = [c.strip() for c in df_bx.columns]
            # 預期：農地群組編號 / 農地地段地號 / 代表農地
            gcol = next((c for c in bcols if "群組" in c), None)
            lcol = next((c for c in bcols if "地段" in c), None)
            rcol = next((c for c in bcols if "代表" in c), None)
            if gcol and lcol:
                with engine.begin() as conn:
                    conn.execute(text("DELETE FROM blocks;"))
                    for _, r in df_bx.iterrows():
                        block_id = safe_str(r[gcol])
                        lot_no = safe_str(r[lcol])
                        is_rep = safe_str(r.get(rcol, "")) if rcol else ""
                        is_rep_bool = True if ("是" in is_rep) else False
                        if block_id and lot_no:
                            conn.execute(text("""
                            INSERT INTO blocks (block_id, lot_no, is_rep)
                            VALUES (:block_id, :lot_no, :is_rep);
                            """), {"block_id": block_id, "lot_no": lot_no, "is_rep": is_rep_bool})

        # 3) lands
        # 對應欄位（Excel→DB）
        # Excel: SGM編號/農地序號/網格編號/鄉鎮市/地段地號/調查方式/代表性/用水種類/TWD97_X/TWD97_Y/目前農地調查現況/農地監測狀態/網格監測頻率/最後調查年分
        with engine.begin() as conn:
            # 不直接 truncate，改 upsert
            for _, r in df_mx.iterrows():
                lot_no = safe_str(r.get("地段地號", ""))
                if not lot_no:
                    continue
                land = {
                    "lot_no": lot_no,
                    "sgm_id": safe_str(r.get("SGM編號", "")) if "SGM編號" in df_mx.columns else None,
                    "land_no": safe_str(r.get("農地序號", "")) if "農地序號" in df_mx.columns else None,
                    "grid_id": safe_str(r.get("網格編號", "")) if "網格編號" in df_mx.columns else None,
                    "township": safe_str(r.get("鄉鎮市", "")) if "鄉鎮市" in df_mx.columns else None,
                    "survey_type": safe_str(r.get("調查方式", "")) if "調查方式" in df_mx.columns else None,
                    "rep": safe_str(r.get("代表性", "")) if "代表性" in df_mx.columns else None,
                    "irrigation_type": safe_str(r.get("用水種類", "")) if "用水種類" in df_mx.columns else None,
                    "x": to_float(r.get("TWD97_X")) if "TWD97_X" in df_mx.columns else None,
                    "y": to_float(r.get("TWD97_Y")) if "TWD97_Y" in df_mx.columns else None,
                    "admin_status": safe_str(r.get("農地監測狀態", "")) if "農地監測狀態" in df_mx.columns else None,
                    "metal_result": safe_str(r.get("目前農地調查現況", "")) if "目前農地調查現況" in df_mx.columns else None,
                    "freq": safe_str(r.get("網格監測頻率", "")) if "網格監測頻率" in df_mx.columns else None,
                    "last_year": int(r.get("最後調查年分")) if str(r.get("最後調查年分", "")).strip().isdigit() else None,
                }
                db_upsert_land(land)

        refresh_data()
        st.success("✅ 匯入完成！DB 已更新。")

# =========================================================
# 13) 頁面：統計首頁
# =========================================================
if menu == "統計首頁":
    st.title("🚜 彰化縣農地監測戰情室")
    st.subheader(f"📅 當前時間：{get_minguo_date()}")

    abs_total, sampling_pts, control_count, build_count, hard_count, normal_count = compute_dashboard_counts(df_master)

    k = st.columns(6)
    k[0].metric("總資料點數", abs_total)
    k[1].metric("總採樣點數(代表+備用)", sampling_pts)
    k[2].metric("管制點數", control_count)
    k[3].metric("建物數量", build_count)
    k[4].metric("難以採樣數量", hard_count)
    k[5].metric("正常退場數量", normal_count)

    st.divider()

    # 管理員工具：匯入 DB
    admin_import_excel_to_db()

# =========================================================
# 14) 頁面：資料庫查詢與下載
# =========================================================
elif menu == "資料庫查詢與下載":
    st.title("🔎 資料庫查詢與下載")

    if "lot_no" in df_master.columns:
        lot_col = "lot_no"
        grid_col = "grid_id"
    else:
        lot_col = "地段地號"
        grid_col = "網格編號"

    q = st.text_input("輸入地段地號/關鍵字搜尋", "")
    view = df_master.copy()
    if q.strip():
        view = view[view[lot_col].astype(str).str.contains(q.strip(), na=False)]

    st.dataframe(view, use_container_width=True, height=520)

    # 下載
    csv = view.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ 下載目前篩選結果 CSV", data=csv, file_name="lands_filtered.csv", mime="text/csv")

# =========================================================
# 15) 頁面：115 年度篩選名 Fletcher (沿用你原本的概念，先做簡化版)
# =========================================================
elif menu == "新年度調查點篩選名單":
    st.title("📅 新年度調查點篩選名單（示範版）")
    st.caption("你原本的 115 名單篩選邏輯我先保留概念：增量必選、延長超過某年要選。後續我們可依你正式規則再精準化。")

    # 兼容欄位
    if "metal_result" in df_master.columns:
        mr_col = "metal_result"
        ly_col = "last_year"
        lot_col = "lot_no"
        grid_col = "grid_id"
        rep_col = "rep"
    else:
        mr_col = "目前農地調查現況"
        ly_col = "最後調查年分"
        lot_col = "地段地號"
        grid_col = "網格編號"
        rep_col = "代表性"

    cutoff = st.number_input("延長狀態：最後調查年分 ≤ 這一年就納入（例：113）", value=113, step=1)
    f_list = df_master[(df_master[mr_col] == '增量') | ((df_master[mr_col] == '延長') & (pd.to_numeric(df_master[ly_col], errors="coerce") <= cutoff))].copy()

    show_cols = [c for c in [grid_col, lot_col, mr_col, rep_col, ly_col] if c in f_list.columns]
    st.dataframe(f_list[show_cols], use_container_width=True, height=560)

# =========================================================
# 16) 頁面：新增年度調查結果（你要的核心功能）
# =========================================================
elif menu == "新增年度調查結果":
    st.title("➕ 新增年度調查結果（寫入 DB）")

    if not DB_ENABLED:
        st.warning("目前尚未設定 DATABASE_URL（Streamlit Secrets），此頁無法寫入 DB。請先完成 Neon + Secrets 設定。")
        st.stop()

    # 查詢欄位
    lot_col = "lot_no" if "lot_no" in df_master.columns else "地段地號"
    grid_col = "grid_id" if "grid_id" in df_master.columns else "網格編號"
    x_col = "x" if "x" in df_master.columns else "TWD97_X"
    y_col = "y" if "y" in df_master.columns else "TWD97_Y"
    rep_col = "rep" if "rep" in df_master.columns else "代表性"
    surv_col = "survey_type" if "survey_type" in df_master.columns else "調查方式"

    q = st.text_input("🔍 搜尋地段地號（例：華南段0159-0000）", "")
    if not q.strip():
        st.info("請先輸入地段地號開始查詢。")
        st.stop()

    matches = df_master[df_master[lot_col].astype(str).str.contains(q.strip(), na=False)].copy()
    if len(matches) == 0:
        st.error("找不到符合的地段地號。")
        st.stop()

    # 若多筆，讓使用者選
    if len(matches) > 1:
        sel = st.selectbox("找到多筆，請選擇一筆", matches[lot_col].astype(str).tolist())
        row = matches[matches[lot_col].astype(str) == sel].iloc[0]
    else:
        row = matches.iloc[0]

    lot_no = safe_str(row[lot_col])
    st.success(f"✅ 已選取：{lot_no}")

    # 顯示基本卡片
    with st.expander("📌 基本資料（主檔）", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        c1.write({"網格": safe_str(row.get(grid_col, "")), "調查方式": safe_str(row.get(surv_col, ""))})
        c2.write({"代表性": safe_str(row.get(rep_col, "")), "用水": safe_str(row.get("irrigation_type", row.get("用水種類","")))})
        c3.write({"TWD97_X": row.get(x_col), "TWD97_Y": row.get(y_col)})
        c4.write({"最後調查年分": row.get("last_year", row.get("最後調查年分","")), "現況": row.get("metal_result", row.get("目前農地調查現況",""))})

    st.divider()

    # 輸入表單
    with st.form("add_sample_form", clear_on_submit=False):
        colA, colB, colC = st.columns([1.1, 1.1, 1.3])
        with colA:
            year = st.number_input("調查年度（民國）", value=int(datetime.now().year - 1911), step=1)
            sample_date = st.date_input("採樣日期", value=date.today())
            admin_input = st.selectbox("本次行政狀態（現勘）", ["（自動判定）", "監測", "正常", "管制", "建物", "難以採樣"], index=0)
        with colB:
            coord_x = st.number_input("採樣點座標 X（TWD97）", value=float(row.get(x_col) or 0.0))
            coord_y = st.number_input("採樣點座標 Y（TWD97）", value=float(row.get(y_col) or 0.0))
            st.caption("若與主檔座標差異 > 3 m，系統會提醒")
        with colC:
            photo_site = st.text_input("現勘照片（URL/路徑/檔名）", "")
            photo_sample = st.text_input("採樣照片（URL/路徑/檔名）", "")

        st.markdown("### 🔬 XRF 八項")
        xrf_inputs = {}
        cols = st.columns(4)
        for i, m in enumerate(METALS):
            with cols[i % 4]:
                xrf_inputs[m] = st.number_input(f"XRF_{m}", value=0.0, step=0.1)

        st.markdown("### 🧪 全量八項（可選；若有填，系統以全量覆蓋 XRF 作計算）")
        total_inputs = {}
        cols2 = st.columns(4)
        for i, m in enumerate(METALS):
            with cols2[i % 4]:
                total_inputs[m] = st.text_input(f"全量_{m}（留空代表無）", "")

        submitted = st.form_submit_button("➕ 加入本次暫存清單")

    # 計算 + 加入 batch
    if submitted:
        # 將全量字串轉 float/None
        total_dict = {}
        for m in METALS:
            s = safe_str(total_inputs.get(m))
            total_dict[m] = float(s) if s and re.match(r"^-?\d+(\.\d+)?$", s) else None

        conc, used_total = choose_concentration(xrf_inputs, total_dict)
        er = compute_er(conc)
        da_pct = compute_da_pct(lot_no, conc)

        metal_result = judge_metal_result(conc, er, da_pct)
        freq = metal_result_to_freq(metal_result)

        admin_status = derive_admin_status(admin_input if admin_input!="（自動判定）" else "", metal_result)

        # 座標偏移提醒
        base_x = to_float(row.get(x_col))
        base_y = to_float(row.get(y_col))
        d = dist_m(base_x, base_y, coord_x, coord_y)
        warn_shift = (d is not None and d > 3.0)

        item = {
            "lot_no": lot_no,
            "year": int(year),
            "sample_date": sample_date,
            "coord_x": float(coord_x),
            "coord_y": float(coord_y),
            "photo_site": photo_site.strip(),
            "photo_sample": photo_sample.strip(),
            "xrf": {m: float(xrf_inputs[m]) for m in METALS},
            "total": {m: total_dict[m] for m in METALS if total_dict[m] is not None},
            "used_total": bool(used_total),
            "admin_status": admin_status,
            "metal_result": metal_result,
            "freq": freq,
            "da_pct": da_pct,
            "er": er,
            "shift_m": d,
            "shift_warn": warn_shift,
            "grid_id": safe_str(row.get(grid_col, "")),
        }
        st.session_state.batch_new_samples.append(item)

        if warn_shift:
            st.warning(f"⚠️ 座標與主檔差異約 {d:.2f} m（> 3m）。請確認是否需要更新主檔座標或是否跨網格。")
        st.success(f"已加入暫存：{lot_no} / {year} / 判定={metal_result} / 頻率={freq}")

    st.divider()
    st.subheader("🧾 本次暫存清單（確認後再寫入 DB）")

    if len(st.session_state.batch_new_samples) == 0:
        st.info("目前暫存清單是空的。")
    else:
        batch_df = pd.DataFrame(st.session_state.batch_new_samples)
        show_cols = ["lot_no","year","metal_result","freq","admin_status","shift_m","grid_id","used_total"]
        st.dataframe(batch_df[show_cols], use_container_width=True, height=360)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🗑️ 清空暫存清單"):
                st.session_state.batch_new_samples = []
                st.rerun()

        with c2:
            if st.button("✅ 確認上傳（寫入 DB）"):
                # 寫入 samples + 回寫 lands（最後調查年分與狀態摘要）
                for item in st.session_state.batch_new_samples:
                    db_insert_sample({
                        "lot_no": item["lot_no"],
                        "year": item["year"],
                        "sample_date": item["sample_date"],
                        "coord_x": item["coord_x"],
                        "coord_y": item["coord_y"],
                        "photo_site": item["photo_site"],
                        "photo_sample": item["photo_sample"],
                        "xrf": item["xrf"],
                        "total": item["total"],
                        "used_total": item["used_total"],
                        "admin_status": item["admin_status"],
                        "metal_result": item["metal_result"],
                        "freq": item["freq"],
                        "da_pct": item["da_pct"],
                        "er": item["er"],
                    })

                    # 回寫 lands 的摘要：last_year / admin_status / metal_result / freq
                    # 如果這次年分比 lands 目前 last_year 新，就更新；或 last_year 空就更新
                    # 先讀 DB 目前值
                    with engine.begin() as conn:
                        cur = conn.execute(text("SELECT last_year FROM lands WHERE lot_no=:lot_no"), {"lot_no": item["lot_no"]}).fetchone()
                        cur_last = cur[0] if cur else None

                    should_update = (cur_last is None) or (item["year"] >= int(cur_last))
                    if should_update:
                        # 若 admin_status 是建物/難以採樣，freq/metal_result 仍保留，但你也可後續改成「不納入計算」
                        land_update = {
                            "lot_no": item["lot_no"],
                            "sgm_id": None,
                            "land_no": None,
                            "grid_id": item.get("grid_id"),
                            "township": None,
                            "survey_type": None,
                            "rep": None,
                            "irrigation_type": None,
                            "x": None,
                            "y": None,
                            "admin_status": item["admin_status"],
                            "metal_result": item["metal_result"],
                            "freq": item["freq"],
                            "last_year": item["year"],
                        }
                        # 先補上既有資料避免被覆蓋為 None
                        with engine.begin() as conn:
                            cur2 = conn.execute(text("""
                                SELECT sgm_id, land_no, grid_id, township, survey_type, rep, irrigation_type, x, y
                                FROM lands WHERE lot_no=:lot_no
                            """), {"lot_no": item["lot_no"]}).fetchone()
                        if cur2:
                            land_update["sgm_id"] = cur2[0]
                            land_update["land_no"] = cur2[1]
                            land_update["grid_id"] = land_update["grid_id"] or cur2[2]
                            land_update["township"] = cur2[3]
                            land_update["survey_type"] = cur2[4]
                            land_update["rep"] = cur2[5]
                            land_update["irrigation_type"] = cur2[6]
                            land_update["x"] = cur2[7]
                            land_update["y"] = cur2[8]

                        db_upsert_land(land_update)

                st.session_state.batch_new_samples = []
                refresh_data()
                st.success("✅ 已寫入 DB 並更新主檔摘要！")
                st.rerun()

# =========================================================
# 17) 頁面：空間地圖檢視（簡化可用版：用 DB/Excel 的座標上點）
# =========================================================
elif menu == "空間地圖檢視":
    st.title("🗺️ 空間地圖檢視（可用版）")
    st.caption("後續你要的『網格上色、點擊資訊卡、符號形狀』都可以做；這裡先讓你確認 DB/資料流與點位可視化正常。")

    # 欄位兼容
    if "lot_no" in df_master.columns:
        lot_col = "lot_no"
        x_col = "x"
        y_col = "y"
        surv_col = "survey_type"
        mr_col = "metal_result"
        admin_col = "admin_status"
    else:
        lot_col = "地段地號"
        x_col = "TWD97_X"
        y_col = "TWD97_Y"
        surv_col = "調查方式"
        mr_col = "目前農地調查現況"
        admin_col = "農地監測狀態"

    # 篩選
    q = st.text_input("快速搜尋地段地號（可空）", "")
    data = df_master.copy()
    if q.strip():
        data = data[data[lot_col].astype(str).str.contains(q.strip(), na=False)]

    # 轉成 WGS84
    pts = []
    for _, r in data.iterrows():
        x = to_float(r.get(x_col))
        y = to_float(r.get(y_col))
        if x is None or y is None:
            continue
        lon, lat = transformer_to_wgs84.transform(x, y)
        pts.append({
            "lat": lat,
            "lon": lon,
            "lot_no": safe_str(r.get(lot_col)),
            "survey": safe_str(r.get(surv_col)),
            "admin": safe_str(r.get(admin_col)),
            "metal": safe_str(r.get(mr_col)),
        })

    if len(pts) == 0:
        st.warning("沒有可用座標點位。")
        st.stop()

    # 地圖中心
    center_lat = float(np.mean([p["lat"] for p in pts]))
    center_lon = float(np.mean([p["lon"] for p in pts]))
    m = folium.Map(location=[center_lat, center_lon], zoom_start=11, tiles="OpenStreetMap")

    def color_for(p):
        # 先用 admin/metal 做簡單上色（你之後有指定形狀與顏色，我們下一步再做）
        if "管制" in p["admin"] or p["metal"] == "管制":
            return "red"
        if "建物" in p["admin"]:
            return "black"
        if "難以採樣" in p["admin"]:
            return "purple"
        if p["metal"] == "增量":
            return "red"
        if p["metal"] == "延長":
            return "blue"
        return "green"

    for p in pts:
        folium.CircleMarker(
            location=[p["lat"], p["lon"]],
            radius=5,
            color=color_for(p),
            fill=True,
            fill_opacity=0.8,
            popup=folium.Popup(
                f"""<b>{p['lot_no']}</b><br/>
                調查方式：{p['survey']}<br/>
                行政狀態：{p['admin']}<br/>
                重金屬判定：{p['metal']}<br/>""",
                max_width=320
            )
        ).add_to(m)

    st_folium(m, use_container_width=True, height=650)


