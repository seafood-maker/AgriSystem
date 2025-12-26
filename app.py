import os
import json
import math
from datetime import datetime, timezone, timedelta
from typing import Dict, Any, List, Optional, Tuple

import pandas as pd
import streamlit as st

from sqlalchemy import create_engine, text
from streamlit_folium import st_folium
import folium


# =========================================================
# 基本設定
# =========================================================
st.set_page_config(page_title="彰化縣農地監測戰情室", layout="wide")
TZ_TW = timezone(timedelta(hours=8))

EXCEL_PATH = "彰化農地管理資料庫.xlsx"

SHEET_MASTER = "農地現況主檔"
SHEET_RECORDS = "歷年調查紀錄"
SHEET_BLOCKS = "同坵塊對照表"
SHEET_STANDARDS = "判定標準表"

METALS = ["汞", "砷", "銅", "鉻", "鎘", "鉛", "鋅", "鎳"]  # 8項（你要銅也納入）
METAL_KEYS = ["Hg", "As", "Cu", "Cr", "Cd", "Pb", "Zn", "Ni"]  # DB JSON keys 對應


# =========================================================
# 小工具
# =========================================================
def now_tw() -> datetime:
    return datetime.now(TZ_TW)


def roc_date_str(dt: Optional[datetime] = None) -> str:
    dt = dt or now_tw()
    roc_year = dt.year - 1911
    return f"民國 {roc_year} 年 {dt.month} 月 {dt.day} 日"


def nstr(x) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    return str(x).strip()


def safe_float(x) -> Optional[float]:
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return None
        s = str(x).strip()
        if s == "":
            return None
        return float(s)
    except Exception:
        return None


def coord_distance_m(x1, y1, x2, y2) -> Optional[float]:
    if x1 is None or y1 is None or x2 is None or y2 is None:
        return None
    try:
        return float(math.hypot(x1 - x2, y1 - y2))
    except Exception:
        return None


def to_bool_rep(x) -> bool:
    s = nstr(x)
    return s in ["是", "代表", "TRUE", "True", "1", "Y", "y", "YES", "Yes"]


def metal_name_to_key(name_zh: str) -> str:
    mapping = {
        "汞": "Hg", "砷": "As", "銅": "Cu", "鉻": "Cr",
        "鎘": "Cd", "鉛": "Pb", "鋅": "Zn", "鎳": "Ni"
    }
    return mapping.get(name_zh, name_zh)


def key_to_metal_name(k: str) -> str:
    mapping = {"Hg": "汞", "As": "砷", "Cu": "銅", "Cr": "鉻", "Cd": "鎘", "Pb": "鉛", "Zn": "鋅", "Ni": "鎳"}
    return mapping.get(k, k)


# =========================================================
# DB 連線
# =========================================================
def get_database_url() -> str:
    if "DATABASE_URL" in st.secrets:
        return st.secrets["DATABASE_URL"]
    if os.getenv("DATABASE_URL"):
        return os.getenv("DATABASE_URL")
    return ""


@st.cache_resource
def get_engine():
    db_url = get_database_url()
    if not db_url:
        return None
    return create_engine(db_url, pool_pre_ping=True)


# =========================================================
# 權限 / 登入
# =========================================================
ROLE_READER = "reader"
ROLE_EDITOR = "editor"
ROLE_ADMIN = "admin"

ROLE_LABELS = {
    ROLE_READER: "讀者",
    ROLE_EDITOR: "編輯者",
    ROLE_ADMIN: "管理者",
}


def get_role() -> str:
    return st.session_state.get("role", ROLE_READER)


def has_editor() -> bool:
    return get_role() in [ROLE_EDITOR, ROLE_ADMIN]


def has_admin() -> bool:
    return get_role() == ROLE_ADMIN


def login_sidebar():
    st.sidebar.markdown("### 🔐 權限登入")
    role = get_role()

    if st.session_state.get("authed", False):
        st.sidebar.success(f"已登入：{ROLE_LABELS.get(role, role)}")
        if st.sidebar.button("登出"):
            st.session_state["authed"] = False
            st.session_state["role"] = ROLE_READER
            st.rerun()
        st.sidebar.divider()
        return

    st.sidebar.caption("讀者可不登入；需要匯入/新增年度結果請登入。")
    target = st.sidebar.selectbox("選擇角色", [ROLE_READER, ROLE_EDITOR, ROLE_ADMIN],
                                  format_func=lambda x: ROLE_LABELS.get(x, x))

    pwd = st.sidebar.text_input("密碼", type="password")
    if st.sidebar.button("登入"):
        if target == ROLE_READER:
            # 讀者允許免密碼（你也可以改成必須密碼）
            st.session_state["authed"] = True
            st.session_state["role"] = ROLE_READER
            st.rerun()

        elif target == ROLE_EDITOR:
            ep = st.secrets.get("EDITOR_PASSWORD", "")
            if not ep:
                st.sidebar.error("未設定 EDITOR_PASSWORD（Secrets）")
                return
            if pwd == ep:
                st.session_state["authed"] = True
                st.session_state["role"] = ROLE_EDITOR
                st.rerun()
            else:
                st.sidebar.error("密碼錯誤")

        elif target == ROLE_ADMIN:
            ap = st.secrets.get("ADMIN_PASSWORD", "")
            if not ap:
                st.sidebar.error("未設定 ADMIN_PASSWORD（Secrets）")
                return
            if pwd == ap:
                st.session_state["authed"] = True
                st.session_state["role"] = ROLE_ADMIN
                st.rerun()
            else:
                st.sidebar.error("密碼錯誤")

    st.sidebar.divider()


# =========================================================
# DB Schema 初始化（不使用 DO $$）
# =========================================================
def init_db(engine):
    if engine is None:
        return

    with engine.begin() as conn:
        # ---------- standards ----------
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS standards (
            item TEXT PRIMARY KEY,
            monitor_std DOUBLE PRECISION,
            control_std DOUBLE PRECISION,
            da_threshold DOUBLE PRECISION
        );
        """))

        # ---------- blocks ----------
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS blocks (
            block_id TEXT NOT NULL,
            地段地號 TEXT NOT NULL,
            is_rep BOOLEAN DEFAULT FALSE,
            PRIMARY KEY (block_id, 地段地號)
        );
        """))

        # ---------- lands (base create) ----------
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS lands (
            地段地號 TEXT PRIMARY KEY,
            sSGM編號 TEXT,
            land_serial TEXT,
            網格編號 TEXT,
            鄉鎮市 TEXT,
            調查方式 TEXT,
            代表性 TEXT,
            用水種類 TEXT,

            coord_x DOUBLE PRECISION,
            coord_y DOUBLE PRECISION,

            initial_metals JSONB,
            農地監測狀態 TEXT,
            農地監測狀態 TEXT,
            網格監測頻率 TEXT,
            最後調查年分 INTEGER,
            year_status JSONB,

            updated_at TIMESTAMP DEFAULT NOW()
        );
        """))

        # ✅ 補欄位：避免你 DB 是舊表缺欄位造成 ProgrammingError
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS sSGM編號 TEXT;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS land_serial TEXT;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS 網格編號 TEXT;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS 鄉鎮市 TEXT;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS 調查方式 TEXT;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS 代表性 TEXT;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS 用水種類 TEXT;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS coord_x DOUBLE PRECISION;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS coord_y DOUBLE PRECISION;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS initial_metals JSONB;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS 農地監測狀態 TEXT;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS 農地監測狀態 TEXT;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS 網格監測頻率 TEXT;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS 最後調查年分 INTEGER;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS year_status JSONB;"))
        conn.execute(text("ALTER TABLE lands ADD COLUMN IF NOT EXISTS updated_at TIMESTAMP DEFAULT NOW();"))

        # ---------- samples ----------
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS samples (
            sample_id SERIAL PRIMARY KEY,
            地段地號 TEXT NOT NULL,
            year INTEGER NOT NULL,
            sample_date DATE,

            coord_x DOUBLE PRECISION,
            coord_y DOUBLE PRECISION,

            photo_site TEXT,
            photo_sample TEXT,

            xrf JSONB,
            total JSONB,
            used_total BOOLEAN DEFAULT FALSE,

            農地監測狀態 TEXT,
            metal_result TEXT,
            網格監測頻率 TEXT,

            da_pct JSONB,
            er JSONB,

            created_at TIMESTAMP DEFAULT NOW()
        );
        """))

        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS sample_date DATE;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS coord_x DOUBLE PRECISION;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS coord_y DOUBLE PRECISION;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS photo_site TEXT;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS photo_sample TEXT;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS xrf JSONB;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS total JSONB;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS used_total BOOLEAN DEFAULT FALSE;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS 農地監測狀態 TEXT;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS metal_result TEXT;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS 網格監測頻率 TEXT;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS da_pct JSONB;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS er JSONB;"))
        conn.execute(text("ALTER TABLE samples ADD COLUMN IF NOT EXISTS created_at TIMESTAMP DEFAULT NOW();"))

        # ---------- indexes ----------
        conn.execute(text("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_samples_lot_year_unique
        ON samples(地段地號, year);
        """))

        conn.execute(text("""
        CREATE INDEX IF NOT EXISTS idx_lands_網格編號
        ON lands(網格編號);
        """))

        conn.execute(text("""
        CREATE INDEX IF NOT EXISTS idx_lands_農地監測狀態
        ON lands(農地監測狀態);
        """))


# =========================================================
# Excel 讀取（cache）
# =========================================================
@st.cache_data(show_spinner=False)
def load_excel_all(path: str) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(path)
    data = {}
    for s in xls.sheet_names:
        data[s] = xls.parse(s)
    return data


# =========================================================
# 規則引擎（集中管理）
# - ER：濃度 / 監測標準
# - DA%：((本次 - baseline) / baseline) * 100
# - metal_result：管制 / 增量 / 延長 / 正常
# - 網格監測頻率：管制 / 持續 / 延長 / 退場
# =========================================================
class RulesEngine:
    def __init__(self, standards: Dict[str, Dict[str, float]]):
        """
        standards[item] = {"monitor":..., "control":..., "da":...}
        item 以中文（汞/砷/銅...）為主
        """
        self.standards = standards

    def compute_er(self, metals: Dict[str, float]) -> Dict[str, Optional[float]]:
        er = {}
        for name, val in metals.items():
            std = self.standards.get(name, {}).get("monitor")
            if std and val is not None:
                er[name] = val / std
            else:
                er[name] = None
        return er

    def compute_da_pct(self, metals: Dict[str, float], baseline: Dict[str, float]) -> Dict[str, Optional[float]]:
        da = {}
        for name, val in metals.items():
            b = baseline.get(name)
            if b is None or b == 0 or val is None:
                da[name] = None
            else:
                da[name] = ((val - b) / b) * 100.0
        return da

    def decide_metal_result(self, metals: Dict[str, float], er: Dict[str, Optional[float]], da: Dict[str, Optional[float]]) -> str:
        # 1) 管制：任何項目 >= 管制標準
        for name, val in metals.items():
            cstd = self.standards.get(name, {}).get("control")
            if cstd is not None and val is not None and val >= cstd:
                return "管制"

        # 2) 若任一項 ER>1 才需要判定增量/延長，否則正常
        any_exceed = any((v is not None and v > 1.0) for v in er.values())
        if not any_exceed:
            return "正常"

        # 3) 增量：任一項 DA% >= DA門檻（且該項 ER>1 更合理）
        for name, da_v in da.items():
            th = self.standards.get(name, {}).get("da")
            er_v = er.get(name)
            if th is not None and da_v is not None and da_v >= th and (er_v is not None and er_v > 1.0):
                return "增量"

        # 4) 其餘 ER>1 -> 延長
        return "延長"

    def metal_result_to_網格監測頻率(self, metal_result: str) -> str:
        mapping = {
            "管制": "管制",
            "增量": "持續",
            "延長": "延長",
            "正常": "退場",
        }
        return mapping.get(metal_result, "")


# =========================================================
# DB 讀取：標準表、主檔統計、查詢
# =========================================================
def db_fetch_standards(engine) -> Dict[str, Dict[str, float]]:
    if engine is None:
        return {}

    with engine.begin() as conn:
        df = pd.read_sql(text("SELECT * FROM standards;"), conn)

    standards = {}
    for _, r in df.iterrows():
        item = nstr(r["item"])
        standards[item] = {
            "monitor": r["monitor_std"] if not pd.isna(r["monitor_std"]) else None,
            "control": r["control_std"] if not pd.isna(r["control_std"]) else None,
            "da": r["da_threshold"] if not pd.isna(r["da_threshold"]) else None,
        }
    return standards


def db_fetch_kpis(engine) -> Dict[str, int]:
    if engine is None:
        return dict(total=0, sample_points=0, control=0, building=0, hard=0, normal_exit=0)

    with engine.begin() as conn:
        total = conn.execute(text("SELECT COUNT(*) FROM lands;")).scalar() or 0
        sample_points = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE COALESCE(NULLIF(TRIM(代表性), ''), '') <> '';
        """)).scalar() or 0
        control = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE 農地監測狀態='管制' OR 農地監測狀態='管制';
        """)).scalar() or 0
        building = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE 農地監測狀態='建物' OR 農地監測狀態='建物';
        """)).scalar() or 0
        hard = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE 農地監測狀態 IN ('難以採樣','無法採樣') OR 農地監測狀態 IN ('難以採樣','無法採樣');
        """)).scalar() or 0
        normal_exit = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE 農地監測狀態='正常' OR 農地監測狀態='正常';
        """)).scalar() or 0

    return dict(
        total=int(total),
        sample_points=int(sample_points),
        control=int(control),
        building=int(building),
        hard=int(hard),
        normal_exit=int(normal_exit),
    )


def db_search_land(engine, query: str) -> pd.DataFrame:
    if engine is None:
        return pd.DataFrame()
    q = query.strip()
    if not q:
        return pd.DataFrame()
    with engine.begin() as conn:
        df = pd.read_sql(text("""
            SELECT *
            FROM lands
            WHERE 地段地號 ILIKE :q OR sSGM編號 ILIKE :q
            ORDER BY 網格編號, 地段地號
            LIMIT 200;
        """), conn, params={"q": f"%{q}%"})
    return df


def db_fetch_land_by_lot(engine, 地段地號: str) -> Optional[Dict[str, Any]]:
    if engine is None:
        return None
    with engine.begin() as conn:
        row = conn.execute(text("SELECT * FROM lands WHERE 地段地號=:lot LIMIT 1;"), {"lot": 地段地號}).mappings().first()
    return dict(row) if row else None


def db_fetch_last_sample(engine, 地段地號: str) -> Optional[Dict[str, Any]]:
    if engine is None:
        return None
    with engine.begin() as conn:
        row = conn.execute(text("""
            SELECT *
            FROM samples
            WHERE 地段地號=:lot
            ORDER BY year DESC, created_at DESC
            LIMIT 1;
        """), {"lot": 地段地號}).mappings().first()
    return dict(row) if row else None


def db_fetch_first_sample(engine, 地段地號: str) -> Optional[Dict[str, Any]]:
    if engine is None:
        return None
    with engine.begin() as conn:
        row = conn.execute(text("""
            SELECT *
            FROM samples
            WHERE 地段地號=:lot
            ORDER BY year ASC, created_at ASC
            LIMIT 1;
        """), {"lot": 地段地號}).mappings().first()
    return dict(row) if row else None


def db_fetch_samples(engine, 地段地號: str) -> pd.DataFrame:
    if engine is None:
        return pd.DataFrame()
    with engine.begin() as conn:
        df = pd.read_sql(text("""
            SELECT year, sample_date, 農地監測狀態, metal_result, 網格監測頻率, xrf, total, used_total, da_pct, er, created_at
            FROM samples
            WHERE 地段地號=:lot
            ORDER BY year DESC, created_at DESC;
        """), conn, params={"lot": 地段地號})
    return df


# =========================================================
# Excel -> DB 匯入（管理者）
# - lands / standards / blocks
# - blocks 使用 ON CONFLICT 避免 IntegrityError
# =========================================================
def admin_import_excel_to_db(engine):
    st.subheader("🛠 管理員工具：Excel → DB 匯入")
    st.caption("此工具會把 repo 裡的 Excel 主檔、標準表、同坵塊表匯入 DB。")

    if not has_admin():
        st.info("需要管理者權限")
        return

    if engine is None:
        st.error("DATABASE_URL 未設定")
        return

    if not os.path.exists(EXCEL_PATH):
        st.error(f"找不到 Excel：{EXCEL_PATH}（請確認放在 repo 根目錄）")
        return

    if st.button("🚀 一鍵匯入（lands / standards / blocks）"):
        with st.spinner("匯入中..."):
            data = load_excel_all(EXCEL_PATH)

            # -------- standards --------
            if SHEET_STANDARDS not in data:
                st.error(f"Excel 缺少分頁：{SHEET_STANDARDS}")
                return
            df_std = data[SHEET_STANDARDS].copy()
            df_std.columns = [nstr(c) for c in df_std.columns]

            col_item = "項目名稱"
            col_m = "監測標準"
            col_c = "管制標準"
            col_da = "上升標準 (DA門檻)"

            with engine.begin() as conn:
                conn.execute(text("DELETE FROM standards;"))
                for _, r in df_std.iterrows():
                    item = nstr(r.get(col_item))
                    if not item:
                        continue
                    conn.execute(text("""
                        INSERT INTO standards (item, monitor_std, control_std, da_threshold)
                        VALUES (:i, :m, :c, :d)
                        ON CONFLICT (item) DO UPDATE
                        SET monitor_std=EXCLUDED.monitor_std,
                            control_std=EXCLUDED.control_std,
                            da_threshold=EXCLUDED.da_threshold;
                    """), {
                        "i": item,
                        "m": safe_float(r.get(col_m)),
                        "c": safe_float(r.get(col_c)),
                        "d": safe_float(r.get(col_da)),
                    })

            # -------- blocks --------
            if SHEET_BLOCKS not in data:
                st.error(f"Excel 缺少分頁：{SHEET_BLOCKS}")
                return
            df_blk = data[SHEET_BLOCKS].copy()
            df_blk.columns = [nstr(c) for c in df_blk.columns]
            gcol = "農地群組編號"
            lcol = "農地地段地號"
            rcol = "代表農地"

            if gcol not in df_blk.columns or lcol not in df_blk.columns:
                st.error("同坵塊對照表欄位不齊（需要：農地群組編號、農地地段地號）")
                return

            df_blk[gcol] = df_blk[gcol].astype(str).str.strip()
            df_blk[lcol] = df_blk[lcol].astype(str).str.strip()
            df_blk = df_blk[(df_blk[gcol] != "") & (df_blk[lcol] != "")]
            df_blk2 = df_blk.drop_duplicates(subset=[gcol, lcol], keep="last")

            with engine.begin() as conn:
                conn.execute(text("DELETE FROM blocks;"))
                for _, r in df_blk2.iterrows():
                    block_id = nstr(r.get(gcol))
                    地段地號 = nstr(r.get(lcol))
                    is_rep = to_bool_rep(r.get(rcol))
                    conn.execute(text("""
                        INSERT INTO blocks (block_id, 地段地號, is_rep)
                        VALUES (:b, :l, :r)
                        ON CONFLICT (block_id, 地段地號) DO UPDATE
                        SET is_rep = EXCLUDED.is_rep;
                    """), {"b": block_id, "l": 地段地號, "r": is_rep})

            # -------- lands --------
            if SHEET_MASTER not in data:
                st.error(f"Excel 缺少分頁：{SHEET_MASTER}")
                return
            df = data[SHEET_MASTER].copy()
            df.columns = [nstr(c) for c in df.columns]

            # 必要欄位
            must_cols = ["地段地號", "網格編號", "鄉鎮市", "用水種類", "調查方式", "代表性", "TWD97_X", "TWD97_Y", "目前農地調查現況"]
            missing = [c for c in must_cols if c not in df.columns]
            if missing:
                st.error(f"主檔缺少欄位：{missing}")
                return

            # 年度狀態欄：xxx狀態
            year_cols = [c for c in df.columns if c.endswith("狀態") and nstr(c).replace("狀態", "").isdigit()]

            # 初始八項
            init_map = {
                "汞": "初始_汞",
                "砷": "初始_砷",
                "銅": "初始_銅",
                "鉻": "初始_鉻",
                "鎘": "初始_鎘",
                "鉛": "初始_鉛",
                "鋅": "初始_鋅",
                "鎳": "初始_鎳",
            }

            df["地段地號"] = df["地段地號"].astype(str).str.strip()
            df = df[df["地段地號"] != ""].drop_duplicates(subset=["地段地號"], keep="last")

            # 若 Excel 還沒新增「農地監測狀態」欄，這裡自動推一個
            if "農地監測狀態" not in df.columns:
                def infer_農地監測狀態(cur):
                    cur = nstr(cur)
                    if cur in ["管制", "建物", "難以採樣", "無法採樣", "正常"]:
                        return "難以採樣" if cur == "無法採樣" else cur
                    if cur in ["增量", "延長"]:
                        return "監測"
                    return ""
                df["農地監測狀態"] = df["目前農地調查現況"].apply(infer_農地監測狀態)

            # 最後調查年分若空，用 year_cols 最新非空推
            if "最後調查年分" not in df.columns:
                df["最後調查年分"] = None
            if "網格監測頻率" not in df.columns:
                df["網格監測頻率"] = ""

            def infer_最後調查年分(row) -> Optional[int]:
                if pd.notna(row.get("最後調查年分")) and str(row.get("最後調查年分")).strip() != "":
                    try:
                        return int(float(row.get("最後調查年分")))
                    except Exception:
                        pass
                latest = None
                for yc in year_cols:
                    val = nstr(row.get(yc))
                    if val:
                        y = int(nstr(yc).replace("狀態", ""))
                        if latest is None or y > latest:
                            latest = y
                return latest

            df["最後調查年分"] = df.apply(infer_最後調查年分, axis=1)

            with engine.begin() as conn:
                for _, r in df.iterrows():
                    地段地號 = nstr(r.get("地段地號"))
                    if not 地段地號:
                        continue

                    initial_metals = {}
                    for mzh, col in init_map.items():
                        if col in df.columns:
                            v = safe_float(r.get(col))
                            if v is not None:
                                initial_metals[metal_name_to_key(mzh)] = v

                    year_status = {}
                    for yc in year_cols:
                        y = nstr(yc).replace("狀態", "")
                        v = nstr(r.get(yc))
                        if v:
                            year_status[y] = v

                    payload = {
                        "地段地號": 地段地號,
                        "sSGM編號": nstr(r.get("SGM編號")),
                        "land_serial": nstr(r.get("農地序號")),
                        "網格編號": nstr(r.get("網格編號")),
                        "鄉鎮市": nstr(r.get("鄉鎮市")),
                        "調查方式": nstr(r.get("調查方式")),
                        "代表性": nstr(r.get("代表性")),
                        "用水種類": nstr(r.get("用水種類")),
                        "coord_x": safe_float(r.get("TWD97_X")),
                        "coord_y": safe_float(r.get("TWD97_Y")),
                        "initial_metals": json.dumps(initial_metals, ensure_ascii=False),
                        "農地監測狀態": nstr(r.get("目前農地調查現況")),
                        "農地監測狀態": nstr(r.get("農地監測狀態")),
                        "網格監測頻率": nstr(r.get("網格監測頻率")),
                        "最後調查年分": safe_float(r.get("最後調查年分")),
                        "year_status": json.dumps(year_status, ensure_ascii=False),
                    }

                    conn.execute(text("""
                        INSERT INTO lands (
                            地段地號, sSGM編號, land_serial, 網格編號, 鄉鎮市, 調查方式,
                            代表性, 用水種類, coord_x, coord_y,
                            initial_metals, 農地監測狀態, 農地監測狀態, 網格監測頻率, 最後調查年分, year_status, updated_at
                        )
                        VALUES (
                            :地段地號, :sSGM編號, :land_serial, :網格編號, :鄉鎮市, :調查方式,
                            :代表性, :用水種類, :coord_x, :coord_y,
                            CAST(:initial_metals AS JSONB), :農地監測狀態, :農地監測狀態, :網格監測頻率,
                            CASE WHEN :最後調查年分 IS NULL THEN NULL ELSE CAST(:最後調查年分 AS INTEGER) END,
                            CAST(:year_status AS JSONB),
                            NOW()
                        )
                        ON CONFLICT (地段地號) DO UPDATE SET
                            sSGM編號=EXCLUDED.sSGM編號,
                            land_serial=EXCLUDED.land_serial,
                            網格編號=EXCLUDED.網格編號,
                            鄉鎮市=EXCLUDED.鄉鎮市,
                            調查方式=EXCLUDED.調查方式,
                            代表性=EXCLUDED.代表性,
                            用水種類=EXCLUDED.用水種類,
                            coord_x=EXCLUDED.coord_x,
                            coord_y=EXCLUDED.coord_y,
                            initial_metals=EXCLUDED.initial_metals,
                            農地監測狀態=EXCLUDED.農地監測狀態,
                            農地監測狀態=EXCLUDED.農地監測狀態,
                            網格監測頻率=EXCLUDED.網格監測頻率,
                            最後調查年分=EXCLUDED.最後調查年分,
                            year_status=EXCLUDED.year_status,
                            updated_at=NOW();
                    """), payload)

        st.success("✅ 匯入完成！請回到首頁確認 KPI 不再是 0。")
        st.cache_data.clear()


# =========================================================
# blocks 管理（可新增/刪除/指定代表點）
# =========================================================
def page_blocks_manage(engine):
    st.subheader("🧩 同坵塊（BLOCK）管理")

    if not has_admin():
        st.info("需要管理者權限")
        return
    if engine is None:
        st.error("DATABASE_URL 未設定")
        return

    with engine.begin() as conn:
        df = pd.read_sql(text("""
            SELECT block_id, 地段地號, is_rep
            FROM blocks
            ORDER BY block_id, 地段地號;
        """), conn)

    st.dataframe(df, use_container_width=True, height=380)

    st.markdown("### ➕ 新增/刪除")
    c1, c2, c3 = st.columns(3)
    with c1:
        new_block = st.text_input("block_id（農地群組編號）")
    with c2:
        new_lot = st.text_input("地段地號（地段地號）")
    with c3:
        new_rep = st.selectbox("是否代表點", ["否", "是"], index=0)

    if st.button("新增到 blocks"):
        if not new_block.strip() or not new_lot.strip():
            st.error("block_id 與 地段地號 不能空白")
        else:
            with engine.begin() as conn:
                conn.execute(text("""
                    INSERT INTO blocks (block_id, 地段地號, is_rep)
                    VALUES (:b,:l,:r)
                    ON CONFLICT (block_id, 地段地號) DO UPDATE
                    SET is_rep=EXCLUDED.is_rep;
                """), {"b": new_block.strip(), "l": new_lot.strip(), "r": (new_rep == "是")})
            st.success("已新增/更新")
            st.rerun()

    st.markdown("### ⭐ 指定某 block 的代表點（只保留一個代表）")
    block_ids = sorted(df["block_id"].unique().tolist()) if not df.empty else []
    bid = st.selectbox("選擇 block_id", [""] + block_ids)
    if bid:
        sub = df[df["block_id"] == bid].copy()
        lots = sub["地段地號"].tolist()
        rep_lot = st.selectbox("指定代表 地段地號", lots)
        if st.button("設定代表點"):
            with engine.begin() as conn:
                conn.execute(text("UPDATE blocks SET is_rep=FALSE WHERE block_id=:b;"), {"b": bid})
                conn.execute(text("""
                    UPDATE blocks SET is_rep=TRUE
                    WHERE block_id=:b AND 地段地號=:l;
                """), {"b": bid, "l": rep_lot})
            st.success("代表點已更新")
            st.rerun()

    st.markdown("### 🗑️ 刪除某筆 (block_id, 地段地號)")
    del_block = st.text_input("要刪除的 block_id")
    del_lot = st.text_input("要刪除的 地段地號")
    if st.button("刪除"):
        if not del_block.strip() or not del_lot.strip():
            st.error("block_id 與 地段地號 不能空白")
        else:
            with engine.begin() as conn:
                conn.execute(text("DELETE FROM blocks WHERE block_id=:b AND 地段地號=:l;"),
                             {"b": del_block.strip(), "l": del_lot.strip()})
            st.success("已刪除")
            st.rerun()


# =========================================================
# 新增年度調查結果（含暫存清單、座標偏差提醒、寫入DB、更新主檔）
# =========================================================
def page_add_annual(engine):
    st.subheader("🧾 新增年度調查結果（寫入 DB + 自動判定）")

    if not has_editor():
        st.info("需要編輯者以上權限（編輯者/管理者）")
        return
    if engine is None:
        st.error("DATABASE_URL 未設定")
        return

    standards = db_fetch_standards(engine)
    if not standards:
        st.warning("目前 standards 為空，請先由管理者 Excel→DB 匯入標準表。")
        return
    rules = RulesEngine(standards)

    # 暫存清單（本次批次新增）
    if "pending_samples" not in st.session_state:
        st.session_state["pending_samples"] = []

    st.markdown("### 1) 搜尋農地")
    q = st.text_input("輸入地段地號（建議）或 SGM", placeholder="例：華南段0159-0000")
    df = db_search_land(engine, q) if q else pd.DataFrame()
    if q and df.empty:
        st.warning("查無資料")
        return

    lot_selected = None
    if not df.empty:
        options = df["地段地號"].tolist()
        lot_selected = st.selectbox("選擇一筆農地（地段地號）", options)

    if not lot_selected:
        st.info("請先搜尋並選擇一筆農地")
        return

    land = db_fetch_land_by_lot(engine, lot_selected)
    if not land:
        st.error("找不到該農地")
        return

    st.markdown("### 2) 填寫本次調查資料")
    c1, c2, c3 = st.columns(3)
    with c1:
        year = st.number_input("調查年度（民國）", min_value=90, max_value=200, value=(now_tw().year - 1911))
        sample_date = st.date_input("採樣日期", value=now_tw().date())
    with c2:
        農地監測狀態 = st.selectbox("本次現勘狀態", ["監測", "正常", "管制", "建物", "難以採樣"], index=0)
    with c3:
        coord_x = st.number_input("採樣點 X（TWD97）", value=float(land.get("coord_x") or 0.0), format="%.3f")
        coord_y = st.number_input("採樣點 Y（TWD97）", value=float(land.get("coord_y") or 0.0), format="%.3f")

    # 照片（先以文字欄位存檔名/URL，雲端不存實體檔避免爆）
    st.caption("照片：Streamlit Cloud 不適合直接存檔案；建議你先填「檔名/雲端連結」。")
    photo_site = st.text_input("現勘照片（檔名或連結）", "")
    photo_sample = st.text_input("採樣照片（檔名或連結）", "")

    st.markdown("### 3) 填 XRF / 全量（八項）")
    st.caption("DA 計算：優先用全量（若有填任一項全量），否則用 XRF。")

    xrf_vals = {}
    total_vals = {}
    cols = st.columns(4)
    for i, mzh in enumerate(METALS):
        with cols[i % 4]:
            xrf_vals[mzh] = st.number_input(f"XRF_{mzh}", value=0.0, step=0.1, format="%.3f")
    cols2 = st.columns(4)
    for i, mzh in enumerate(METALS):
        with cols2[i % 4]:
            total_vals[mzh] = st.text_input(f"全量_{mzh}（可空白）", "")

    # 判定是否使用全量
    used_total = any(nstr(v) != "" for v in total_vals.values())
    chosen = {}
    for mzh in METALS:
        if used_total:
            v = safe_float(total_vals[mzh])
            chosen[mzh] = v if v is not None else float(xrf_vals[mzh])
        else:
            chosen[mzh] = float(xrf_vals[mzh])

    # baseline：採第一次納入定期監測時的值（這裡用「該 lot 第一筆 samples」作 baseline）
    baseline = {}
    first = db_fetch_first_sample(engine, lot_selected)
    if first and first.get("used_total") and first.get("total"):
        b = first["total"]
    elif first and first.get("xrf"):
        b = first["xrf"]
    else:
        b = None

    if isinstance(b, dict):
        for k, v in b.items():
            baseline[key_to_metal_name(k)] = v
    else:
        # 沒 baseline 的話，先用本次當 baseline（DA=0）
        baseline = {m: chosen[m] for m in METALS}

    # 計算 ER / DA
    er = rules.compute_er(chosen)
    da = rules.compute_da_pct(chosen, baseline)
    metal_result = rules.decide_metal_result(chosen, er, da)

    # 但若現勘是 建物/難以採樣/管制 → 直接把 lands/農地監測狀態 走現勘結果
    # 金屬結果 metal_result 仍存 samples（便於追溯），但主檔現況以現勘優先
    網格監測頻率 = rules.metal_result_to_網格監測頻率(metal_result)

    st.markdown("### 4) 系統自動判定結果（預覽）")
    cA, cB, cC = st.columns(3)
    cA.metric("金屬判定（metal_result）", metal_result)
    cB.metric("對應策略頻率（網格監測頻率）", 網格監測頻率)
    cC.metric("是否使用全量", "是" if used_total else "否")

    # 座標偏差提醒
    dist = coord_distance_m(land.get("coord_x"), land.get("coord_y"), coord_x, coord_y)
    if dist is not None and dist > 3.0:
        st.warning(f"⚠️ 座標偏差 {dist:.2f} 公尺（>3m）。建議確認是否要更新主檔座標。")

    # 加入暫存清單
    if st.button("➕ 加入本次暫存清單"):
        st.session_state["pending_samples"].append({
            "地段地號": lot_selected,
            "year": int(year),
            "sample_date": str(sample_date),
            "coord_x": float(coord_x),
            "coord_y": float(coord_y),
            "photo_site": photo_site.strip(),
            "photo_sample": photo_sample.strip(),
            "xrf": {metal_name_to_key(k): float(xrf_vals[k]) for k in METALS},
            "total": {metal_name_to_key(k): safe_float(total_vals[k]) for k in METALS} if used_total else {},
            "used_total": bool(used_total),
            "農地監測狀態": 農地監測狀態,
            "metal_result": metal_result,
            "網格監測頻率": 網格監測頻率,
            "da_pct": {metal_name_to_key(k): da.get(k) for k in METALS},
            "er": {metal_name_to_key(k): er.get(k) for k in METALS},
            "update_master_coord": (dist is not None and dist > 3.0),
        })
        st.success("已加入暫存清單（可在下方檢查後一次上傳）")

    # 顯示暫存清單
    st.markdown("### 5) 本次暫存清單（確認後一次寫入 DB）")
    pending = st.session_state.get("pending_samples", [])
    if pending:
        st.dataframe(pd.DataFrame(pending), use_container_width=True, height=240)

        colx, coly = st.columns(2)
        with colx:
            if st.button("🧹 清空暫存清單"):
                st.session_state["pending_samples"] = []
                st.rerun()

        with coly:
            if st.button("📤 上傳暫存清單到 DB（寫入 samples + 更新 lands）"):
                with engine.begin() as conn:
                    for rec in pending:
                        # 1) upsert samples（同 地段地號+year）
                        conn.execute(text("""
                            INSERT INTO samples (
                                地段地號, year, sample_date,
                                coord_x, coord_y,
                                photo_site, photo_sample,
                                xrf, total, used_total,
                                農地監測狀態, metal_result, 網格監測頻率,
                                da_pct, er
                            )
                            VALUES (
                                :地段地號, :year, :sample_date,
                                :coord_x, :coord_y,
                                :photo_site, :photo_sample,
                                CAST(:xrf AS JSONB), CAST(:total AS JSONB), :used_total,
                                :農地監測狀態, :metal_result, :網格監測頻率,
                                CAST(:da_pct AS JSONB), CAST(:er AS JSONB)
                            )
                            ON CONFLICT (地段地號, year) DO UPDATE SET
                                sample_date=EXCLUDED.sample_date,
                                coord_x=EXCLUDED.coord_x,
                                coord_y=EXCLUDED.coord_y,
                                photo_site=EXCLUDED.photo_site,
                                photo_sample=EXCLUDED.photo_sample,
                                xrf=EXCLUDED.xrf,
                                total=EXCLUDED.total,
                                used_total=EXCLUDED.used_total,
                                農地監測狀態=EXCLUDED.農地監測狀態,
                                metal_result=EXCLUDED.metal_result,
                                網格監測頻率=EXCLUDED.網格監測頻率,
                                da_pct=EXCLUDED.da_pct,
                                er=EXCLUDED.er,
                                created_at=NOW();
                        """), {
                            "地段地號": rec["地段地號"],
                            "year": rec["year"],
                            "sample_date": rec["sample_date"],
                            "coord_x": rec["coord_x"],
                            "coord_y": rec["coord_y"],
                            "photo_site": rec["photo_site"],
                            "photo_sample": rec["photo_sample"],
                            "xrf": json.dumps(rec["xrf"], ensure_ascii=False),
                            "total": json.dumps(rec["total"], ensure_ascii=False),
                            "used_total": rec["used_total"],
                            "農地監測狀態": rec["農地監測狀態"],
                            "metal_result": rec["metal_result"],
                            "網格監測頻率": rec["網格監測頻率"],
                            "da_pct": json.dumps(rec["da_pct"], ensure_ascii=False),
                            "er": json.dumps(rec["er"], ensure_ascii=False),
                        })

                        # 2) 更新 lands：年度狀態 + 最後年分 + 現況/頻率/座標（可選）
                        # 年度狀態欄：依你規則：若任一金屬 > 監測標準 -> "監測"；全正常 -> "正常"；
                        # 現勘若建物/難以採樣/管制 -> 對應填入
                        status_for_year = "監測" if rec["metal_result"] in ["增量", "延長"] else "正常"
                        if rec["農地監測狀態"] in ["建物", "難以採樣", "管制"]:
                            status_for_year = rec["農地監測狀態"]

                        # 取現有 year_status
                        land_row = conn.execute(text("SELECT year_status FROM lands WHERE 地段地號=:lot;"),
                                                {"lot": rec["地段地號"]}).mappings().first()
                        ys = land_row["year_status"] if land_row and land_row["year_status"] else {}
                        if isinstance(ys, str):
                            try:
                                ys = json.loads(ys)
                            except Exception:
                                ys = {}
                        if not isinstance(ys, dict):
                            ys = {}

                        ys[str(rec["year"])] = status_for_year

                        # 主檔狀態：以現勘優先
                        new_admin = rec["農地監測狀態"]  # 監測/正常/管制/建物/難以採樣
                        # 目前農地調查現況：建物/難以採樣/管制 => 同字；否則用 metal_result
                        if new_admin in ["管制", "建物", "難以採樣"]:
                            new_current = new_admin
                            new_網格監測頻率 = "管制" if new_admin == "管制" else ""  # 建物/難以採樣可留空或自訂
                        else:
                            new_current = rec["metal_result"]
                            new_網格監測頻率 = rec["網格監測頻率"]

                        update_coord = rec.get("update_master_coord", False)
                        if update_coord:
                            conn.execute(text("""
                                UPDATE lands
                                SET
                                    coord_x=:x,
                                    coord_y=:y,
                                    農地監測狀態=:農地監測狀態,
                                    農地監測狀態=:cur,
                                    網格監測頻率=:網格監測頻率,
                                    最後調查年分=:year,
                                    year_status=CAST(:ys AS JSONB),
                                    updated_at=NOW()
                                WHERE 地段地號=:lot;
                            """), {
                                "x": rec["coord_x"],
                                "y": rec["coord_y"],
                                "農地監測狀態": new_admin,
                                "cur": new_current,
                                "網格監測頻率": new_網格監測頻率,
                                "year": rec["year"],
                                "ys": json.dumps(ys, ensure_ascii=False),
                                "lot": rec["地段地號"],
                            })
                        else:
                            conn.execute(text("""
                                UPDATE lands
                                SET
                                    農地監測狀態=:農地監測狀態,
                                    農地監測狀態=:cur,
                                    網格監測頻率=:網格監測頻率,
                                    最後調查年分=:year,
                                    year_status=CAST(:ys AS JSONB),
                                    updated_at=NOW()
                                WHERE 地段地號=:lot;
                            """), {
                                "農地監測狀態": new_admin,
                                "cur": new_current,
                                "網格監測頻率": new_網格監測頻率,
                                "year": rec["year"],
                                "ys": json.dumps(ys, ensure_ascii=False),
                                "lot": rec["地段地號"],
                            })

                st.success("✅ 已寫入 DB 並更新主檔")
                st.session_state["pending_samples"] = []
                st.cache_data.clear()
                st.rerun()
    else:
        st.info("暫存清單目前是空的")

    st.divider()
    st.markdown("### 6) 該筆農地歷史紀錄（samples）")
    hist = db_fetch_samples(engine, lot_selected)
    if hist.empty:
        st.info("目前尚無 samples 記錄")
    else:
        st.dataframe(hist, use_container_width=True, height=260)


# =========================================================
# 歷史紀錄查詢
# =========================================================
def page_history(engine):
    st.subheader("🗂️ 歷史紀錄查詢（samples）")
    if engine is None:
        st.error("DATABASE_URL 未設定")
        return

    q = st.text_input("輸入地段地號或 SGM")
    df = db_search_land(engine, q) if q else pd.DataFrame()
    if q and df.empty:
        st.warning("查無資料")
        return
    lot_selected = None
    if not df.empty:
        lot_selected = st.selectbox("選擇 地段地號", df["地段地號"].tolist())

    if lot_selected:
        hist = db_fetch_samples(engine, lot_selected)
        if hist.empty:
            st.info("尚無歷史紀錄")
        else:
            st.dataframe(hist, use_container_width=True, height=420)


# =========================================================
# 總表清單
# =========================================================
def page_master_list(engine):
    st.subheader("📋 總表清單（lands）")
    if engine is None:
        st.error("DATABASE_URL 未設定")
        return

    with engine.begin() as conn:
        df = pd.read_sql(text("""
            SELECT
                地段地號 AS 地段地號,
                sSGM編號 AS SGM編號,
                網格編號 AS 網格編號,
                鄉鎮市 AS 鄉鎮市,
                調查方式 AS 調查方式,
                代表性 AS 代表性,
                用水種類 AS 用水種類,
                農地監測狀態 AS 農地監測狀態,
                農地監測狀態 AS 目前農地調查現況,
                網格監測頻率 AS 網格監測頻率,
                最後調查年分 AS 最後調查年分
            FROM lands
            ORDER BY 網格編號, 地段地號
            LIMIT 5000;
        """), conn)

    st.dataframe(df, use_container_width=True, height=520)


# =========================================================
# 資料查詢（SGM / 地段地號）
# =========================================================
def page_search(engine):
    st.subheader("🔎 資料查詢（SGM 或 地段地號）")
    if engine is None:
        st.error("DATABASE_URL 未設定")
        return

    q = st.text_input("請輸入 SGM 或 地段地號", placeholder="例：華南段0159-0000")
    if not q:
        st.info("輸入後即可查詢")
        return

    df = db_search_land(engine, q)
    if df.empty:
        st.warning("查無資料")
        return

    st.dataframe(df, use_container_width=True, height=320)

    row = df.iloc[0].to_dict()
    st.markdown("### 🧾 資訊卡（第一筆）")
    c1, c2, c3 = st.columns(3)
    c1.write({"地段地號": row.get("地段地號"), "SGM": row.get("sSGM編號"), "網格": row.get("網格編號")})
    c2.write({"現勘狀態": row.get("農地監測狀態"), "現況": row.get("農地監測狀態"), "頻率": row.get("網格監測頻率")})
    c3.write({"X": row.get("coord_x"), "Y": row.get("coord_y"), "最後年分": row.get("最後調查年分")})


# =========================================================
# 空間地圖（folium）
# - 顏色/形狀依你指定
# - 點擊顯示卡片資訊
# =========================================================
def pick_marker_style(row: Dict[str, Any]) -> Tuple[str, int, str]:
    """
    return (shape, sides, color)
    shape: "polygon" or "circle"
    """
    農地監測狀態 = nstr(row.get("農地監測狀態"))
    cur = nstr(row.get("農地監測狀態"))
    rep = nstr(row.get("代表性"))
    method = nstr(row.get("調查方式")) + " " + nstr(row.get("用水種類"))

    # 備用點：菱形 白色
    if "備用" in rep:
        return ("polygon", 4, "white")  # diamond 近似

    # 六角形：管制/建物/難以採樣
    if 農地監測狀態 == "管制" or cur == "管制":
        return ("polygon", 6, "red")
    if 農地監測狀態 == "建物" or cur == "建物":
        return ("polygon", 6, "black")
    if 農地監測狀態 in ["難以採樣", "無法採樣"] or cur in ["難以採樣", "無法採樣"]:
        return ("polygon", 6, "purple")

    # 系統型 vs 個案型（用調查方式/用水種類文字猜）
    is_system = ("系統" in method)
    is_case = ("個案" in method)

    # 以現況判定增量/延長/正常
    # （若主檔現況是增量/延長/正常）
    if cur == "增量":
        if is_system:
            return ("polygon", 3, "red")   # triangle red
        if is_case:
            return ("polygon", 4, "red")   # square red
        return ("polygon", 3, "red")
    if cur == "延長":
        if is_system:
            return ("polygon", 3, "blue")
        if is_case:
            return ("polygon", 4, "blue")
        return ("polygon", 3, "blue")
    if cur == "正常":
        if is_system:
            return ("polygon", 3, "green")
        if is_case:
            return ("polygon", 4, "green")
        return ("polygon", 3, "green")

    # default
    return ("circle", 0, "gray")


def page_map(engine):
    st.subheader("🗺️ 空間地圖（點擊看資訊卡）")
    if engine is None:
        st.error("DATABASE_URL 未設定")
        return

    with engine.begin() as conn:
        df = pd.read_sql(text("""
            SELECT 地段地號, sSGM編號, 網格編號, 鄉鎮市, 調查方式, 代表性, 用水種類,
                   coord_x, coord_y, 農地監測狀態, 農地監測狀態, 網格監測頻率, 最後調查年分
            FROM lands
            WHERE coord_x IS NOT NULL AND coord_y IS NOT NULL
            LIMIT 5000;
        """), conn)

    if df.empty:
        st.info("尚無座標資料（請先匯入 lands 或更新座標）")
        return

    # map center
    cx = float(df["coord_y"].mean())
    cy = float(df["coord_x"].mean())

    m = folium.Map(location=[cx, cy], zoom_start=11, tiles="OpenStreetMap")

    for _, r in df.iterrows():
        row = r.to_dict()
        x = row.get("coord_x")
        y = row.get("coord_y")
        if x is None or y is None:
            continue

        shape, sides, color = pick_marker_style(row)
        popup_html = f"""
        <b>地段地號：</b>{row.get('地段地號','')}<br>
        <b>SGM：</b>{row.get('sSGM編號','')}<br>
        <b>網格：</b>{row.get('網格編號','')}<br>
        <b>鄉鎮：</b>{row.get('鄉鎮市','')}<br>
        <b>現勘狀態：</b>{row.get('農地監測狀態','')}<br>
        <b>現況：</b>{row.get('農地監測狀態','')}<br>
        <b>頻率：</b>{row.get('網格監測頻率','')}<br>
        <b>最後年分：</b>{row.get('最後調查年分','')}<br>
        """

        if shape == "circle":
            folium.CircleMarker(
                location=[y, x],
                radius=5,
                color=color,
                fill=True,
                fill_opacity=0.8,
                popup=folium.Popup(popup_html, max_width=320),
            ).add_to(m)
        else:
            # RegularPolygonMarker：sides=3 triangle / 4 square / 6 hexagon
            folium.RegularPolygonMarker(
                location=[y, x],
                number_of_sides=sides,
                radius=7,
                color=color,
                fill=True,
                fill_color=color,
                fill_opacity=0.85,
                popup=folium.Popup(popup_html, max_width=360),
            ).add_to(m)

    st_folium(m, use_container_width=True, height=620)


# =========================================================
# Dashboard
# =========================================================
def page_dashboard(engine):
    st.title("🚜 彰化縣農地監測戰情室")
    st.markdown(f"### 🗓️ 當前時間：{roc_date_str()}")

    kpi = db_fetch_kpis(engine)
    cols = st.columns(6)
    cols[0].metric("總資料點數", kpi["total"])
    cols[1].metric("總採樣點數(代表+備用)", kpi["sample_points"])
    cols[2].metric("管制點數", kpi["control"])
    cols[3].metric("建物數量", kpi["building"])
    cols[4].metric("難以採樣數量", kpi["hard"])
    cols[5].metric("正常退場數量", kpi["normal_exit"])

    st.divider()

    # 管理者才顯示匯入工具
    if has_admin():
        admin_import_excel_to_db(engine)
    else:
        st.info("匯入工具需要管理者權限（側邊欄登入）")


# =========================================================
# Router
# =========================================================
def main():
    engine = get_engine()
    if engine is None:
        st.warning("⚠️ 尚未設定 DATABASE_URL。請到 Streamlit Secrets 設定 DATABASE_URL。")
    else:
        init_db(engine)

    login_sidebar()

    st.sidebar.markdown("### 📌 功能選單")
    page = st.sidebar.radio("前往頁面", [
        "首頁 Dashboard",
        "總表清單",
        "資料查詢",
        "新增年度調查結果",
        "歷史紀錄查詢",
        "同坵塊管理",
        "空間地圖",
    ])

    if page == "首頁 Dashboard":
        page_dashboard(engine)
    elif page == "總表清單":
        page_master_list(engine)
    elif page == "資料查詢":
        page_search(engine)
    elif page == "新增年度調查結果":
        page_add_annual(engine)
    elif page == "歷史紀錄查詢":
        page_history(engine)
    elif page == "同坵塊管理":
        page_blocks_manage(engine)
    elif page == "空間地圖":
        page_map(engine)


if __name__ == "__main__":
    main()





