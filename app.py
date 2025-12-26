# app.py
# -*- coding: utf-8 -*-

import base64
import datetime as dt
import json
import math
from dataclasses import dataclass
from typing import Dict, Any, Optional, Tuple, List

import pandas as pd
import numpy as np
import streamlit as st

from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

import folium
from streamlit_folium import st_folium


# =========================
# 0) 基本設定
# =========================
APP_TITLE = "彰化縣農地監測戰情室"
DEFAULT_TIMEZONE = "Asia/Taipei"
METALS_8 = ["汞", "砷", "銅", "鉻", "鎘", "鉛", "鋅", "鎳"]

# 你的判定對應
# metal_result: 增量/延長/正常/管制 (重金屬濃度調查結果)
# freq: 持續/延長/退場/管制 (策略監測頻率)
METAL_TO_FREQ = {
    "增量": "持續",
    "延長": "延長",
    "正常": "退場",
    "管制": "管制",
}

ADMIN_STATUS_SET = ["監測", "管制", "建物", "正常", "難以採樣"]


# =========================
# 1) 權限（讀者/編輯/管理）
# =========================
@dataclass
class Role:
    name: str
    can_edit: bool = False
    can_admin: bool = False


def get_role_from_sidebar() -> Role:
    st.sidebar.markdown("## 🔐 權限")
    role_name = st.sidebar.selectbox("角色", ["閱讀者", "編輯者", "管理者"], index=0)

    if role_name == "閱讀者":
        return Role("閱讀者", can_edit=False, can_admin=False)

    pwd = st.sidebar.text_input("密碼", type="password", placeholder="請輸入密碼")
    if role_name == "編輯者":
        ok = (pwd == st.secrets.get("EDITOR_PASSWORD", ""))
        if not ok and pwd:
            st.sidebar.error("密碼錯誤")
        return Role("編輯者", can_edit=ok, can_admin=False)

    if role_name == "管理者":
        ok = (pwd == st.secrets.get("ADMIN_PASSWORD", ""))
        if not ok and pwd:
            st.sidebar.error("密碼錯誤")
        return Role("管理者", can_edit=ok, can_admin=ok)

    return Role("閱讀者", can_edit=False, can_admin=False)


# =========================
# 2) DB 連線
# =========================
def get_engine() -> Optional[Engine]:
    """
    你要在 Streamlit Cloud -> App -> Settings -> Secrets 設：
    DB_URL = "postgresql+psycopg2://user:pass@host/db?sslmode=require"
    """
    db_url = st.secrets.get("DB_URL", "").strip()
    if not db_url:
        return None
    return create_engine(db_url, pool_pre_ping=True)


# =========================
# 3) DB schema + migration（避免舊表缺欄位）
# =========================
def init_db(engine: Engine) -> None:
    if engine is None:
        return

    with engine.begin() as conn:
        # ---- standards ----
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS standards (
            item TEXT PRIMARY KEY,
            monitor_std DOUBLE PRECISION,
            control_std DOUBLE PRECISION,
            da_threshold DOUBLE PRECISION
        );
        """))

        # ---- blocks（同坵塊對照）----
        # block_id + lot_no unique (複合主鍵)
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS blocks (
            block_id TEXT NOT NULL,
            lot_no TEXT NOT NULL,
            is_rep BOOLEAN DEFAULT FALSE,
            PRIMARY KEY (block_id, lot_no)
        );
        """))

        # ---- lands（農地現況主檔）----
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS lands (
            lot_no TEXT PRIMARY KEY,

            sgm_no TEXT,
            land_serial TEXT,
            grid_id TEXT,
            township TEXT,
            land_addr TEXT,
            survey_method TEXT,   -- 系統型/個案型（或你原本 Excel 的調查方式）
            rep_role TEXT,        -- 代表點/備用點/空白
            water_type TEXT,

            coord_x DOUBLE PRECISION,
            coord_y DOUBLE PRECISION,

            initial_metals JSONB,         -- 初始八項
            admin_status TEXT,            -- 監測/管制/建物/正常/難以採樣（行政狀態/農地監測狀態）
            current_metal_result TEXT,    -- 增量/延長/正常/管制（重金屬結果）
            freq TEXT,                    -- 持續/延長/退場/管制（策略監測頻率）
            last_year INTEGER,
            year_status JSONB,            -- {"101":"監測","102":"正常",...}

            updated_at TIMESTAMP DEFAULT NOW()
        );
        """))

        # ---- samples（歷年調查紀錄）----
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS samples (
            sample_id SERIAL PRIMARY KEY,
            lot_no TEXT NOT NULL,
            year INTEGER NOT NULL,
            sample_date DATE,

            coord_x DOUBLE PRECISION,
            coord_y DOUBLE PRECISION,

            photo_site_b64 TEXT,
            photo_sample_b64 TEXT,

            xrf JSONB,
            total JSONB,
            used_total BOOLEAN DEFAULT FALSE,

            admin_status TEXT,     -- 監測/管制/建物/正常/難以採樣（本次現勘）
            metal_result TEXT,     -- 增量/延長/正常/管制（本次重金屬結果）
            freq TEXT,             -- 持續/延長/退場/管制（本次策略頻率）

            da_pct JSONB,          -- {"砷": 12.3, ...}
            er JSONB,              -- {"砷": 1.12, ...}

            created_at TIMESTAMP DEFAULT NOW(),

            UNIQUE (lot_no, year)
        );
        """))

        # ========== Migration：補欄位（防舊表）==========
        # lands
        for col_sql in [
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS sgm_no TEXT;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS land_serial TEXT;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS grid_id TEXT;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS township TEXT;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS land_addr TEXT;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS survey_method TEXT;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS rep_role TEXT;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS water_type TEXT;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS coord_x DOUBLE PRECISION;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS coord_y DOUBLE PRECISION;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS initial_metals JSONB;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS admin_status TEXT;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS current_metal_result TEXT;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS freq TEXT;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS last_year INTEGER;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS year_status JSONB;",
            "ALTER TABLE lands ADD COLUMN IF NOT EXISTS updated_at TIMESTAMP DEFAULT NOW();",
        ]:
            conn.execute(text(col_sql))

        # samples
        for col_sql in [
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS sample_date DATE;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS coord_x DOUBLE PRECISION;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS coord_y DOUBLE PRECISION;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS photo_site_b64 TEXT;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS photo_sample_b64 TEXT;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS xrf JSONB;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS total JSONB;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS used_total BOOLEAN DEFAULT FALSE;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS admin_status TEXT;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS metal_result TEXT;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS freq TEXT;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS da_pct JSONB;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS er JSONB;",
            "ALTER TABLE samples ADD COLUMN IF NOT EXISTS created_at TIMESTAMP DEFAULT NOW();",
        ]:
            conn.execute(text(col_sql))

        # indexes
        conn.execute(text("CREATE INDEX IF NOT EXISTS idx_lands_grid_id ON lands(grid_id);"))
        conn.execute(text("CREATE INDEX IF NOT EXISTS idx_lands_admin_status ON lands(admin_status);"))
        conn.execute(text("CREATE INDEX IF NOT EXISTS idx_samples_lot_no ON samples(lot_no);"))


# =========================
# 4) 工具：圖片 b64
# =========================
def file_to_b64(file) -> Optional[str]:
    if file is None:
        return None
    data = file.read()
    if not data:
        return None
    return base64.b64encode(data).decode("utf-8")


def b64_to_image_tag(b64: str, width: int = 240) -> str:
    return f'<img src="data:image/jpeg;base64,{b64}" width="{width}"/>'


# =========================
# 5) standards 讀取
# =========================
def db_get_standards(engine: Engine) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    with engine.begin() as conn:
        rows = conn.execute(text("SELECT item, monitor_std, control_std, da_threshold FROM standards;")).fetchall()
    for r in rows:
        out[str(r[0])] = {
            "monitor_std": float(r[1]) if r[1] is not None else np.nan,
            "control_std": float(r[2]) if r[2] is not None else np.nan,
            "da_threshold": float(r[3]) if r[3] is not None else np.nan,
        }
    return out


# =========================
# 6) 計算：ER / DA / 判定
# =========================
def safe_float(x) -> Optional[float]:
    try:
        if x is None:
            return None
        if isinstance(x, str) and x.strip() == "":
            return None
        v = float(x)
        if math.isnan(v):
            return None
        return v
    except Exception:
        return None


def calc_er(values: Dict[str, Any], standards: Dict[str, Dict[str, float]]) -> Dict[str, Optional[float]]:
    er = {}
    for m in METALS_8:
        v = safe_float(values.get(m))
        std = standards.get(m, {}).get("monitor_std", np.nan)
        if v is None or std is None or (isinstance(std, float) and math.isnan(std)) or std == 0:
            er[m] = None
        else:
            er[m] = v / std
    return er


def pick_baseline_for_da(engine: Engine, lot_no: str, metals: Dict[str, Any]) -> Dict[str, Optional[float]]:
    """
    你的 DA 定義需要「納入定期監測時填報之重金屬污染物檢測值」作為 baseline。
    在平台上我們先用「此地段地號歷史 samples 最早一筆」當 baseline（你之後可改成指定年度）。
    """
    with engine.begin() as conn:
        row = conn.execute(text("""
            SELECT xrf, total, used_total
            FROM samples
            WHERE lot_no = :lot_no
            ORDER BY year ASC
            LIMIT 1;
        """), {"lot_no": lot_no}).fetchone()

    if not row:
        return {m: None for m in METALS_8}

    xrf0 = row[0] or {}
    total0 = row[1] or {}
    used_total0 = bool(row[2]) if row[2] is not None else False
    base = total0 if used_total0 else xrf0

    out = {}
    for m in METALS_8:
        out[m] = safe_float(base.get(m))
    return out


def calc_da_pct(current: Dict[str, Any], baseline: Dict[str, Optional[float]]) -> Dict[str, Optional[float]]:
    """
    DA% = ((current - baseline) / baseline) * 100
    """
    da = {}
    for m in METALS_8:
        cur = safe_float(current.get(m))
        base = baseline.get(m)
        if cur is None or base is None or base == 0:
            da[m] = None
        else:
            da[m] = ((cur - base) / base) * 100.0
    return da


def decide_metal_result(
    current_vals: Dict[str, Any],
    er: Dict[str, Optional[float]],
    da: Dict[str, Optional[float]],
    standards: Dict[str, Dict[str, float]],
    admin_status: str
) -> str:
    """
    決定本次重金屬結果：增量/延長/正常/管制
    - 若 admin_status 為 管制/建物/難以採樣：直接回對應（管制/建物/難以採樣）
      但為了你定義一致，我們在 metal_result 只回 增量/延長/正常/管制（建物/難以採樣屬 admin_status）
      => 若 admin_status=管制 -> metal_result=管制；其他 admin_status 不強制影響 metal_result
    - 若任一金屬 >= control_std -> 管制
    - 若任一金屬 > monitor_std（ER>1）：
        * 若任一金屬 DA% >= da_threshold -> 增量
        * else -> 延長
    - else -> 正常
    """
    if admin_status == "管制":
        return "管制"

    # 控制標準優先
    for m in METALS_8:
        v = safe_float(current_vals.get(m))
        ctl = standards.get(m, {}).get("control_std", np.nan)
        if v is not None and ctl is not None and not (isinstance(ctl, float) and math.isnan(ctl)):
            if v >= float(ctl):
                return "管制"

    any_over_monitor = any((er.get(m) is not None and er[m] > 1.0) for m in METALS_8)
    if not any_over_monitor:
        return "正常"

    # 增量判定：任一 DA% >= 門檻
    for m in METALS_8:
        th = standards.get(m, {}).get("da_threshold", np.nan)
        if th is None or (isinstance(th, float) and math.isnan(th)):
            continue
        if da.get(m) is not None and da[m] >= float(th):
            return "增量"

    return "延長"


# =========================
# 7) DB 寫入：samples + 同步更新 lands
# =========================
def db_upsert_sample_and_sync_land(
    engine: Engine,
    lot_no: str,
    year: int,
    sample_date: Optional[dt.date],
    coord_x: Optional[float],
    coord_y: Optional[float],
    admin_status: str,
    xrf_vals: Dict[str, Any],
    total_vals: Dict[str, Any],
    used_total: bool,
    photo_site_b64: Optional[str],
    photo_sample_b64: Optional[str],
) -> Tuple[str, str, Dict[str, Any], Dict[str, Any]]:
    """
    寫入 samples（同年同 lot_no 會更新）
    並同步更新 lands：
      - last_year
      - admin_status（農地監測狀態）
      - current_metal_result（增量/延長/正常/管制）
      - freq（持續/延長/退場/管制）
      - year_status JSONB (加上該年狀態：監測/正常/管制/建物/難以採樣)
    """
    standards = db_get_standards(engine)

    # 選擇用於計算的濃度
    vals_for_calc = total_vals if used_total else xrf_vals

    # baseline (DA)
    baseline = pick_baseline_for_da(engine, lot_no, vals_for_calc)
    er = calc_er(vals_for_calc, standards)
    da = calc_da_pct(vals_for_calc, baseline)

    metal_result = decide_metal_result(vals_for_calc, er, da, standards, admin_status)
    freq = METAL_TO_FREQ.get(metal_result, "退場")

    # 本年度狀態（你說要寫進 101~114 那種狀態欄；DB 用 year_status JSON 取代動態新增欄）
    # 規則：若任何金屬>監測標準 -> "監測"，若全部正常->"正常"，若 admin_status=管制/建物/難以採樣->對應寫入
    if admin_status in ["管制", "建物", "難以採樣"]:
        year_state = admin_status
    else:
        any_over = any((er.get(m) is not None and er[m] > 1.0) for m in METALS_8)
        year_state = "監測" if any_over else "正常"

    with engine.begin() as conn:
        # samples upsert
        conn.execute(text("""
            INSERT INTO samples (
                lot_no, year, sample_date,
                coord_x, coord_y,
                photo_site_b64, photo_sample_b64,
                xrf, total, used_total,
                admin_status, metal_result, freq,
                da_pct, er, created_at
            )
            VALUES (
                :lot_no, :year, :sample_date,
                :coord_x, :coord_y,
                :photo_site_b64, :photo_sample_b64,
                :xrf::jsonb, :total::jsonb, :used_total,
                :admin_status, :metal_result, :freq,
                :da_pct::jsonb, :er::jsonb, NOW()
            )
            ON CONFLICT (lot_no, year) DO UPDATE SET
                sample_date = EXCLUDED.sample_date,
                coord_x = EXCLUDED.coord_x,
                coord_y = EXCLUDED.coord_y,
                photo_site_b64 = EXCLUDED.photo_site_b64,
                photo_sample_b64 = EXCLUDED.photo_sample_b64,
                xrf = EXCLUDED.xrf,
                total = EXCLUDED.total,
                used_total = EXCLUDED.used_total,
                admin_status = EXCLUDED.admin_status,
                metal_result = EXCLUDED.metal_result,
                freq = EXCLUDED.freq,
                da_pct = EXCLUDED.da_pct,
                er = EXCLUDED.er,
                created_at = NOW();
        """), {
            "lot_no": lot_no,
            "year": int(year),
            "sample_date": sample_date.isoformat() if sample_date else None,
            "coord_x": coord_x,
            "coord_y": coord_y,
            "photo_site_b64": photo_site_b64,
            "photo_sample_b64": photo_sample_b64,
            "xrf": json.dumps(xrf_vals or {}, ensure_ascii=False),
            "total": json.dumps(total_vals or {}, ensure_ascii=False),
            "used_total": bool(used_total),
            "admin_status": admin_status,
            "metal_result": metal_result,
            "freq": freq,
            "da_pct": json.dumps(da or {}, ensure_ascii=False),
            "er": json.dumps(er or {}, ensure_ascii=False),
        })

        # lands 同步 year_status json
        prev = conn.execute(text("SELECT year_status FROM lands WHERE lot_no=:lot_no;"), {"lot_no": lot_no}).fetchone()
        year_status = prev[0] if prev and prev[0] else {}
        if isinstance(year_status, str):
            try:
                year_status = json.loads(year_status)
            except Exception:
                year_status = {}
        year_status[str(year)] = year_state

        conn.execute(text("""
            UPDATE lands
            SET last_year = :last_year,
                admin_status = :admin_status,
                current_metal_result = :current_metal_result,
                freq = :freq,
                year_status = :year_status::jsonb,
                updated_at = NOW()
            WHERE lot_no = :lot_no;
        """), {
            "lot_no": lot_no,
            "last_year": int(year),
            "admin_status": admin_status,
            "current_metal_result": metal_result,
            "freq": freq,
            "year_status": json.dumps(year_status, ensure_ascii=False),
        })

    return metal_result, freq, er, da


# =========================
# 8) Excel 匯入（lands / standards / blocks）
# =========================
LANDS_COL_MAP = {
    "地段地號": "lot_no",
    "SGM編號": "sgm_no",
    "農地序號": "land_serial",
    "網格編號": "grid_id",
    "鄉鎮市": "township",
    "地段地號(地址)": "land_addr",  # 可有可無
    "地段地號 ": "lot_no",  # 防尾空白欄名
    "地段地號　": "lot_no",
    "調查方式": "survey_method",
    "代表性": "rep_role",
    "用水種類": "water_type",
    "TWD97_X": "coord_x",
    "TWD97_Y": "coord_y",
    "目前農地調查現況": "current_metal_result",
    "農地監測狀態": "admin_status",
    "網格監測頻率": "freq",
    "最後調查年分": "last_year",
}

def read_excel_sheet(excel_path: str, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def import_standards_from_excel(engine: Engine, excel_path: str, sheet_name="判定標準表") -> int:
    df = read_excel_sheet(excel_path, sheet_name=sheet_name).copy()
    df.columns = [str(c).strip() for c in df.columns]

    # 期望欄：項目名稱/監測標準/管制標準/上升標準 (DA門檻)
    rename = {
        "項目名稱": "item",
        "監測標準": "monitor_std",
        "管制標準": "control_std",
        "上升標準 (DA門檻)": "da_threshold",
        "上升標準(DA門檻)": "da_threshold",
        "上升標準": "da_threshold",
    }
    df = df.rename(columns=rename)

    need = ["item", "monitor_std", "control_std", "da_threshold"]
    for c in need:
        if c not in df.columns:
            df[c] = None

    df["item"] = df["item"].astype(str).str.strip()
    df = df[df["item"].notna() & (df["item"] != "")]
    df = df[need].copy()

    df["monitor_std"] = pd.to_numeric(df["monitor_std"], errors="coerce")
    df["control_std"] = pd.to_numeric(df["control_std"], errors="coerce")
    df["da_threshold"] = pd.to_numeric(df["da_threshold"], errors="coerce")

    rows = df.to_dict(orient="records")

    with engine.begin() as conn:
        for r in rows:
            conn.execute(text("""
                INSERT INTO standards (item, monitor_std, control_std, da_threshold)
                VALUES (:item, :monitor_std, :control_std, :da_threshold)
                ON CONFLICT (item) DO UPDATE SET
                    monitor_std = EXCLUDED.monitor_std,
                    control_std = EXCLUDED.control_std,
                    da_threshold = EXCLUDED.da_threshold;
            """), r)
    return len(rows)

def import_blocks_from_excel(engine: Engine, excel_path: str, sheet_name="同坵塊對照表") -> int:
    df = read_excel_sheet(excel_path, sheet_name=sheet_name).copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename = {
        "農地群組編號": "block_id",
        "農地地段地號": "lot_no",
        "代表農地": "is_rep",
    }
    df = df.rename(columns=rename)

    for c in ["block_id", "lot_no", "is_rep"]:
        if c not in df.columns:
            df[c] = None

    df["block_id"] = df["block_id"].astype(str).str.strip()
    df["lot_no"] = df["lot_no"].astype(str).str.strip()

    def to_bool(x):
        if pd.isna(x):
            return False
        s = str(x).strip()
        return (s in ["是", "是 (代表)", "代表", "TRUE", "True", "1", "Y", "y"])

    df["is_rep"] = df["is_rep"].apply(to_bool)

    df = df[df["block_id"].notna() & (df["block_id"] != "") & df["lot_no"].notna() & (df["lot_no"] != "")]
    rows = df[["block_id", "lot_no", "is_rep"]].to_dict(orient="records")

    with engine.begin() as conn:
        # 這裡不清空，改為 upsert（避免 IntegrityError）
        for r in rows:
            conn.execute(text("""
                INSERT INTO blocks (block_id, lot_no, is_rep)
                VALUES (:block_id, :lot_no, :is_rep)
                ON CONFLICT (block_id, lot_no) DO UPDATE SET
                    is_rep = EXCLUDED.is_rep;
            """), r)
    return len(rows)

def import_lands_from_excel(engine: Engine, excel_path: str, sheet_name="農地現況主檔") -> int:
    df = read_excel_sheet(excel_path, sheet_name=sheet_name).copy()
    df.columns = [str(c).strip() for c in df.columns]

    # 中文欄位 -> DB 欄位
    rename = {}
    for c in df.columns:
        if c in LANDS_COL_MAP:
            rename[c] = LANDS_COL_MAP[c]
    df = df.rename(columns=rename)

    # 最少要有 lot_no
    if "lot_no" not in df.columns:
        raise ValueError("Excel 主檔缺少『地段地號』欄位（或欄名不一致）。")

    # 基本清理
    df["lot_no"] = df["lot_no"].astype(str).str.strip()
    df = df[df["lot_no"].notna() & (df["lot_no"] != "")].copy()

    # 數值欄
    for c in ["coord_x", "coord_y", "last_year"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # 初始八項（若有）
    initial_cols = {
        "初始_汞": "汞",
        "初始_砷": "砷",
        "初始_銅": "銅",
        "初始_鉻": "鉻",
        "初始_鎘": "鎘",
        "初始_鉛": "鉛",
        "初始_鋅": "鋅",
        "初始_鎳": "鎳",
    }

    initial_metals = None
    if any(c in df.columns for c in initial_cols.keys()):
        initial_metals = []
        for _, row in df.iterrows():
            obj = {}
            for excel_c, m in initial_cols.items():
                if excel_c in df.columns:
                    obj[m] = safe_float(row.get(excel_c))
            initial_metals.append(obj)
        df["initial_metals"] = initial_metals

    # year_status（如果你 Excel 有 101狀態~114狀態，我們把它收進 json）
    ycols = [c for c in df.columns if str(c).endswith("狀態")]
    def build_year_status(r):
        ys = {}
        for c in ycols:
            s = str(r.get(c)).strip()
            if s and s != "nan":
                year_txt = str(c).replace("狀態", "").replace(".", "").strip()
                if year_txt.isdigit():
                    ys[year_txt] = s
        return ys

    df["year_status"] = df.apply(build_year_status, axis=1)

    # 欄位補齊
    want = [
        "lot_no","sgm_no","land_serial","grid_id","township","land_addr",
        "survey_method","rep_role","water_type",
        "coord_x","coord_y",
        "initial_metals",
        "admin_status","current_metal_result","freq","last_year",
        "year_status"
    ]
    for c in want:
        if c not in df.columns:
            df[c] = None

    rows = df[want].to_dict(orient="records")

    with engine.begin() as conn:
        for r in rows:
            conn.execute(text("""
                INSERT INTO lands (
                    lot_no, sgm_no, land_serial, grid_id, township, land_addr,
                    survey_method, rep_role, water_type,
                    coord_x, coord_y,
                    initial_metals,
                    admin_status, current_metal_result, freq, last_year,
                    year_status,
                    updated_at
                )
                VALUES (
                    :lot_no, :sgm_no, :land_serial, :grid_id, :township, :land_addr,
                    :survey_method, :rep_role, :water_type,
                    :coord_x, :coord_y,
                    :initial_metals::jsonb,
                    :admin_status, :current_metal_result, :freq, :last_year,
                    :year_status::jsonb,
                    NOW()
                )
                ON CONFLICT (lot_no) DO UPDATE SET
                    sgm_no = EXCLUDED.sgm_no,
                    land_serial = EXCLUDED.land_serial,
                    grid_id = EXCLUDED.grid_id,
                    township = EXCLUDED.township,
                    land_addr = EXCLUDED.land_addr,
                    survey_method = EXCLUDED.survey_method,
                    rep_role = EXCLUDED.rep_role,
                    water_type = EXCLUDED.water_type,
                    coord_x = EXCLUDED.coord_x,
                    coord_y = EXCLUDED.coord_y,
                    initial_metals = EXCLUDED.initial_metals,
                    admin_status = EXCLUDED.admin_status,
                    current_metal_result = EXCLUDED.current_metal_result,
                    freq = EXCLUDED.freq,
                    last_year = EXCLUDED.last_year,
                    year_status = EXCLUDED.year_status,
                    updated_at = NOW();
            """), {
                **r,
                "initial_metals": json.dumps(r.get("initial_metals") or {}, ensure_ascii=False),
                "year_status": json.dumps(r.get("year_status") or {}, ensure_ascii=False),
            })
    return len(rows)


# =========================
# 9) KPI / 查詢
# =========================
def db_fetch_kpis(engine: Engine) -> Dict[str, int]:
    with engine.begin() as conn:
        total = conn.execute(text("SELECT COUNT(*) FROM lands;")).scalar() or 0
        ctrl = conn.execute(text("SELECT COUNT(*) FROM lands WHERE admin_status='管制';")).scalar() or 0
        bldg = conn.execute(text("SELECT COUNT(*) FROM lands WHERE admin_status='建物';")).scalar() or 0
        hard = conn.execute(text("SELECT COUNT(*) FROM lands WHERE admin_status='難以採樣';")).scalar() or 0
        normal = conn.execute(text("SELECT COUNT(*) FROM lands WHERE admin_status='正常';")).scalar() or 0

        # 採樣點：代表點 + 備用點（rep_role 非空）
        sample_points = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE COALESCE(NULLIF(TRIM(rep_role), ''), '') <> '';
        """)).scalar() or 0

        grids_total = conn.execute(text("""
            SELECT COUNT(DISTINCT grid_id) FROM lands WHERE grid_id IS NOT NULL AND TRIM(grid_id) <> '';
        """)).scalar() or 0

        freq_continue = conn.execute(text("""
            SELECT COUNT(DISTINCT grid_id) FROM lands
            WHERE freq='持續' AND grid_id IS NOT NULL AND TRIM(grid_id) <> '';
        """)).scalar() or 0

        freq_extend = conn.execute(text("""
            SELECT COUNT(DISTINCT grid_id) FROM lands
            WHERE freq='延長' AND grid_id IS NOT NULL AND TRIM(grid_id) <> '';
        """)).scalar() or 0

        freq_exit = conn.execute(text("""
            SELECT COUNT(DISTINCT grid_id) FROM lands
            WHERE freq='退場' AND grid_id IS NOT NULL AND TRIM(grid_id) <> '';
        """)).scalar() or 0

        # 個案型（survey_method 包含 個案）
        case_total = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE survey_method LIKE '%個案%';
        """)).scalar() or 0
        case_continue = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE survey_method LIKE '%個案%' AND freq='持續';
        """)).scalar() or 0
        case_extend = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE survey_method LIKE '%個案%' AND freq='延長';
        """)).scalar() or 0
        case_normal = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE survey_method LIKE '%個案%' AND freq='退場';
        """)).scalar() or 0

    return {
        "total": int(total),
        "sample_points": int(sample_points),
        "ctrl": int(ctrl),
        "bldg": int(bldg),
        "hard": int(hard),
        "normal": int(normal),
        "grids_total": int(grids_total),
        "grids_continue": int(freq_continue),
        "grids_extend": int(freq_extend),
        "grids_exit": int(freq_exit),
        "case_total": int(case_total),
        "case_continue": int(case_continue),
        "case_extend": int(case_extend),
        "case_normal": int(case_normal),
    }


def db_query_lands(engine: Engine, keyword: str, limit: int = 200) -> pd.DataFrame:
    kw = f"%{keyword.strip()}%"
    with engine.begin() as conn:
        rows = conn.execute(text("""
            SELECT lot_no, sgm_no, grid_id, township, survey_method, rep_role, water_type,
                   admin_status, current_metal_result, freq, last_year
            FROM lands
            WHERE lot_no ILIKE :kw OR COALESCE(sgm_no,'') ILIKE :kw
            ORDER BY lot_no
            LIMIT :limit;
        """), {"kw": kw, "limit": int(limit)}).fetchall()

    df = pd.DataFrame(rows, columns=[
        "地段地號","sgm編號","網格編號","鄉鎮市","調查方式","代表性","用水種類",
        "農地監測狀態","目前農地調查現況","網格監測頻率","最後調查年分"
    ])
    return df


def db_list_lands(engine: Engine, limit: int = 500) -> pd.DataFrame:
    with engine.begin() as conn:
        rows = conn.execute(text("""
            SELECT lot_no, sgm_no, grid_id, township, survey_method, rep_role, water_type,
                   admin_status, current_metal_result, freq, last_year
            FROM lands
            ORDER BY lot_no
            LIMIT :limit;
        """), {"limit": int(limit)}).fetchall()

    return pd.DataFrame(rows, columns=[
        "地段地號","sgm編號","網格編號","鄉鎮市","調查方式","代表性","用水種類",
        "農地監測狀態","目前農地調查現況","網格監測頻率","最後調查年分"
    ])


def db_history_samples(engine: Engine, lot_no: str) -> pd.DataFrame:
    with engine.begin() as conn:
        rows = conn.execute(text("""
            SELECT year, sample_date, admin_status, metal_result, freq, used_total, er, da_pct
            FROM samples
            WHERE lot_no = :lot_no
            ORDER BY year DESC;
        """), {"lot_no": lot_no}).fetchall()

    df = pd.DataFrame(rows, columns=["年度","日期","本次現勘","重金屬結果","策略頻率","用全量計算","ER","DA%"])
    return df


def db_get_land(engine: Engine, lot_no: str) -> Optional[Dict[str, Any]]:
    with engine.begin() as conn:
        row = conn.execute(text("""
            SELECT lot_no, sgm_no, grid_id, township, survey_method, rep_role, water_type,
                   coord_x, coord_y, admin_status, current_metal_result, freq, last_year, year_status
            FROM lands WHERE lot_no=:lot_no;
        """), {"lot_no": lot_no}).fetchone()
    if not row:
        return None
    keys = ["lot_no","sgm_no","grid_id","township","survey_method","rep_role","water_type",
            "coord_x","coord_y","admin_status","current_metal_result","freq","last_year","year_status"]
    d = dict(zip(keys, row))
    if isinstance(d.get("year_status"), str):
        try:
            d["year_status"] = json.loads(d["year_status"])
        except Exception:
            d["year_status"] = {}
    return d


# =========================
# 10) UI Pages
# =========================
def page_dashboard(engine: Engine):
    st.markdown(f"# 🚜 {APP_TITLE}")

    today = dt.date.today()
    st.markdown(f"### 🗓️ 當前時間：民國 {today.year - 1911} 年 {today.month} 月 {today.day} 日")

    if engine is None:
        st.warning("尚未設定 DB_URL（Secrets），目前為展示模式（不會寫入資料庫）。")
        return

    kpi = db_fetch_kpis(engine)

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("總資料點數", kpi["total"])
    c2.metric("總採樣點數(代表+備用)", kpi["sample_points"])
    c3.metric("管制點數", kpi["ctrl"])
    c4.metric("建物數量", kpi["bldg"])
    c5.metric("難以採樣數量", kpi["hard"])
    c6.metric("正常退場數量", kpi["normal"])

    st.divider()

    g1, g2, g3, g4 = st.columns(4)
    g1.metric("網格總數", kpi["grids_total"])
    g2.metric("持續網格", kpi["grids_continue"])
    g3.metric("延長網格", kpi["grids_extend"])
    g4.metric("退場網格", kpi["grids_exit"])

    st.divider()

    a1, a2, a3, a4 = st.columns(4)
    a1.metric("個案型總數", kpi["case_total"])
    a2.metric("個案-持續", kpi["case_continue"])
    a3.metric("個案-延長", kpi["case_extend"])
    a4.metric("個案-正常", kpi["case_normal"])


def page_admin_import(engine: Engine, role: Role):
    st.markdown("## 🛠️ 管理員工具：Excel → DB 匯入")

    if engine is None:
        st.error("未設定 DB_URL，無法匯入。")
        return
    if not role.can_admin:
        st.info("請在左側選擇『管理者』並輸入密碼。")
        return

    st.write("此工具會把 repo 裡的 Excel 主檔、標準表、同坵塊表匯入 DB（可重複匯入，會 UPSERT 更新）。")

    excel_path = st.text_input("Excel 路徑（repo 內檔名）", value="彰化農地管理資料庫.xlsx")
    sheet_lands = st.text_input("主檔 Sheet 名稱", value="農地現況主檔")
    sheet_std = st.text_input("標準表 Sheet 名稱", value="判定標準表")
    sheet_blocks = st.text_input("同坵塊 Sheet 名稱", value="同坵塊對照表")

    if st.button("🚀 一鍵匯入（lands / standards / blocks）"):
        try:
            n_std = import_standards_from_excel(engine, excel_path, sheet_name=sheet_std)
            n_blk = import_blocks_from_excel(engine, excel_path, sheet_name=sheet_blocks)
            n_lnd = import_lands_from_excel(engine, excel_path, sheet_name=sheet_lands)
            st.success(f"完成匯入：standards={n_std}，blocks={n_blk}，lands={n_lnd}")
        except Exception as e:
            st.exception(e)

    st.caption("💡 匯入後，回首頁 KPI 應該不再是 0。")


def page_list(engine: Engine):
    st.markdown("## 📋 總表清單")
    if engine is None:
        st.warning("未設定 DB_URL")
        return
    limit = st.slider("顯示筆數", 50, 2000, 500, step=50)
    df = db_list_lands(engine, limit=limit)
    st.dataframe(df, use_container_width=True, height=520)


def page_search(engine: Engine):
    st.markdown("## 🔎 資料查詢（SGM / 地段地號）")
    if engine is None:
        st.warning("未設定 DB_URL")
        return
    kw = st.text_input("輸入關鍵字（例如：華南段0159-0000）")
    if not kw.strip():
        return
    df = db_query_lands(engine, kw, limit=500)
    st.dataframe(df, use_container_width=True, height=520)


def page_history(engine: Engine):
    st.markdown("## 🕰️ 歷史紀錄查詢")
    if engine is None:
        st.warning("未設定 DB_URL")
        return
    lot_no = st.text_input("輸入地段地號（lot_no）")
    if not lot_no.strip():
        return
    land = db_get_land(engine, lot_no.strip())
    if not land:
        st.error("查無此地段地號")
        return

    st.write("### 農地基本資訊")
    st.json({
        "地段地號": land["lot_no"],
        "SGM": land.get("sgm_no"),
        "網格": land.get("grid_id"),
        "鄉鎮": land.get("township"),
        "調查方式": land.get("survey_method"),
        "代表性": land.get("rep_role"),
        "用水": land.get("water_type"),
        "座標": (land.get("coord_x"), land.get("coord_y")),
        "農地監測狀態": land.get("admin_status"),
        "目前農地調查現況": land.get("current_metal_result"),
        "網格監測頻率": land.get("freq"),
        "最後調查年分": land.get("last_year"),
    })

    st.write("### 歷年 samples")
    df = db_history_samples(engine, lot_no.strip())
    st.dataframe(df, use_container_width=True, height=420)

    st.write("### 歷年狀態 year_status（JSON）")
    st.json(land.get("year_status") or {})


def page_add_year(engine: Engine, role: Role):
    st.markdown("## ➕ 新增年度調查結果")
    if engine is None:
        st.warning("未設定 DB_URL")
        return
    if not role.can_edit:
        st.info("請在左側切換『編輯者』或『管理者』並輸入密碼。")
        return

    lot_no = st.text_input("搜尋地段地號（lot_no）", placeholder="例如：華南段0159-0000")
    if not lot_no.strip():
        return

    land = db_get_land(engine, lot_no.strip())
    if not land:
        st.error("查無此地段地號（請先匯入 lands）")
        return

    st.write("### ✅ 找到農地")
    st.json({
        "地段地號": land["lot_no"],
        "網格": land.get("grid_id"),
        "鄉鎮": land.get("township"),
        "調查方式": land.get("survey_method"),
        "用水種類": land.get("water_type"),
        "既有座標": (land.get("coord_x"), land.get("coord_y")),
        "目前狀態": {
            "農地監測狀態": land.get("admin_status"),
            "重金屬結果": land.get("current_metal_result"),
            "策略頻率": land.get("freq"),
            "最後年度": land.get("last_year"),
        }
    })

    st.divider()

    colA, colB, colC = st.columns(3)
    year = colA.number_input("調查年度（民國）", min_value=90, max_value=200, value=114, step=1)
    date_input = colB.date_input("採樣日期", value=dt.date.today())
    admin_status = colC.selectbox("本次現勘（行政狀態）", ADMIN_STATUS_SET, index=0)

    col1, col2 = st.columns(2)
    coord_x = col1.number_input("本次採樣座標 X（TWD97）", value=float(land.get("coord_x") or 0.0))
    coord_y = col2.number_input("本次採樣座標 Y（TWD97）", value=float(land.get("coord_y") or 0.0))

    # 座標偏移提醒（>3m）
    try:
        ox = float(land.get("coord_x") or 0.0)
        oy = float(land.get("coord_y") or 0.0)
        dist = math.sqrt((coord_x - ox) ** 2 + (coord_y - oy) ** 2)
        if ox != 0.0 and oy != 0.0 and dist > 3.0:
            st.warning(f"⚠️ 座標偏移約 {dist:.2f} m（>3m）。請確認是否需要更新點位，並評估是否跨網格。")
    except Exception:
        pass

    st.write("### 📷 照片上傳")
    p1, p2 = st.columns(2)
    photo_site = p1.file_uploader("現勘照片", type=["jpg","jpeg","png"])
    photo_sample = p2.file_uploader("採樣照片", type=["jpg","jpeg","png"])

    st.divider()

    st.write("### 🧪 重金屬結果（XRF / 全量）")
    st.caption("DA/ER 計算：預設用 XRF；若你有填『全量』且勾選使用全量，則以全量計算。")

    xrf_vals = {}
    total_vals = {}
    cX, cT = st.columns(2)

    with cX:
        st.write("#### XRF（八項）")
        for m in METALS_8:
            xrf_vals[m] = st.text_input(f"XRF_{m}", value="")
    with cT:
        st.write("#### 全量（八項）")
        for m in METALS_8:
            total_vals[m] = st.text_input(f"全量_{m}", value="")

    used_total = st.checkbox("✅ 優先使用全量進行 DA/ER 計算（若有填）", value=False)

    # 把輸入轉 float
    xrf_vals = {m: safe_float(v) for m, v in xrf_vals.items()}
    total_vals = {m: safe_float(v) for m, v in total_vals.items()}

    if used_total:
        # 若全量都沒填，就自動退回 XRF
        if all(total_vals.get(m) is None for m in METALS_8):
            st.info("你勾選使用全量，但全量未填，將自動用 XRF 計算。")
            used_total = False

    st.divider()

    if "pending_samples" not in st.session_state:
        st.session_state["pending_samples"] = []

    if st.button("✅ 加入暫存清單（尚未寫入 DB）"):
        st.session_state["pending_samples"].append({
            "lot_no": lot_no.strip(),
            "year": int(year),
            "sample_date": date_input,
            "coord_x": float(coord_x) if coord_x else None,
            "coord_y": float(coord_y) if coord_y else None,
            "admin_status": admin_status,
            "xrf_vals": xrf_vals,
            "total_vals": total_vals,
            "used_total": bool(used_total),
            "photo_site_b64": file_to_b64(photo_site),
            "photo_sample_b64": file_to_b64(photo_sample),
        })
        st.success("已加入暫存清單。你可以再檢查後按『上傳寫入 DB』。")

    st.write("### 🧾 本次暫存清單")
    pending = st.session_state.get("pending_samples", [])
    if pending:
        st.dataframe(pd.DataFrame([
            {
                "lot_no": p["lot_no"],
                "year": p["year"],
                "admin_status": p["admin_status"],
                "used_total": p["used_total"],
                "coord_x": p["coord_x"],
                "coord_y": p["coord_y"],
            } for p in pending
        ]), use_container_width=True, height=220)

        colU, colC = st.columns(2)
        if colC.button("🗑️ 清空暫存"):
            st.session_state["pending_samples"] = []
            st.rerun()

        if colU.button("⬆️ 上傳寫入 DB（批次）"):
            ok = 0
            for p in list(st.session_state["pending_samples"]):
                try:
                    metal_result, freq, er, da = db_upsert_sample_and_sync_land(
                        engine=engine,
                        lot_no=p["lot_no"],
                        year=p["year"],
                        sample_date=p["sample_date"],
                        coord_x=p["coord_x"],
                        coord_y=p["coord_y"],
                        admin_status=p["admin_status"],
                        xrf_vals=p["xrf_vals"],
                        total_vals=p["total_vals"],
                        used_total=p["used_total"],
                        photo_site_b64=p["photo_site_b64"],
                        photo_sample_b64=p["photo_sample_b64"],
                    )
                    ok += 1
                except Exception as e:
                    st.error(f"寫入失敗：{p['lot_no']} {p['year']}")
                    st.exception(e)

            st.session_state["pending_samples"] = []
            st.success(f"完成寫入 {ok} 筆。")
            st.rerun()
    else:
        st.info("目前暫存清單是空的。")


def page_map(engine: Engine):
    st.markdown("## 🗺️ 空間地圖檢視")
    st.caption("依照狀態顯示點位（顏色/圖示）。若未顯示，通常是座標欄位為空或為 0。")

    if engine is None:
        st.warning("未設定 DB_URL")
        return

    with engine.begin() as conn:
        rows = conn.execute(text("""
            SELECT lot_no, grid_id, survey_method, admin_status, current_metal_result, freq, coord_x, coord_y, rep_role
            FROM lands
            WHERE coord_x IS NOT NULL AND coord_y IS NOT NULL
              AND coord_x <> 0 AND coord_y <> 0
            LIMIT 5000;
        """)).fetchall()

    if not rows:
        st.info("沒有可用座標資料（coord_x/coord_y）。")
        return

    df = pd.DataFrame(rows, columns=["lot_no","grid_id","survey_method","admin_status","metal_result","freq","x","y","rep_role"])

    # 以平均座標置中
    cx = float(df["y"].mean())
    cy = float(df["x"].mean())
    m = folium.Map(location=[cx, cy], zoom_start=11, control_scale=True)

    # 顏色規則
    # 系統型-增量/延長/正常：紅/藍/綠
    # 個案型-增量/延長/正常：紅/藍/綠
    # 管制：紅；建物：黑；難以採樣：紫
    # 備用點：白（用灰邊框表示）
    def marker_style(r):
        admin = r["admin_status"]
        metal = r["metal_result"]
        survey = str(r["survey_method"] or "")
        rep = str(r["rep_role"] or "")

        # 特殊行政狀態
        if admin == "管制":
            return ("red", "ban")   # 六角形在 folium 不好做，先用 icon
        if admin == "建物":
            return ("black", "home")
        if admin == "難以採樣":
            return ("purple", "question")

        # 備用點
        is_backup = ("備用" in rep)

        # 增量/延長/正常
        if metal == "增量":
            color = "red"
        elif metal == "延長":
            color = "blue"
        else:
            color = "green"

        icon = "triangle" if ("系統" in survey) else "square"
        if is_backup:
            # 備用點：用淡色 + 圖示
            return ("lightgray", "diamond")
        return (color, icon)

    # folium icon 受限：用 FontAwesome icon 近似
    icon_map = {
        "triangle": "caret-up",
        "square": "stop",
        "diamond": "certificate",
        "ban": "ban",
        "home": "home",
        "question": "question",
    }

    for _, r in df.iterrows():
        color, icon_key = marker_style(r)
        popup = folium.Popup(html=f"""
            <b>{r['lot_no']}</b><br/>
            網格：{r['grid_id']}<br/>
            調查方式：{r['survey_method']}<br/>
            農地監測狀態：{r['admin_status']}<br/>
            重金屬結果：{r['metal_result']}<br/>
            策略頻率：{r['freq']}<br/>
            代表性：{r['rep_role']}<br/>
        """, max_width=360)

        folium.Marker(
            location=[float(r["y"]), float(r["x"])],
            popup=popup,
            tooltip=r["lot_no"],
            icon=folium.Icon(color=color, icon=icon_map.get(icon_key, "info-sign"), prefix="fa"),
        ).add_to(m)

    st_folium(m, width=1100, height=650)


# =========================
# 11) 主程式
# =========================
def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")

    role = get_role_from_sidebar()

    engine = get_engine()
    if engine is not None:
        init_db(engine)

    # Sidebar nav
    st.sidebar.markdown("---")
    page = st.sidebar.radio(
        "選單",
        ["首頁", "總表清單", "資料查詢", "歷史紀錄查詢", "新增年度調查結果", "空間地圖檢視", "管理者：Excel 匯入"],
        index=0
    )

    if page == "首頁":
        page_dashboard(engine)
    elif page == "總表清單":
        page_list(engine)
    elif page == "資料查詢":
        page_search(engine)
    elif page == "歷史紀錄查詢":
        page_history(engine)
    elif page == "新增年度調查結果":
        page_add_year(engine, role)
    elif page == "空間地圖檢視":
        page_map(engine)
    elif page == "管理者：Excel 匯入":
        page_admin_import(engine, role)


if __name__ == "__main__":
    main()





