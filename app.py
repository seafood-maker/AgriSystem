import os
import json
import pandas as pd
import streamlit as st
from datetime import datetime, timezone, timedelta

from sqlalchemy import create_engine, text

# =========================
# 基本設定
# =========================
st.set_page_config(page_title="彰化縣農地監測戰情室", layout="wide")

TZ_TW = timezone(timedelta(hours=8))

EXCEL_PATH = "彰化農地管理資料庫.xlsx"

# Excel Sheet Names
SHEET_MASTER = "農地現況主檔"
SHEET_RECORDS = "歷年調查紀錄"
SHEET_BLOCKS = "同坵塊對照表"
SHEET_STANDARDS = "判定標準表"


# =========================
# 連線：DATABASE_URL
# =========================
def get_database_url() -> str:
    # Streamlit Cloud 建議放在 Secrets：DATABASE_URL
    if "DATABASE_URL" in st.secrets:
        return st.secrets["DATABASE_URL"]
    # 本機測試可用環境變數
    if os.getenv("DATABASE_URL"):
        return os.getenv("DATABASE_URL")
    return ""


@st.cache_resource
def get_engine():
    db_url = get_database_url()
    if not db_url:
        return None
    # Neon / Supabase 通常給的是 postgres:// 或 postgresql:// 都可
    return create_engine(db_url, pool_pre_ping=True)


# =========================
# DB Schema 初始化
# =========================
def init_db(engine):
    if engine is None:
        return

    with engine.begin() as conn:
        # lands：主檔（以 lot_no=地段地號 作唯一鍵）
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS lands (
            lot_no TEXT PRIMARY KEY,
            sgm_no TEXT,
            land_serial TEXT,
            grid_id TEXT,
            township TEXT,
            section_no TEXT,
            survey_method TEXT,
            rep_role TEXT,           -- 代表性（代表/備用/空白）
            water_type TEXT,

            coord_x DOUBLE PRECISION,
            coord_y DOUBLE PRECISION,

            initial_metals JSONB,     -- 初始八項
            current_metal_result TEXT, -- 目前農地調查現況（增量/延長/正常/管制/建物/難以採樣）
            admin_status TEXT,         -- 農地監測狀態（監測/管制/建物/正常/難以採樣）
            freq TEXT,                 -- 網格監測頻率（持續/延長/退場/管制等）
            last_year INTEGER,         -- 最後調查年分（民國年）

            year_status JSONB,         -- {"101":"監測","102":"正常",...}

            updated_at TIMESTAMP DEFAULT NOW()
        );
        """))

        # standards：判定標準表
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS standards (
            item TEXT PRIMARY KEY,
            monitor_std DOUBLE PRECISION,
            control_std DOUBLE PRECISION,
            da_threshold DOUBLE PRECISION
        );
        """))

        # blocks：同坵塊對照表（block_id + lot_no 為複合主鍵）
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS blocks (
            block_id TEXT NOT NULL,
            lot_no TEXT NOT NULL,
            is_rep BOOLEAN DEFAULT FALSE,
            PRIMARY KEY (block_id, lot_no)
        );
        """))

        # samples：歷年調查紀錄（可後續在「新增年度調查結果」頁面寫入）
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
            freq TEXT,               -- 持續/延長/退場/管制（策略頻率）

            da_pct JSONB,            -- {"Cd": 12.3, ...}
            er JSONB,                -- {"Cd": 0.8, ...}

            created_at TIMESTAMP DEFAULT NOW()
        );
        """))

        # samples 防重：同一 lot_no + year 只允許一筆（你也可以改成允許多筆）
        conn.execute(text("""
        DO $$
        BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM pg_indexes
                WHERE indexname = 'idx_samples_lot_year_unique'
            ) THEN
                CREATE UNIQUE INDEX idx_samples_lot_year_unique ON samples(lot_no, year);
            END IF;
        END $$;
        """))


# =========================
# 小工具：讀 Excel
# =========================
@st.cache_data(show_spinner=False)
def load_excel(path: str):
    if not os.path.exists(path):
        return None
    xls = pd.ExcelFile(path)
    sheets = xls.sheet_names
    data = {}
    for s in sheets:
        data[s] = xls.parse(s)
    return data


def roc_today_str():
    now = datetime.now(TZ_TW)
    roc_year = now.year - 1911
    return f"民國 {roc_year} 年 {now.month} 月 {now.day} 日"


def normalize_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def to_bool_rep(x: str) -> bool:
    s = normalize_str(x)
    return s in ["是", "代表", "True", "1", "Y", "y", "YES", "Yes"]


# =========================
# 統計：首頁 KPI
# =========================
def fetch_kpis(engine):
    if engine is None:
        return {
            "total": 0,
            "sample_points": 0,
            "control": 0,
            "building": 0,
            "hard": 0,
            "normal_exit": 0,
        }

    with engine.begin() as conn:
        total = conn.execute(text("SELECT COUNT(*) FROM lands;")).scalar() or 0

        # 採樣點數：rep_role 非空（代表/備用都算）
        sample_points = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE COALESCE(NULLIF(TRIM(rep_role), ''), '') <> '';
        """)).scalar() or 0

        control = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE admin_status = '管制' OR current_metal_result = '管制';
        """)).scalar() or 0

        building = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE admin_status = '建物' OR current_metal_result = '建物';
        """)).scalar() or 0

        hard = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE admin_status IN ('難以採樣', '無法採樣') OR current_metal_result IN ('難以採樣', '無法採樣');
        """)).scalar() or 0

        normal_exit = conn.execute(text("""
            SELECT COUNT(*) FROM lands
            WHERE admin_status = '正常' OR current_metal_result = '正常';
        """)).scalar() or 0

    return {
        "total": total,
        "sample_points": sample_points,
        "control": control,
        "building": building,
        "hard": hard,
        "normal_exit": normal_exit,
    }


# =========================
# 管理員：登入/驗證
# =========================
def is_admin():
    return st.session_state.get("is_admin", False)


def admin_login_box():
    st.sidebar.markdown("### 🔐 管理員登入")
    if is_admin():
        st.sidebar.success("已登入管理員")
        if st.sidebar.button("登出"):
            st.session_state["is_admin"] = False
            st.rerun()
        return

    pwd = st.sidebar.text_input("管理員密碼", type="password")
    if st.sidebar.button("登入"):
        admin_pwd = st.secrets.get("ADMIN_PASSWORD", "")
        if not admin_pwd:
            st.sidebar.error("Secrets 尚未設定 ADMIN_PASSWORD")
            return
        if pwd == admin_pwd:
            st.session_state["is_admin"] = True
            st.sidebar.success("登入成功")
            st.rerun()
        else:
            st.sidebar.error("密碼錯誤")


# =========================
# Excel -> DB 匯入（核心）
# =========================
def admin_import_excel_to_db(engine):
    st.subheader("🛠 管理員工具：Excel → DB 匯入")
    st.caption("此工具會把 repo 裡的 Excel 主檔、標準表、同坵塊表匯入 DB。")

    if engine is None:
        st.error("DATABASE_URL 未設定，無法連線 DB。請先到 Streamlit Secrets 設定 DATABASE_URL。")
        return

    if not os.path.exists(EXCEL_PATH):
        st.error(f"找不到 Excel：{EXCEL_PATH}（請確認檔案在 repo 根目錄）")
        return

    if not is_admin():
        st.info("請先在側邊欄登入管理員")
        return

    if st.button("🚀 一鍵匯入（lands / standards / blocks）"):
        with st.spinner("正在匯入..."):
            data = load_excel(EXCEL_PATH)
            if data is None:
                st.error("Excel 讀取失敗")
                return

            # ---------- 讀三張 ----------
            if SHEET_MASTER not in data:
                st.error(f"Excel 缺少分頁：{SHEET_MASTER}")
                return
            if SHEET_STANDARDS not in data:
                st.error(f"Excel 缺少分頁：{SHEET_STANDARDS}")
                return
            if SHEET_BLOCKS not in data:
                st.error(f"Excel 缺少分頁：{SHEET_BLOCKS}")
                return

            df_master = data[SHEET_MASTER].copy()
            df_std = data[SHEET_STANDARDS].copy()
            df_blk = data[SHEET_BLOCKS].copy()

            # ---------- 匯入 standards ----------
            df_std.columns = [normalize_str(c) for c in df_std.columns]
            col_item = "項目名稱"
            col_m = "監測標準"
            col_c = "管制標準"
            col_da = "上升標準 (DA門檻)"

            with engine.begin() as conn:
                conn.execute(text("DELETE FROM standards;"))

                for _, r in df_std.iterrows():
                    item = normalize_str(r.get(col_item))
                    if not item:
                        continue
                    monitor_std = r.get(col_m)
                    control_std = r.get(col_c)
                    da_th = r.get(col_da)

                    conn.execute(text("""
                        INSERT INTO standards (item, monitor_std, control_std, da_threshold)
                        VALUES (:item, :m, :c, :da)
                        ON CONFLICT (item) DO UPDATE
                        SET monitor_std = EXCLUDED.monitor_std,
                            control_std = EXCLUDED.control_std,
                            da_threshold = EXCLUDED.da_threshold;
                    """), {
                        "item": item,
                        "m": None if pd.isna(monitor_std) else float(monitor_std),
                        "c": None if pd.isna(control_std) else float(control_std),
                        "da": None if pd.isna(da_th) else float(da_th),
                    })

            # ---------- 匯入 blocks（含去重 + upsert） ----------
            df_blk.columns = [normalize_str(c) for c in df_blk.columns]
            gcol = "農地群組編號"
            lcol = "農地地段地號"
            rcol = "代表農地"

            # 清理
            df_blk[gcol] = df_blk[gcol].astype(str).str.strip()
            df_blk[lcol] = df_blk[lcol].astype(str).str.strip()

            # 去除空 lot_no / block_id
            df_blk = df_blk[(df_blk[gcol] != "") & (df_blk[lcol] != "")]

            # 去重（同 block+lot 留最後一筆）
            df_blk2 = df_blk.drop_duplicates(subset=[gcol, lcol], keep="last")

            with engine.begin() as conn:
                conn.execute(text("DELETE FROM blocks;"))
                for _, r in df_blk2.iterrows():
                    block_id = normalize_str(r.get(gcol))
                    lot_no = normalize_str(r.get(lcol))
                    is_rep_bool = to_bool_rep(r.get(rcol))

                    conn.execute(text("""
                        INSERT INTO blocks (block_id, lot_no, is_rep)
                        VALUES (:block_id, :lot_no, :is_rep)
                        ON CONFLICT (block_id, lot_no) DO UPDATE
                        SET is_rep = EXCLUDED.is_rep;
                    """), {
                        "block_id": block_id,
                        "lot_no": lot_no,
                        "is_rep": is_rep_bool
                    })

            # ---------- 匯入 lands ----------
            df_master.columns = [normalize_str(c) for c in df_master.columns]

            # 你的欄位（依你描述）
            # 若 Excel 欄名有小差異，你可以在這裡調整對應
            col_lot = "地段地號"
            col_sgm = "SGM編號"
            col_land_serial = "農地序號"
            col_grid = "網格編號"
            col_town = "鄉鎮市"
            col_section = "地段地號"  # 若你有地段欄就改掉
            col_method = "調查方式"
            col_rep = "代表性"
            col_water = "用水種類"
            col_x = "TWD97_X"
            col_y = "TWD97_Y"

            col_current = "目前農地調查現況"
            col_admin = "農地監測狀態"
            col_freq = "網格監測頻率"
            col_last_year = "最後調查年分"

            # 初始金屬欄位（依你主檔）
            metal_cols = {
                "Hg": "初始_汞",
                "As": "初始_砷",
                "Cr": "初始_鉻",
                "Cd": "初始_鎘",
                "Pb": "初始_鉛",
                "Zn": "初始_鋅",
                "Ni": "初始_鎳"
                # 你主檔若有「銅」可再加
            }

            # 年度狀態欄（可能是 101狀態～114狀態）
            year_cols = [c for c in df_master.columns if c.endswith("狀態")]
            # 只保留像「101狀態」這種
            year_cols = [c for c in year_cols if normalize_str(c).replace("狀態", "").isdigit()]

            # lot_no 清理
            df_master[col_lot] = df_master[col_lot].astype(str).str.strip()
            df_master = df_master[df_master[col_lot] != ""]

            # 去重 lot_no（同地段地號留最後一筆）
            df_master = df_master.drop_duplicates(subset=[col_lot], keep="last")

            with engine.begin() as conn:
                # 不先 DELETE lands，改用 upsert，避免誤刪
                for _, r in df_master.iterrows():
                    lot_no = normalize_str(r.get(col_lot))
                    if not lot_no:
                        continue

                    initial_metals = {}
                    for k, c in metal_cols.items():
                        v = r.get(c)
                        if pd.isna(v):
                            continue
                        try:
                            initial_metals[k] = float(v)
                        except Exception:
                            pass

                    year_status = {}
                    for yc in year_cols:
                        y = normalize_str(yc).replace("狀態", "")
                        val = normalize_str(r.get(yc))
                        if val:
                            year_status[y] = val

                    payload = {
                        "lot_no": lot_no,
                        "sgm_no": normalize_str(r.get(col_sgm)),
                        "land_serial": normalize_str(r.get(col_land_serial)),
                        "grid_id": normalize_str(r.get(col_grid)),
                        "township": normalize_str(r.get(col_town)),
                        "section_no": "",  # 若你有獨立地段欄可改
                        "survey_method": normalize_str(r.get(col_method)),
                        "rep_role": normalize_str(r.get(col_rep)),
                        "water_type": normalize_str(r.get(col_water)),
                        "coord_x": None if pd.isna(r.get(col_x)) else float(r.get(col_x)),
                        "coord_y": None if pd.isna(r.get(col_y)) else float(r.get(col_y)),
                        "initial_metals": json.dumps(initial_metals, ensure_ascii=False),
                        "current_metal_result": normalize_str(r.get(col_current)),
                        "admin_status": normalize_str(r.get(col_admin)),
                        "freq": normalize_str(r.get(col_freq)),
                        "last_year": None if pd.isna(r.get(col_last_year)) else int(float(r.get(col_last_year))),
                        "year_status": json.dumps(year_status, ensure_ascii=False),
                    }

                    conn.execute(text("""
                        INSERT INTO lands (
                            lot_no, sgm_no, land_serial, grid_id, township, section_no, survey_method,
                            rep_role, water_type, coord_x, coord_y,
                            initial_metals, current_metal_result, admin_status, freq, last_year, year_status, updated_at
                        )
                        VALUES (
                            :lot_no, :sgm_no, :land_serial, :grid_id, :township, :section_no, :survey_method,
                            :rep_role, :water_type, :coord_x, :coord_y,
                            CAST(:initial_metals AS JSONB), :current_metal_result, :admin_status, :freq, :last_year,
                            CAST(:year_status AS JSONB), NOW()
                        )
                        ON CONFLICT (lot_no) DO UPDATE SET
                            sgm_no = EXCLUDED.sgm_no,
                            land_serial = EXCLUDED.land_serial,
                            grid_id = EXCLUDED.grid_id,
                            township = EXCLUDED.township,
                            section_no = EXCLUDED.section_no,
                            survey_method = EXCLUDED.survey_method,
                            rep_role = EXCLUDED.rep_role,
                            water_type = EXCLUDED.water_type,
                            coord_x = EXCLUDED.coord_x,
                            coord_y = EXCLUDED.coord_y,
                            initial_metals = EXCLUDED.initial_metals,
                            current_metal_result = EXCLUDED.current_metal_result,
                            admin_status = EXCLUDED.admin_status,
                            freq = EXCLUDED.freq,
                            last_year = EXCLUDED.last_year,
                            year_status = EXCLUDED.year_status,
                            updated_at = NOW();
                    """), payload)

        st.success("✅ 匯入完成！請回到首頁確認 KPI 是否更新（不再是 0）。")


# =========================
# 總表清單 / 查詢
# =========================
def page_master_list(engine):
    st.subheader("📋 總表清單")

    if engine is None:
        st.error("DATABASE_URL 未設定")
        return

    with engine.begin() as conn:
        df = pd.read_sql(text("""
            SELECT
                lot_no AS 地段地號,
                sgm_no AS SGM編號,
                grid_id AS 網格編號,
                township AS 鄉鎮市,
                rep_role AS 代表性,
                water_type AS 用水種類,
                admin_status AS 農地監測狀態,
                current_metal_result AS 目前農地調查現況,
                freq AS 網格監測頻率,
                last_year AS 最後調查年分
            FROM lands
            ORDER BY grid_id, lot_no
            LIMIT 5000;
        """), conn)

    st.dataframe(df, use_container_width=True)


def page_search(engine):
    st.subheader("🔎 資料查詢（SGM 或 地段地號）")

    if engine is None:
        st.error("DATABASE_URL 未設定")
        return

    q = st.text_input("請輸入 SGM 編號或地段地號（例：華南段0159-0000）")
    if not q:
        st.info("輸入後即可查詢")
        return

    q = q.strip()

    with engine.begin() as conn:
        df = pd.read_sql(text("""
            SELECT *
            FROM lands
            WHERE lot_no ILIKE :q OR sgm_no ILIKE :q
            LIMIT 50;
        """), conn, params={"q": f"%{q}%"})

    if df.empty:
        st.warning("查無資料")
        return

    st.dataframe(df, use_container_width=True)

    # 顯示第一筆詳細卡
    row = df.iloc[0].to_dict()
    st.markdown("### 🧾 資訊卡（第一筆）")
    c1, c2, c3 = st.columns(3)
    c1.write({"地段地號": row.get("lot_no"), "SGM": row.get("sgm_no"), "網格": row.get("grid_id")})
    c2.write({"狀態": row.get("admin_status"), "現況": row.get("current_metal_result"), "頻率": row.get("freq")})
    c3.write({"座標X": row.get("coord_x"), "座標Y": row.get("coord_y"), "最後年分": row.get("last_year")})


# =========================
# 主頁 Dashboard
# =========================
def page_dashboard(engine):
    st.title("🚜 彰化縣農地監測戰情室")
    st.markdown(f"### 🗓️ 當前時間：{roc_today_str()}")

    kpi = fetch_kpis(engine)

    cols = st.columns(6)
    cols[0].metric("總資料點數", kpi["total"])
    cols[1].metric("總採樣點數(代表+備用)", kpi["sample_points"])
    cols[2].metric("管制點數", kpi["control"])
    cols[3].metric("建物數量", kpi["building"])
    cols[4].metric("難以採樣數量", kpi["hard"])
    cols[5].metric("正常退場數量", kpi["normal_exit"])

    st.divider()
    admin_import_excel_to_db(engine)


# =========================
# App Router
# =========================
def main():
    engine = get_engine()
    if engine is None:
        st.warning("⚠️ 尚未設定 DATABASE_URL。請到 Streamlit Secrets 設定 DATABASE_URL 與 ADMIN_PASSWORD。")
    else:
        init_db(engine)

    admin_login_box()

    st.sidebar.markdown("### 📌 功能選單")
    page = st.sidebar.radio("前往頁面", [
        "首頁 Dashboard",
        "總表清單",
        "資料查詢"
    ])

    if page == "首頁 Dashboard":
        page_dashboard(engine)
    elif page == "總表清單":
        page_master_list(engine)
    elif page == "資料查詢":
        page_search(engine)


if __name__ == "__main__":
    main()




