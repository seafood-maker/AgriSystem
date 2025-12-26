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
from datetime import datetime
import plotly.express as px
import io

# --- 1. 系統權限與基礎設定 ---
ADMIN_PASSWORD = "ET23597010"
st.set_page_config(page_title="彰化農地智慧管理系統", layout="wide", page_icon="🌾")

EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"

transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)

def clean_id(val):
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

def get_minguo_date():
    now = datetime.now()
    return f"民國 {now.year - 1911} 年 {now.month} 月 {now.day} 日"

# --- 2. 資料讀取引擎 (具備自動偵測分頁功能) ---
@st.cache_data
def load_all_data():
    if not os.path.exists(EXCEL_PATH): 
        st.error(f"❌ 找不到 Excel 檔案：{EXCEL_PATH}")
        return None, None, None, None
    try:
        xl = pd.ExcelFile(EXCEL_PATH)
        actual_sheets = xl.sheet_names
        
        # 定義一個輔助函式來尋找最接近的分頁名稱 (忽略空格)
        def get_best_sheet(target, sheets):
            for s in sheets:
                if target == s.strip():
                    return s
            return None

        s1 = get_best_sheet("農地現況主檔", actual_sheets)
        s2 = get_best_sheet("歷年調查紀錄", actual_sheets)
        s3 = get_best_sheet("同坵塊對照表", actual_sheets)
        s4 = get_best_sheet("判定標準表", actual_sheets)

        if not s1:
            st.error(f"❌ Excel 讀取錯誤：找不到『農地現況主檔』分頁。目前偵測到的分頁有：{actual_sheets}")
            return None, None, None, None

        df_m = pd.read_excel(xl, sheet_name=s1)
        df_m.columns = df_m.columns.str.strip()
        df_m['網格編號'] = df_m['網格編號'].apply(clean_id)
        
        df_h = pd.read_excel(xl, sheet_name=s2) if s2 else pd.DataFrame()
        df_b = pd.read_excel(xl, sheet_name=s3) if s3 else pd.DataFrame()
        df_s = pd.read_excel(xl, sheet_name=s4).set_index('項目名稱') if s4 else pd.DataFrame()
        
        return df_m, df_h, df_b, df_s
    except Exception as e:
        st.error(f"❌ Excel 讀取發生嚴重錯誤: {e}")
        return None, None, None, None

@st.cache_data
def load_grid_shp():
    if os.path.exists(SHP_PATH):
        try:
            gdf = None
            for enc in ['utf-8', 'cp950', 'big5']:
                try: gdf = gpd.read_file(SHP_PATH, encoding=enc); break
                except: continue
            if gdf is not None:
                if gdf.crs is None or gdf.crs.to_epsg() != 3826:
                    gdf.set_crs(epsg=3826, allow_override=True, inplace=True)
                if '網格號' in gdf.columns: gdf['網格號'] = gdf['網格號'].apply(clean_id)
                return gdf
        except: return None
    return None

df_master, df_history, df_block, df_settings = load_all_data()
gdf_grid = load_grid_shp()

# --- 3. 側邊欄 ---
st.sidebar.title("🌿 系統選單")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新增年度調查結果", "空間地圖檢視"])

# --- 4. 主程式邏輯 ---
if df_master is not None:
    # 全域數據統計 (2453筆基準)
    abs_total = len(df_master)
    sampling_pts = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])

    # --- A. 統計首頁 ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.info(f"📅 當前時間：{get_minguo_date()}")

        # 1. 頂列統計指標
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("總資料點數", abs_total)
        m2.metric("總採樣點數", sampling_pts)
        m3.metric("管制點數", control_count)
        m4.metric("建物數量", build_count)
        m5.metric("難以採樣數量", hard_count)
        m6.metric("正常退場數量", normal_count)

        st.divider()

        # 2. 系統型網格統計
        st.subheader("🌐 系統型網格現況統計")
        grid_df = df_master.drop_duplicates('網格編號').copy()
        grid_df['網格監測頻率'] = grid_df['網格監測頻率'].fillna('無網格狀態').astype(str).str.strip()
        g_cont = len(grid_df[grid_df['網格監測頻率'] == '持續'])
        g_ext = len(grid_df[grid_df['網格監測頻率'] == '延長'])
        g_exited = len(grid_df[grid_df['網格監測頻率'] == '退場'])
        g_total_active = g_cont + g_ext + g_exited
        g_none = len(grid_df) - g_total_active

        g_cols = st.columns(5)
        g_cols[0].metric("持續網格", g_cont); g_cols[1].metric("延長網格", g_ext)
        g_cols[2].metric("退場網格", g_exited); g_cols[3].metric("有效網格合計", g_total_active)
        g_cols[4].metric("無網格狀態", g_none)

        st.divider()

        # 3. 個案型農地統計 (修正：排除系統型即個案型，確保抓到 16筆建物與6筆難以採樣)
        st.subheader("📦 個案型農地現況統計")
        case_data = df_master[~df_master['調查方式'].astype(str).str.contains('系統', na=False)].copy()
        case_data['目前農地調查現況'] = case_data['目前農地調查現況'].fillna('未知').astype(str).str.strip()
        
        c_stats = {
            "持續": len(case_data[case_data['目前農地調查現況'] == '增量']),
            "延長": len(case_data[case_data['目前農地調查現況'] == '延長']),
            "退場": len(case_data[case_data['目前農地調查現況'] == '正常']),
            "管制": len(case_data[case_data['目前農地調查現況'] == '管制']),
            "難以採樣": len(case_data[case_data['目前農地調查現況'] == '難以採樣']),
            "建物": len(case_data[case_data['目前農地調查現況'] == '建物'])
        }

        c_cols = st.columns(6)
        c_cols[0].metric("持續", c_stats["持續"]); c_cols[1].metric("延長", c_stats["延長"])
        c_cols[2].metric("退場", c_stats["退場"]); c_cols[3].metric("管制", c_stats["管制"])
        c_cols[4].metric("難以採樣", c_stats["難以採樣"]); c_cols[5].metric("建物", c_stats["建物"])

        st.divider()
        # --- 網格查詢與樹狀圖 (保留 8.1 版邏輯) ---
        # ... (後續代碼與 8.1 一致)






