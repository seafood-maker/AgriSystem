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

# --- 1. 系統設定 ---
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

# --- 2. 資料讀取引擎 ---
@st.cache_data
def load_all_data():
    if not os.path.exists(EXCEL_PATH): return None, None, None, None
    try:
        xl = pd.ExcelFile(EXCEL_PATH)
        actual_sheets = xl.sheet_names
        def get_sheet(name, sheets):
            for s in sheets:
                if name == s.strip(): return s
            return None
        
        s1 = get_sheet("農地現況主檔", actual_sheets)
        s2 = get_sheet("歷年調查紀錄", actual_sheets)
        s3 = get_sheet("同坵塊對照表", actual_sheets)
        s4 = get_sheet("判定標準表", actual_sheets)

        if not s1: return None, None, None, None

        df_m = pd.read_excel(xl, sheet_name=s1)
        df_m.columns = df_m.columns.str.strip()
        df_m['網格編號'] = df_m['網格編號'].apply(clean_id)
        df_m['調查方式'] = df_m['調查方式'].fillna('未知').astype(str).str.strip()
        df_m['目前農地調查現況'] = df_m['目前農地調查現況'].fillna('未知').astype(str).str.strip()
        
        df_h = pd.read_excel(xl, sheet_name=s2) if s2 else pd.DataFrame()
        df_b = pd.read_excel(xl, sheet_name=s3) if s3 else pd.DataFrame()
        df_s = pd.read_excel(xl, sheet_name=s4).set_index('項目名稱') if s4 else pd.DataFrame()
        
        return df_m, df_h, df_b, df_s
    except Exception as e:
        st.error(f"Excel 讀取錯誤: {e}")
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
st.sidebar.title("🌿 系統導覽")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新增年度調查結果", "空間地圖檢視"])

if df_master is not None:
    # --- 全域統計數據計算 ---
    abs_total = len(df_master)
    sampling_pts = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])

    # --- A. 統計首頁 ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")

        # 1. 頂列統計指標 (水平一列)
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("總資料點數", abs_total)
        m2.metric("總採樣點數", sampling_pts)
        m3.metric("管制點數", control_count)
        m4.metric("建物數量", build_count)
        m5.metric("難以採樣數量", hard_count)
        m6.metric("正常退場數量", normal_count)

        st.divider()

        # 2. 系統型網格現況統計 (水平一列)
        st.subheader("🌐 系統型網格現況統計")
        grid_df = df_master.drop_duplicates('網格編號').copy()
        grid_df['網格監測頻率'] = grid_df['網格監測頻率'].fillna('無網格狀態').astype(str).str.strip()
        g_cont = len(grid_df[grid_df['網格監測頻率'] == '持續'])
        g_ext = len(grid_df[grid_df['網格監測頻率'] == '延長'])
        g_exit = len(grid_df[grid_df['網格監測頻率'] == '退場'])
        g_total_active = g_cont + g_ext + g_exit
        g_none = len(grid_df) - g_total_active

        g_cols = st.columns(5)
        g_cols[0].metric("持續網格", g_cont)
        g_cols[1].metric("延長網格", g_ext)
        g_cols[2].metric("退場網格", g_exit)
        g_cols[3].metric("有效網格合計", g_total_active)
        g_cols[4].metric("無網格狀態", g_none)

        st.divider()

        # 3. 個案型農地統計 (修正階梯式排版：水平一列)
        st.subheader("📦 個案型農地現況統計")
        # 排除系統型即為個案型
        case_data = df_master[~df_master['調查方式'].astype(str).str.contains('系統', na=False)].copy()
        c_map = {"增量":"持續", "延長":"延長", "正常":"退場", "難以採樣":"難以採樣", "管制":"管制", "建物":"建物"}
        case_data['顯示狀態'] = case_data['目前農地調查現況'].map(c_map).fillna(case_data['目前農地調查現況'])
        c_counts = case_data['顯示狀態'].value_counts()
        
        # 關鍵修正：先宣告 columns 容器，再依序填入
        c_cols = st.columns(6)
        c_labels = ["持續", "延長", "退場", "管制", "難以採樣", "建物"]
        for i, lab in enumerate(c_labels):
            c_cols[i].metric(lab, c_counts.get(lab, 0))

        st.divider()

        # 4. 網格查詢系統
        st.subheader("🔍 網格查詢系統")
        grid_search = st.text_input("輸入網格號碼 (例如: G001)")
        if grid_search:
            grid_res = df_master[df_master['網格編號'] == clean_id(grid_search)]
            if not grid_res.empty:
                def highlight(row):
                    return ['background-color: #FFFFCC' if row.代表性 == '代表點' else '' for _ in row]
                st.dataframe(grid_res.style.apply(highlight, axis=1), use_container_width=True)
            else: st.warning("查無此網格編號")

        st.divider()

        # 5. 近三年樹狀圖
        st.subheader("📊 近三年調查分佈 (樹狀圖)")
        years = [112, 113, 114]
        tree_cols = st.columns(3)
        for i, y in enumerate(years):
            col_name = f"{y}狀態"
            if col_name in df_master.columns:
                y_df = df_master[df_master[col_name].notna()].copy()
                if not y_df.empty:
                    def map_tree_label(row):
                        status = str(row[col_name]).strip()
                        current_inv = str(row['目前農地調查現況']).strip()
                        if status == "監測":
                            if "增量" in current_inv: return "持續"
                            if "延長" in current_inv: return "延長"
                            return "持續"
                        elif status == "正常": return "退場"
                        else: return status
                    y_df['最終判定'] = y_df.apply(map_tree_label, axis=1)
                    y_counts = y_df.groupby(['調查方式', '最終判定']).size().reset_index(name='筆數')
                    y_counts['根'] = f"{y}年 (共{len(y_df)}筆)"
                    fig = px.treemap(y_counts, path=['根', '調查方式', '最終判定'], values='筆數', color='最終判定',
                                     color_discrete_map={'持續':'#FFB6C1','延長':'#ADD8E6','退場':'#90EE90','管制':'#FF3333','建物':'#D3D3D3'})
                    tree_cols[i].plotly_chart(fig, use_container_width=True)
                else: tree_cols[i].info(f"{y} 年無調查數據")

    # --- B. 資料庫查詢與下載 ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 資料庫查詢與下載中心")
        t1, t2, t3 = st.tabs(["總表清單", "歷年調查結果", "同坵塊關聯"])
        with t1:
            st.dataframe(df_master, use_container_width=True)
            towrite = io.BytesIO()
            df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載全量總表", data=towrite.getvalue(), file_name="農地總表.xlsx")
        with t2:
            if not df_history.empty:
                y_sel = st.selectbox("選擇下載年度", sorted(df_history['調查年度'].unique()))
                st.dataframe(df_history[df_history['調查年度']==y_sel])
        with t3:
            s_lot = st.text_input("輸入地段地號搜尋同群組")
            if s_lot and not df_block.empty:
                match = df_block[df_block['農地地段地號'] == s_lot.strip()]
                if not match.empty:
                    gid = match.iloc[0]['農地群組編號']
                    rel = df_block[df_block['農地群組編號'] == gid]['農地地段地號'].tolist()
                    st.dataframe(df_master[df_master['地段地號'].isin(rel)])

    # --- C. 新增年度調查結果 ---
    elif menu == "新增年度調查結果":
        st.title("➕ 新增調查紀錄 (管理員後台)")
        pwd = st.sidebar.text_input("密碼", type="password")
        if pwd == ADMIN_PASSWORD:
            st.success("✅ 通過驗證")
            # 此處保留您之前的錄入與判定邏輯代碼...
        else: st.warning("請於左側輸入密碼以解鎖編輯權限")

    # --- D. 空間地圖檢視 ---
    elif menu == "空間地圖檢視":
        st.title("🗺️ 空間地圖檢視")
        # 建立衛星地圖
        m = folium.Map(location=[24.05, 120.5], zoom_start=11, 
                       tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', 
                       attr='Esri')
        if gdf_grid is not None:
            grid_status = df_master.drop_duplicates('網格編號')[['網格編號', '網格監測頻率']]
            merged = gdf_grid.to_crs(epsg=4326).merge(grid_status, left_on='網格號', right_on='網格編號', how='left')
            def get_c(f):
                f = str(f)
                return '#FFB6C1' if '持續' in f else '#ADD8E6' if '延長' in f else '#90EE90' if '退場' in f else '#F8F8F8'
            folium.GeoJson(merged, style_function=lambda x: {
                'fillColor': get_c(x['properties'].get('網格監測頻率')),
                'color': 'white', 'weight': 1, 'fillOpacity': 0.4
            }).add_to(m)
        
        # 繪製點位
        sample = df_master.sample(min(800, len(df_master)))
        for _, r in sample.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                s = 4 if "個案" in str(r['調查方式']) else 3
                mon_s = str(r['農地監測狀態'])
                if mon_s == "管制": s = 6; c = "red"
                elif mon_s == "建物": s = 6; c = "black"
                elif mon_s == "難以採樣": s = 6; c = "purple"
                elif str(r['代表性']) == "備用點": s = 4; c = "white"
                else:
                    inv = str(r['目前農地調查現況'])
                    c = "red" if "增量" in inv else "blue" if "延長" in inv else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=s, radius=6, color=c, fill=True, popup=f"{r['地段地號']}").add_to(m)
            except: continue
        st_folium(m, width=1100, height=700)
else:
    st.error("❌ 讀取 Excel 失敗")




