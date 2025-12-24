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

# --- 1. 系統權限與基礎設定 ---
ADMIN_PASSWORD = "admin_changhua"  # 這是您的後台管理密碼
st.set_page_config(page_title="彰化農地智慧管理系統", layout="wide", page_icon="🌾")

EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"

# 座標轉換器 (TWD97 -> WGS84)
transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ['汞', '砷', '銅', '鉻', '鎘', '鉛', '鋅', '鎳']

# 數據清洗工具
def clean_id(val):
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

# --- 2. 資料讀取引擎 ---
@st.cache_data
def load_all_data():
    if not os.path.exists(EXCEL_PATH): return None, None, None, None
    try:
        xl = pd.ExcelFile(EXCEL_PATH)
        df_m = pd.read_excel(xl, "農地現況主檔")
        df_m.columns = df_m.columns.str.strip()
        df_m['網格編號'] = df_m['網格編號'].apply(clean_id)
        
        df_h = pd.read_excel(xl, "歷年調查紀錄")
        df_h.columns = df_h.columns.str.strip()
        df_b = pd.read_excel(xl, "同坵塊對照表")
        df_b.columns = df_b.columns.str.strip()
        df_s = pd.read_excel(xl, "判定標準表").set_index('項目名稱')
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
                try:
                    gdf = gpd.read_file(SHP_PATH, encoding=enc)
                    break
                except: continue
            if gdf is not None:
                if gdf.crs is None or gdf.crs.to_epsg() != 3826:
                    gdf.set_crs(epsg=3826, allow_override=True, inplace=True)
                if '網格號' in gdf.columns:
                    gdf['網格號'] = gdf['網格號'].apply(clean_id)
                if '狀態' in gdf.columns:
                    gdf['狀態'] = gdf['狀態'].apply(clean_id)
                return gdf
        except Exception as e:
            st.error(f"Shapefile 讀取失敗: {e}")
    return None

df_master, df_history, df_block, df_settings = load_all_data()
gdf_grid = load_grid_shp()

# --- 3. 空間運算：判斷新座標落在哪個網格 ---
def find_grid_by_coords(x, y, gdf):
    if gdf is None: return "未知"
    p = Point(x, y)
    match = gdf[gdf.contains(p)]
    if not match.empty:
        return str(match.iloc[0]['網格號'])
    return "範圍外"

# --- 4. 側邊欄與導覽 ---
st.sidebar.title("🌿 彰化農地管理系統")
menu = st.sidebar.radio("請選擇功能：", ["📊 統計首頁", "📋 總表清單", "➕ 新增年度調查", "🔍 歷年數據查詢", "🗺️ 空間地圖檢視"])

if df_master is not None:
    # --- 全域統計數值計算 (防止報錯) ---
    abs_total = len(df_master)
    sampling_pts = df_master[df_master['代表性'].isin(['代表點', '備用點'])]
    control_count = len(df_master[df_master['農地監測狀態'] == '管制'])
    build_count = len(df_master[df_master['農地監測狀態'] == '建物'])
    hard_count = len(df_master[df_master['農地監測狀態'] == '難以採樣'])
    normal_count = len(df_master[df_master['農地監測狀態'] == '正常'])

    # --- A. 統計首頁 ---
    if menu == "📊 統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("總資料筆數", abs_total)
        k2.metric("總採樣點", len(sampling_pts))
        k3.metric("管制點位", control_count)
        k4.metric("建物數量", build_count)
        k5.metric("難以採樣", hard_count)
        k6.metric("正常退場", normal_count)
        
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("🌐 系統型網格現況 (Excel)")
            grid_summary = df_master.drop_duplicates('網格編號')
            st.write(f"總網格數: {len(grid_summary)}")
            st.dataframe(grid_summary['網格監測頻率'].value_counts())
        with c2:
            st.subheader("📦 個案型農地統計")
            st.dataframe(df_master[df_master['調查方式'] == '個案型農地']['農地監測狀態'].value_counts())

    # --- B. 總表清單 ---
    elif menu == "📋 總表清單":
        st.header("📋 農地現況總表 (全量資料)")
        search_q = st.text_input("輸入地段地號或 SGM 編號快速搜尋")
        df_show = df_master
        if search_q:
            df_show = df_master[df_master.astype(str).apply(lambda x: x.str.contains(search_q)).any(axis=1)]
        st.dataframe(df_show, use_container_width=True)

    # --- C. 新增年度調查 (加上密碼保護) ---
    elif menu == "➕ 新增年度調查":
        st.header("➕ 新增年度調查與 DA 判定")
        pwd = st.sidebar.text_input("請輸入管理員密碼", type="password")
        
        if pwd == ADMIN_PASSWORD:
            st.success("身分驗證成功，您可以開始錄入數據。")
            search_id = st.text_input("🔍 第一步：請輸入地段地號搜尋 (例: 華南段0159-0000)")
            
            if search_id:
                targets = df_master[df_master['地段地號'] == search_id]
                if not targets.empty:
                    # 如果有多筆點位在同地號，讓使用者選一筆
                    selected_sgm = st.selectbox("確認 SGM 編號", targets['SGM編號'].unique())
                    row = targets[targets['SGM編號'] == selected_sgm].iloc[0]
                    
                    with st.form("survey_entry"):
                        st.subheader(f"📍 錄入對象：{search_id} ({selected_sgm})")
                        col_y, col_admin = st.columns(2)
                        target_year = col_y.selectbox("調查年度", [114, 115, 116])
                        admin_status = col_admin.selectbox("行政狀態", ["監測", "建物", "管制", "難以採樣", "正常"])
                        
                        st.write("---")
                        st.write("🧪 重金屬檢測值輸入 (DA 判定優先取全量)")
                        current_vals = {}
                        for m in METALS:
                            c1, c2 = st.columns(2)
                            xrf = c1.number_input(f"{m} (XRF)", min_value=0.0, key=f"xrf_{m}")
                            tot = c2.number_input(f"{m} (全量)", min_value=0.0, key=f"tot_{m}")
                            current_vals[m] = tot if tot > 0 else xrf
                        
                        if st.form_submit_button("執行判定"):
                            st.subheader("💡 判定結果預覽")
                            da_results = {}
                            final_invest = "正常"
                            
                            for m in METALS:
                                init = row.get(f'初始_{m}', 0)
                                s2 = df_settings.loc[m, '管制標準']
                                da_limit = df_settings.loc[m, '上升標準 (DA門檻)']
                                curr = current_vals[m]
                                
                                # DA = (本次 - 初始) / 初始 * 100
                                da_val = ((curr - init) / init * 100) if init > 0 else 0
                                da_results[m] = f"{da_val:.1f}%"
                                
                                # 核心邏輯
                                if curr > s2: final_invest = "管制"
                                elif init > s2 and da_val > da_limit:
                                    if final_invest != "管制": final_invest = "增量"
                                elif curr > df_settings.loc[m, '監測標準']:
                                    if final_invest not in ["管制", "增量"]: final_invest = "延長"
                            
                            st.write(f"最新判定狀態：**{final_invest}**")
                            st.json(da_results)
                            st.info("💡 提醒：在雲端版本中，請將結果記錄後更新至您的 Excel 檔案並重新上傳至 GitHub 以永久保存。")
                else:
                    st.error("找不到該地號，請確認格式。")
        else:
            st.warning("請在左側選單下方輸入密碼以解鎖新增功能。")

    # --- D. 空間地圖檢視 ---
    elif menu == "🗺️ 空間地圖檢視":
        st.header("🗺️ 衛星影像監測圖 (網格頻率著色)")
        
        m = folium.Map(
            location=[24.05, 120.5], zoom_start=11,
            tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
            attr='Esri World Imagery'
        )
        
        if gdf_grid is not None:
            # 準備對接資料
            grid_status = df_master.drop_duplicates('網格編號')[['網格編號', '網格監測頻率']]
            gdf_wgs84 = gdf_grid.to_crs(epsg=4326)
            merged = gdf_wgs84.merge(grid_status, left_on='網格號', right_on='網格編號', how='left')

            def get_color(f):
                f = str(f)
                if '持續' in f: return '#FFB6C1'
                if '延長' in f: return '#ADD8E6'
                if '退場' in f: return '#90EE90'
                return '#F8F8F8'

            folium.GeoJson(
                merged,
                style_function=lambda x: {
                    'fillColor': get_color(x['properties'].get('網格監測頻率')),
                    'color': 'white', 'weight': 1, 'fillOpacity': 0.4
                },
                tooltip=folium.GeoJsonTooltip(fields=['網格號', '網格監測頻率'])
            ).add_to(m)

        # 繪製點位
        sample_df = df_master.sample(min(800, len(df_master)))
        for _, r in sample_df.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                method = str(r['調查方式']); inv_s = str(r['目前農地調查現況'])
                mon_s = str(r['農地監測狀態'])
                
                sides = 4; color = "gray"
                if "個案" in method: sides = 4
                else: sides = 3
                
                if mon_s == "管制": sides = 6; color = "red"
                elif mon_s == "建物": sides = 6; color = "black"
                elif mon_s == "難以採樣": sides = 6; color = "purple"
                elif str(r['代表性']) == "備用點": sides = 4; color = "white"
                else:
                    if "增量" in inv_s: color = "red"
                    elif "延長" in inv_s: color = "blue"
                    elif "正常" in inv_s: color = "green"

                folium.RegularPolygonMarker(
                    location=[lat, lon], number_of_sides=sides, radius=6,
                    color=color, fill=True, popup=f"地號: {r['地段地號']}"
                ).add_to(m)
            except: continue
        
        st_folium(m, width=1100, height=700)
else:
    st.error("Excel 讀取失敗，請確認檔案。")