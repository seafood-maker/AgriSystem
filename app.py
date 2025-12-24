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

# --- 1. 系統權限與基礎設定 ---
ADMIN_PASSWORD = "ET23597010"  # 已更新為您的專屬密碼
st.set_page_config(page_title="彰化農地智慧管理系統", layout="wide", page_icon="🌾")

EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"

# 座標轉換器 (TWD97 EPSG:3826 -> WGS84 EPSG:4326)
transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ['汞', '砷', '銅', '鉻', '鎘', '鉛', '鋅', '鎳']

# 數據清洗工具 (解決對接不上的核心)
def clean_id_format(val):
    s = str(val).strip()
    s = re.sub(r'\.0$', '', s) # 刪除數字後綴的 .0
    return s

# --- 2. 資料讀取引擎 ---
@st.cache_data
def load_all_data():
    if not os.path.exists(EXCEL_PATH): return None, None, None, None
    try:
        xl = pd.ExcelFile(EXCEL_PATH)
        df_m = pd.read_excel(xl, "農地現況主檔")
        df_m.columns = df_m.columns.str.strip()
        df_m['網格編號'] = df_m['網格編號'].apply(clean_id_format)
        
        df_h = pd.read_excel(xl, "歷年調查紀錄")
        df_b = pd.read_excel(xl, "同坵塊對照表")
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
                
                # SHP 欄位對齊: 網格號
                if '網格號' in gdf.columns:
                    gdf['網格號'] = gdf['網格號'].apply(clean_id_format)
                return gdf
        except Exception as e:
            st.error(f"Shapefile 讀取失敗: {e}")
    return None

df_master, df_history, df_block, df_settings = load_all_data()
gdf_grid = load_grid_shp()

# 空間判定功能
def find_grid_by_coords(x, y, gdf):
    if gdf is None: return "未知"
    p = Point(x, y)
    match = gdf[gdf.contains(p)]
    if not match.empty:
        return str(match.iloc[0]['網格號'])
    return "範圍外"

# --- 3. 側邊欄 ---
st.sidebar.title("🌿 彰化農地管理 7.0")
menu = st.sidebar.radio("功能導覽", ["📊 統計首頁", "📋 總表清單", "➕ 新增年度調查", "🗺️ 空間地圖檢視"])

# --- 4. 主程式邏輯 ---
if df_master is not None:
    # 全域數據統計 (修復之前的 NameError)
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
            st.subheader("🌐 系統型網格現況")
            grid_uniq = df_master.drop_duplicates('網格編號')
            st.write(f"總網格數: {len(grid_uniq)}")
            st.dataframe(grid_uniq['網格監測頻率'].value_counts())
        with c2:
            st.subheader("📦 個案型農地現況")
            st.dataframe(df_master[df_master['調查方式'] == '個案型農地']['農地監測狀態'].value_counts())

    # --- B. 總表清單 ---
    elif menu == "📋 總表清單":
        st.header("📋 農地資料全表")
        search = st.text_input("搜尋地段地號、SGM 或網格編號")
        df_show = df_master
        if search:
            df_show = df_master[df_master.astype(str).apply(lambda x: x.str.contains(search)).any(axis=1)]
        st.dataframe(df_show, use_container_width=True)

    # --- C. 新增年度調查 (含 DA 與 3M 判定) ---
    elif menu == "➕ 新增年度調查":
        st.header("➕ 新增調查結果與 DA 判定")
        pwd = st.sidebar.text_input("管理員授權碼", type="password")
        
        if pwd == ADMIN_PASSWORD:
            st.success("✅ 權限已解鎖")
            search_lot = st.text_input("🔍 第一步：輸入地段地號搜尋農地 (例: 華南段0159-0000)")
            if search_lot:
                targets = df_master[df_master['地段地號'] == search_lot]
                if not targets.empty:
                    selected_sgm = st.selectbox("確認 SGM 編號", targets['SGM編號'].unique())
                    row = targets[targets['SGM編號'] == selected_sgm].iloc[0]
                    
                    with st.form("survey_form"):
                        st.subheader(f"📍 編輯對象：{search_lot}")
                        col_y, col_admin = st.columns(2)
                        target_year = col_y.selectbox("調查年度", [114, 115, 116])
                        admin_status = col_admin.selectbox("行政狀態", ["監測", "建物", "管制", "難以採樣", "正常"])
                        
                        st.write("---")
                        st.subheader("1. 座標核對 (誤差 3M 警示)")
                        c_x = st.number_input("實測座標 X (TWD97)", value=float(row['TWD97_X']))
                        c_y = st.number_input("實測座標 Y (TWD97)", value=float(row['TWD97_Y']))
                        dist = np.sqrt((c_x - row['TWD97_X'])**2 + (c_y - row['TWD97_Y'])**2)
                        if dist > 3:
                            st.warning(f"⚠️ 座標偏移 {dist:.2f} 公尺！")
                            st.info(f"當前位處網格: {find_grid_by_coords(c_x, c_y, gdf_grid)}")
                        
                        st.write("---")
                        st.subheader("2. 重金屬檢測值錄入")
                        current_vals = {}
                        for m in METALS:
                            m_c1, m_c2 = st.columns(2)
                            xrf = m_c1.number_input(f"{m} (XRF)", min_value=0.0, key=f"x_{m}")
                            tot = m_c2.number_input(f"{m} (全量)", min_value=0.0, key=f"t_{m}")
                            current_vals[m] = tot if tot > 0 else xrf

                        if st.form_submit_button("執行自動判定"):
                            st.subheader("💡 判定結果報告")
                            final_invest = "正常"
                            da_results = {}
                            
                            for m in METALS:
                                init = row.get(f'初始_{m}', 0)
                                s1 = df_settings.loc[m, '監測標準']
                                s2 = df_settings.loc[m, '管制標準']
                                da_limit = df_settings.loc[m, '上升標準 (DA門檻)']
                                curr = current_vals[m]
                                
                                da = ((curr - init) / init * 100) if init > 0 else 0
                                da_results[m] = f"{da:.1f}%"
                                
                                if curr > s2: final_invest = "管制"
                                elif init > s2 and da > da_limit: 
                                    if final_invest != "管制": final_invest = "增量"
                                elif curr > s1:
                                    if final_invest not in ["管制", "增量"]: final_invest = "延長"
                            
                            st.write(f"最終判定：**{final_invest}**")
                            st.write("DA 值清單：", da_results)
                            st.info(f"待上傳動作：將結果填入 {target_year}狀態，並更新最新年度為 {target_year}")
                else: st.error("查無此地號")
        else: st.warning("請在左側輸入密碼以進行資料錄入")

    # --- D. 空間地圖檢視 (衛星航照 + 網格變色) ---
    elif menu == "🗺️ 空間地圖檢視":
        st.header("🗺️ 衛星影像監測圖 (網格邏輯著色)")
        m = folium.Map(
            location=[24.05, 120.5], zoom_start=11,
            tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
            attr='Esri World Imagery'
        )
        
        if gdf_grid is not None:
            grid_status = df_master.drop_duplicates('網格編號')[['網格編號', '網格監測頻率']]
            gdf_w = gdf_grid.to_crs(epsg=4326)
            merged = gdf_w.merge(grid_status, left_on='網格號', right_on='網格編號', how='left')

            def get_color(f):
                f = str(f)
                if '持續' in f: return '#FFB6C1' # 淡紅
                if '延長' in f: return '#ADD8E6' # 淡藍
                if '退場' in f: return '#90EE90' # 淡綠
                return '#F8F8F8' # 淡白

            folium.GeoJson(
                merged,
                style_function=lambda x: {
                    'fillColor': get_color(x['properties'].get('網格監測頻率')),
                    'color': 'white', 'weight': 1, 'fillOpacity': 0.4
                },
                tooltip=folium.GeoJsonTooltip(fields=['網格號', '網格監測頻率'])
            ).add_to(m)

        # 繪製農地點 (三角/正方/六角)
        sample_df = df_master.sample(min(1000, len(df_master)))
        for _, r in sample_df.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                method = str(r['調查方式']); mon_s = str(r['農地監測狀態']); inv_s = str(r['目前農地調查現況'])
                
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
    st.error("Excel 載入失敗。")