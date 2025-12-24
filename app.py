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
ADMIN_PASSWORD = "admin_changhua" # 您可以自行修改此管理員密碼
st.set_page_config(page_title="彰化農地智慧管理系統", layout="wide", page_icon="🌾")

# 檔案路徑 (需確保上傳 GitHub 時檔名一致)
EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"

# 座標轉換器 (TWD97 EPSG:3826 -> WGS84 EPSG:4326)
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
            # 支援多種編碼讀取
            gdf = None
            for enc in ['utf-8', 'cp950', 'big5']:
                try:
                    gdf = gpd.read_file(SHP_PATH, encoding=enc)
                    break
                except: continue
            
            if gdf is not None:
                if gdf.crs is None or gdf.crs.to_epsg() != 3826:
                    gdf.set_crs(epsg=3826, allow_override=True, inplace=True)
                # 您的 SHP 網格編號欄位為『網格號』
                if '網格號' in gdf.columns:
                    gdf['網格號'] = gdf['網格號'].apply(clean_id)
                return gdf
        except Exception as e:
            st.error(f"Shapefile 讀取失敗: {e}")
    return None

df_master, df_history, df_block, df_settings = load_all_data()
gdf_grid = load_grid_shp()

# --- 3. 核心運算：空間網格判定 ---
def find_grid_by_coords(x, y, gdf):
    if gdf is None: return "未知"
    p = Point(x, y)
    match = gdf[gdf.contains(p)]
    if not match.empty:
        return str(match.iloc[0]['網格號'])
    return "範圍外"

# --- 4. 側邊欄與導覽 ---
st.sidebar.title("🌿 彰化農地管理系統")
menu = st.sidebar.radio("功能選單", ["📊 統計首頁", "📋 總表清單", "🔍 歷年數據查詢", "🗺️ 空間地圖檢視", "🔐 管理員後台"])

if df_master is not None:
    # 統計基礎數據
    total_spots = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_pts = len(df_master[df_master['農地監測狀態'] == '管制'])
    build_pts = len(df_master[df_master['農地監測狀態'] == '建物'])
    hard_pts = len(df_master[df_master['農地監測狀態'] == '難以採樣'])
    normal_pts = len(df_master[df_master['農地監測狀態'] == '正常'])

    # --- A. 統計首頁 ---
    if menu == "📊 統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("資料總筆數", len(df_master))
        k2.metric("總採樣點", total_spots)
        k3.metric("管制點位", control_pts)
        k4.metric("建物數量", build_pts)
        k5.metric("難以採樣", hard_pts)
        k6.metric("正常(退場)", normal_count)
        
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("🌐 系統型網格現況 (Excel)")
            grid_uniq = df_master.drop_duplicates('網格編號')
            st.write(f"總網格數: {len(grid_uniq)}")
            st.dataframe(grid_uniq['網格監測頻率'].value_counts())
        with c2:
            st.subheader("📦 個案型農地現況")
            st.dataframe(df_master[df_master['調查方式'] == '個案型農地']['農地監測狀態'].value_counts())

    # --- B. 總表清單 ---
    elif menu == "📋 總表清單":
        st.header("📋 農地資料總表")
        search = st.text_input("輸入地段地號或 SGM 編號快速查詢")
        df_show = df_master
        if search:
            df_show = df_master[df_master.astype(str).apply(lambda x: x.str.contains(search)).any(axis=1)]
        st.dataframe(df_show, use_container_width=True)

    # --- C. 歷年查詢 ---
    elif menu == "🔍 歷年數據查詢":
        st.header("🔍 歷史調查數據追蹤")
        lot_query = st.text_input("請輸入要查詢的地段地號")
        if lot_query:
            # 此處可加入顯示該地號歷史濃度趨勢圖的代碼
            st.write(f"顯示 {lot_query} 之歷史軌跡...")

    # --- D. 空間地圖檢視 ---
    elif menu == "🗺️ 空間地圖檢視":
        st.header("🗺️ 衛星影像監測圖 (網格頻率著色)")
        
        m = folium.Map(
            location=[24.05, 120.5], zoom_start=11,
            tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
            attr='Esri World Imagery'
        )
        
        # 繪製網格面 (Shapefile)
        if gdf_grid is not None:
            grid_status = df_master.drop_duplicates('網格編號')[['網格編號', '網格監測頻率']]
            gdf_wgs84 = gdf_grid.to_crs(epsg=4326)
            merged = gdf_wgs84.merge(grid_status, left_on='網格號', right_on='網格編號', how='left')

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

        # 繪製農地點 (農地形狀顏色規則)
        sample_df = df_master.sample(min(800, len(df_master)))
        for _, r in sample_df.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                sides = 4; color = "gray"
                if "個案" in str(r['調查方式']): sides = 4
                else: sides = 3
                
                # 形狀與顏色判定
                mon_s = str(r['農地監測狀態'])
                inv_s = str(r['目前農地調查現況'])
                
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
                    color=color, fill=True, fill_opacity=0.9,
                    popup=f"地號: {r['地段地號']}"
                ).add_to(m)
            except: continue
        
        st_folium(m, width=1100, height=700)

    # --- E. 管理員後台 (計算 DA 值) ---
    elif menu == "🔐 管理員後台":
        st.header("🔑 管理員數據錄入 (需授權)")
        pwd = st.text_input("請輸入管理密碼", type="password")
        if pwd == ADMIN_PASSWORD:
            st.success("身分驗證成功")
            lot_id = st.text_input("🔍 搜尋欲錄入之地段地號")
            if lot_id:
                targets = df_master[df_master['地段地號'] == lot_id]
                if not targets.empty:
                    row = targets.iloc[0]
                    with st.form("input_form"):
                        col1, col2 = st.columns(2)
                        new_x = col1.number_input("採樣 X (TWD97)", value=float(row['TWD97_X']))
                        new_y = col2.number_input("採樣 Y (TWD97)", value=float(row['TWD97_Y']))
                        
                        # 3米誤差與網格判定
                        dist = np.sqrt((new_x - row['TWD97_X'])**2 + (new_y - row['TWD97_Y'])**2)
                        if dist > 3:
                            st.warning(f"⚠️ 座標偏移 {dist:.2f} 米")
                            new_grid = find_grid_by_coords(new_x, new_y, gdf_grid)
                            st.error(f"目前位於網格: {new_grid}")
                        
                        # DA 判定 (僅示範銅)
                        cu_val = st.number_input("銅 Cu 濃度", value=0.0)
                        
                        if st.form_submit_button("執行判定"):
                            init_cu = row.get('初始_銅', 0)
                            s2_cu = df_settings.loc['銅', '管制標準']
                            da_cu = ((cu_val - init_cu) / init_cu * 100) if init_cu > 0 else 0
                            st.write(f"DA 增量: {da_cu:.2f}%")
                            st.info("計算完成。請注意：GitHub 模式下存檔需手動更新 Excel 並重新上傳。")
        elif pwd != "":
            st.error("密碼錯誤")