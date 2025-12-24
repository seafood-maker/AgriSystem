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

# --- 1. 系統權限與路徑設定 ---
ADMIN_PASSWORD = "admin_changhua"  # 您可以自行修改此管理員密碼
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
        # 讀取主檔
        df_m = pd.read_excel(xl, "農地現況主檔")
        df_m.columns = df_m.columns.str.strip()
        df_m['網格編號'] = df_m['網格編號'].apply(clean_id_format)
        
        # 讀取其他分頁
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
                
                # SHP 欄位對齊: 網格號
                if '網格號' in gdf.columns:
                    gdf['網格號'] = gdf['網格號'].apply(clean_id_format)
                # SHP 欄位對齊: 狀態 (用於診斷)
                if '狀態' in gdf.columns:
                    gdf['狀態'] = gdf['狀態'].astype(str).str.strip()
                return gdf
        except Exception as e:
            st.error(f"Shapefile 讀取失敗: {e}")
    return None

# 執行資料載入
df_master, df_history, df_block, df_settings = load_all_data()
gdf_grid = load_grid_shp()

# --- 3. 核心空間運算功能 ---
def find_grid_by_coords(x, y, gdf):
    if gdf is None: return "未知"
    p = Point(x, y)
    match = gdf[gdf.contains(p)]
    if not match.empty:
        return str(match.iloc[0]['網格號'])
    return "範圍外"

# --- 4. 側邊欄與選單 ---
st.sidebar.title("🌿 彰化農地管理系統")
menu = st.sidebar.radio("功能選單", ["📊 統計首頁", "📋 總表清單", "🔍 歷年數據查詢", "🗺️ 空間地圖檢視", "🔐 管理員後台"])

# --- 5. 程式主邏輯 ---
if df_master is not None:
    # 先計算全域統計數值，避免各分頁報錯 (例如之前的 normal_count 錯誤)
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
            st.subheader("📦 個案型農地現況")
            st.dataframe(df_master[df_master['調查方式'] == '個案型農地']['農地監測狀態'].value_counts())

    # --- B. 總表清單 ---
    elif menu == "📋 總表清單":
        st.header("📋 農地現況總表 (全量資料)")
        search_q = st.text_input("輸入地段地號、SGM 或網格編號快速搜尋")
        df_show = df_master
        if search_q:
            df_show = df_master[df_master.astype(str).apply(lambda x: x.str.contains(search_q)).any(axis=1)]
        st.write(f"顯示資料筆數: {len(df_show)}")
        st.dataframe(df_show, use_container_width=True)

    # --- C. 歷年數據查詢 ---
    elif menu == "🔍 歷年數據查詢":
        st.header("🔍 歷史調查數據追蹤")
        target_lot = st.text_input("請輸入地段地號查詢")
        if target_lot:
            history_res = df_history[df_history['SGM編號'].isin(df_master[df_master['地段地號']==target_lot]['SGM編號'])]
            st.dataframe(history_res)

    # --- D. 空間地圖檢視 ---
    elif menu == "🗺️ 空間地圖檢視":
        st.header("🗺️ 衛星影像監測圖 (網格頻率著色)")
        
        # 建立衛星底圖
        m = folium.Map(
            location=[24.05, 120.5], zoom_start=11,
            tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
            attr='Esri World Imagery'
        )
        
        # 1. 網格連動與變色
        if gdf_grid is not None:
            grid_status = df_master.drop_duplicates('網格編號')[['網格編號', '網格監測頻率']]
            gdf_wgs84 = gdf_grid.to_crs(epsg=4326)
            merged = gdf_wgs84.merge(grid_status, left_on='網格號', right_on='網格編號', how='left')

            def get_grid_style(feature):
                freq = str(feature['properties'].get('網格監測頻率', 'nan')).strip()
                fill = '#F8F8F8' # 預設淡白
                if '延長' in freq: fill = '#ADD8E6' # 淡藍
                elif '持續' in freq: fill = '#FFB6C1' # 淡紅
                elif '退場' in freq: fill = '#90EE90' # 淡綠
                return {
                    'fillColor': fill, 'color': 'white', 'weight': 1, 
                    'fillOpacity': 0.5 if fill != '#F8F8F8' else 0.1
                }

            folium.GeoJson(
                merged,
                style_function=get_grid_style,
                tooltip=folium.GeoJsonTooltip(fields=['網格號', '網格監測頻率'], aliases=['網格號:', 'Excel頻率:'])
            ).add_to(m)

        # 2. 農地點位 (照要求之形狀與顏色)
        sample_df = df_master.sample(min(800, len(df_master)))
        for _, r in sample_df.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                method = str(r['調查方式']); inv_status = str(r['目前農地調查現況'])
                mon_status = str(r['農地監測狀態']); rep = str(r['代表性'])
                
                sides = 4; color = "gray"
                if "個案" in method: sides = 4 # 正方形
                else: sides = 3 # 三角形
                
                if mon_status == "管制": sides = 6; color = "red"
                elif mon_status == "建物": sides = 6; color = "black"
                elif mon_status == "難以採樣": sides = 6; color = "purple"
                elif rep == "備用點": sides = 4; color = "white"
                else:
                    if "增量" in inv_status: color = "red"
                    elif "延長" in inv_status: color = "blue"
                    elif "正常" in inv_status: color = "green"

                folium.RegularPolygonMarker(
                    location=[lat, lon], number_of_sides=sides, radius=6,
                    color=color, fill=True, fill_opacity=0.9,
                    popup=f"地號: {r['地段地號']}<br>狀態: {inv_status}"
                ).add_to(m)
            except: continue
        
        st_folium(m, width=1100, height=700)

    # --- E. 管理員後台 (含 DA 計算與 3M 預警) ---
    elif menu == "🔐 管理員後台":
        st.header("🔑 管理員數據錄入")
        pwd = st.text_input("請輸入後台管理密碼", type="password")
        if pwd == ADMIN_PASSWORD:
            st.success("密碼正確")
            search_id = st.text_input("🔍 輸入地段地號查詢農地")
            if search_id:
                targets = df_master[df_master['地段地號'] == search_id]
                if not targets.empty:
                    row = targets.iloc[0]
                    with st.form("input_form"):
                        st.subheader("1. 座標核對")
                        c_x = st.number_input("實測 X (TWD97)", value=float(row['TWD97_X']))
                        c_y = st.number_input("實測 Y (TWD97)", value=float(row['TWD97_Y']))
                        dist = np.sqrt((c_x - row['TWD97_X'])**2 + (c_y - row['TWD97_Y'])**2)
                        if dist > 3:
                            st.warning(f"⚠️ 座標偏移 {dist:.2f} 公尺！")
                            st.info(f"當前位處網格: {find_grid_by_coords(c_x, c_y, gdf_grid)}")
                        
                        st.subheader("2. 重金屬濃度錄入")
                        st.write("DA = (本次 - 初期) / 初期 * 100")
                        # 僅以銅、鋅為例，實際應用可迴圈八項
                        cu_in = st.number_input("銅 (Cu) 濃度", value=0.0)
                        
                        if st.form_submit_button("執行 DA 判定"):
                            init_cu = row.get('初始_銅', 0)
                            s2_cu = df_settings.loc['銅', '管制標準']
                            da_limit = df_settings.loc['銅', '上升標準 (DA門檻)']
                            da_val = ((cu_val - init_cu)/init_cu*100) if init_cu > 0 else 0
                            st.write(f"銅 DA 值: {da_val:.2f}%")
        elif pwd != "":
            st.error("密碼錯誤")

else:
    st.error("找不到資料庫檔案，請檢查 GitHub 檔案名稱。")