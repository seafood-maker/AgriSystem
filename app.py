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
ADMIN_PASSWORD = "ET23597010"  # 您的專屬密碼
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

# 民國日期轉換
def get_minguo_date():
    now = datetime.now()
    return f"民國 {now.year - 1911} 年 {now.month} 月 {now.day} 日"

# --- 2. 資料讀取引擎 (自動偵測分頁) ---
@st.cache_data
def load_all_data():
    if not os.path.exists(EXCEL_PATH): return None, None, None, None
    try:
        xl = pd.ExcelFile(EXCEL_PATH)
        actual_sheets = xl.sheet_names
        
        def get_best_sheet(target, sheets):
            for s in sheets:
                if target == s.strip(): return s
            return None

        s1 = get_best_sheet("農地現況主檔", actual_sheets)
        s2 = get_best_sheet("歷年調查紀錄", actual_sheets)
        s3 = get_best_sheet("同坵塊對照表", actual_sheets)
        s4 = get_best_sheet("判定標準表", actual_sheets)

        if not s1:
            st.error(f"❌ 找不到『農地現況主檔』分頁，請檢查 Excel 名稱。")
            return None, None, None, None

        df_m = pd.read_excel(xl, sheet_name=s1)
        df_m.columns = df_m.columns.str.strip()
        df_m['網格編號'] = df_m['網格編號'].apply(clean_id)
        
        df_h = pd.read_excel(xl, sheet_name=s2) if s2 else pd.DataFrame()
        df_b = pd.read_excel(xl, sheet_name=s3) if s3 else pd.DataFrame()
        df_s = pd.read_excel(xl, sheet_name=s4).set_index('項目名稱') if s4 else pd.DataFrame()
        
        return df_m, df_h, df_b, df_s
    except Exception as e:
        st.error(f"❌ Excel 讀取錯誤: {e}")
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

def find_grid_by_coords(x, y, gdf):
    if gdf is None: return "未知"
    p = Point(x, y)
    match = gdf[gdf.contains(p)]
    return str(match.iloc[0]['網格號']) if not match.empty else "範圍外"

# --- 3. 側邊欄 ---
st.sidebar.title("🌿 系統導覽")
menu = st.sidebar.radio("功能選單", ["統計首頁", "資料庫查詢與下載", "新增年度調查結果", "空間地圖檢視"])

# --- 4. 主程式邏輯 ---
if df_master is not None:
    # 全域數據統計 (確保各頁面變數可用)
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

        # 1. 頂列統計
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("總資料點數", abs_total)
        m2.metric("總採樣點數", sampling_pts)
        m3.metric("管制點數", control_count)
        m4.metric("建物數量", build_count)
        m5.metric("難以採樣數量", hard_count)
        m6.metric("正常退場數量", normal_count)

        st.divider()

        # 2. 系統型網格
        st.subheader("🌐 系統型網格現況統計")
        grid_df = df_master.drop_duplicates('網格編號').copy()
        grid_df['網格監測頻率'] = grid_df['網格監測頻率'].fillna('無網格狀態').astype(str).str.strip()
        g_cont = len(grid_df[grid_df['網格監測頻率'] == '持續'])
        g_ext = len(grid_df[grid_df['網格監測頻率'] == '延長'])
        g_exited = len(grid_df[grid_df['網格監測頻率'] == '退場'])
        g_none = len(grid_df) - (g_cont + g_ext + g_exited)
        
        g_cols = st.columns(5)
        g_cols[0].metric("持續網格", g_cont); g_cols[1].metric("延長網格", g_ext)
        g_cols[2].metric("退場網格", g_exited); g_cols[3].metric("有效網格合計", g_cont+g_ext+g_exited)
        g_cols[4].metric("無網格狀態", g_none)

        st.divider()

        # 3. 個案型農地 (修正: 排除系統型即為個案型)
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
        for i, lab in enumerate(["持續", "延長", "退場", "管制", "難以採樣", "建物"]):
            c_cols[i].metric(lab, c_stats[lab])

        st.divider()

        # 4. 網格查詢系統
        st.subheader("🔍 網格查詢系統")
        grid_search = st.text_input("輸入網格號碼搜尋內部農地 (如: G001)")
        if grid_search:
            grid_res = df_master[df_master['網格編號'] == clean_id(grid_search)]
            if not grid_res.empty:
                def highlight_rep(row):
                    return ['background-color: #FFFFCC' if row.代表性 == '代表點' else '' for _ in row]
                st.dataframe(grid_res.style.apply(highlight_rep, axis=1))
            else: st.warning("查無此網格")

        st.divider()

        # 5. 近三年樹狀圖
        st.subheader("📊 近三年調查分佈")
        years = [112, 113, 114]
        tree_cols = st.columns(3)
        for i, y in enumerate(years):
            col_name = f"{y}狀態"
            if col_name in df_master.columns:
                plot_data = df_master[df_master[col_name].notna()].copy()
                if not plot_data.empty:
                    plot_data['筆數'] = 1
                    fig = px.treemap(plot_data, path=[px.Constant(f"{y}年"), '調查方式', col_name], values='筆數',
                                     color=col_name, color_discrete_map={'監測':'#ADD8E6', '正常':'#90EE90', '管制':'#FFB6C1', '建物':'#D3D3D3'})
                    tree_cols[i].plotly_chart(fig, use_container_width=True)

    # --- B. 資料庫查詢與下載 ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 資料庫中心")
        t1, t2, t3 = st.tabs(["總表清單", "歷年調查結果", "同坵塊關聯"])
        with t1:
            st.dataframe(df_master)
            towrite = io.BytesIO()
            df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載全量總表", data=towrite.getvalue(), file_name="農地總表.xlsx")
        with t2:
            s_year = st.selectbox("選擇年度", sorted(df_history['調查年度'].unique()) if not df_history.empty else [114])
            hist_res = df_history[df_history['調查年度'] == s_year]
            st.dataframe(hist_res)
        with t3:
            s_lot = st.text_input("搜尋地號查看同群組農地")
            if s_lot:
                match = df_block[df_block['農地地段地號'] == s_lot]
                if not match.empty:
                    gid = match.iloc[0]['農地群組編號']
                    rel_lots = df_block[df_block['農地群組編號'] == gid]['農地地段地號'].tolist()
                    st.dataframe(df_master[df_master['地段地號'].isin(rel_lots)])

    # --- C. 新增年度調查 (含 DA 判定與 3M 預警) ---
    elif menu == "新增年度調查結果":
        st.title("➕ 新增調查紀錄 (管理員專區)")
        pwd = st.sidebar.text_input("請輸入密碼", type="password")
        if pwd == ADMIN_PASSWORD:
            st.success("✅ 驗證通過")
            search_id = st.text_input("🔍 第一步：輸入地段地號搜尋農地")
            if search_id:
                targets = df_master[df_master['地段地號'] == search_id]
                if not targets.empty:
                    row = targets.iloc[0]
                    with st.form("data_form"):
                        col_a, col_b = st.columns(2)
                        c_x = col_a.number_input("實測 X (TWD97)", value=float(row['TWD97_X']))
                        c_y = col_b.number_input("實測 Y (TWD97)", value=float(row['TWD97_Y']))
                        dist = np.sqrt((c_x - row['TWD97_X'])**2 + (c_y - row['TWD97_Y'])**2)
                        if dist > 3:
                            st.warning(f"⚠️ 座標偏移 {dist:.2f} 米！位於網格: {find_grid_by_coords(c_x, c_y, gdf_grid)}")
                        
                        st.write("🧪 重金屬錄入 (DA 判定)")
                        user_v = {}
                        for m in METALS:
                            u_col1, u_col2 = st.columns(2)
                            xrf = u_col1.number_input(f"{m} (XRF)", min_value=0.0)
                            tot = u_col2.number_input(f"{m} (全量)", min_value=0.0)
                            user_v[m] = tot if tot > 0 else xrf
                        
                        if st.form_submit_button("執行判定"):
                            final_st = "正常"; da_list = {}
                            for m in METALS:
                                init = row.get(f'初始_{m}', 0)
                                da = ((user_v[m]-init)/init*100) if init > 0 else 0
                                da_list[m] = f"{da:.1f}%"
                                if user_v[m] > df_settings.loc[m, '管制標準']: final_st = "管制"
                                elif init > df_settings.loc[m, '管制標準'] and da > df_settings.loc[m, '上升標準 (DA門檻)']:
                                    if final_st != "管制": final_st = "增量"
                            st.write(f"💡 建議判定: **{final_st}**", da_list)
        else: st.warning("請解鎖後進行錄入")

    # --- D. 空間地圖檢視 ---
    elif menu == "空間地圖檢視":
        st.header("🗺️ 衛星影像與網格監測圖")
        m = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri Imagery')
        if gdf_grid is not None:
            grid_status = df_master.drop_duplicates('網格編號')[['網格編號', '網格監測頻率']]
            merged = gdf_grid.to_crs(epsg=4326).merge(grid_status, left_on='網格號', right_on='網格編號', how='left')
            def get_c(f):
                f = str(f)
                return '#FFB6C1' if '持續' in f else '#ADD8E6' if '延長' in f else '#90EE90' if '退場' in f else '#F8F8F8'
            folium.GeoJson(merged, style_function=lambda x: {'fillColor': get_c(x['properties'].get('網格監測頻率')), 'color': 'white', 'weight': 1, 'fillOpacity': 0.4}).add_to(m)
        
        sample = df_master.sample(min(800, len(df_master)))
        for _, r in sample.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                s = 4 if "個案" in str(r['調查方式']) else 3
                if str(r['農地監測狀態']) == "管制": s = 6; c = "red"
                elif str(r['農地監測狀態']) == "建物": s = 6; c = "black"
                elif str(r['農地監測狀態']) == "難以採樣": s = 6; c = "purple"
                elif str(r['代表性']) == "備用點": s = 4; c = "white"
                else:
                    stv = str(r['目前農地調查現況'])
                    c = "red" if "增量" in stv else "blue" if "延長" in stv else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=s, radius=6, color=c, fill=True, popup=f"地號: {r['地段地號']}").add_to(m)
            except: continue
        st_folium(m, width=1100, height=700)
else:
    st.error("系統啟動失敗，請檢查 Excel 資料夾與分頁。")





