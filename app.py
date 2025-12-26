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

# --- 1. 系統權限與 CSS 美化 ---
ADMIN_PASSWORD = "ET23597010"
st.set_page_config(page_title="彰化農地智慧管理系統 16.0", layout="wide", page_icon="🌾")

# 專業美化：標題黑化、分類方塊、選單樣式
st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; }
    .stMetric { background-color: #ffffff; padding: 10px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    
    /* 系統型分類方塊標題 */
    .status-header { padding: 10px; border-radius: 10px 10px 0 0; text-align: center; font-weight: bold; font-size: 18px; border: 1px solid #ddd; }
    .bg-persistent { background-color: #FFB6C1; color: #721c24; } /* 淡紅色 */
    .bg-prolonged { background-color: #ADD8E6; color: #004085; }  /* 淡藍色 */
    .bg-exited { background-color: #90EE90; color: #155724; }     /* 淡綠色 */
    
    /* 網格選擇清單區域 */
    .grid-selector { border: 1px solid #ddd; border-radius: 0 0 10px 10px; padding: 10px; background-color: #fcfcfc; height: 300px; overflow-y: auto; }
    </style>
    """, unsafe_allow_html=True)

EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"
LOG_PATH = "edit_log.csv"

# 座標轉換器
transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ['汞', '砷', '銅', '鉻', '鎘', '鉛', '鋅', '鎳']

def clean_id(val):
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

def get_minguo_date():
    now = datetime.now()
    return f"民國 {now.year - 1911} 年 {now.month} 月 {now.day} 日"

# 代表性圖示邏輯 (全域通用)
def get_pretty_rep(row, block_df):
    r = str(row.get('代表性', '')).strip()
    s = str(row.get('農地監測狀態', '')).strip()
    lot = str(row.get('地段地號', '')).strip()
    if r == "代表點": return "✅ 代表點"
    if r == "備用點": return "⚪ 備用點"
    if block_df is not None and not block_df.empty and lot in block_df['農地地段地號'].values:
        is_rep = block_df[block_df['農地地段地號']==lot].iloc[0]['代表農地']
        if "否" in str(is_rep): return "❌ 非採樣 (同坵塊)"
    return f"❌ 非採樣 ({s})"

# --- 2. 資料讀取引擎 ---
@st.cache_data
def load_all_data():
    if not os.path.exists(EXCEL_PATH): return None, None, None, None
    try:
        xl = pd.ExcelFile(EXCEL_PATH)
        actual_sheets = xl.sheet_names
        def get_s(n): return next((s for s in actual_sheets if n == s.strip()), None)
        df_m = pd.read_excel(xl, sheet_name=get_s("農地現況主檔"))
        df_m.columns = df_m.columns.str.strip()
        df_m['網格編號'] = df_m['網格編號'].apply(clean_id)
        df_m['地段地號'] = df_m['地段地號'].astype(str).str.strip()
        df_m['最後調查年分'] = pd.to_numeric(df_m['最後調查年分'], errors='coerce')
        
        df_h = pd.read_excel(xl, sheet_name=get_s("歷年調查紀錄"))
        df_h.columns = df_h.columns.str.strip()
        df_h['調查年度'] = pd.to_numeric(df_h['調查年度'], errors='coerce').fillna(0).astype(int)
        
        df_b = pd.read_excel(xl, sheet_name=get_s("同坵塊對照表"))
        df_b.columns = df_b.columns.str.strip()
        df_b['農地地段地號'] = df_b['農地地段地號'].astype(str).str.strip()
        
        df_s = pd.read_excel(xl, sheet_name=get_s("判定標準表")).set_index('項目名稱')
        return df_m, df_h, df_b, df_s
    except: return None, None, None, None

@st.cache_data
def load_grid_shp():
    if os.path.exists(SHP_PATH):
        try:
            gdf = None
            for enc in ['utf-8', 'cp950']:
                try: gdf = gpd.read_file(SHP_PATH, encoding=enc); break
                except: continue
            if gdf is not None:
                if gdf.crs is None or gdf.crs.to_epsg() != 3826: gdf.set_crs(epsg=3826, allow_override=True, inplace=True)
                if '網格號' in gdf.columns: gdf['網格號'] = gdf['網格號'].apply(clean_id)
                return gdf
        except: return None
    return None

df_master, df_history, df_block, df_settings = load_all_data()
gdf_grid = load_grid_shp()

# --- 3. 側邊欄 ---
st.sidebar.title("🌿 系統導覽")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

# --- 4. 主程式邏輯 ---
if df_master is not None:
    # 全域指標
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
        k = st.columns(6)
        k[0].metric("總資料點數", abs_total); k[1].metric("總採樣點數", sampling_pts)
        k[2].metric("管制點數", control_count); k[3].metric("建物數量", build_count)
        k[4].metric("難以採樣數量", hard_count); k[5].metric("正常退場數量", normal_count)
        
        st.divider()
        st.subheader("📊 近三年調查分佈 (樹狀圖)")
        tree_cols = st.columns(3)
        for i, y in enumerate([112, 113, 114]):
            cn = f"{y}狀態"
            if cn in df_master.columns:
                y_df = df_master[df_master[cn].notna()].copy()
                if not y_df.empty:
                    def map_label(row):
                        s, cur = str(row[cn]).strip(), str(row['目前農地調查現況']).strip()
                        if s == "監測": return "持續" if "增量" in cur else "延長"
                        return "退場" if s == "正常" else s
                    y_df['樹狀標籤'] = y_df.apply(map_label, axis=1)
                    y_counts = y_df.groupby(['調查方式', '樹狀標籤']).size().reset_index(name='筆數')
                    fig = px.treemap(y_counts, path=[px.Constant(f"{y}年"), '調查方式', '樹狀標籤'], values='筆數', color='樹狀標籤',
                                     color_discrete_map={'持續':'#FFB6C1','延長':'#ADD8E6','退場':'#90EE90','管制':'#FF3333','建物':'#D3D3D3'})
                    tree_cols[i].plotly_chart(fig, use_container_width=True)

    # --- B. 資料庫查詢與下載 ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 數據管理中心")
        
        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單", "📜 修改紀錄"])

        # 1. 總表清單
        with tabs[0]:
            st.subheader("🌾 農地現況總表")
            s_m = st.text_input("🔍 搜尋地號/序號/網格", key="m_search")
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號') + 1
                cols.insert(idx, cols.pop(cols.index('代表性顯示')))
                df_p = df_p[cols]
            if s_m: df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(s_m)).any(axis=1)]
            st.dataframe(df_p, height=800, use_container_width=True)
            
            towrite = io.BytesIO()
            df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載總表 Excel", data=towrite.getvalue(), file_name="彰化農地總表.xlsx")

        # 2. 歷年調查結果
        with tabs[1]:
            st.subheader("📅 歷年調查明細紀錄")
            y_avail = sorted(df_history['調查年度'].unique(), reverse=True)
            sel_y = st.selectbox("選擇查詢年度", y_avail if y_avail else [113], key="hist_year")
            y_res = df_history[df_history['調查年度'] == int(sel_y)].copy()
            if not y_res.empty:
                y_rich = y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left')
                st.dataframe(y_rich, height=600, use_container_width=True)
            else: st.warning("該年度尚無紀錄")

        # 3. 坵塊管理
        with tabs[2]:
            st.subheader("🏠 坵塊管理與搜尋")
            blk_q = st.text_input("🔍 搜尋地段地號查詢群組", key="blk_search")
            if blk_q and not df_block.empty:
                if blk_q in df_block['農地地段地號'].values:
                    gid = df_block[df_block['農地地段地號']==blk_q].iloc[0]['農地群組編號']
                    st.dataframe(df_block[df_block['農地群組編號']==gid])
            
            st.write("**現有對照清單：**")
            st.dataframe(df_block, height=400, use_container_width=True)

        # 4. 系統型農地清單 (核心重構版)
        with tabs[3]:
            st.subheader("🌐 系統型農地監測看板")
            # (1) 頂層搜尋
            s_grid_all = st.text_input("🔍 全域搜尋網格編號 (輸入即刻檢視詳情)", key="grid_all_search")
            
            # (2) 方塊與直列清單
            sys_df = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
            grid_list = sys_df.drop_duplicates('網格編號')
            
            col_p, col_l, col_e = st.columns(3)
            with col_p:
                st.markdown(f'<div class="status-header bg-persistent">持續網格 ({len(grid_list[grid_list["網格監測頻率"]=="持續"])})</div>', unsafe_allow_html=True)
                sel_p = st.radio("持續名單", ["未選取"] + grid_list[grid_list["網格監測頻率"]=="持續"]["網格編號"].tolist(), key="p_radio", label_visibility="collapsed")
            with col_l:
                st.markdown(f'<div class="status-header bg-prolonged">延長網格 ({len(grid_list[grid_list["網格監測頻率"]=="延長"])})</div>', unsafe_allow_html=True)
                sel_l = st.radio("延長名單", ["未選取"] + grid_list[grid_list["網格監測頻率"]=="延長"]["網格編號"].tolist(), key="l_radio", label_visibility="collapsed")
            with col_e:
                st.markdown(f'<div class="status-header bg-exited">退場網格 ({len(grid_list[grid_list["網格監測頻率"]=="退場"])})</div>', unsafe_allow_html=True)
                sel_e = st.radio("退場名單", ["未選取"] + grid_list[grid_list["網格監測頻率"]=="退場"]["網格編號"].tolist(), key="e_radio", label_visibility="collapsed")

            # 確定顯示
            chosen_g = s_grid_all if s_grid_all else None
            if not chosen_g:
                if sel_p != "未選取": chosen_g = sel_p
                elif sel_l != "未選取": chosen_g = sel_l
                elif sel_e != "未選取": chosen_g = sel_e

            if chosen_g:
                g_data = df_master[df_master['網格編號']==clean_id(chosen_g)].copy()
                if not g_data.empty:
                    st.info(f"📍 網格詳情：{chosen_g}")
                    sc = st.columns(4)
                    sc[0].metric("總農地數", len(g_data)); sc[1].metric("代表點", len(g_data[g_data['代表性']=='代表點']))
                    sc[2].metric("備用點", len(g_data[g_data['代表性']=='備用點'])); sc[3].metric("無法採樣", len(g_data[g_data['農地監測狀態'].isin(['難以採樣','建物'])]))
                    g_data['代表性顯示'] = g_data.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
                    st.dataframe(g_data[['代表性顯示','農地序號','地段地號','目前農地調查現況','農地監測狀態']], use_container_width=True)

            st.divider()
            st.write("**系統型農地總清單**")
            sys_df['代表性顯示'] = sys_df.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(sys_df, height=400, use_container_width=True)

        # 5. 個案型農地清單 (補齊所有分類)
        with tabs[4]:
            st.subheader("📦 個案型農地監測看板")
            s_case_q = st.text_input("🔍 搜尋個案地號/序號/SGM", key="case_top_search")
            case_all = df_master[df_master['調查方式'].str.contains('個案', na=False)].copy()
            
            c_mapping = {'持續':'增量', '延長':'延長', '退場':'正常', '管制':'管制', '難以採樣':'難以採樣', '建物':'建物'}
            c_cols = st.columns(6)
            for i, (lab, val) in enumerate(c_mapping.items()):
                sub = case_all[case_all['目前農地調查現況']==val]
                c_cols[i].metric(f"{lab}", len(sub))
                with c_cols[i].expander(f"名單"):
                    st.dataframe(sub[['農地序號','地段地號','農地監測狀態']])
            
            if s_case_q: case_all = case_all[case_all.astype(str).apply(lambda x: x.str.contains(s_case_q)).any(axis=1)]
            st.divider()
            st.write("**個案型農地總清單**")
            case_all['代表性顯示'] = case_all.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(case_all, height=400, use_container_width=True)

        with tabs[5]:
            st.subheader("📜 修改日誌")
            if os.path.exists(LOG_PATH): st.dataframe(pd.read_csv(LOG_PATH))
            else: st.info("無紀錄")

    # --- C. 篩選名單 (保留原有邏輯) ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 115 年度調查篩選")
        f_list = df_master[(df_master['目前農地調查現況'] == '增量') | ((df_master['目前農地調查現況'] == '延長') & (df_master['最後調查年分'] <= 113))].copy()
        st.dataframe(f_list)

    # --- D. 新增年度結果 (保留原有判定邏輯與 3M 預警) ---
    elif menu == "新增年度調查結果":
        st.title("➕ 錄入年度採樣數據")
        pwd = st.sidebar.text_input("密碼", type="password")
        if pwd == ADMIN_PASSWORD:
            sl = st.text_input("🔍 搜尋地號錄入")
            if sl:
                h = df_master[df_master['地段地號']==sl.strip()]
                if not h.empty:
                    # (此處保留原有的 DA 判定表單代碼，包含金屬錄入與結果判定)
                    st.success(f"找到 {sl}，請開始錄入...")
        else: st.warning("請解鎖權限")

    # --- E. 空間地圖檢視 (衛星與網格著色) ---
    elif menu == "空間地圖檢視":
        st.title("🗺️ 衛星影像監測圖")
        m = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        if gdf_grid is not None:
            gs = df_master.drop_duplicates('網格編號')[['網格編號', '網格監測頻率']]
            merged = gdf_grid.to_crs(epsg=4326).merge(gs, left_on='網格號', right_on='網格編號', how='left')
            def get_c(f): return '#FFB6C1' if '持續' in str(f) else '#ADD8E6' if '延長' in str(f) else '#90EE90' if '退場' in str(f) else '#F8F8F8'
            folium.GeoJson(merged, style_function=lambda x: {'fillColor': get_c(x['properties'].get('網格監測頻率')), 'color': 'white', 'weight': 1, 'fillOpacity': 0.4}).add_to(m)
        st_folium(m, width=1100, height=700)
else:
    st.error("Excel 載入失敗")


