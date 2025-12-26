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

# --- 1. 系統設定與權限 ---
ADMIN_PASSWORD = "ET23597010"
st.set_page_config(page_title="彰化農地智慧管理系統", layout="wide", page_icon="🌾")

# 專業美化 CSS
st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; }
    .stMetric { background-color: #ffffff; padding: 10px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .stDownloadButton button { width: 100%; background-color: #2e7d32; color: white; border-radius: 5px; }
    
    /* 系統型分類看板樣式 */
    .status-header { padding: 10px; border-radius: 10px; text-align: center; font-weight: bold; font-size: 18px; border: 1px solid #ddd; margin-bottom: 5px; }
    .bg-p { background-color: #FFB6C1; color: #721c24; } /* 持續-淡紅 */
    .bg-l { background-color: #ADD8E6; color: #004085; } /* 延長-淡藍 */
    .bg-e { background-color: #90EE90; color: #155724; } /* 退場-淡綠 */
    
    /* 選單清單區域 */
    .grid-list-box { height: 250px; overflow-y: auto; border: 1px solid #eee; padding: 10px; background: #fff; border-radius: 5px; }
    </style>
    """, unsafe_allow_html=True)

EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"
LOG_PATH = "edit_log.csv"

transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ["汞", "砷", "銅", "鉻", "鎘", "鉛", "鋅", "鎳"]

def clean_id(val):
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

def get_minguo_date():
    now = datetime.now()
    return f"民國 {now.year - 1911} 年 {now.month} 月 {now.day} 日"

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
        df_h = pd.read_excel(xl, sheet_name=get_s("歷年調查紀錄"))
        df_h['調查年度'] = pd.to_numeric(df_h['調查年度'], errors='coerce').fillna(0).astype(int)
        df_b = pd.read_excel(xl, sheet_name=get_s("同坵塊對照表"))
        df_b['農地地段地號'] = df_b['農地地段地號'].astype(str).str.strip()
        df_s = pd.read_excel(xl, sheet_name=get_s("判定標準表")).set_index('項目名稱')
        return df_m, df_h, df_b, df_s
    except: return None, None, None, None

df_master, df_history, df_block, df_settings = load_all_data()

# --- 3. [核心修正]：全局數據預處理 (解決 NameError) ---
if df_master is not None:
    # 全域指標
    abs_total = len(df_master)
    sampling_pts_count = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])

    # 系統型預處理
    grid_sys_master = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
    grid_uniq_list = grid_sys_master.drop_duplicates('網格編號').copy()
    grid_uniq_list['網格監測頻率'] = grid_uniq_list['網格監測頻率'].fillna('無狀態')
    
    # 個案型預處理 (修正：排除系統型即個案型)
    case_data_master = df_master[~df_master['調查方式'].str.contains('系統', na=False)].copy()
    c_map = {"增量":"持續", "延長":"延長", "正常":"退場", "管制":"管制", "難以採樣":"難以採樣", "建物":"建物"}
    case_data_master['顯示狀態'] = case_data_master['目前農地調查現況'].map(c_map).fillna(case_data_master['目前農地調查現況'])

# --- 4. 側邊欄與導覽 ---
st.sidebar.title("🌿 系統選單")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

# --- 5. 頁面邏輯 ---
if df_master is not None:
    # --- A. 統計首頁 ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        k = st.columns(6)
        k[0].metric("總資料點數", abs_total); k[1].metric("總採樣點數", sampling_pts_count)
        k[2].metric("管制點數", control_pts_count); k[3].metric("建物數量", build_pts_count)
        k[4].metric("難以採樣數量", hard_pts_count); k[5].metric("正常退場數量", normal_pts_count)
        
        st.divider()
        st.subheader("📊 近三年調查分佈 (樹狀圖)")
        tree_cols = st.columns(3)
        for i, y in enumerate([112, 113, 114]):
            cn = f"{y}狀態"
            if cn in df_master.columns:
                y_df = df_master[df_master[cn].notna()].copy()
                if not y_df.empty:
                    def ml(r):
                        s, cur = str(r[cn]).strip(), str(r['目前農地調查現況']).strip()
                        if s == "監測": return "持續" if "增量" in cur else "延長"
                        return "退場" if s == "正常" else s
                    y_df['標籤'] = y_df.apply(ml, axis=1)
                    y_counts = y_df.groupby(['調查方式', '標籤']).size().reset_index(name='筆數')
                    fig = px.treemap(y_counts, path=[px.Constant(f"{y}年"), '調查方式', '標籤'], values='筆數', color='標籤',
                                     color_discrete_map={'持續':'#FFB6C1','延長':'#ADD8E6','退場':'#90EE90','管制':'#FF3333'})
                    tree_cols[i].plotly_chart(fig, use_container_width=True)

    # --- B. 資料庫查詢與下載 ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 數據管理中心")
        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單", "📜 修改紀錄"])

        # 1. 總表
        with tabs[0]:
            st.subheader("🌾 農地現況總表")
            s_m = st.text_input("🔍 搜尋地號/序號/網格", key="master_search_box")
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號') + 1
                cols.insert(idx, cols.pop(cols.index('代表性顯示')))
                df_p = df_p[cols]
            if s_m: df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(s_m)).any(axis=1)]
            st.dataframe(df_p, height=800, use_container_width=True)

        # 2. 歷年結果
        with tabs[1]:
            st.subheader("📅 歷年調查紀錄")
            y_avail = sorted(df_history['調查年度'].unique(), reverse=True)
            sel_y = st.selectbox("選擇查詢年度", y_avail if y_avail else [113])
            y_res = df_history[df_history['調查年度'] == int(sel_y)].copy()
            if not y_res.empty:
                y_rich = y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left')
                st.dataframe(y_rich, height=600, use_container_width=True)

        # 3. 坵塊管理
        with tabs[2]:
            st.subheader("🏠 坵塊管理")
            blk_q = st.text_input("🔍 搜尋地號確認群組")
            if blk_q and blk_q in df_block['農地地段地號'].values:
                gid = df_block[df_block['農地地段地號']==blk_q].iloc[0]['農地群組編號']
                st.dataframe(df_block[df_block['農地群組編號']==gid])
            st.write("**現有對照清單：**")
            st.dataframe(df_block, height=400, use_container_width=True)

        # 4. 系統型農地清單 (方塊選單強化版)
        with tabs[3]:
            st.subheader("🌐 系統型農地監測看板")
            s_grid_top = st.text_input("🔍 全域搜尋網格編號 (例如: G2405)", key="s_grid_tab")
            
            f_cols = st.columns(3)
            # 定義顯示清單
            p_list = grid_uniq_list[grid_uniq_list['網格監測頻率'] == '持續']['網格編號'].tolist()
            l_list = grid_uniq_list[grid_uniq_list['網格監測頻率'] == '延長']['網格編號'].tolist()
            e_list = grid_uniq_list[grid_uniq_list['網格監測頻率'] == '退場']['網格編號'].tolist()

            with f_cols[0]:
                st.markdown('<div class="status-header bg-p">持續網格</div>', unsafe_allow_html=True)
                sel_p = st.selectbox("選取持續網格", ["--請選擇--"] + p_list, key="sel_p_tab")
            with f_cols[1]:
                st.markdown('<div class="status-header bg-l">延長網格</div>', unsafe_allow_html=True)
                sel_l = st.selectbox("選取延長網格", ["--請選擇--"] + l_list, key="sel_l_tab")
            with f_cols[2]:
                st.markdown('<div class="status-header bg-e">退場網格</div>', unsafe_allow_html=True)
                sel_e = st.selectbox("選取退場網格", ["--請選擇--"] + e_list, key="sel_e_tab")

            final_gid = s_grid_top if s_grid_top else None
            if not final_gid:
                if sel_p != "--請選擇--": final_gid = sel_p
                elif sel_l != "--請選擇--": final_gid = sel_l
                elif sel_e != "--請選擇--": final_gid = sel_e

            if final_gid:
                g_data = df_master[df_master['網格編號'] == clean_id(final_gid)].copy()
                if not g_data.empty:
                    st.info(f"📍 網格 {final_gid} 統計資訊")
                    sc = st.columns(4)
                    sc[0].metric("農地總筆數", len(g_data))
                    sc[1].metric("代表點", len(g_data[g_data['代表性']=='代表點']))
                    sc[2].metric("備用點", len(g_data[g_data['代表性']=='備用點']))
                    sc[3].metric("無法採樣/建物", len(g_data[g_data['農地監測狀態'].isin(['難以採樣','建物'])]))
                    g_data['代表性顯示'] = g_data.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
                    st.dataframe(g_data[['代表性顯示','農地序號','地段地號','目前農地調查現況','農地監測狀態']], use_container_width=True)

            st.divider()
            st.write("**系統型農地總清單**")
            grid_sys_master['代表性顯示'] = grid_sys_master.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(grid_sys_master, height=400, use_container_width=True)

        # 5. 個案型農地清單 (補齊指標與搜尋)
        with tabs[4]:
            st.subheader("📦 個案型農地監測看板")
            s_case_top = st.text_input("🔍 搜尋個案地號/序號/SGM", key="s_case_tab_box")
            
            # 六大看板 (修正階梯問題)
            c_labels = ["持續", "延長", "退場", "管制", "難以採樣", "建物"]
            c_metrics = st.columns(6)
            counts_map = case_data_master['顯示狀態'].value_counts()
            for i, lab in enumerate(c_labels):
                c_metrics[i].metric(lab, counts_map.get(lab, 0))
                with c_metrics[i].expander("名單"):
                    st.dataframe(case_data_master[case_data_master['顯示狀態']==lab][['農地序號','地段地號']])

            st.divider()
            if s_case_top: 
                case_data_master = case_data_master[case_data_master.astype(str).apply(lambda x: x.str.contains(s_case_top)).any(axis=1)]
            st.write("**個案型農地總清單**")
            case_data_master['代表性顯示'] = case_data_master.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(case_data_master, height=400, use_container_width=True)

        # 6. 修改紀錄
        with tabs[5]:
            st.subheader("📜 系統修改紀錄回溯")
            if os.path.exists(LOG_PATH): st.dataframe(pd.read_csv(LOG_PATH))
            else: st.info("無紀錄")

    # --- C/D/E 頁面保留 ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 115 年度預計調查篩選")
        f_list = df_master[(df_master['目前農地調查現況'] == '增量') | ((df_master['目前農地調查現況'] == '延長') & (df_master['最後調查年分'] <= 113))].copy()
        st.dataframe(f_list)

    elif menu == "新增年度調查結果":
        st.title("➕ 錄入年度數據")
        pwd = st.sidebar.text_input("密碼", type="password", key="pwd_add")
        if pwd == ADMIN_PASSWORD:
            sl = st.text_input("🔍 搜尋地號錄入")
            if sl:
                h = df_master[df_master['地段地號']==sl.strip()]
                if not h.empty: st.success(f"找到 {sl}，請填寫數據並判定...")
        else: st.warning("請輸入密碼")

    elif menu == "空間地圖檢視":
        st.title("🗺️ 衛星影像與網格監測圖")
        m = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        st_folium(m, width=1100, height=700)

else:
    st.error("Excel 載入失敗")


