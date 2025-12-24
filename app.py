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

# 座標轉換器
transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)

# 數據清洗
def clean_id(val):
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

# 民國日期轉換
def get_minguo_date():
    now = datetime.now()
    return f"民國 {now.year - 1911} 年 {now.month} 月 {now.day} 日"

# --- 2. 資料讀取 ---
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

df_master, df_history, df_block, df_settings = load_all_data()

# --- 3. 功能導覽選單 ---
st.sidebar.title("🌿 系統選單")
menu = st.sidebar.radio(
    "功能導覽", 
    ["統計首頁", "資料庫查詢與下載", "新增年度調查結果", "空間地圖檢視"]
)

if df_master is not None:
    # 預先計算基礎指標 (避免 NameError)
    abs_total = len(df_master)
    sampling_pts = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制')])
    build_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物')])
    hard_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣')])
    normal_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常')])

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

        # 2. 系統型網格現況
        st.subheader("🌐 系統型網格現況統計")
        grid_df = df_master.drop_duplicates('網格編號')
        g_cont = len(grid_df[grid_df['網格監測頻率'] == '持續'])
        g_ext = len(grid_df[grid_df['網格監測頻率'] == '延長'])
        g_exited = len(grid_df[grid_df['網格監測頻率'] == '退場'])
        g_total_active = g_cont + g_ext + g_exited

        g1, g2, g3, g4, g5 = st.columns(5)
        g1.metric("網格總數", len(grid_df))
        g2.metric("持續網格", g_cont)
        g3.metric("延長網格", g_ext)
        g4.metric("退場網格", g_exited)
        g5.metric("有效網格合計", g_total_active)

        # 3. 個案型農地現況 (對照映射)
        st.subheader("📦 個案型農地現況統計")
        case_df = df_master[df_master['調查方式'] == '個案型農地'].copy()
        status_map = {'增量': '持續', '延長': '延長', '正常': '退場', '難以採樣': '難以採樣', '管制': '管制', '建物': '建物'}
        case_df['顯示狀態'] = case_df['目前農地調查現況'].map(status_map)
        case_counts = case_df['顯示狀態'].value_counts()

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        for idx, lab in enumerate(['持續', '延長', '退場', '管制', '難以採樣', '建物']):
            st.columns(6)[idx].metric(lab, case_counts.get(lab, 0))

        st.divider()

        # 4. 網格查詢系統
        st.subheader("🔍 網格查詢系統")
        grid_search = st.text_input("請輸入網格號碼 (例如: G001)")
        if grid_search:
            grid_res = df_master[df_master['網格編號'] == clean_id(grid_search)]
            if not grid_res.empty:
                st.dataframe(grid_res)
            else:
                st.warning("查無此網格編號")

        st.divider()

        # 5. 近三年調查結果樹狀圖 (修正版)
        st.subheader("📊 近三年調查分佈")
        years = [112, 113, 114]
        tree_cols = st.columns(3)
        
        for i, year in enumerate(years):
            with tree_cols[i]:
                col_name = f"{year}狀態"
                if col_name in df_master.columns:
                    # 數據清洗：排除空值並確保型態為字串
                    plot_data = df_master[df_master[col_name].notna()].copy()
                    if not plot_data.empty:
                        plot_data['調查方式'] = plot_data['調查方式'].astype(str)
                        plot_data[col_name] = plot_data[col_name].astype(str)
                        # 增加一個計數欄位
                        plot_data['筆數'] = 1
                        
                        try:
                            fig = px.treemap(
                                plot_data, 
                                path=[px.Constant(f"{year}年度總計"), '調查方式', col_name], 
                                values='筆數',
                                color=col_name,
                                color_discrete_map={'監測':'#ADD8E6', '正常':'#90EE90', '管制':'#FFB6C1', '建物':'#D3D3D3', '(?)': '#EEEEEE'}
                            )
                            fig.update_layout(margin=dict(t=30, l=10, r=10, b=10))
                            st.plotly_chart(fig, use_container_width=True)
                        except Exception as e:
                            st.error(f"{year} 繪圖失敗: 請檢查該年度狀態欄位內容")
                    else:
                        st.write(f"⚪ {year} 年度無調查數據")

    # --- B. 資料庫查詢與下載 ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 資料庫查詢與下載中心")
        tab1, tab2, tab3 = st.tabs(["📊 總表清單", "📅 歷年調查結果", "🔗 同坵塊關聯表"])

        with tab1:
            st.dataframe(df_master, use_container_width=True)
            # 下載總表
            towrite = io.BytesIO()
            df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載全量總表 (.xlsx)", data=towrite.getvalue(), file_name="彰化農地總表.xlsx")

        with tab2:
            s_year = st.selectbox("選擇下載年度", sorted(df_history['調查年度'].unique() if not df_history.empty else [114]))
            hist_res = df_history[df_history['調查年度'] == s_year]
            st.dataframe(hist_res)
            
            towrite_h = io.BytesIO()
            hist_res.to_excel(towrite_h, index=False, engine='xlsxwriter')
            st.download_button(f"📥 下載 {s_year} 調查紀錄", data=towrite_h.getvalue(), file_name=f"{s_year}_調查結果.xlsx")

        with tab3:
            search_lot = st.text_input("搜尋地段地號查看關聯坵塊")
            if search_lot:
                match = df_block[df_block['農地地段地號'] == search_lot]
                if not match.empty:
                    gid = match.iloc[0]['農地群組編號']
                    all_plots = df_block[df_block['農地群組編號'] == gid]['農地地段地號'].tolist()
                    st.dataframe(df_master[df_master['地段地號'].isin(all_plots)])
                else:
                    st.warning("查無同坵塊關聯資料")

    # 保留原本的 新增 與 地圖 邏輯
    elif menu == "新增年度調查結果":
        st.title("➕ 新增年度調查結果")
        pwd = st.sidebar.text_input("後台密碼", type="password")
        if pwd == ADMIN_PASSWORD:
            st.success("權限解鎖")
            # (此處插入您原本的新增資料表單邏輯)
        else:
            st.warning("請輸入密碼以進行操作")

    elif menu == "空間地圖檢視":
        st.title("🗺️ 空間地圖檢視")
        # (此處插入您原本的 Folium 衛星圖邏輯)

else:
    st.error("❌ 讀取 Excel 失敗")
