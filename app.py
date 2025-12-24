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
        df_m = pd.read_excel(xl, "農地現況主檔").applymap(lambda x: x.strip() if isinstance(x, str) else x)
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
    # --- A. 統計首頁 ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.info(f"📅 當前時間：{get_minguo_date()}")

        # 1. 頂列統計指標
        abs_total = len(df_master)
        sampling_pts = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
        control_count = len(df_master[df_master['農地監測狀態'] == '管制'])
        build_count = len(df_master[df_master['農地監測狀態'] == '建物'])
        hard_count = len(df_master[df_master['農地監測狀態'] == '難以採樣'])
        normal_count = len(df_master[df_master['農地監測狀態'] == '正常'])

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
        g_none = len(grid_df[grid_df['網格監測頻率'].isna() | (grid_df['網格監測頻率'] == '無網格狀態')])
        g_total_active = g_cont + g_ext + g_exited

        g1, g2, g3, g4, g5 = st.columns(5)
        g1.metric("網格總數 (含無狀態)", len(grid_df))
        g2.metric("持續網格", g_cont)
        g3.metric("延長網格", g_ext)
        g4.metric("退場網格", g_exited)
        g5.metric("有效網格總計", g_total_active, help="持續+延長+退場")

        # 3. 個案型農地現況 (對照映射)
        st.subheader("📦 個案型農地現況統計")
        case_df = df_master[df_master['調查方式'] == '個案型農地'].copy()
        
        # 映射轉換
        status_map = {
            '增量': '持續', '延長': '延長', '正常': '退場', 
            '難以採樣': '難以採樣', '管制': '管制', '建物': '建物'
        }
        case_df['顯示狀態'] = case_df['目前農地調查現況'].map(status_map)
        case_counts = case_df['顯示狀態'].value_counts()

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("持續", case_counts.get('持續', 0))
        c2.metric("延長", case_counts.get('延長', 0))
        c3.metric("退場", case_counts.get('退場', 0))
        c4.metric("管制", case_counts.get('管制', 0))
        c5.metric("難以採樣", case_counts.get('難以採樣', 0))
        c6.metric("建物", case_counts.get('建物', 0))

        st.divider()

        # 4. 網格查詢系統
        st.subheader("🔍 網格查詢系統")
        grid_search = st.text_input("請輸入網格號碼 (例如: G001)")
        if grid_search:
            grid_res = df_master[df_master['網格編號'] == grid_search]
            if not grid_res.empty:
                # 標示代表點
                def color_rep(val):
                    color = 'background-color: #FFFFCC' if val == '代表點' else ''
                    return color
                st.dataframe(grid_res.style.applymap(color_rep, subset=['代表性']))
            else:
                st.warning("查無此網格編號")

        st.divider()

        # 5. 近三年調查結果樹狀圖
        st.subheader("📊 近三年調查分佈 (樹狀圖)")
        # 假設當前民國 114 年，近三年為 112, 113, 114
        years = [112, 113, 114]
        tree_cols = st.columns(3)
        
        for i, year in enumerate(years):
            with tree_cols[i]:
                # 抓取該年度有調查的筆數 (從主檔狀態欄位判斷)
                col_name = f"{year}狀態"
                if col_name in df_master.columns:
                    year_data = df_master[df_master[col_name].notna()].copy()
                    year_data['調查類別'] = year_data['調查方式']
                    year_data['結果'] = year_data[col_name]
                    
                    if not year_data.empty:
                        fig = px.treemap(
                            year_data, 
                            path=['調查類別', '結果'], 
                            title=f"{year}年度 (總筆數: {len(year_data)})",
                            color='結果',
                            color_discrete_map={'監測':'#ADD8E6', '正常':'#90EE90', '管制':'#FFB6C1', '建物':'#D3D3D3'}
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.write(f"{year} 年度無調查紀錄")

    # --- B. 資料庫查詢與下載 ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 資料庫查詢與下載中心")
        st.info("💡 此頁面資料為唯讀模式。如需更新資料，請至『新增年度調查結果』。")

        tab1, tab2, tab3 = st.tabs(["📊 總表清單", "📅 歷年調查結果", "🔗 同坵塊關聯表"])

        with tab1:
            st.subheader("農地現況總表")
            st.dataframe(df_master, use_container_width=True)
            # 下載按鈕
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_master.to_excel(writer, index=False, sheet_name='總表')
            st.download_button("📥 下載總表 Excel", data=output.getvalue(), file_name="彰化農地總表.xlsx")

        with tab2:
            st.subheader("歷年調查結果查詢")
            search_year = st.selectbox("請選擇年度", [112, 113, 114, 115])
            # 從 History 表篩選
            hist_res = df_history[df_history['調查年度'] == search_year]
            st.dataframe(hist_res)
            
            output_h = io.BytesIO()
            with pd.ExcelWriter(output_h, engine='xlsxwriter') as writer:
                hist_res.to_excel(writer, index=False, sheet_name=f'{search_year}調查紀錄')
            st.download_button(f"📥 下載 {search_year} 年度調查結果", data=output_h.getvalue(), file_name=f"{search_year}_調查結果.xlsx")

        with tab3:
            st.subheader("多筆農地同坵塊清單")
            search_plot = st.text_input("輸入地段地號搜尋同坵塊農地")
            if search_plot:
                # 1. 找出該地號屬於哪個群組
                group_match = df_block[df_block['農地地段地號'] == search_plot]
                if not group_match.empty:
                    gid = group_match.iloc[0]['農地群組編號']
                    # 2. 找出該群組所有地號
                    all_lots = df_block[df_block['農地群組編號'] == gid]['農地地段地號'].tolist()
                    # 3. 從主檔抓出這些地號的所有農地
                    related_plots = df_master[df_master['地段地號'].isin(all_lots)]
                    st.success(f"找到同群組 ({gid}) 關聯農地：")
                    st.dataframe(related_plots)
                else:
                    st.warning("此地號無關聯同坵塊紀錄")

    # 其他頁面保留原有架構，暫不改動
    elif menu == "新增年度調查結果":
        st.title("➕ 新增年度調查結果")
        st.warning("請輸入管理員密碼解鎖後台...")
        # (此處保留 7.0 版本的密碼鎖與錄入邏輯)

    elif menu == "空間地圖檢視":
        st.title("🗺️ 空間地圖檢視")
        # (此處保留 7.0 版本的衛星地圖與網格變色邏輯)

else:
    st.error("❌ 系統偵測不到 Excel 檔案，請確認『彰化農地管理資料庫.xlsx』已上傳。")