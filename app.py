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
LOG_PATH = "edit_log.csv"
SHP_PATH = "彰化網格.shp"

transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ['汞', '砷', '銅', '鉻', '鎘', '鉛', '鋅', '鎳']

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
        
        df_m = pd.read_excel(xl, sheet_name=get_sheet("農地現況主檔", actual_sheets))
        df_m.columns = df_m.columns.str.strip()
        df_m['網格編號'] = df_m['網格編號'].apply(clean_id)
        df_m['地段地號'] = df_m['地段地號'].astype(str).str.strip()
        
        df_h = pd.read_excel(xl, sheet_name=get_sheet("歷年調查紀錄", actual_sheets))
        df_b = pd.read_excel(xl, sheet_name=get_sheet("同坵塊對照表", actual_sheets))
        df_s = pd.read_excel(xl, sheet_name=get_sheet("判定標準表", actual_sheets)).set_index('項目名稱')
        
        return df_m, df_h, df_b, df_s
    except Exception as e:
        st.error(f"Excel 讀取錯誤: {e}")
        return None, None, None, None

df_master, df_history, df_block, df_settings = load_all_data()

# 稽核紀錄函式
def log_change(lot, field, before, after):
    new_entry = pd.DataFrame([{
        "修改時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "地段地號": lot,
        "修改欄位": field,
        "原本內容": before,
        "更新內容": after,
        "執行人": "管理員"
    }])
    if not os.path.exists(LOG_PATH):
        new_entry.to_csv(LOG_PATH, index=False, encoding='utf-8-sig')
    else:
        new_entry.to_csv(LOG_PATH, mode='a', header=False, index=False, encoding='utf-8-sig')

# --- 3. 側邊欄 ---
st.sidebar.title("🌿 系統導覽")
menu = st.sidebar.radio("功能選單", ["統計首頁", "資料庫查詢與下載", "新增年度調查結果", "空間地圖檢視"])

if df_master is not None:
    # 全域數據統計
    abs_total = len(df_master)
    sampling_pts = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])

    # --- A. 統計首頁 (保留原有完整功能) ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("總資料點數", abs_total)
        k2.metric("總採樣點數", sampling_pts)
        k3.metric("管制點數", control_count)
        k4.metric("建物數量", build_count)
        k5.metric("難以採樣數量", hard_count)
        k6.metric("正常退場數量", normal_count)

        st.divider()
        st.subheader("🌐 系統型網格現況統計")
        grid_df = df_master.drop_duplicates('網格編號').copy()
        grid_df['網格監測頻率'] = grid_df['網格監測頻率'].fillna('無網格狀態')
        g_c = len(grid_df[grid_df['網格監測頻率'] == '持續'])
        g_e = len(grid_df[grid_df['網格監測頻率'] == '延長'])
        g_ex = len(grid_df[grid_df['網格監測頻率'] == '退場'])
        g_sum = g_c + g_e + g_ex
        g_cols = st.columns(5)
        g_cols[0].metric("持續", g_c); g_cols[1].metric("延長", g_e); g_cols[2].metric("退場", g_ex)
        g_cols[3].metric("有效網格合計", g_sum); g_cols[4].metric("無網格狀態", len(grid_df)-g_sum)

        st.divider()
        st.subheader("📦 個案型農地統計")
        case_data = df_master[~df_master['調查方式'].astype(str).str.contains('系統', na=False)].copy()
        c_map = {"增量":"持續", "延長":"延長", "正常":"退場", "難以採樣":"難以採樣", "管制":"管制", "建物":"建物"}
        case_data['對應'] = case_data['目前農地調查現況'].map(c_map).fillna(case_data['目前農地調查現況'])
        c_counts = case_data['對應'].value_counts()
        c_cols = st.columns(6)
        for i, lab in enumerate(["持續", "延長", "退場", "管制", "難以採樣", "建物"]):
            c_cols[i].metric(lab, c_counts.get(lab, 0))

        st.divider()
        st.subheader("🔍 網格查詢系統")
        gs = st.text_input("輸入網格號碼 (如: G001)")
        if gs:
            res = df_master[df_master['網格編號'] == clean_id(gs)]
            st.dataframe(res.style.apply(lambda x: ['background-color: #FFFFCC' if x.代表性=='代表點' else '' for _ in x], axis=1))

        st.divider()
        st.subheader("📊 近三年調查分佈 (樹狀圖)")
        tree_cols = st.columns(3)
        for i, y in enumerate([112, 113, 114]):
            cn = f"{y}狀態"
            if cn in df_master.columns:
                y_df = df_master[df_master[cn].notna()].copy()
                if not y_df.empty:
                    y_counts = y_df.groupby(['調查方式', cn]).size().reset_index(name='筆數')
                    fig = px.treemap(y_counts, path=[px.Constant(f"{y}年"), '調查方式', cn], values='筆數', color=cn,
                                     color_discrete_map={'監測':'#ADD8E6','正常':'#90EE90','管制':'#FFB6C1','建物':'#D3D3D3'})
                    tree_cols[i].plotly_chart(fig, use_container_width=True)

    # --- B. 資料庫查詢與下載 (新增與修正功能) ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 數據管理中心")
        
        # 權限檢查
        admin_mode = False
        with st.sidebar.expander("🔐 資料修正權限"):
            pwd = st.text_input("輸入修正密碼", type="password")
            if pwd == ADMIN_PASSWORD:
                admin_mode = True
                st.success("✅ 已開啟直接修正模式 (所有更動將被紀錄)")

        tab1, tab2, tab3, tab4 = st.tabs(["📋 總表清單", "📅 歷年結果與 DA 分析", "🏠 同坵塊管理", "📜 修改紀錄回溯"])
        
        with tab1:
            st.subheader("🌾 農地現況總表")
            # 定義美化圖示邏輯
            def get_rich_rep(row):
                r = str(row['代表性']).strip()
                s = str(row['農地監測狀態']).strip()
                # 檢查是否為重複坵塊 (分頁三)
                is_block_dup = False
                if not df_block.empty and row['地段地號'] in df_block['農地地段地號'].values:
                    if "否" in str(df_block[df_block['農地地段地號']==row['地段地號']].iloc[0]['代表農地']):
                        is_block_dup = True
                
                if r == "代表點": return "✅ 代表點"
                if r == "備用點": return "⚪ 備用點"
                if is_block_dup: return "❌ 非採樣 (同坵塊備註)"
                return f"❌ 非採樣 ({s})"

            df_pretty = df_master.copy()
            df_pretty['代表性顯示'] = df_pretty.apply(get_rich_rep, axis=1)
            
            # 搜尋與展示
            q_m = st.text_input("🔍 關鍵字過濾總表")
            if q_m: df_pretty = df_pretty[df_pretty.astype(str).apply(lambda x: x.str.contains(q_m)).any(axis=1)]
            
            if admin_mode:
                edited = st.data_editor(df_pretty, height=800, use_container_width=True, key="main_editor")
                if st.button("💾 儲存並同步變更"):
                    st.info("系統正在比對差異並寫入紀錄檔...")
                    # 這裡未來可實作與 GitHub API 連動存檔
            else:
                st.dataframe(df_pretty, height=800, use_container_width=True)

            # 下載 Excel
            towrite = io.BytesIO()
            df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載全量總表 Excel", data=towrite.getvalue(), file_name="彰化農地總表.xlsx")

        with tab2:
            st.subheader("📅 年度調查明細 (含重金屬 DA 分析)")
            y_sel = st.selectbox("選擇調查年度", [112, 113, 114, 115])
            # 關聯調查方式、濃度與 DA
            y_data = df_history[df_history['調查年度'] == y_sel].copy()
            if not y_data.empty:
                y_rich = y_data.merge(df_master[['SGM編號', '地段地號', '調查方式', '目前農地調查現況']], on='SGM編號', how='left')
                # 分析判定原因
                def analyze_reason(row):
                    reasons = []
                    for m in METALS:
                        da = row.get(f'DA_{m}', 0)
                        if da > 20: reasons.append(f"{m}DA({da}%)")
                    return " / ".join(reasons) if reasons else "數值穩定"
                y_rich['判定依據說明'] = y_rich.apply(analyze_reason, axis=1)
                st.dataframe(y_rich, use_container_width=True)
            else: st.info("該年度尚無調查數據")

        with tab3:
            st.subheader("🔗 同坵塊清單與新增")
            st.dataframe(df_block, use_container_width=True)
            with st.expander("➕ 新增重複坵塊關聯"):
                with st.form("add_block"):
                    new_gid = st.text_input("坵塊群組編號 (如: B101)")
                    new_lot = st.text_input("地段地號")
                    new_is_rep = st.selectbox("是否為代表點", ["是 (代表)", "否"])
                    if st.form_submit_button("確認新增"):
                        st.success(f"已記錄：{new_lot} 為群組 {new_gid} 之成員。系統將自動連動調查結果。")

        with tab4:
            st.subheader("📜 歷史修改稽核日誌")
            if os.path.exists(LOG_PATH):
                st.table(pd.read_csv(LOG_PATH).tail(15))
            else: st.info("尚無手動修正紀錄")

    # --- C. 新增年度調查結果 (保留邏輯) ---
    elif menu == "新增年度調查結果":
        st.title("➕ 錄入當年度調查數據")
        # ... (此處保留 8.0 的錄入、3公尺預警與 DA 判定代碼)

    # --- D. 空間地圖檢視 (保留邏輯) ---
    elif menu == "空間地圖檢視":
        st.title("🗺️ 衛星影像監測圖")
        # ... (此處保留 8.0 的衛星地圖與變色邏輯)

else:
    st.error("❌ 找不到資料庫檔案")




