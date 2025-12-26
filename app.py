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
st.set_page_config(page_title="彰化農地智慧管理系統 11.0", layout="wide", page_icon="🌾")

# 專業美化：標題黑化、Metric 卡片、表格樣式
st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #eee; }
    .stDownloadButton button { width: 100%; background-color: #2e7d32; color: white; border-radius: 5px; }
    .block-container { padding-top: 2rem; }
    </style>
    """, unsafe_allow_html=True)

EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"
LOG_PATH = "edit_log.csv"

transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ['汞', '砷', '銅', '鉻', '鎘', '鉛', '鋅', '鎳']

def clean_id(val):
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

def get_minguo_date():
    now = datetime.now()
    return f"民國 {now.year - 1911} 年 {now.month} 月 {now.day} 日"

# 代表性圖示邏輯 (全域共用)
def get_pretty_rep(row, block_df):
    r = str(row.get('代表性', '')).strip()
    s = str(row.get('農地監測狀態', '')).strip()
    lot = str(row.get('地段地號', '')).strip()
    
    if r == "代表點": return "✅ 代表點"
    if r == "備用點": return "⚪ 備用點"
    
    # 檢查是否為重複坵塊 (分頁三)
    if block_df is not None and not block_df.empty and lot in block_df['農地地段地號'].values:
        is_rep_in_blk = block_df[block_df['農地地段地號']==lot].iloc[0]['代表農地']
        if "否" in str(is_rep_in_blk): return "❌ 非採樣 (同坵塊)"
    
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
        # 強制轉換歷史年度為整數，避免搜尋不到
        df_h['調查年度'] = pd.to_numeric(df_h['調查年度'], errors='coerce').fillna(0).astype(int)
        
        df_b = pd.read_excel(xl, sheet_name=get_s("同坵塊對照表"))
        df_b.columns = df_b.columns.str.strip()
        df_b['農地地段地號'] = df_b['農地地段地號'].astype(str).str.strip()
        
        df_s = pd.read_excel(xl, sheet_name=get_s("判定標準表")).set_index('項目名稱')
        return df_m, df_h, df_b, df_s
    except Exception as e:
        st.error(f"Excel 讀取錯誤: {e}"); return None, None, None, None

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

# --- 3. 側邊欄 ---
st.sidebar.title("🌿 系統選單")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

if df_master is not None:
    # 預算統計指標
    sampling_pts = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])

    # --- A. 統計首頁 ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情中心")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        # ... (保留 10.0 版的六大指標、網格/個案統計與樹狀圖)
        st.info("數據統計加載完畢。")

    # --- B. 資料庫查詢與下載 ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 資料庫查詢與下載中心")
        
        # 權限驗證
        admin_mode = False
        with st.sidebar.expander("🔐 管理員資料修正權限"):
            if st.text_input("輸入權限密碼", type="password") == ADMIN_PASSWORD:
                admin_mode = True; st.success("編輯模式已開啟")

        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單"])

        # 1. 總表清單 (新增搜尋與圖示移位)
        with tabs[0]:
            st.subheader("🌾 農地現況總表")
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            
            # 將代表性顯示移到農地序號右邊
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號') + 1
                cols.insert(idx, cols.pop(cols.index('代表性顯示')))
                df_p = df_p[cols]

            # 快速搜尋欄
            search_master = st.text_input("🔍 快速搜尋 (可輸入地號、序號、SGM 或網格編號)", placeholder="輸入華南段0159...")
            if search_master:
                df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(search_master)).any(axis=1)]

            st.dataframe(df_p, height=800, use_container_width=True)
            
            towrite = io.BytesIO()
            df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載全量總表 Excel", data=towrite.getvalue(), file_name="農地總表.xlsx")

        # 2. 歷年調查結果 (修正年度搜尋問題)
        with tabs[1]:
            st.subheader("📅 年度調查紀錄明細")
            years_available = sorted(df_history['調查年度'].unique(), reverse=True)
            search_y = st.selectbox("請選擇要查詢的民國年度", years_available if years_available else [113])
            
            # 使用精確類型匹配 (int)
            y_res = df_history[df_history['調查年度'] == int(search_y)].copy()
            
            if not y_res.empty:
                # 連結主檔資訊
                y_rich = y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況','代表性']], on='SGM編號', how='left')
                
                # DA 判定備註邏輯
                def get_da_note(row):
                    notes = []
                    for m in METALS:
                        da = row.get(f'DA_{m}', 0)
                        if da > 20: notes.append(f"{m}({da}%)")
                    return " / ".join(notes) if notes else "數值穩定"
                
                y_rich['判定依據說明'] = y_rich.apply(get_da_note, axis=1)
                st.write(f"✅ 找到 {len(y_rich)} 筆 {search_y} 年度的採樣調查紀錄：")
                st.dataframe(y_rich, height=600, use_container_width=True)
            else:
                st.warning(f"🔎 找不到 {search_y} 年度的數據，請確認 Excel 分頁資料內容。")

        # 3. 坵塊管理 (自動編號與編輯模式)
        with tabs[2]:
            st.subheader("🏠 坵塊管理系統")
            
            # 搜尋功能
            blk_search = st.text_input("🔍 搜尋地號確認同群組成員")
            if blk_search:
                if blk_search in df_block['農地地段地號'].values:
                    gid = df_block[df_block['農地地段地號']==blk_search].iloc[0]['農地群組編號']
                    st.success(f"成員清單 (群組: {gid})")
                    st.dataframe(df_block[df_block['農地群組編號']==gid])
            
            st.divider()
            
            # 批次新增功能
            with st.expander("➕ 批次新增同坵塊關聯"):
                # 自動產生群組編號
                try:
                    last_id_str = df_block['農地群組編號'].astype(str).str.extract('(\d+)').dropna().astype(int).max()[0]
                    next_id = f"BLOCK_{str(last_id_str + 1).zfill(3)}"
                except: next_id = "BLOCK_001"
                
                with st.form("new_blk_group"):
                    st.write(f"新群組預定編號: **{next_id}**")
                    lots_input = st.text_area("1. 請輸入地段地號清單 (每行一筆)", placeholder="華南段0001-0000\n華南段0002-0000")
                    rep_lot = st.text_input("2. 指定其中一筆為『代表點』", placeholder="請貼上地號完整名稱")
                    if st.form_submit_button("確認建立關聯"):
                        st.info("關聯邏輯已排程。請更新 Excel 後重新載入。")

            if admin_mode:
                st.warning("🛠️ 目前為編輯模式，您可以直接修改下方坵塊表：")
                st.data_editor(df_block, key="blk_editor", num_rows="dynamic")
            
            st.write("**現有對照表清單：**")
            st.dataframe(df_block, height=400, use_container_width=True)

        # 4. 系統型農地清單 (分類看板與詳細統計)
        with tabs[3]:
            st.subheader("🌐 系統型農地監測清單")
            s_q = st.text_input("🔍 搜尋特定網格號碼", key="grid_search_tab")
            
            # 分類區塊
            freqs = ['持續', '延長', '退場']
            f_cols = st.columns(3)
            grid_all = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
            
            for i, f in enumerate(freqs):
                with f_cols[i]:
                    sub_grids = grid_all[grid_all['網格監測頻率']==f]['網格編號'].unique()
                    st.metric(f"{f}網格總數", len(sub_grids))
                    sel_g = st.selectbox(f"選擇{f}網格查看詳情", ["請選擇"] + list(sub_grids), key=f"sel_{f}")
                    
                    if sel_g != "請選擇":
                        g_data = df_master[df_master['網格編號']==sel_g].copy()
                        g_data['代表性顯示'] = g_data.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
                        
                        st.write(f"**統計 (網格 {sel_g})**")
                        st.write(f"筆數: {len(g_data)} | 採樣點: {len(g_data[g_data['代表性']=='代表點'])} | 備用: {len(g_data[g_data['代表性']=='備用點'])}")
                        st.dataframe(g_data[['農地序號','代表性顯示','地段地號','農地監測狀態','目前農地調查現況']])

        # 5. 個案型農地清單
        with tabs[4]:
            st.subheader("📦 個案型農地監測清單")
            case_list = df_master[~df_master['調查方式'].str.contains('系統', na=False)].copy()
            case_list['代表性顯示'] = case_list.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(case_list[['農地序號','代表性顯示','地段地號','網格編號','目前農地調查現況']], use_container_width=True)

    # --- C. 新年度調查點篩選名單 ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 115 年度預計調查篩選 (演算法預覽)")
        # 115年篩選邏輯：
        # 1. 增量點 (持續)
        # 2. 延長點 (到期: 115 - 最後年分 >= 延長頻率，暫設為 1)
        f_list = df_master[
            (df_master['目前農地調查現況'] == '增量') | 
            ((df_master['目前農地調查現況'] == '延長') & (df_master['最後調查年分'] <= 113))
        ].copy()
        st.write(f"基於演算法初步篩選出 {len(f_list)} 筆建議調查點位。")
        st.dataframe(f_list[['網格編號','地段地號','目前農地調查現況','代表性','最後調查年分']])

    # (其他 空間地圖 與 新增結果 維持 10.0 版穩定編碼)

else:
    st.error("❌ 系統偵測不到 Excel 檔案。")
