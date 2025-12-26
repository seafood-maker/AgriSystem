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
import base64
from sqlalchemy import create_engine, text

# --- 1. 系統設定與權限 ---
ADMIN_PASSWORD = "ET23597010"
EDITOR_PASSWORD = "editor_changhua"
st.set_page_config(page_title="彰化農地智慧管理系統 20.0", layout="wide", page_icon="🌾")

# 專業美化 CSS
st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .grid-card { padding: 15px; border-radius: 15px; text-align: center; margin-bottom: 10px; border: 1px solid #ddd; font-weight: bold; }
    .bg-p { background-color: #FFB6C1; color: #721c24; } /* 持續-淡紅 */
    .bg-l { background-color: #ADD8E6; color: #004085; } /* 延長-淡藍 */
    .bg-e { background-color: #90EE90; color: #155724; } /* 退場-淡綠 */
    .stDownloadButton button { width: 100%; background-color: #2e7d32; color: white; border-radius: 5px; }
    </style>
    """, unsafe_allow_html=True)

EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"
LOG_PATH = "edit_log.csv"

# 座標轉換引擎 (TWD97 -> WGS84)
transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ["汞", "砷", "銅", "鉻", "鎘", "鉛", "鋅", "鎳"]

# --- 2. 輔助函式 ---
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

# --- 3. 資料載入與資料庫連動 ---
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
        df_h['調查年度'] = pd.to_numeric(df_h['調查年度'], errors='coerce').fillna(0).astype(int)
        df_b = pd.read_excel(xl, sheet_name=get_s("同坵塊對照表"))
        df_b['農地地段地號'] = df_b['農地地段地號'].astype(str).str.strip()
        df_s = pd.read_excel(xl, sheet_name=get_s("判定標準表")).set_index('項目名稱')
        return df_m, df_h, df_b, df_s
    except: return None, None, None, None

@st.cache_data
def load_grid_shp():
    if os.path.exists(SHP_PATH):
        try:
            gdf = gpd.read_file(SHP_PATH)
            if gdf.crs is None or gdf.crs.to_epsg() != 3826: gdf.set_crs(epsg=3826, allow_override=True, inplace=True)
            if '網格號' in gdf.columns: gdf['網格號'] = gdf['網格號'].apply(clean_id)
            return gdf
        except: return None
    return None

df_master, df_history, df_block, df_settings = load_all_data()
gdf_grid = load_grid_shp()

# --- 4. 側邊欄權限與導覽 ---
def get_role():
    st.sidebar.markdown("### 🔐 系統權限控管")
    role = st.sidebar.selectbox("切換角色", ["一般閱讀者", "後台編輯者", "系統管理者"])
    if role == "一般閱讀者": return "viewer"
    pwd = st.sidebar.text_input("輸入權限密碼", type="password")
    if role == "後台編輯者" and pwd == EDITOR_PASSWORD: return "editor"
    if role == "系統管理者" and pwd == ADMIN_PASSWORD: return "admin"
    if pwd: st.sidebar.error("密碼錯誤")
    return "viewer"

current_role = get_role()
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

# --- 5. 主程式頁面邏輯 ---
if df_master is not None:
    # 全域指標
    abs_total = len(df_master)
    sampling_pts = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])

    # --- 頁面 A：統計首頁 ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        
        # 1. 六大頂列指標
        k = st.columns(6)
        k[0].metric("總資料點數", abs_total); k[1].metric("總採樣點數", sampling_pts)
        k[2].metric("管制點數", control_count); k[3].metric("建物數量", build_count)
        k[4].metric("難以採樣數量", hard_count); k[5].metric("正常退場數量", normal_count)

        st.divider()
        # 2. 系統型網格區塊 (彩色方塊清單)
        st.subheader("🌐 系統型網格現況 (點選看詳情)")
        grid_uniq = df_master.drop_duplicates('網格編號')
        gr_p = grid_uniq[grid_uniq['網格監測頻率']=='持續']['網格編號'].tolist()
        gr_l = grid_uniq[grid_uniq['網格監測頻率']=='延長']['網格編號'].tolist()
        gr_e = grid_uniq[grid_uniq['網格監測頻率']=='退場']['網格編號'].tolist()

        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f'<div class="grid-card bg-p">持續網格 ({len(gr_p)})</div>', unsafe_allow_html=True)
            sel_p = st.selectbox("持續網格清單", ["--未選擇--"] + gr_p)
        with c2:
            st.markdown(f'<div class="grid-card bg-l">延長網格 ({len(gr_l)})</div>', unsafe_allow_html=True)
            sel_l = st.selectbox("延長網格清單", ["--未選擇--"] + gr_l)
        with c3:
            st.markdown(f'<div class="grid-card bg-e">退場網格 ({len(gr_e)})</div>', unsafe_allow_html=True)
            sel_e = st.selectbox("退場網格清單", ["--未選擇--"] + gr_e)

        chosen_g = sel_p if sel_p != "--未選擇--" else sel_l if sel_l != "--未選擇--" else sel_e if sel_e != "--未選擇--" else None
        if chosen_g:
            g_data = df_master[df_master['網格編號']==chosen_g].copy()
            st.info(f"📍 網格 {chosen_g} 統計：農地 {len(g_data)} 筆 | 代表點 {len(g_data[g_data['代表性']=='代表點'])} | 備用點 {len(g_data[g_data['代表性']=='備用點'])}")
            g_data['代表性顯示'] = g_data.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(g_data[['代表性顯示','農地序號','地段地號','目前農地調查現況','農地監測狀態']])

        st.divider()
        # 3. 個案型 6 大看板
        st.subheader("📦 個案型農地監測看板")
        case_data = df_master[~df_master['調查方式'].astype(str).str.contains('系統', na=False)].copy()
        c_map = {"增量":"持續", "延長":"延長", "正常":"退場", "管制":"管制", "建物":"建物", "難以採樣":"難以採樣"}
        case_data['對應'] = case_data['目前農地調查現況'].map(c_map).fillna(case_data['目前農地調查現況'])
        counts = case_data['對應'].value_counts()
        cc = st.columns(6)
        labels = ["持續", "延長", "退場", "管制", "建物", "難以採樣"]
        for i, lab in enumerate(labels):
            cc[i].metric(lab, counts.get(lab, 0))

        st.divider()
        # 4. 樹狀圖 (112-114)
        st.subheader("📊 近三年調查分佈 (樹狀圖)")
        t_cols = st.columns(3)
        for i, y in enumerate([112, 113, 114]):
            col_n = f"{y}狀態"
            if col_n in df_master.columns:
                y_df = df_master[df_master[col_n].notna()].copy()
                if not y_df.empty:
                    y_df['樹狀標籤'] = y_df.apply(lambda r: "持續" if "增量" in str(r['目前農地調查現況']) and r[col_n]=="監測" else ("退場" if r[col_n]=="正常" else r[col_n]), axis=1)
                    y_counts = y_df.groupby(['調查方式', '樹狀標籤']).size().reset_index(name='筆數')
                    fig = px.treemap(y_counts, path=[px.Constant(f"{y}年"), '調查方式', '樹狀標籤'], values='筆數', color='樹狀標籤',
                                     color_discrete_map={'持續':'#FFB6C1','延長':'#ADD8E6','退場':'#90EE90','管制':'#FF3333'})
                    t_cols[i].plotly_chart(fig, use_container_width=True)

    # --- 頁面 B：資料庫查詢與下載 ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 數據管理中心")
        t = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單", "📜 修改紀錄"])
        
        with t[0]: # 總表
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號') + 1
                cols.insert(idx, cols.pop(cols.index('代表性顯示')))
                df_p = df_p[cols]
            s_q = st.text_input("🔍 關鍵字查詢總表")
            if s_q: df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(s_q)).any(axis=1)]
            st.dataframe(df_p, height=800, use_container_width=True)
            # 下載
            towrite = io.BytesIO()
            df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載全量總表", data=towrite.getvalue(), file_name="彰化農地總表.xlsx")

        with t[1]: # 歷年 (修正年度)
            y_avail = sorted(df_history['調查年度'].unique(), reverse=True)
            sel_y = st.selectbox("選擇調查年度", y_avail if y_avail else [113])
            y_res = df_history[df_history['調查年度']==int(sel_y)].copy()
            if not y_res.empty:
                st.dataframe(y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left'))
            else: st.warning("無數據")

        with t[2]: # 坵塊 (多筆新增)
            st.subheader("🏠 同坵塊群組管理")
            blk_search = st.text_input("🔍 搜尋地號找群組")
            if blk_search and blk_search in df_block['農地地段地號'].values:
                gid = df_block[df_block['農地地段地號']==blk_search].iloc[0]['農地群組編號']
                st.dataframe(df_block[df_block['農地群組編號']==gid])
            with st.expander("➕ 批次新增同坵塊關聯"):
                try: 
                    last_id = df_block['農地群組編號'].str.extract('(\d+)').dropna().astype(int).max()[0]
                    next_id = f"BLOCK_{str(last_id+1).zfill(3)}"
                except: next_id = "BLOCK_001"
                with st.form("blk_form"):
                    st.write(f"預計編號: **{next_id}**")
                    lots_in = st.text_area("1. 貼上地號 (每行一筆)"); rep_in = st.text_input("2. 指定代表點")
                    if st.form_submit_button("建立關聯"): st.info("已加入排程，請手動更新 Excel。")
            st.write("**對照清單：**")
            st.dataframe(df_block, height=400)

        with t[3]: # 系統型看板 (底端總表)
            st.subheader("🌐 系統型詳細清單")
            sys_df = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
            sys_df['代表性顯示'] = sys_df.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(sys_df, height=600)

        with t[4]: # 個案型看板 (含搜尋)
            st.subheader("📦 個案型詳細清單")
            s_c = st.text_input("🔍 搜尋個案地號")
            c_df = case_data.copy()
            c_df['代表性顯示'] = c_df.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            if s_c: c_df = c_df[c_df['地段地號'].str.contains(s_c)]
            st.dataframe(c_df, height=600)

    # --- 頁面 C：篩選名單 ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 115 年度篩選預演")
        f_list = df_master[(df_master['目前農地調查現況'] == '增量') | ((df_master['目前農地調查現況'] == '延長') & (df_master['最後調查年分'] <= 113))].copy()
        st.dataframe(f_list[['網格編號','地段地號','目前農地調查現況','代表性','最後調查年分']])

    # --- 頁面 D：新增調查 (DA 判定與照片) ---
    elif menu == "新增年度調查結果":
        st.title("➕ 錄入採樣數據")
        if current_role == "viewer": st.warning("請登入後台權限")
        else:
            sl = st.text_input("🔍 搜尋地號錄入")
            if sl:
                h = df_master[df_master['地段地號']==sl.strip()]
                if not h.empty:
                    r = h.iloc[0]
                    with st.form("entry"):
                        st.subheader(f"📍 編輯：{sl}")
                        c1, c2 = st.columns(2)
                        nx, ny = c1.number_input("實測 X", value=float(r['TWD97_X'])), c2.number_input("實測 Y", value=float(r['TWD97_Y']))
                        if np.sqrt((nx-r['TWD97_X'])**2 + (ny-r['TWD97_Y'])**2) > 3:
                            st.warning("⚠️ 偏移 > 3M"); st.info(f"網格判定: {find_grid_by_coords(nx, ny, gdf_grid)}")
                        st.file_uploader("現勘/採樣照片上傳")
                        v = {}
                        for m in METALS:
                            col1, col2 = st.columns(2)
                            v[m] = col2.number_input(f"{m}(全量)", min_value=0.0) or col1.number_input(f"{m}(XRF)", min_value=0.0)
                        if st.form_submit_button("執行判定"):
                            final = "正常"; das = {}
                            for m in METALS:
                                init = r.get(f'初始_{m}', 0); da = ((v[m]-init)/init*100) if init > 0 else 0
                                if v[m] > df_settings.loc[m,'管制標準']: final = "管制"
                                elif init > df_settings.loc[m,'管制標準'] and da > df_settings.loc[m,'上升標準 (DA門檻)']:
                                    if final != "管制": final = "增量"
                            st.success(f"建議判定: {final}")

    # --- 頁面 E：空間地圖 ---
    elif menu == "空間地圖檢視":
        st.title("🗺️ 衛星影像監測圖")
        m = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        if gdf_grid is not None:
            gs = df_master.drop_duplicates('網格編號')[['網格編號', '網格監測頻率']]
            merged = gdf_grid.to_crs(epsg=4326).merge(gs, left_on='網格號', right_on='網格編號', how='left')
            def get_c(f): return '#FFB6C1' if '持續' in str(f) else '#ADD8E6' if '延長' in str(f) else '#90EE90' if '退場' in str(f) else '#F8F8F8'
            folium.GeoJson(merged, style_function=lambda x: {'fillColor': get_c(x['properties'].get('網格監測頻率')), 'color': 'white', 'weight': 1, 'fillOpacity': 0.4}).add_to(m)
        smp = df_master.sample(min(800, len(df_master)))
        for _, r in smp.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                sd, c = (4, "blue") if "個案" in str(r['調查方式']) else (3, "green")
                if str(r['農地監測狀態']) == "管制": sd, c = 6, "red"
                elif str(r['代表性']) == "備用點": sd, c = 4, "white"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=sd, radius=6, color=c, fill=True, popup=f"{r['地段地號']}").add_to(m)
            except: continue
        st_folium(m, width=1100, height=700)

else:
    st.error("Excel 載入失敗，請確認檔案。")

