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

# 專業美化：標題黑化、Metric 卡片、下載按鈕
st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .stDownloadButton button { width: 100%; background-color: #2e7d32; color: white; border-radius: 5px; }
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

# 代表性圖示邏輯
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

def find_grid_by_coords(x, y, gdf):
    if gdf is None: return "未知"
    p = Point(x, y)
    match = gdf[gdf.contains(p)]
    return str(match.iloc[0]['網格號']) if not match.empty else "範圍外"

# --- 3. 側邊欄 ---
st.sidebar.title("🌿 系統選單")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

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
        st.subheader("🌐 系統型網格現況統計")
        grid_df = df_master.drop_duplicates('網格編號').copy()
        grid_df['網格監測頻率'] = grid_df['網格監測頻率'].fillna('無網格狀態')
        g_c = len(grid_df[grid_df['網格監測頻率'] == '持續'])
        g_e = len(grid_df[grid_df['網格監測頻率'] == '延長'])
        g_ex = len(grid_df[grid_df['網格監測頻率'] == '退場'])
        g_sum = g_c + g_e + g_ex
        gc = st.columns(5)
        gc[0].metric("持續", g_c); gc[1].metric("延長", g_e); gc[2].metric("退場", g_ex)
        gc[3].metric("有效網格合計", g_sum); gc[4].metric("無網格狀態", len(grid_df)-g_sum)

        st.divider()
        st.subheader("📦 個案型農地現況統計")
        case_data = df_master[~df_master['調查方式'].astype(str).str.contains('系統', na=False)].copy()
        c_map = {"增量":"持續", "延長":"延長", "正常":"退場", "難以採樣":"難以採樣", "管制":"管制", "建物":"建物"}
        case_data['顯示'] = case_data['目前農地調查現況'].map(c_map).fillna(case_data['目前農地調查現況'])
        cc = case_data['顯示'].value_counts()
        cc_cols = st.columns(6)
        for i, lab in enumerate(["持續", "延長", "退場", "管制", "難以採樣", "建物"]):
            cc_cols[i].metric(lab, cc.get(lab, 0))

        st.divider()
        st.subheader("🔍 網格查詢系統")
        gs = st.text_input("輸入網格號碼搜尋 (如: G001)")
        if gs:
            res = df_master[df_master['網格編號'] == clean_id(gs)]
            if not res.empty:
                st.dataframe(res.style.apply(lambda x: ['background-color: #FFFFCC' if x.代表性=='代表點' else '' for _ in x], axis=1))
            else: st.warning("查無此網格")

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

    # --- B. 資料庫查詢與下載 ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 資料庫查詢與下載中心")
        admin_mode = False
        with st.sidebar.expander("🔐 管理員資料修正權限"):
            if st.text_input("輸入權限密碼", type="password") == ADMIN_PASSWORD:
                admin_mode = True; st.success("編輯模式已開啟")

        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單"])

        with tabs[0]:
            st.subheader("🌾 農地現況總表")
            search_master = st.text_input("🔍 快速搜尋 (地號/序號/網格)", key="m_search")
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號') + 1
                cols.insert(idx, cols.pop(cols.index('代表性顯示')))
                df_p = df_p[cols]
            if search_master: df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(search_master)).any(axis=1)]
            st.dataframe(df_p, height=800, use_container_width=True)
            towrite = io.BytesIO()
            df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載全量總表 Excel", data=towrite.getvalue(), file_name="農地總表.xlsx")

        with tabs[1]:
            st.subheader("📅 年度調查紀錄明細")
            years = sorted(df_history['調查年度'].unique(), reverse=True)
            search_y = st.selectbox("請選擇查詢年度", years if years else [113])
            y_res = df_history[df_history['調查年度'] == int(search_y)].copy()
            if not y_res.empty:
                y_rich = y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left')
                def get_da_note(row):
                    notes = [f"{m}({row.get(f'DA_{m}',0)}%)" for m in METALS if row.get(f'DA_{m}',0) > 20]
                    return " / ".join(notes) if notes else "穩定"
                y_rich['判定依據說明'] = y_rich.apply(get_da_note, axis=1)
                st.dataframe(y_rich, height=600, use_container_width=True)
            else: st.warning("找不到數據")

        with tabs[2]:
            st.subheader("🏠 坵塊管理系統")
            blk_q = st.text_input("🔍 搜尋地號確認群組成員")
            if blk_q and not df_block.empty:
                if blk_q in df_block['農地地段地號'].values:
                    gid = df_block[df_block['農地地段地號']==blk_q].iloc[0]['農地群組編號']
                    st.success(f"群組: {gid}"); st.dataframe(df_block[df_block['農地群組編號']==gid])
            
            with st.expander("➕ 批次新增同坵塊關聯"):
                try:
                    last_id = df_block['農地群組編號'].str.extract('(\d+)').dropna().astype(int).max()[0]
                    next_id = f"BLOCK_{str(last_id + 1).zfill(3)}"
                except: next_id = "BLOCK_001"
                with st.form("new_blk"):
                    st.write(f"預計編號: {next_id}")
                    lots_in = st.text_area("地號清單 (每行一筆)"); rep_in = st.text_input("指定代表點")
                    if st.form_submit_button("建立關聯"): st.info("已排程，請更新 Excel。")

            st.write("**現有對照清單：**")
            if admin_mode: st.data_editor(df_block, key="blk_edit", num_rows="dynamic")
            else: st.dataframe(df_block, height=400, use_container_width=True)

        with tabs[3]:
            st.subheader("🌐 系統型農地清單看板")
            f_cols = st.columns(3); grid_all = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
            for i, f in enumerate(['持續', '延長', '退場']):
                with f_cols[i]:
                    sub_gs = grid_all[grid_all['網格監測頻率']==f]['網格編號'].unique()
                    st.metric(f"{f}網格數", len(sub_gs))
                    sel_g = st.selectbox(f"檢視{f}網格", ["請選擇"] + list(sub_gs), key=f"sys_{f}")
                    if sel_g != "請選擇":
                        g_data = df_master[df_master['網格編號']==sel_g].copy()
                        g_data['代表性圖示'] = g_data.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
                        st.write(f"農地數: {len(g_data)} | 採樣: {len(g_data[g_data['代表性']=='代表點'])} | 備用: {len(g_data[g_data['代表性']=='備用點'])}")
                        st.dataframe(g_data[['代表性圖示','地段地號','目前農地調查現況']])

        with tabs[4]:
            st.subheader("📦 個案型農地監測清單")
            case_list = df_master[~df_master['調查方式'].str.contains('系統', na=False)].copy()
            case_list['代表性圖示'] = case_list.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(case_list[['代表性圖示','地段地號','網格編號','目前農地調查現況']], use_container_width=True)

    # --- C. 新年度篩選 ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 115 年度篩選預演")
        f_list = df_master[(df_master['目前農地調查現況'] == '增量') | ((df_master['目前農地調查現況'] == '延長') & (df_master['最後調查年分'] <= 113))].copy()
        st.dataframe(f_list[['網格編號','地段地號','目前農地調查現況','代表性','最後調查年分']])

    # --- D. 新增結果 ---
    elif menu == "新增年度調查結果":
        st.title("➕ 錄入年度數據")
        pwd = st.sidebar.text_input("管理員密碼", type="password")
        if pwd == ADMIN_PASSWORD:
            search_lot = st.text_input("🔍 搜尋地號")
            if search_lot:
                hits = df_master[df_master['地段地號']==search_lot.strip()]
                if not hits.empty:
                    row = hits.iloc[0]
                    with st.form("entry"):
                        c1, c2 = st.columns(2)
                        nx, ny = c1.number_input("實測 X", value=float(row['TWD97_X'])), c2.number_input("實測 Y", value=float(row['TWD97_Y']))
                        if np.sqrt((nx-row['TWD97_X'])**2 + (ny-row['TWD97_Y'])**2) > 3:
                            st.warning("誤差 > 3M"); st.info(f"網格: {find_grid_by_coords(nx, ny, gdf_grid)}")
                        v = {}
                        for m in METALS:
                            mc1, mc2 = st.columns(2)
                            v[m] = mc2.number_input(f"{m}(全量)", min_value=0.0) or mc1.number_input(f"{m}(XRF)", min_value=0.0)
                        if st.form_submit_button("執行判定"):
                            final_st = "正常"; da_list = {}
                            for m in METALS:
                                init = row.get(f'初始_{m}', 0); da = ((v[m]-init)/init*100) if init > 0 else 0
                                if v[m] > df_settings.loc[m,'管制標準']: final_st = "管制"
                                elif init > df_settings.loc[m,'管制標準'] and da > df_settings.loc[m,'上升標準 (DA門檻)']:
                                    if final_st != "管制": final_st = "增量"
                            st.write(f"建議: {final_st}")
        else: st.warning("請解鎖")

    # --- E. 空間地圖 ---
    elif menu == "空間地圖檢視":
        st.title("🗺️ 衛星影像監測圖")
        m = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        if gdf_grid is not None:
            grid_status = df_master.drop_duplicates('網格編號')[['網格編號', '網格監測頻率']]
            merged = gdf_grid.to_crs(epsg=4326).merge(grid_status, left_on='網格號', right_on='網格編號', how='left')
            def get_c(f): return '#FFB6C1' if '持續' in str(f) else '#ADD8E6' if '延長' in str(f) else '#90EE90' if '退場' in str(f) else '#F8F8F8'
            folium.GeoJson(merged, style_function=lambda x: {'fillColor': get_c(x['properties'].get('網格監測頻率')), 'color': 'white', 'weight': 1, 'fillOpacity': 0.4}).add_to(m)
        sample = df_master.sample(min(1000, len(df_master)))
        for _, r in sample.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                s = 4 if "個案" in str(r['調查方式']) else 3
                if str(r['農地監測狀態']) == "管制": s, c = 6, "red"
                elif str(r['農地監測狀態']) == "建物": s, c = 6, "black"
                elif str(r['農地監測狀態']) == "難以採樣": s, c = 6, "purple"
                elif str(r['代表性']) == "備用點": s, c = 4, "white"
                else:
                    stv = str(r['目前農地調查現況'])
                    c = "red" if "增量" in stv else "blue" if "延長" in stv else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=s, radius=6, color=c, fill=True, popup=f"{r['地段地號']}").add_to(m)
            except: continue
        st_folium(m, width=1100, height=700)
else:
    st.error("Excel 載入失敗")
