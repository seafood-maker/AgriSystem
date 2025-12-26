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
st.set_page_config(page_title="彰化農地智慧管理系統", layout="wide", page_icon="🌾")

# 專業美化：表格標題黑化、Metric 卡片、按鈕
st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border: 1px solid #eee; }
    .stDownloadButton button { width: 100%; background-color: #2e7d32; color: white; border-radius: 5px; }
    .grid-info-box { padding: 15px; border: 1px solid #e6e9ef; border-radius: 10px; background-color: #fcfcfc; margin-bottom: 10px; }
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

# --- 2. 資料讀取引擎 (自動容錯) ---
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

def find_grid_by_coords(x, y, gdf):
    if gdf is None: return "未知"
    p = Point(x, y)
    match = gdf[gdf.contains(p)]
    return str(match.iloc[0]['網格號']) if not match.empty else "範圍外"

# --- 3. 側邊欄 ---
st.sidebar.title("🌿 系統導覽")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

# --- 4. 主程式邏輯 ---
if df_master is not None:
    # 全域指標計算
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
        col_g, col_c = st.columns(2)
        with col_g:
            st.subheader("🌐 系統型網格統計")
            grid_df = df_master.drop_duplicates('網格編號').copy()
            grid_df['網格監測頻率'] = grid_df['網格監測頻率'].fillna('無網格狀態')
            g_c = len(grid_df[grid_df['網格監測頻率'] == '持續'])
            g_e = len(grid_df[grid_df['網格監測頻率'] == '延長'])
            g_ex = len(grid_df[grid_df['網格監測頻率'] == '退場'])
            g_sum = g_c + g_e + g_ex
            gc = st.columns(5)
            gc[0].metric("持續", g_c); gc[1].metric("延長", g_e); gc[2].metric("退場", g_ex)
            gc[3].metric("有效網格合計", g_sum); gc[4].metric("無網格狀態", len(grid_df)-g_sum)
        
        with col_c:
            st.subheader("📦 個案型農地統計")
            case_data = df_master[~df_master['調查方式'].astype(str).str.contains('系統', na=False)].copy()
            c_map = {"增量":"持續", "延長":"延長", "正常":"退場", "難以採樣":"難以採樣", "管制":"管制", "建物":"建物"}
            case_data['顯示'] = case_data['目前農地調查現況'].map(c_map).fillna(case_data['目前農地調查現況'])
            cc = case_data['顯示'].value_counts()
            cc_cols = st.columns(6)
            for i, lab in enumerate(["持續", "延長", "退場", "管制", "難以採樣", "建物"]):
                cc_cols[i].metric(lab, cc.get(lab, 0))

        st.divider()
        st.subheader("🔍 網格快速查詢系統")
        gs = st.text_input("輸入網格號碼搜尋 (如: G001)")
        if gs:
            res = df_master[df_master['網格編號'] == clean_id(gs)]
            if not res.empty:
                st.write(f"網格 {gs} 內農地清單：")
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
                    def map_label(row):
                        s, cur = str(row[cn]).strip(), str(row['目前農地調查現況']).strip()
                        if s == "監測": return "持續" if "增量" in cur else "延長"
                        return "退場" if s == "正常" else s
                    y_df['樹狀標籤'] = y_df.apply(map_label, axis=1)
                    y_counts = y_df.groupby(['調查方式', '樹狀標籤']).size().reset_index(name='筆數')
                    fig = px.treemap(y_counts, path=[px.Constant(f"{y}年"), '調查方式', '樹狀標籤'], values='筆數', color='樹狀標籤',
                                     color_discrete_map={'持續':'#FFB6C1','延長':'#ADD8E6','退場':'#90EE90','管制':'#FF3333','建物':'#D3D3D3'})
                    tree_cols[i].plotly_chart(fig, use_container_width=True)

    # --- B. 資料庫查詢與下載 (深度整合版) ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 數據管理與查詢中心")
        admin_mode = False
        with st.sidebar.expander("🔐 管理權限驗證"):
            if st.text_input("輸入修正密碼", type="password") == ADMIN_PASSWORD:
                admin_mode = True; st.success("編輯模式已開啟")

        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單", "📜 修改紀錄"])

        with tabs[0]:
            st.subheader("🌾 農地現況總表")
            search_m = st.text_input("🔍 搜尋地號/序號/網格", key="master_search")
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            # 欄位移位：將代表性顯示移到農地序號右邊
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號') + 1
                cols.insert(idx, cols.pop(cols.index('代表性顯示')))
                df_p = df_p[cols]
            if search_m: df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(search_m)).any(axis=1)]
            
            if admin_mode: st.data_editor(df_p, height=800, use_container_width=True)
            else: st.dataframe(df_p, height=800, use_container_width=True)
            
            towrite = io.BytesIO()
            df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載全量總表 Excel", data=towrite.getvalue(), file_name="彰化農地總表.xlsx")

        with tabs[1]:
            st.subheader("📅 年度調查明細")
            years = sorted(df_history['調查年度'].unique(), reverse=True)
            sel_y = st.selectbox("請選擇查詢年度", years if years else [113])
            y_res = df_history[df_history['調查年度'] == int(sel_y)].copy()
            if not y_res.empty:
                y_rich = y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left')
                def da_note(row):
                    notes = [f"{m}({row.get(f'DA_{m}',0)}%)" for m in METALS if row.get(f'DA_{m}',0) > 20]
                    return " / ".join(notes) if notes else "穩定"
                y_rich['DA判定說明'] = y_rich.apply(da_note, axis=1)
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
                    last_num = df_block['農地群組編號'].str.extract('(\d+)').dropna().astype(int).max()[0]
                    next_gid = f"BLOCK_{str(last_num + 1).zfill(3)}"
                except: next_gid = "BLOCK_001"
                with st.form("new_blk"):
                    st.write(f"預定編號: **{next_gid}**")
                    lots_in = st.text_area("1. 請輸入地段地號清單 (每行一筆)")
                    rep_in = st.text_input("2. 指定哪一筆為『代表點』？")
                    if st.form_submit_button("確認建立"): st.info("已排程，請手動更新 Excel。")
            st.write("**現有對照清單：**")
            st.dataframe(df_block, height=400, use_container_width=True)

        with tabs[3]:
            st.subheader("🌐 系統型農地清單看板")
            # 全域搜尋
            sg_q = st.text_input("🔍 搜尋特定網格號碼 (不限頻率)", key="sg_q")
            sys_all = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
            
            # 分類區塊
            f_cols = st.columns(3)
            for i, f in enumerate(['持續', '延長', '退場']):
                with f_cols[i]:
                    sub_ids = sys_all[sys_all['網格監測頻率']==f]['網格編號'].unique()
                    st.metric(f"{f}網格數", len(sub_ids))
                    sel_g = st.selectbox(f"選取{f}網格", ["請選擇"] + list(sub_ids), key=f"sys_box_{f}")
            
            target_g = sg_q if sg_q else None
            if not target_g:
                for f in ['持續', '延長', '退場']:
                    if st.session_state.get(f"sys_box_{f}") and st.session_state.get(f"sys_box_{f}") != "請選擇":
                        target_g = st.session_state.get(f"sys_box_{f}")
            
            if target_g:
                g_data = df_master[df_master['網格編號']==target_g].copy()
                if not g_data.empty:
                    st.markdown(f"#### 📋 網格 {target_g} 詳細統計")
                    sc = st.columns(4)
                    sc[0].metric("總數", len(g_data)); sc[1].metric("代表點", len(g_data[g_data['代表性']=='代表點']))
                    sc[2].metric("備用點", len(g_data[g_data['代表性']=='備用點'])); sc[3].metric("無法採樣", len(g_data[g_data['農地監測狀態']=='難以採樣']))
                    g_data['代表性顯示'] = g_data.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
                    st.dataframe(g_data[['代表性顯示','農地序號','地段地號','農地監測狀態','目前農地調查現況']], use_container_width=True)

        with tabs[4]:
            st.subheader("📦 個案型農地監測看板")
            case_all = df_master[~df_master['調查方式'].str.contains('系統', na=False)].copy()
            c_f_cols = st.columns(3)
            c_t = {'持續':'增量', '延長':'延長', '退場':'正常'}
            for i, (lab, val) in enumerate(c_t.items()):
                with c_f_cols[i]:
                    sub_c = case_all[case_all['目前農地調查現況']==val]
                    st.metric(f"{lab}個案數", len(sub_c))
                    if st.checkbox(f"展開{lab}名單"): st.dataframe(sub_c[['農地序號','地段地號','農地監測狀態']])
            st.write("**個案型總清單：**")
            case_all['代表性顯示'] = case_all.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(case_all[['代表性顯示','地段地號','網格編號','目前農地調查現況']], height=400)

        with tabs[5]:
            st.subheader("📜 系統修改紀錄回溯")
            if os.path.exists(LOG_PATH): st.dataframe(pd.read_csv(LOG_PATH).sort_values(by="修改時間", ascending=False))
            else: st.info("無紀錄")

    # --- C. 新年度篩選 ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 115 年度調查初步篩選")
        f_list = df_master[(df_master['目前農地調查現況'] == '增量') | ((df_master['目前農地調查現況'] == '延長') & (df_master['最後調查年分'] <= 113))].copy()
        st.dataframe(f_list[['網格編號','地段地號','目前農地調查現況','代表性','最後調查年分']])

    # --- D. 新增年度調查結果 (DA 與 3M 判定) ---
    elif menu == "新增年度調查結果":
        st.title("➕ 錄入年度採樣數據")
        pwd = st.sidebar.text_input("管理員密碼", type="password")
        if pwd == ADMIN_PASSWORD:
            sl = st.text_input("🔍 輸入地號搜尋")
            if sl:
                h = df_master[df_master['地段地號']==sl.strip()]
                if not h.empty:
                    r = h.iloc[0]
                    with st.form("en"):
                        c1, c2 = st.columns(2)
                        nx, ny = c1.number_input("實測 X", value=float(r['TWD97_X'])), c2.number_input("實測 Y", value=float(r['TWD97_Y']))
                        if np.sqrt((nx-r['TWD97_X'])**2 + (ny-r['TWD97_Y'])**2) > 3:
                            st.warning("⚠️ 偏移 > 3M"); st.info(f"網格判定: {find_grid_by_coords(nx, ny, gdf_grid)}")
                        v = {}
                        for m in METALS:
                            col1, col2 = st.columns(2)
                            v[m] = col2.number_input(f"{m}(全量)", min_value=0.0) or col1.number_input(f"{m}(XRF)", min_value=0.0)
                        if st.form_submit_button("執行自動判定"):
                            final = "正常"; das = {}
                            for m in METALS:
                                init = r.get(f'初始_{m}', 0); da = ((v[m]-init)/init*100) if init > 0 else 0
                                das[m] = f"{da:.1f}%"
                                if v[m] > df_settings.loc[m,'管制標準']: final = "管制"
                                elif init > df_settings.loc[m,'管制標準'] and da > df_settings.loc[m,'上升標準 (DA門檻)']:
                                    if final != "管制": final = "增量"
                            st.success(f"建議判定: {final}"); st.write("DA細節:", das)
        else: st.warning("請解鎖權限")

    # --- E. 空間地圖 ---
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
                sd = 4 if "個案" in str(r['調查方式']) else 3
                mon_s = str(r['農地監測狀態'])
                if mon_s == "管制": sd, c = 6, "red"
                elif mon_s == "建物": sd, c = 6, "black"
                elif mon_s == "難以採樣": sd, c = 6, "purple"
                elif str(r['代表性']) == "備用點": sd, c = 4, "white"
                else:
                    inv = str(r['目前農地調查現況'])
                    c = "red" if "增量" in inv else "blue" if "延長" in inv else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=sd, radius=6, color=c, fill=True, popup=f"{r['地段地號']}").add_to(m)
            except: continue
        st_folium(m, width=1100, height=700)
else:
    st.error("系統初始化失敗")
