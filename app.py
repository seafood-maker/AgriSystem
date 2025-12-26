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

# 專業美化 CSS：包含標題黑化、Metric 卡片、分類方塊、篩選卡片樣式
st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .stDownloadButton button { width: 100%; background-color: #1b5e20; color: white; border-radius: 5px; }
    /* 系統型方塊樣式 */
    .status-header { padding: 10px; border-radius: 10px; text-align: center; font-weight: bold; font-size: 18px; border: 1px solid #ddd; margin-bottom: 5px; }
    .bg-p { background-color: #FFB6C1; color: #721c24; } /* 持續-淡紅 */
    .bg-l { background-color: #ADD8E6; color: #004085; } /* 延長-淡藍 */
    .bg-e { background-color: #90EE90; color: #155724; } /* 退場-淡綠 */
    /* 篩選計畫卡片 */
    .filter-card { padding: 20px; border-radius: 15px; background-color: #f0f4f8; border-left: 5px solid #2e7d32; margin-bottom: 20px; }
    </style>
    """, unsafe_allow_html=True)

EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"
LOG_PATH = "edit_log.csv"

# 座標轉換器 (TWD97 -> WGS84)
transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ["汞", "砷", "銅", "鉻", "鎘", "鉛", "鋅", "鎳"]

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
        if '延長頻率' not in df_m.columns: df_m['延長頻率'] = 2 # 預設頻率
        
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

# --- 4. 全局數據預處理 (防 NameError) ---
if df_master is not None:
    # 全域指標
    abs_total = len(df_master)
    sampling_pts_count = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])

    # 系統型看板數據
    grid_sys_all = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
    grid_unique_df = grid_sys_all.drop_duplicates('網格編號').copy()
    grid_unique_df['網格監測頻率'] = grid_unique_df['網格監測頻率'].fillna('無狀態')
    
    # 個案型看板數據
    case_data_master = df_master[~df_master['調查方式'].str.contains('系統', na=False)].copy()
    c_map = {"增量":"持續", "延長":"延長", "正常":"退場", "管制":"管制", "建物":"建物", "難以採樣":"難以採樣"}
    case_data_master['顯示狀態'] = case_data_master['目前農地調查現況'].map(c_map).fillna(case_data_master['目前農地調查現況'])

    # --- A. 統計首頁 ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        k = st.columns(6)
        k[0].metric("總資料點數", abs_total); k[1].metric("總採樣點數", sampling_pts_count)
        k[2].metric("管制點數", control_pts_count); k[3].metric("建物數量", build_pts_count)
        k[4].metric("難以採樣數量", hard_pts_count); k[5].metric("正常退場數量", normal_pts_count)
        
        st.divider()
        st.subheader("🌐 系統型網格現況統計")
        g_c = len(grid_unique_df[grid_unique_df['網格監測頻率'] == '持續'])
        g_l = len(grid_unique_df[grid_unique_df['網格監測頻率'] == '延長'])
        g_e = len(grid_unique_df[grid_unique_df['網格監測頻率'] == '退場'])
        g_sum = g_c + g_l + g_e
        g_none = len(grid_unique_df) - g_sum
        gc = st.columns(5)
        gc[0].metric("持續網格", g_c); gc[1].metric("延長網格", g_l); gc[2].metric("退場網格", g_e)
        gc[3].metric("有效網格合計", g_sum); gc[4].metric("無網格狀態", g_none)

        st.divider()
        st.subheader("📦 個案型農地現況統計")
        cc_map = case_data_master['顯示狀態'].value_counts()
        cc_cols = st.columns(6)
        for i, lab in enumerate(["持續", "延長", "退場", "管制", "難以採樣", "建物"]):
            cc_cols[i].metric(lab, cc_map.get(lab, 0))

        st.divider()
        st.subheader("🔍 網格查詢系統")
        gs = st.text_input("輸入網格號碼搜尋 (如: G2405)")
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
        admin_mode = False
        with st.sidebar.expander("🔐 管理權限"):
            if st.text_input("修正密碼", type="password") == ADMIN_PASSWORD: admin_mode = True; st.success("編輯模式開啟")

        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單", "📜 修改紀錄"])

        with tabs[0]: # 1. 總表
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號') + 1
                cols.insert(idx, cols.pop(cols.index('代表性顯示')))
                df_p = df_p[cols]
            sm = st.text_input("🔍 搜尋地號/序號/網格", key="m_s")
            if sm: df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(sm)).any(axis=1)]
            st.dataframe(df_p, height=800, use_container_width=True)
            # 下載
            towrite = io.BytesIO()
            df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載 Excel", data=towrite.getvalue(), file_name="彰化農地總表.xlsx")

        with tabs[1]: # 2. 歷年結果
            st.subheader("📅 歷年調查紀錄明細")
            y_avail = sorted(df_history['調查年度'].unique(), reverse=True)
            sel_y = st.selectbox("選擇查詢年度", y_avail if y_avail else [113])
            y_res = df_history[df_history['調查年度'] == int(sel_y)].copy()
            if not y_res.empty:
                y_rich = y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left')
                st.dataframe(y_rich, height=600, use_container_width=True)
            else: st.warning("查無紀錄")

        with tabs[2]: # 3. 坵塊管理 (自動編號)
            st.subheader("🏠 坵塊關聯搜尋")
            blk_q = st.text_input("🔍 搜尋地號確認群組")
            if blk_q and blk_q in df_block['農地地段地號'].values:
                gid = df_block[df_block['農地地段地號']==blk_q].iloc[0]['農地群組編號']
                st.dataframe(df_block[df_block['農地群組編號']==gid])
            with st.expander("➕ 批次新增同坵塊關聯"):
                try: 
                    last_num = df_block['農地群組編號'].str.extract('(\d+)').dropna().astype(int).max()[0]
                    next_id = f"BLOCK_{str(last_num + 1).zfill(3)}"
                except: next_id = "BLOCK_001"
                with st.form("new_blk"):
                    st.write(f"預計編號: **{next_id}**")
                    lots_in = st.text_area("1. 請輸入地號清單 (每行一筆)"); rep_in = st.text_input("2. 指定代表點")
                    if st.form_submit_button("確認建立"): st.info("已加入排程。")
            st.write("**現有對照清單：**")
            st.dataframe(df_block, height=400)

        with tabs[3]: # 4. 系統型看板 (分類與搜尋)
            st.subheader("🌐 系統型農地監測看板")
            s_grid_top = st.text_input("🔍 搜尋特定網格號碼", key="s_gt")
            grid_list = grid_unique_df
            p_ids = grid_list[grid_list['網格監測頻率'] == '持續']['網格編號'].tolist()
            l_ids = grid_list[grid_list['網格監測頻率'] == '延長']['網格編號'].tolist()
            e_ids = grid_list[grid_list['網格監測頻率'] == '退場']['網格編號'].tolist()
            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                st.markdown('<div class="status-header bg-p">持續網格</div>', unsafe_allow_html=True)
                sel_p = st.selectbox("選取持續網格", ["未選"] + p_ids, key="sp")
            with sc2:
                st.markdown('<div class="status-header bg-l">延長網格</div>', unsafe_allow_html=True)
                sel_l = st.selectbox("選取延長網格", ["未選"] + l_ids, key="sl")
            with sc3:
                st.markdown('<div class="status-header bg-e">退場網格</div>', unsafe_allow_html=True)
                sel_e = st.selectbox("選取退場網格", ["未選"] + e_ids, key="se")
            
            chosen = s_grid_top if s_grid_top else (sel_p if sel_p != "未選" else sel_l if sel_l != "未選" else sel_e if sel_e != "未選" else None)
            if chosen:
                g_data = df_master[df_master['網格編號']==clean_id(chosen)].copy()
                if not g_data.empty:
                    st.info(f"📍 網格 {chosen} 詳情看板")
                    ssc = st.columns(4)
                    ssc[0].metric("總筆數", len(g_data)); ssc[1].metric("代表點", len(g_data[g_data['代表性']=='代表點']))
                    ssc[2].metric("備用點", len(g_data[g_data['代表性']=='備用點'])); ssc[3].metric("無法採樣", len(g_data[g_data['農地監測狀態']=='難以採樣']))
                    g_data['代表性顯示'] = g_data.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
                    st.dataframe(g_data[['代表性顯示','農地序號','地段地號','目前農地調查現況','農地監測狀態']], use_container_width=True)
            st.divider()
            st.write("**系統型農地總表**")
            grid_sys_all['代表性顯示'] = grid_sys_all.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(grid_sys_all, height=400)

        with tabs[4]: # 5. 個案型看板 (搜尋與分類)
            st.subheader("📦 個案型農地監測看板")
            scq = st.text_input("🔍 搜尋個案地號/序號", key="scq")
            clabs = ["持續", "延長", "退場", "管制", "難以採樣", "建物"]
            c_cols = st.columns(6)
            for i, lab in enumerate(clabs):
                sub = case_data_master[case_data_master['顯示狀態']==lab]
                c_cols[i].metric(lab, len(sub))
                with c_cols[i].expander("名單"): st.dataframe(sub[['農地序號','地段地號']])
            if scq: case_data_master = case_data_master[case_data_master.astype(str).apply(lambda x: x.str.contains(scq)).any(axis=1)]
            st.divider()
            st.write("**個案型農地總表**")
            case_data_master['代表性顯示'] = case_data_master.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(case_data_master, height=400)

        with tabs[5]: # 6. 修改紀錄
            st.subheader("📜 系統稽核日誌")
            if os.path.exists(LOG_PATH): st.dataframe(pd.read_csv(LOG_PATH))
            else: st.info("尚無紀錄")

    # --- C. 新年度調查點篩選名單 (重點演算法) ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 115 年度調查點位篩選")
        target_year = 115
        st.markdown(f'<div class="filter-card"><b>115 年度篩選邏輯：</b><br>1. 持續型(增量) 必選<br>2. 延長型(到期) 依頻率篩選<br>3. 若代表點失效則自動標示需遞補。</div>', unsafe_allow_html=True)
        
        # 1. 個案型名單
        c_plan = case_data_master[
            (case_data_master['顯示狀態'] == '持續') | 
            ((case_data_master['顯示狀態'] == '延長') & (target_year - case_data_master['最後調查年分'] >= case_data_master['延長頻率']))
        ].copy()
        c_plan['篩選分類'] = '個案型獨立監測'
        
        # 2. 系統型網格名單 (遞補邏輯)
        g_plan_list = []
        for gid in grid_unique_df['網格編號'].unique():
            g_data = df_master[df_master['網格編號']==gid]
            freq = str(g_data['網格監測頻率'].iloc[0])
            last_y = g_data['最後調查年分'].max()
            if freq == '持續' or (freq == '延長' and (target_year - last_y >= 2)):
                reps = g_data[g_data['代表性'] == '代表點']
                active_reps = reps[~reps['農地監測狀態'].isin(['管制','建物','難以採樣'])]
                if len(active_reps) < 3:
                    backups = g_data[g_data['代表性']=='備用點'].sort_values('農地序號').head(3-len(active_reps))
                    final_g = pd.concat([active_reps, backups])
                    final_g['篩選分類'] = f'網格({freq})-需遞補'
                else:
                    final_g = active_reps.copy(); final_g['篩選分類'] = f'網格({freq})'
                g_plan_list.append(final_g)
        
        s_plan = pd.concat(g_plan_list) if g_plan_list else pd.DataFrame()
        full_plan = pd.concat([c_plan, s_plan])
        st.subheader(f"📍 115 年度建議清單 (共 {len(full_plan)} 筆)")
        st.dataframe(full_plan[['網格編號','地段地號','目前農地調查現況','代表性','篩選分類']], height=600, use_container_width=True)

    # --- D. 新增年度調查結果 ---
    elif menu == "新增年度調查結果":
        st.title("➕ 錄入採樣數據與 DA 判定")
        pwd = st.sidebar.text_input("管理員密碼", type="password")
        if pwd == ADMIN_PASSWORD:
            sl = st.text_input("🔍 第一步：搜尋欲錄入地號")
            if sl:
                h = df_master[df_master['地段地號']==sl.strip()]
                if not h.empty:
                    r = h.iloc[0]
                    with st.form("en"):
                        st.subheader(f"📍 編輯：{sl}")
                        c1, c2 = st.columns(2)
                        nx, ny = c1.number_input("實測 X", value=float(r['TWD97_X'])), c2.number_input("實測 Y", value=float(r['TWD97_Y']))
                        if np.sqrt((nx-r['TWD97_X'])**2 + (ny-r['TWD97_Y'])**2) > 3:
                            st.warning("⚠️ 座標偏移 > 3M"); st.info(f"網格判定: {find_grid_by_coords(nx, ny, gdf_grid)}")
                        v = {}
                        for m in METALS:
                            mc1, mc2 = st.columns(2)
                            v[m] = mc2.number_input(f"{m}(全量)", min_value=0.0) or mc1.number_input(f"{m}(XRF)", min_value=0.0)
                        if st.form_submit_button("執行判定"):
                            final_st = "正常"; das = {}
                            for m in METALS:
                                init = r.get(f'初始_{m}', 0); da = ((v[m]-init)/init*100) if init > 0 else 0
                                das[m] = f"{da:.1f}%"
                                if v[m] > df_settings.loc[m,'管制標準']: final_st = "管制"
                                elif init > df_settings.loc[m,'管制標準'] and da > df_settings.loc[m,'上升標準 (DA門檻)']:
                                    if final_st != "管制": final_st = "增量"
                            st.success(f"建議判定: {final_st}"); st.write("DA細節:", das)
        else: st.warning("請輸入密碼解鎖權限")

    # --- E. 空間地圖檢視 ---
    elif menu == "空間地圖檢視":
        st.title("🗺️ 衛星影像與網格著色圖")
        m = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri Imagery')
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
                else:
                    stv = str(r['目前農地調查現況'])
                    c = "red" if "增量" in stv else "blue" if "延長" in stv else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=sd, radius=6, color=c, fill=True, popup=f"{r['地段地號']}").add_to(m)
            except: continue
        st_folium(m, width=1100, height=700)
else:
    st.error("❌ 讀取資料庫失敗，請確認 Excel 檔案與分頁正確。")



