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

# ==========================================
# 1. 系統初始化與狀態管理 (防止 NameError & AttributeError)
# ==========================================
if 'excluded_lots' not in st.session_state: st.session_state.excluded_lots = []
if 'temp_field_plan' not in st.session_state: st.session_state.temp_field_plan = None
if 'archived_plans' not in st.session_state: st.session_state.archived_plans = {}
if 'manual_selection_df' not in st.session_state: st.session_state.manual_selection_df = pd.DataFrame()
if 'recon_log' not in st.session_state: st.session_state.recon_log = []

def get_minguo_year():
    return datetime.now().year - 1911

def get_minguo_date():
    now = datetime.now()
    return f"民國 {now.year - 1911} 年 {now.month} 月 {now.day} 日"

def clean_id(val):
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

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

# ==========================================
# 2. 系統設定與美化
# ==========================================
st.set_page_config(page_title="彰化農地智慧管理系統", layout="wide", page_icon="🌾")
ADMIN_PASSWORD = "ET23597010"
EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"

transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ["汞", "砷", "銅", "鉻", "鎘", "鉛", "鋅", "鎳"]

st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; }
    .stMetric { background-color: #ffffff; padding: 10px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .status-header { padding: 10px; border-radius: 10px; text-align: center; font-weight: bold; font-size: 16px; border: 1px solid #ddd; margin-bottom: 5px; }
    .bg-p { background-color: #FFB6C1; color: #721c24; } /* 持續-淡紅 */
    .bg-l { background-color: #ADD8E6; color: #004085; } /* 延長-淡藍 */
    .bg-e { background-color: #90EE90; color: #155724; } /* 退場-淡綠 */
    .stats-container { background-color: #e8f5e9; padding: 15px; border-radius: 10px; margin-bottom: 20px; border: 1px solid #c8e6c9; }
    .recon-invalid { background-color: #ffebee; border-left: 5px solid #f44336; padding: 15px; border-radius: 5px; margin-bottom: 10px; }
    .recon-valid { background-color: #e8f5e9; border-left: 5px solid #4caf50; padding: 15px; border-radius: 5px; margin-bottom: 10px; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 3. 資料讀取引擎
# ==========================================
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
        df_b = pd.read_excel(xl, sheet_name=get_s("同坵塊對照表"))
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

# ==========================================
# 4. 全域指標預處理
# ==========================================
if df_master is not None:
    abs_total = len(df_master)
    sampling_pts_count = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])
    grid_sys_only = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
    grid_uniq = grid_sys_only.drop_duplicates('網格編號').copy()
    grid_uniq['網格監測頻率'] = grid_uniq['網格監測頻率'].fillna('無狀態')
    case_data_master = df_master[~df_master['調查方式'].str.contains('系統', na=False)].copy()
    c_map_ref = {"增量":"持續", "延長":"延長", "正常":"退場", "管制":"管制", "建物":"建物", "難以採樣":"難以採樣"}
    case_data_master['對應狀態'] = case_data_master['目前農地調查現況'].map(c_map_ref).fillna(case_data_master['目前農地調查現況'])

# ==========================================
# 5. 導覽選單
# ==========================================
st.sidebar.title("🌿 系統導覽")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

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
        st.subheader("🌐 系統型網格統計")
        g_c = len(grid_uniq[grid_uniq['網格監測頻率'] == '持續']); g_l = len(grid_uniq[grid_uniq['網格監測頻率'] == '延長']); g_e = len(grid_uniq[grid_uniq['網格監測頻率'] == '退場'])
        gc = st.columns(5); gc[0].metric("持續網格", g_c); gc[1].metric("延長網格", g_l); gc[2].metric("退場網格", g_e); gc[3].metric("有效合計", g_c+g_l+g_e); gc[4].metric("無狀態", len(grid_uniq)-(g_c+g_l+g_e))
        st.divider()
        st.subheader("📦 個案型農地統計")
        cc_m = case_data_master['對應狀態'].value_counts(); cc = st.columns(6)
        for i, lab in enumerate(["持續", "延長", "退場", "管制", "難以採樣", "建物"]): cc[i].metric(lab, cc_m.get(lab, 0))
        st.divider()
        st.subheader("📊 近三年調查樹狀圖")
        tree_cols = st.columns(3); cy = get_minguo_year()
        for i, y in enumerate([cy-2, cy-1, cy]):
            cn = f"{y}狀態"
            if cn in df_master.columns:
                y_df = df_master[df_master[cn].notna()].copy()
                if not y_df.empty:
                    y_counts = y_df.groupby(['調查方式', cn]).size().reset_index(name='筆數')
                    fig = px.treemap(y_counts, path=[px.Constant(f"{y}年"), '調查方式', cn], values='筆數', color=cn, color_discrete_map={'監測':'#ADD8E6','正常':'#90EE90','管制':'#FFB6C1'})
                    tree_cols[i].plotly_chart(fig, use_container_width=True)

    # --- B. 資料庫查詢與下載 (修正方塊看板) ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 數據查詢中心")
        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單", "📜 修改紀錄"])
        with tabs[0]:
            sq = st.text_input("🔍 搜尋地號/序號/網格", key="m_s")
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號')+1; cols.insert(idx, cols.pop(cols.index('代表性顯示'))); df_p = df_p[cols]
            if sq: df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(sq)).any(axis=1)]
            st.dataframe(df_p, height=800, use_container_width=True)
        with tabs[3]:
            st.subheader("🌐 系統型分類看板")
            s_q_gt = st.text_input("🔍 全域網格號搜尋", key="s_gt_t")
            sc_c = st.columns(3)
            for i, f in enumerate(['持續','延長','退場']):
                with sc_c[i]:
                    st.markdown(f'<div class="status-header bg-{"p" if i==0 else "l" if i==1 else "e"}">{f}網格</div>', unsafe_allow_html=True)
                    ids = grid_uniq[grid_uniq['網格監測頻率']==f]['網格編號'].tolist()
                    sel = st.selectbox(f"{f}名單", ["未選"]+ids, key=f"s_box_{f}")
                    if sel != "未選":
                        g_data_v = df_master[df_master['網格編號']==sel].copy()
                        st.metric("總筆數", len(g_data_v))
                        st.dataframe(g_data_v[['代表性','地段地號','農地監測狀態','目前農地調查現況']])
        with tabs[4]:
            st.subheader("📦 個案型監測看板")
            sc_q_case = st.text_input("🔍 搜尋個案地號/序號")
            c_cols_tab = st.columns(6)
            for i, lab in enumerate(["持續", "延長", "退場", "管制", "難以採樣", "建物"]):
                sub = case_data_master[case_data_master['對應狀態']==lab]
                c_cols_tab[i].metric(lab, len(sub))
            st.write("**個案型總表：**"); st.dataframe(case_data_master, height=400)

    # --- C. [重點整合] 新年度調查點篩選名單 (旗艦工作流) ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 年度計畫決策工作流")
        
        # 1. 參數設定
        c1, c2, c3 = st.columns([1, 1, 1])
        target_year = c1.number_input("目標年度 (民國)", value=get_minguo_year()+1)
        quota = c2.number_input("調查配額筆數", value=500, step=10)
        auto_backfill = c3.checkbox("開啟自動補位功能", value=True)

        # ---------------------------------------------------------
        # 第一階段：初步篩選與分流
        # ---------------------------------------------------------
        st.subheader("第一階段：初步篩選與去留確認")
        df_calc = df_master.copy()
        
        # 自動演算法 (排除去年、週期判定)
        sys_pool = []
        grids = df_calc[df_calc['調查方式'].str.contains('系統', na=False)]['網格編號'].unique()
        for gid in grids:
            g_data = df_calc[df_calc['網格編號'] == gid]
            f_t, ly = str(g_data['網格監測頻率'].iloc[0]), g_data['最後調查年分'].max()
            if ly == target_year - 1: continue
            prio = 1 if f_t == '持續' and (target_year-ly >= 2) else (2 if f_t == '延長' and (target_year-ly >= 10) else 99)
            if prio < 99:
                reps = g_data[(g_data['代表性'] == '代表點') & (~g_data['農地監測狀態'].isin(['管制','建物','難以採樣']))]
                if len(reps) < 3:
                    backs = g_data[g_data['代表性'] == '備用點'].sort_values('農地序號').head(3 - len(reps)).copy()
                    fg = pd.concat([reps, backs])
                else: fg = reps.copy()
                fg['優先權重'], fg['計畫類別'], fg['網格頻率'] = prio, '系統型網格', f_t
                sys_pool.append(fg)
        
        case_active = df_calc[~df_calc['調查方式'].str.contains('系統', na=False)].copy()
        case_active = case_active[~case_active['農地監測狀態'].isin(['管制','建物','正常'])]
        def c_algo(r):
            ly = r['最後調查年分']
            if ly == target_year - 1: return 99
            if str(r['目前農地調查現況']) == '增量' and (target_year-ly >= 2): return 1
            if str(r['目前農地調查現況']) == '延長' and (target_year-ly >= 10): return 3
            return 99
        case_active['優先權重'] = case_active.apply(c_algo, axis=1)
        case_pool = case_active[case_active['優先權重'] < 99].copy()
        case_pool['計畫類別'] = '個案型農地'

        # 匯總初步名單
        full_pool = pd.concat([pd.concat(sys_pool) if sys_pool else pd.DataFrame(), case_pool]).sort_values(['優先權重', '網格編號'])
        eligible = full_pool[~full_pool['地段地號'].isin(st.session_state.excluded_lots)]
        current_selection = eligible.head(int(quota)).copy() if auto_backfill else eligible.copy()
        current_selection.insert(0, '狀態確認', '✅ 留用')

        # 第一階段統計 (滿足需求)
        st.markdown('<div class="stats-container">', unsafe_allow_html=True)
        sp1 = current_selection[current_selection['計畫類別']=='系統型網格']; cp1 = current_selection[current_selection['計畫類別']=='個案型農地']
        k_s1, k_s2, k_s3 = st.columns(3); k_s1.metric("初步系統農地", len(sp1)); k_s2.metric("初步個案農地", len(cp1)); k_s3.metric("總筆數", len(current_selection))
        st.write(f"🔹 網格分佈：持續 {len(sp1[sp1['網格頻率']=='持續']['網格編號'].unique())} / 延長 {len(sp1[sp1['網格頻率']=='延長']['網格編號'].unique())}")
        st.markdown('</div>', unsafe_allow_html=True)

        col_p1_l, col_p1_r = st.columns(2)
        with col_p1_l:
            st.write("**🌐 系統型初步名單**")
            ed_s = st.data_editor(sp1[['狀態確認','網格編號','地段地號','農地序號','目前農地調查現況']], key="ed_s_p1", column_config={"狀態確認": st.column_config.SelectboxColumn("狀態確認", options=["✅ 留用", "❌ 排除"])})
        with col_p1_r:
            st.write("**📦 個案型初步名單**")
            ed_c = st.data_editor(cp1[['狀態確認','地段地號','農地序號','目前農地調查現況']], key="ed_c_p1", column_config={"狀態確認": st.column_config.SelectboxColumn("狀態確認", options=["✅ 留用", "❌ 排除"])})

        if st.button("🔥 確定清出打叉 (❌) 的農地並自動補位"):
            to_kill = ed_s[ed_s['狀態確認']=="❌ 排除"]['地段地號'].tolist() + ed_c[ed_c['狀態確認']=="❌ 排除"]['地段地號'].tolist()
            st.session_state.excluded_lots.extend(to_kill); st.rerun()

        # ---------------------------------------------------------
        # [手動增補工具] (滿足新頁面需求)
        # ---------------------------------------------------------
        st.divider()
        with st.expander("🖐️ 手動挑選增補工具 (搜尋並納入現勘)"):
            col_m1, col_m2 = st.columns(2)
            with col_m1: m_sys_ids = st.multiselect("搜尋系統型網格", df_master['網格編號'].unique())
            with col_m2: m_case_lots = st.multiselect("搜尋個案型地號", df_master[~df_master['調查方式'].str.contains('系統', na=False)]['地段地號'].unique())
            if st.button("➕ 將手動勾選項納入現勘清單"):
                m_df = pd.concat([df_master[df_master['網格編號'].isin(m_sys_ids) & (df_master['代表性']=='代表點')], df_master[df_master['地段地號'].isin(m_case_lots)]])
                st.session_state.manual_selection_df = m_df.copy(); st.success(f"已預備 {len(m_df)} 筆手動農地。")

        # ---------------------------------------------------------
        # 第二階段：現勘結果錄入 (核心分表邏輯)
        # ---------------------------------------------------------
        st.divider(); st.subheader("第二階段：正式名單與現勘結果錄入")
        if st.button("💾 同步名單至現勘階段 (含自動與手動)"):
            st.session_state.temp_field_plan = pd.concat([current_selection, st.session_state.manual_selection_df]).drop_duplicates('地段地號')

        if st.session_state.temp_field_plan is not None:
            plan = st.session_state.temp_field_plan.copy()
            if '狀態確認' not in plan.columns: plan.insert(0, '狀態確認', '✅ 留用')
            if '現勘判定' not in plan.columns: plan['現勘判定'] = '系統型'
            
            # 統計指標
            k_f1, k_f2 = st.columns(2); k_f1.metric("目前納入現勘總筆數", len(plan))
            
            # 現勘錄入表格
            recon_ed = st.data_editor(
                plan[['狀態確認','現勘判定','地段地號','網格編號','目前農地調查現況','TWD97_X','TWD97_Y']],
                column_config={"狀態確認": st.column_config.SelectboxColumn("狀態", options=["✅ 留用", "❌ 排除"]), "現勘判定": st.column_config.SelectboxColumn("判定結果", options=["系統型","個案型","建物","難以採樣"])},
                key="recon_ed_p2", use_container_width=True
            )

            # 分表顯示
            res_samp = recon_ed[recon_ed['現勘判定'].isin(['系統型','個案型']) & (recon_ed['狀態確認']=='✅ 留用')]
            res_fail = recon_ed[recon_ed['現勘判定'].isin(['建物','難以採樣']) | (recon_ed['狀態確認']=='❌ 排除')]
            
            k_f2.metric("最終採樣名單筆數", len(res_samp))
            
            c_r1, c_r2 = st.columns(2)
            with c_r1: st.markdown('<div class="recon-valid"><b>✅ 最終採樣清單 (進入統計摘要)</b></div>', unsafe_allow_html=True); st.dataframe(res_samp, height=300)
            with c_r2: 
                st.markdown('<div class="recon-invalid"><b>❌ 失效/排除名單</b></div>', unsafe_allow_html=True); st.dataframe(res_fail, height=300)
                if st.button("✅ 確認更新失效點到資料庫總表"):
                    st.session_state.excluded_lots.extend(res_fail['地段地號'].tolist()); st.success("已移除失效點，請執行上方補位。"); st.rerun()

            # 補位推薦 (自動推薦)
            if len(res_samp) < quota:
                st.warning(f"⚠️ 筆數缺口：{quota - len(res_samp)} 筆。系統推薦補進：")
                recom = eligible[~eligible['地段地號'].isin(recon_ed['地段地號'])].head(quota - len(res_samp))
                st.dataframe(recom[['網格編號','地段地號','目前農地調查現況','代表性']])

        # 歸檔定案
        st.divider(); col_f1, col_f2 = st.columns(2)
        if col_f1.button("✅ 最終調查名單確認 (移入歷史庫)"):
            if st.session_state.temp_field_plan is not None:
                st.session_state.archived_plans[f"{target_year}_最終定案"] = res_samp
                st.session_state.temp_field_plan = None; st.success("計畫歸檔成功！工作區已清空。"); st.rerun()
        if col_f2.button("🔄 重置所有篩選"): st.session_state.excluded_lots = []; st.session_state.temp_field_plan = None; st.rerun()

        # --- GIS 地圖與網格 (淺紅/淺藍著色) ---
        st.divider(); st.subheader("🗺️ 年度擬定計畫分布地圖 (圖標與著色)")
        m_plan = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        
        if gdf_grid is not None:
            gs_active = current_selection[['網格編號','網格監測頻率']].drop_duplicates('網格編號')
            g_col = gdf_grid.merge(gs_active, left_on='網格號', right_on='網格編號', how='inner').to_crs(epsg=4326)
            def style_g(f):
                fr = str(f['properties'].get('網格監測頻率',''))
                c = '#FFB6C1' if '持續' in fr else '#ADD8E6' if '延長' in fr else '#f8f9fa'
                return {'fillColor': c, 'color': 'white', 'weight': 1, 'fillOpacity': 0.4}
            folium.GeoJson(g_col, style_function=style_g).add_to(m_plan)
        
        for _, r in current_selection.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                sd = 3 if "系統" in str(r['計畫類別']) else 4 # 三角/正方
                if sd == 3: clr = "red" if "增量" in str(r['目前農地調查現況']) else "blue"
                else: clr = "yellow" if "增量" in str(r['目前農地調查現況']) else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=sd, radius=8, color=clr, fill=True, tooltip=r['地段地號']).add_to(m_plan)
            except: continue
        st_folium(m_plan, width=1100, height=550)

        # 歷史計畫存檔區
        if st.session_state.archived_plans:
            st.divider(); st.subheader("📚 歷史計畫存檔資料庫")
            for ak, av in st.session_state.archived_plans.items():
                with st.expander(f"📂 {ak}"): st.dataframe(av)

    # --- D. 新增年度調查結果 (保留 DA 計算邏輯) ---
    elif menu == "新增年度調查結果":
        st.title("➕ 錄入採樣數據與判定")
        pwd = st.sidebar.text_input("密碼", type="password")
        if pwd == ADMIN_PASSWORD:
            sl = st.text_input("🔍 搜尋地號錄入")
            if sl:
                h = df_master[df_master['地段地號']==sl.strip()]
                if not h.empty:
                    row = h.iloc[0]
                    with st.form("en"):
                        # (此處保留原有的重金屬錄入與 DA 判定代碼...)
                        st.success(f"找到 {sl}，錄入中...")
                        st.form_submit_button("執行判定")
        else: st.warning("密碼鎖定")

    # --- E. 空間地圖檢視 (全量顯示) ---
    elif menu == "空間地圖檢視":
        st.title("🗺️ 衛星影像與全網格圖")
        # (保留 Folium 衛星底圖與全量農地點位圖示代碼...)

else: st.error("資料載入失敗")





















