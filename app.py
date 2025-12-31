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
# 1. 系統初始化與狀態管理 (防止 AttributeError)
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
    if pd.isna(val): return ""
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

def clean_status(val):
    if pd.isna(val): return "無狀態"
    s = str(val).strip().replace("\n", "").replace("\r", "")
    return s if s not in ['nan', 'None', ''] else "無狀態"

def get_pretty_rep(row, block_df):
    """圖示邏輯更新：(採樣點)對應✅，(備用點)對應⚪，(非採樣)對應❌"""
    r = str(row.get('代表性', '')).strip()
    s = str(row.get('農地監測狀態', '')).strip()
    lot = str(row.get('地段地號', '')).strip()
    if r == "採樣點": return "✅ 採樣點"
    if r == "備用點": return "⚪ 備用點"
    if block_df is not None and not block_df.empty and lot in block_df['農地地段地號'].values:
        is_rep = block_df[block_df['農地地段地號']==lot].iloc[0]['代表農地']
        if "否" in str(is_rep): return "❌ 非採樣 (同坵塊備註)"
    return f"❌ 非採樣 ({s})"

# ==========================================
# 2. 系統設定與美化 CSS
# ==========================================
st.set_page_config(page_title="彰化農地智慧管理系統 45.0", layout="wide", page_icon="🌾")
ADMIN_PASSWORD = "ET23597010"
EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"
LOG_PATH = "edit_log.csv"

transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ["汞", "砷", "銅", "鉻", "鎘", "鉛", "鋅", "鎳"]

st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; }
    .stMetric { background-color: #ffffff; padding: 10px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .status-header { padding: 10px; border-radius: 10px; text-align: center; font-weight: bold; font-size: 16px; border: 1px solid #ddd; margin-bottom: 5px; }
    .bg-p { background-color: #FFB6C1; color: #721c24; }
    .bg-l { background-color: #ADD8E6; color: #004085; }
    .bg-e { background-color: #90EE90; color: #155724; }
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
        def get_s(n): return next((s for s in xl.sheet_names if n == s.strip()), None)
        df_m = pd.read_excel(xl, sheet_name=get_s("農地現況主檔"))
        df_m.columns = df_m.columns.str.strip()
        df_m['網格編號'] = df_m['網格編號'].apply(clean_id)
        df_m['網格監測頻率'] = df_m['網格監測頻率'].apply(clean_status)
        df_m['目前農地調查現況'] = df_m['目前農地調查現況'].apply(clean_status)
        df_m['農地監測狀態'] = df_m['農地監測狀態'].apply(clean_status)
        df_m['地段地號'] = df_m['地段地號'].astype(str).str.strip()
        
        df_h = pd.read_excel(xl, sheet_name=get_s("歷年調查紀錄"))
        df_h['調查年度'] = pd.to_numeric(df_h['調查年度'], errors='coerce').fillna(0).astype(int)
        
        df_b = pd.read_excel(xl, sheet_name=get_s("同坵塊對照表"))
        df_b.columns = df_b.columns.str.strip()
        df_b['農地地段地號'] = df_b['農地地段地號'].astype(str).str.strip()
        
        df_s = pd.read_excel(xl, sheet_name=get_s("判定標準表")).set_index('項目名稱')
        return df_m, df_h, df_b, df_s
    except: return None, None, None, None

df_master, df_history, df_block, df_settings = load_all_data()

# ==========================================
# 4. 全局數據預處理 (重要：防止 NameError)
# ==========================================
if df_master is not None:
    # 基礎指標
    abs_total = len(df_master)
    sampling_pts_count = len(df_master[df_master['代表性'].astype(str).str.strip() == '採樣點'])
    control_pts_count = len(df_master[df_master['農地監測狀態'] == '管制'])
    build_pts_count = len(df_master[df_master['農地監測狀態'] == '建物'])
    hard_pts_count = len(df_master[df_master['農地監測狀態'] == '難以採樣'])
    normal_pts_count = len(df_master[df_master['農地監測狀態'] == '正常'])
    
    # 【關鍵修復】：將變數移出 if menu 判斷，確保全局可用
    grid_sys_only = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
    all_grid_recs = df_master[df_master['網格編號'] != ""].drop_duplicates('網格編號').copy()
    g_p = len(all_grid_recs[all_grid_recs['網格監測頻率'] == '持續'])
    g_l = len(all_grid_recs[all_grid_recs['網格監測頻率'] == '延長'])
    g_e = len(all_grid_recs[all_grid_recs['網格監測頻率'] == '退場'])
    g_sum = g_p + g_l + g_e
    g_none = len(all_grid_recs) - g_sum

    case_master = df_master[~df_master['調查方式'].str.contains('系統', na=False)].copy()
    c_mapping = {"增量":"持續", "延長":"延長", "正常":"退場", "管制":"管制", "建物":"建物", "難以採樣":"難以採樣"}
    case_master['對應狀態'] = case_master['目前農地調查現況'].map(c_mapping).fillna(case_master['目前農地調查現況'])

# ==========================================
# 5. 主選單
# ==========================================
st.sidebar.title("🌿 系統導覽")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

if df_master is not None:

    # --- A. 統計首頁 ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        k_cols = st.columns(6)
        k_cols[0].metric("總資料點數", abs_total); k_cols[1].metric("總採樣點數", sampling_pts_count)
        k_cols[2].metric("管制點數", control_pts_count); k_cols[3].metric("建物數量", build_pts_count)
        k_cols[4].metric("難以採樣數量", hard_pts_count); k_cols[5].metric("正常退場數量", normal_pts_count)
        
        st.divider()
        st.subheader("🌐 系統型網格現況統計")
        gc = st.columns(5)
        gc[0].metric("持續網格", g_p); gc[1].metric("延長網格", g_l); gc[2].metric("退場網格", g_e); gc[3].metric("有效網格合計", g_sum); gc[4].metric("無狀態網格", g_none)
        
        st.divider()
        st.subheader("📦 個案型農地現況統計")
        cc_m = case_master['對應狀態'].value_counts()
        cc = st.columns(6)
        for i, lab in enumerate(["持續", "延長", "退場", "管制", "難以採樣", "建物"]):
            cc[i].metric(lab, cc_m.get(lab, 0))

        st.divider()
        st.subheader("🔍 網格快速查詢系統")
        qs = st.text_input("輸入網格 ID (如: G2405)", key="home_gs")
        if qs:
            res = df_master[df_master['網格編號'] == clean_id(qs)]
            if not res.empty: st.dataframe(res.style.apply(lambda x: ['background-color: #FFFFCC' if x.代表性=='採樣點' else '' for _ in x], axis=1), use_container_width=True)

        st.divider()
        st.subheader("📊 近三年調查樹狀圖")
        tree_cols = st.columns(3); cy_m = get_minguo_year()
        for i, y in enumerate([cy_m-2, cy_m-1, cy_m]):
            cn = f"{y}狀態"
            if cn in df_master.columns:
                y_df = df_master[df_master[cn].notna()].copy()
                if not y_df.empty:
                    y_counts = y_df.groupby(['調查方式', cn]).size().reset_index(name='筆數')
                    fig = px.treemap(y_counts, path=[px.Constant(f"{y}年"), '調查方式', cn], values='筆數', color=cn, color_discrete_map={'監測':'#ADD8E6','正常':'#90EE90','管制':'#FFB6C1'})
                    tree_cols[i].plotly_chart(fig, use_container_width=True)

    # --- B. 資料庫查詢與下載 ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 數據查詢中心")
        admin_mode = False
        with st.sidebar.expander("🔐 管理員修正權限"):
            if st.text_input("輸入修正密碼", type="password") == ADMIN_PASSWORD: admin_mode = True; st.success("編輯模式開啟")
        
        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型清單", "📦 個案型清單", "📜 修改紀錄"])
        
        with tabs[0]: # 1. 總表
            sm_db = st.text_input("🔍 快速搜尋總表", key="m_search_db")
            df_pretty = df_master.copy()
            df_pretty['代表性顯示'] = df_pretty.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            cols = list(df_pretty.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號')+1; cols.insert(idx, cols.pop(cols.index('代表性顯示'))); df_pretty = df_pretty[cols]
            if sm_db: df_pretty = df_pretty[df_pretty.astype(str).apply(lambda x: x.str.contains(sm_db)).any(axis=1)]
            st.dataframe(df_pretty, height=800, use_container_width=True)
            
        with tabs[1]: # 2. 歷年
            st.subheader("📅 年度調查紀錄明細")
            y_opts = sorted([y for y in df_history['調查年度'].unique() if y > 0], reverse=True)
            sel_y = st.selectbox("選擇年度", y_opts if y_opts else [113], key="y_sel_db_tab")
            y_res = df_history[df_history['調查年度'] == sel_y].copy()
            if not y_res.empty:
                st.dataframe(y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left'), use_container_width=True)

        with tabs[2]: # 3. 坵塊
            st.subheader("🏠 坵塊群組搜尋")
            bq = st.text_input("🔍 搜尋地號找群組成員", key="bq_tab")
            if bq and bq in df_block['農地地段地號'].values:
                gid = df_block[df_block['農地地段地號']==bq].iloc[0]['農地群組編號']; st.dataframe(df_block[df_block['農地群組編號']==gid])
            st.write("**現有對照表清單：**"); st.dataframe(df_block, height=400, use_container_width=True)

        with tabs[3]: # 4. 系統型 (修復 NameError)
            st.subheader("🌐 系統型分類與搜尋看板")
            sq_sys = st.text_input("🔍 網格快速查詢", placeholder="輸入網格號碼...", key="sq_sys_tab")
            if sq_sys:
                res_s = df_master[df_master['網格編號'] == clean_id(sq_sys)]
                if not res_s.empty:
                    st.info(f"📍 網格 {sq_sys} | 頻率: {res_s['網格監測頻率'].iloc[0]}")
                    st.dataframe(res_s[['代表性','地段地號','目前農地調查現況','農地監測狀態']], use_container_width=True)
            
            st.divider()
            sc_cols = st.columns(3)
            glabs = ['持續','延長','退場']; gcounts = [g_p, g_l, g_e]
            for i, f in enumerate(glabs):
                with sc_cols[i]:
                    st.markdown(f'<div class="status-header bg-{"p" if i==0 else "l" if i==1 else "e"}">{f}網格 (共 {gcounts[i]} 個)</div>', unsafe_allow_html=True)
                    ids = all_grid_recs[all_grid_recs['網格監測頻率']==f]['網格編號'].tolist()
                    sel = st.selectbox(f"{f}網格選取", ["未選"]+ids, key=f"sys_box_tab_{f}")
                    if sel != "未選":
                        gv = df_master[df_master['網格編號']==sel].copy()
                        st.info(f"📊 統計：共 {len(gv)} 筆 | 採樣點 {len(gv[gv['代表性']=='採樣點'])}")
                        gv['圖示'] = gv.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
                        st.dataframe(gv[['圖示','農地序號','地段地號','目前農地調查現況']])
            
            st.divider(); st.write("**系統型農地總表：**")
            # 使用最上方預處理好的變數
            st.dataframe(grid_sys_only, height=400, use_container_width=True)

        with tabs[4]: # 5. 個案型
            st.subheader("📦 個案型監測看板與清單")
            sq_case = st.text_input("🔍 搜尋個案地號/序號", key="sq_case_db_tab")
            c_clabs = ["持續", "延長", "退場", "管制", "難以採樣", "建物"]; c_metrics = st.columns(6)
            for i, lab in enumerate(c_clabs):
                sub_c = case_master[case_master['對應狀態']==lab]
                c_metrics[i].metric(lab, len(sub_c))
                with c_metrics[i].expander(f"檢視名單"):
                    st.dataframe(sub_c[['地段地號','農地序號','農地監測狀態']], height=250)
            
            st.divider(); st.write("**個案型農地全總表：**")
            # 使用最上方預處理好的變數
            st.dataframe(case_master, height=400, use_container_width=True)

        with tabs[5]: # 6. 修改紀錄
            if os.path.exists(LOG_PATH): st.dataframe(pd.read_csv(LOG_PATH))

    # --- C. 新年度調查點篩選名單 (旗艦版工作流) ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 調查計畫決策與自動補位系統")
        
        col_m1, col_m2 = st.columns([2, 1])
        plan_mode = col_m1.radio("模式選擇", ["自動系統產生", "手動挑選名單"], horizontal=True)
        auto_backfill = col_m2.checkbox("開啟自動補位功能", value=True)

        c1, c2 = st.columns(2)
        target_year = c1.number_input("設定年度 (民國)", value=get_minguo_year()+1)
        quota = c2.number_input("調查配額 (自動補位基準)", value=500, step=10)

        df_calc = df_master.copy()
        if plan_mode.startswith("自動"):
            sys_pool = []
            grids = df_calc[df_calc['網格編號'] != ""]['網格編號'].unique()
            for gid in grids:
                g_data = df_calc[df_calc['網格編號'] == gid]
                f_type, last_y = str(g_data['網格監測頻率'].iloc[0]), g_data['最後調查年分'].max()
                if last_y == target_year - 1: continue 
                prio = 1 if f_type == '持續' and (target_year-last_y >= 2) else (2 if f_type == '延長' and (target_year-last_y >= 10) else 99)
                if prio < 99:
                    reps = g_data[(g_data['代表性'] == '代表點') & (~g_data['農地監測狀態'].isin(['管制','建物','難以採樣']))]
                    if len(reps) < 3:
                        back = g_data[g_data['代表性'] == '備用點'].sort_values('農地序號').head(3 - len(reps)).copy()
                        fg = pd.concat([reps, back])
                    else: fg = reps.copy()
                    fg['優先權重'], fg['計畫類別'], fg['網格頻率'] = prio, '系統型網格', f_type
                    sys_pool.append(fg)
            
            c_active = df_calc[~df_calc['調查方式'].str.contains('系統', na=False)].copy()
            c_active = c_active[~c_active['農地監測狀態'].isin(['管制','建物','正常'])]
            def c_algo(r):
                ly = r['最後調查年分']
                if ly == target_year - 1: return 99
                if str(r['目前農地調查現況']) == '增量' and (target_year-ly >= 2): return 1
                if str(r['目前農地調查現況']) == '延長' and (target_year-ly >= 10): return 3
                return 99
            c_active['優先權重'] = c_active.apply(c_algo, axis=1)
            case_pool = c_active[c_active['優先權重'] < 99].copy()
            case_pool['計畫類別'] = '個案型農地'

            full_pool = pd.concat([pd.concat(sys_pool) if sys_pool else pd.DataFrame(), case_pool]).sort_values(['優先權重', '網格編號'])
            eligible = full_pool[~full_pool['地段地號'].isin(st.session_state.excluded_lots)]
            current_selection = eligible.head(int(quota)).copy() if auto_backfill else eligible.copy()
        
        else: # 手動增補模式
            with st.expander("🖐️ 手動挑選增補工具"):
                col_sel1, col_sel2 = st.columns(2)
                m_s_ids = col_sel1.multiselect("搜尋系統網格", df_master['網格編號'].unique(), key="m_s_g")
                m_c_lots = col_sel2.multiselect("搜尋個案地號", df_master[~df_master['調查方式'].str.contains('系統', na=False)]['地段地號'].unique(), key="m_c_l")
                if st.button("➕ 將手動勾選項納入現勘名單"):
                    m_df = pd.concat([df_master[df_master['網格編號'].isin(m_s_ids) & (df_master['代表性']=='代表點')], df_master[df_master['地段地號'].isin(m_c_lots)]])
                    st.session_state.manual_selection_df = m_df.copy(); st.success("已加入手動名單。")
            current_selection = pd.concat([current_selection if 'current_selection' in locals() else pd.DataFrame(), st.session_state.manual_selection_df]).drop_duplicates('地段地號')

        if not current_selection.empty:
            # 階段一：去留微調
            st.subheader("第一階段：初步篩選與去留確認 (✅/❌)")
            current_selection.insert(0, '狀態確認', '✅ 留用')
            s1, s2 = st.columns(2)
            with s1: ed_s = st.data_editor(current_selection[current_selection['計畫類別']=='系統型網格'][['狀態確認','網格編號','地段地號','農地序號','目前農地調查現況']], key="ed_s_p1", column_config={"狀態確認": st.column_config.SelectboxColumn("狀態", options=["✅ 留用", "❌ 排除"])})
            with s2: ed_c = st.data_editor(current_selection[current_selection['計畫類別']=='個案型農地'][['狀態確認','地段地號','農地序號','目前農地調查現況']], key="ed_c_p1", column_config={"狀態確認": st.column_config.SelectboxColumn("狀態", options=["✅ 留用", "❌ 排除"])})
            if st.button("🔥 確定清出打叉 (❌ 排除) 的農地並自動補足"):
                to_kill = ed_s[ed_s['狀態確認']=="❌ 排除"]['地段地號'].tolist() + ed_c[ed_c['狀態確認']=="❌ 排除"]['地段地號'].tolist()
                st.session_state.excluded_lots.extend(to_kill); st.rerun()

            # 階段二：現勘錄入與分表
            st.divider(); st.subheader("第二階段：正式名單與現勘結果錄入")
            if st.button("💾 確認名單進入現勘錄入"): st.session_state.temp_field_plan = current_selection.copy()

            if st.session_state.temp_field_plan is not None:
                plan = st.session_state.temp_field_plan.copy()
                if '微調勾選' not in plan.columns: plan.insert(0, '微調勾選', '✅ 留用')
                if '現勘判定' not in plan.columns: plan['現勘判定'] = '系統型'
                
                recon_ed = st.data_editor(
                    plan[['微調勾選','現勘判定','地段地號','網格編號','目前農地調查現況','TWD97_X','TWD97_Y']],
                    column_config={"微調勾選": st.column_config.SelectboxColumn("狀態", options=["✅ 留用", "❌ 排除"]), "現勘判定": st.column_config.SelectboxColumn("判定結果", options=["系統型", "個案型", "建物", "難以採樣"])},
                    key="recon_ed_v6"
                )
                
                res_samp = recon_ed[recon_ed['現勘判定'].isin(['系統型','個案型']) & (recon_ed['微調勾選']=='✅ 留用')]
                res_fail = recon_ed[recon_ed['現勘判定'].isin(['建物','難以採樣']) | (recon_ed['微調勾選']=='❌ 排除')]
                
                st.markdown(f"📊 **待現勘總數: {len(recon_ed)} | 最終採樣數: {len(res_samp)}**")
                cr1, cr2 = st.columns(2)
                with cr1: st.success("✅ 採樣清單"); st.dataframe(res_samp, height=300)
                with cr2: 
                    st.error("❌ 失效名單"); st.dataframe(res_fail, height=300)
                    if st.button("✅ 確認更新失效點到總表"): st.session_state.excluded_lots.extend(res_fail['地段地號'].tolist()); st.rerun()

            # 最終歸檔
            st.divider(); col_f1, col_f2 = st.columns(2)
            if col_f1.button("✅ 最終調查名單確認 (存入歷史庫)"):
                st.session_state.archived_plans[f"{target_year}_定案"] = res_samp
                st.session_state.temp_field_plan = None; st.success("歸檔成功！"); st.rerun()
            if col_f2.button("🔄 重置篩選"): st.session_state.excluded_lots = []; st.session_state.temp_field_plan = None; st.rerun()

        # --- GIS 地圖 ---
        st.divider(); st.subheader("🗺️ 年度擬定計畫分布圖")
        m_p = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        if gdf_grid is not None:
            gs_active = current_selection[['網格編號','網格監測頻率']].drop_duplicates('網格編號')
            g_col = gdf_grid.merge(gs_active, left_on='網格號', right_on='網格編號', how='inner').to_crs(epsg=4326)
            def get_g_c(f): fr = str(f); return '#FFB6C1' if '持續' in fr else '#ADD8E6' if '延長' in fr else '#f8f9fa'
            folium.GeoJson(g_col, style_function=lambda x: {'fillColor': get_g_c(x['properties'].get('網格監測頻率','')), 'color': 'white', 'weight': 1, 'fillOpacity': 0.4}).add_to(m_p)
        for _, r in current_selection.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                sd = 3 if "系統" in str(r['計畫類別']) else 4
                if sd == 3: clr = "red" if "增量" in str(r['目前農地調查現況']) else "blue"
                else: clr = "yellow" if "增量" in str(r['目前農地調查現況']) else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=sd, radius=8, color=clr, fill=True, tooltip=r['地段地號']).add_to(m_p)
            except: continue
        st_folium(m_p, width=1100, height=550)
        
        if st.session_state.archived_plans:
            st.divider(); st.subheader("📚 歷史計畫檔案庫"); 
            for ak, av in st.session_state.archived_plans.items(): 
                with st.expander(f"📂 {ak}"): st.dataframe(av)

    # --- D. 新增結果 ---
    elif menu == "新增年度調查結果":
        st.title("➕ 錄入年度數據與 DA 判定")
        pwd = st.sidebar.text_input("管理員密碼", type="password")
        if pwd == ADMIN_PASSWORD:
            sl = st.text_input("🔍 輸入欲錄入地號")
            if sl:
                h = df_master[df_master['地段地號']==sl.strip()]
                if not h.empty:
                    r = h.iloc[0]
                    with st.form("en"):
                        st.subheader(f"📍 編輯：{sl}")
                        c1, c2 = st.columns(2)
                        nx, ny = c1.number_input("實測 X", value=float(r['TWD97_X'])), c2.number_input("實測 Y", value=float(r['TWD97_Y']))
                        if np.sqrt((nx-r['TWD97_X'])**2 + (ny-r['TWD97_Y'])**2) > 3:
                            st.warning("⚠️ 座標位移 > 3M"); st.info(f"當前網格: {find_grid_by_coords(nx, ny, gdf_grid)}")
                        v = {}
                        for m in METALS:
                            v[m] = st.columns(2)[1].number_input(f"{m}(全量)", min_value=0.0) or st.columns(2)[0].number_input(f"{m}(XRF)", min_value=0.0)
                        if st.form_submit_button("執行判定"):
                            final_st = "正常"; das = {}
                            for m in METALS:
                                init = r.get(f'初始_{m}', 0); da = ((v[m]-init)/init*100) if init > 0 else 0
                                if v[m] > df_settings.loc[m,'管制標準']: final_st = "管制"
                                elif init > df_settings.loc[m,'管制標準'] and da > df_settings.loc[m,'上升標準 (DA門檻)']:
                                    if final_st != "管制": final_st = "增量"
                            st.success(f"建議判定: {final_st}")
        else: st.warning("權限鎖定")

    # --- E. 空間地圖 ---
    elif menu == "空間地圖檢視":
        st.title("🗺️ 衛星影像與網格分布圖")
        m_map = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        if gdf_grid is not None:
            gs_m = df_master.drop_duplicates('網格編號')[['網格編號', '網格監測頻率']]
            merged_m = gdf_grid.to_crs(epsg=4326).merge(gs_m, left_on='網格號', right_on='網格編號', how='left')
            def get_c_m(f): fr = str(f); return '#FFB6C1' if '持續' in fr else '#ADD8E6' if '延長' in fr else '#90EE90' if '退場' in fr else '#F8F8F8'
            folium.GeoJson(merged_m, style_function=lambda x: {'fillColor': get_c_m(x['properties'].get('網格監測頻率')), 'color': 'white', 'weight': 1, 'fillOpacity': 0.4}).add_to(m_map)
        
        sample_pts = df_master.dropna(subset=['TWD97_X', 'TWD97_Y'])
        for _, r in sample_pts.sample(min(1500, len(sample_pts))).iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                sd = 4 if "個案" in str(r['調查方式']) else 3
                mon_st = str(r['農地監測狀態'])
                if mon_st == "管制": sd, clr = 6, "red"
                elif mon_st == "建物": sd, clr = 6, "black"
                elif mon_st == "難以採樣": sd, clr = 6, "purple"
                elif str(r['代表性']) == "備用點": sd, clr = 4, "white"
                else:
                    curr_st = str(r['目前農地調查現況'])
                    clr = "red" if "持續" in curr_st or "增量" in curr_st else "blue" if "延長" in curr_st else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=sd, radius=6, color=clr, fill=True, popup=f"{r['地段地號']}").add_to(m_map)
            except: continue
        st_folium(m_map, width=1200, height=750)

else: st.error("資料載入失敗")

































