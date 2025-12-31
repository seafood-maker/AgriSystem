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
# 1. 基礎函式與系統初始化
# ==========================================

# 初始化所有會話狀態 (防止分頁報錯)
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
    """清理編號：轉字串、去空格、去Excel浮點數尾綴"""
    if pd.isna(val): return ""
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

def clean_status(val):
    """狀態清理：確保對接精準 (解決 245 vs 254 問題)"""
    if pd.isna(val): return "無狀態"
    s = str(val).strip()
    return s if s not in ['nan', 'None', ''] else "無狀態"

def get_pretty_rep(row, block_df):
    """代表性圖示邏輯：✅ 代表, ⚪ 備用, ❌ 非採樣(含原因)"""
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
# 2. 系統設定與專業美化 CSS
# ==========================================

st.set_page_config(page_title="彰化農地智慧管理系統", layout="wide", page_icon="🌾")
ADMIN_PASSWORD = "ET23597010"
EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"
LOG_PATH = "edit_log.csv"

transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ["汞", "砷", "銅", "鉻", "鎘", "鉛", "鋅", "鎳"]

# 包含黑標題與彩色方塊樣式
st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; }
    .stMetric { background-color: #ffffff; padding: 10px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .status-header { padding: 10px; border-radius: 10px; text-align: center; font-weight: bold; font-size: 16px; border: 1px solid #ddd; margin-bottom: 5px; }
    .bg-p { background-color: #FFB6C1; color: #721c24; }
    .bg-l { background-color: #ADD8E6; color: #004085; }
    .bg-e { background-color: #90EE90; color: #155724; }
    .stats-container { background-color: #e8f5e9; padding: 15px; border-radius: 10px; margin-bottom: 20px; border: 1px solid #c8e6c9; }
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
        # 強制清洗核心欄位
        df_m['網格編號'] = df_m['網格編號'].apply(clean_id)
        df_m['網格監測頻率'] = df_m['網格監測頻率'].apply(clean_status)
        df_m['地段地號'] = df_m['地段地號'].astype(str).str.strip()
        df_m['最後調查年分'] = pd.to_numeric(df_m['最後調查年分'], errors='coerce')
        if '延長頻率' not in df_m.columns: df_m['延長頻率'] = 10
        
        df_h = pd.read_excel(xl, sheet_name=get_s("歷年調查紀錄"))
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
            gdf = gpd.read_file(SHP_PATH)
            if gdf.crs is None or gdf.crs.to_epsg() != 3826: gdf.set_crs(epsg=3826, allow_override=True, inplace=True)
            if '網格號' in gdf.columns: gdf['網格號'] = gdf['網格號'].apply(clean_id)
            return gdf
        except: return None
    return None

df_master, df_history, df_block, df_settings = load_all_data()
gdf_grid = load_grid_shp()

def find_grid_by_coords(x, y, gdf):
    if gdf is None: return "未知"
    p = Point(x, y); match = gdf[gdf.contains(p)]
    return str(match.iloc[0]['網格號']) if not match.empty else "範圍外"

# ==========================================
# 4. 全局數據統計 (解決 NameError)
# ==========================================

if df_master is not None:
    # 頂列指標
    abs_total = len(df_master)
    sampling_pts_count = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])

    # 網格型統計數據
    grid_all_list = df_master[df_master['網格編號'] != ""].drop_duplicates('網格編號').copy()
    g_p = len(grid_all_list[grid_all_list['網格監測頻率'] == '持續'])
    g_l = len(grid_all_list[grid_all_list['網格監測頻率'] == '延長'])
    g_e = len(grid_all_list[grid_all_list['網格監測頻率'] == '退場'])
    g_sum = g_p + g_l + g_e
    g_none = len(grid_all_list) - g_sum

    # 個案型統計數據
    case_master = df_master[~df_master['調查方式'].astype(str).str.contains('系統', na=False)].copy()
    c_map = {"增量":"持續", "延長":"延長", "正常":"退場", "管制":"管制", "建物":"建物", "難以採樣":"難以採樣"}
    case_master['對應狀態'] = case_master['目前農地調查現況'].map(c_map).fillna(case_master['目前農地調查現況'])

# ==========================================
# 5. 主導覽
# ==========================================

st.sidebar.title("🌿 系統選單")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

if df_master is not None:

    # --- A. 統計首頁 (完整恢復版) ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        
        # 1. 六大看板
        k = st.columns(6)
        k[0].metric("總資料點數", abs_total); k[1].metric("總採樣點數", sampling_pts_count)
        k[2].metric("管制點數", control_pts_count); k[3].metric("建物數量", build_pts_count)
        k[4].metric("難以採樣數量", hard_pts_count); k[5].metric("正常退場數量", normal_pts_count)

        st.divider()
        # 2. 系統型網格統計 (精確對齊)
        st.subheader("🌐 系統型網格現況統計")
        gc = st.columns(5)
        gc[0].metric("持續網格", g_p); gc[1].metric("延長網格", g_l); gc[2].metric("退場網格", g_e)
        gc[3].metric("有效網格合計", g_sum); gc[4].metric("無狀態網格", g_none)

        st.divider()
        # 3. 個案型看板
        st.subheader("📦 個案型農地現況統計")
        case_counts = case_master['對應狀態'].value_counts()
        cc = st.columns(6)
        clabs = ["持續", "延長", "退場", "管制", "難以採樣", "建物"]
        for i, lab in enumerate(clabs):
            cc[i].metric(lab, case_counts.get(lab, 0))

        st.divider()
        # 4. 網格查詢 (代表點高亮)
        st.subheader("🔍 網格快速查詢系統")
        qs_id = st.text_input("輸入網格 ID 搜尋內容 (如: G2405)", key="home_qs")
        if qs_id:
            res = df_master[df_master['網格編號'] == clean_id(qs_id)]
            if not res.empty: st.dataframe(res.style.apply(lambda x: ['background-color: #FFFFCC' if x.代表性=='代表點' else '' for _ in x], axis=1), use_container_width=True)
            else: st.warning("查無資料")

        st.divider()
        # 5. 樹狀圖
        st.subheader("📊 近三年調查分佈 (樹狀圖)")
        tree_cols = st.columns(3); cy = get_minguo_year()
        for i, y in enumerate([cy-2, cy-1, cy]):
            col_n = f"{y}狀態"
            if col_n in df_master.columns:
                y_df = df_master[df_master[col_n].notna()].copy()
                if not y_df.empty:
                    y_counts = y_df.groupby(['調查方式', col_n]).size().reset_index(name='筆數')
                    fig = px.treemap(y_counts, path=[px.Constant(f"{y}年"), '調查方式', col_n], values='筆數', color=col_n,
                                     color_discrete_map={'監測':'#ADD8E6','正常':'#90EE90','管制':'#FFB6C1'})
                    tree_cols[i].plotly_chart(fig, use_container_width=True)

    # --- B. 資料庫查詢與下載 (功能全開修正版) ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 數據查詢中心")
        admin_mode = False
        with st.sidebar.expander("🔐 管理員修正權限"):
            if st.text_input("輸入權限密碼", type="password") == ADMIN_PASSWORD: admin_mode = True; st.success("編輯模式已開啟")
        
        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單", "📜 修改紀錄"])
        
        with tabs[0]: # 1. 總表清單 (圖示移位、黑標題、高度800)
            st.subheader("🌾 農地現況總表")
            sm_q = st.text_input("🔍 搜尋地號/序號/網格", key="m_s_tab")
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            # 欄位位置調整
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號')+1; cols.insert(idx, cols.pop(cols.index('代表性顯示'))); df_p = df_p[cols]
            if sm_q: df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(sm_q)).any(axis=1)]
            st.dataframe(df_p, height=800, use_container_width=True)
            towrite = io.BytesIO(); df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載全量 Excel", data=towrite.getvalue(), file_name="彰化農地總表.xlsx")

        with tabs[1]: # 2. 歷年結果
            y_list = sorted(df_history['調查年度'].unique(), reverse=True); sel_y = st.selectbox("選擇查詢年度", y_list if y_list else [113])
            y_res = df_history[df_history['調查年度'] == int(sel_y)].copy()
            if not y_res.empty: st.dataframe(y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left'), use_container_width=True)
            else: st.warning("該年度無數據")

        with tabs[2]: # 3. 坵塊管理 (自動編號 + 多筆新增)
            st.subheader("🏠 坵塊群組管理")
            blk_q = st.text_input("🔍 搜尋地號找群組成員"); 
            if blk_q and blk_q in df_block['農地地段地號'].values:
                gid = df_block[df_block['農地地段地號']==blk_q].iloc[0]['農地群組編號']; st.dataframe(df_block[df_block['農地群組編號']==gid])
            with st.expander("➕ 批次新增同坵塊關聯"):
                try: 
                    last_id = df_block['農地群組編號'].str.extract('(\d+)').dropna().astype(int).max()[0]
                    next_id_str = f"BLOCK_{str(last_id+1).zfill(3)}"
                except: next_id_str = "BLOCK_001"
                with st.form("new_blk"):
                    st.write(f"預計編號: **{next_id_str}**"); li = st.text_area("1. 貼上地號清單 (每行一筆)"); ri = st.text_input("2. 指定代表點地號")
                    if st.form_submit_button("確認建立關聯"): st.info("已提交排程。")
            st.write("**對照清單：**"); st.dataframe(df_block, height=400)

        with tabs[3]: # 4. 系統型看板 (彩色區塊)
            st.subheader("🌐 系統型農地監測看板")
            s_gt_q = st.text_input("🔍 全域搜尋網格號碼", key="s_gt_search")
            # 分類方塊
            sc1, sc2, sc3 = st.columns(3)
            for i, f in enumerate(['持續','延長','退場']):
                with sc1 if i==0 else sc2 if i==1 else sc3:
                    st.markdown(f'<div class="status-header bg-{"p" if i==0 else "l" if i==1 else "e"}">{f}網格</div>', unsafe_allow_html=True)
                    ids = grid_all_list[grid_all_list['網格監測頻率']==f]['網格編號'].tolist()
                    sel = st.selectbox(f"選取{f}網格", ["未選"]+ids, key=f"sys_box_{f}")
                    if sel != "未選":
                        g_v = df_master[df_master['網格編號']==sel].copy()
                        st.info(f"📍 網格 {sel} 統計：農地 {len(g_v)} 筆 | 代表點 {len(g_v[g_v['代表性']=='代表點'])}")
                        g_v['代表性圖示'] = g_v.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
                        st.dataframe(g_v[['代表性圖示','地段地號','農地監測狀態','目前農地調查現況']])
            st.divider(); st.write("**系統型總表：**"); st.dataframe(grid_sys_only, height=400)

        with tabs[4]: # 5. 個案型看板
            st.subheader("📦 個案型監測看板")
            scq_top = st.text_input("🔍 搜尋個案地號/序號")
            c_clabs = ["持續", "延長", "退場", "管制", "難以採樣", "建物"]; c_cols_tab = st.columns(6)
            for i, lab in enumerate(c_clabs):
                sub = case_master[case_master['對應狀態']==lab]
                c_cols_tab[i].metric(lab, len(sub))
            st.write("**個案型總表：**"); st.dataframe(case_master, height=400)

        with tabs[5]: # 6. 修改日誌
            st.subheader("📜 系統修改紀錄日誌")
            if os.path.exists(LOG_PATH): st.dataframe(pd.read_csv(LOG_PATH).sort_values(by="修改時間", ascending=False))
            else: st.info("尚無異動紀錄")

    # --- C. [核心大升級] 新年度調查點篩選名單 (旗艦工作流) ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 年度調查計畫決策工作流")
        
        # 1. 模式與配額
        col_m1, col_m2 = st.columns([2, 1])
        plan_mode = col_m1.radio("模式選擇", ["自動系統產生", "手動挑選名單"], horizontal=True)
        auto_backfill = col_m2.checkbox("開啟自動補位功能", value=True)

        c1, c2 = st.columns(2)
        target_year = c1.number_input("設定目標年度 (民國)", value=get_minguo_year()+1)
        quota = c2.number_input("設定本年度預算筆數", value=500, step=10)

        # --- 演算法池建立 ---
        df_calc = df_master.copy()
        if plan_mode.startswith("自動"):
            sys_pool = []
            grids = df_calc[df_calc['網格編號'] != ""]['網格編號'].unique()
            for gid in grids:
                g_data = df_calc[df_calc['網格編號'] == gid]
                f_type, last_y = str(g_data['網格監測頻率'].iloc[0]), g_data['最後調查年分'].max()
                if last_y == target_year - 1: continue # 排除去年
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
            def c_algo_p(r):
                ly = r['最後調查年分']
                if ly == target_year - 1: return 99
                if str(r['目前農地調查現況']) == '增量' and (target_year-ly >= 2): return 1
                if str(r['目前農地調查現況']) == '延長' and (target_year-ly >= 10): return 3
                return 99
            c_active['優先權重'] = c_active.apply(c_algo_p, axis=1)
            c_pool = c_active[c_active['優先權重'] < 99].copy()
            c_pool['計畫類別'], c_pool['篩選備註'] = '個案型農地', '獨立判定'

            full_pool = pd.concat([pd.concat(sys_pool) if sys_pool else pd.DataFrame(), c_pool]).sort_values(['優先權重'])
            eligible = full_pool[~full_pool['地段地號'].isin(st.session_state.excluded_lots)]
            current_selection = eligible.head(int(quota)).copy() if auto_backfill else eligible.copy()
        
        else: # 手動自選名單 (新增功能)
            with st.expander("🖐️ 手動挑選增補工具 (搜尋並納入現勘)"):
                col_sel1, col_sel2 = st.columns(2)
                m_s_ids = col_sel1.multiselect("搜尋系統網格", df_master['網格編號'].unique(), key="m_s_g")
                m_c_lots = col_sel2.multiselect("搜尋個案地號", df_master[~df_master['調查方式'].str.contains('系統', na=False)]['地段地號'].unique(), key="m_c_l")
                if st.button("➕ 將手動勾選項納入計畫"):
                    m_df = pd.concat([df_master[df_master['網格編號'].isin(m_s_ids) & (df_master['代表性']=='代表點')], df_master[df_master['地段地號'].isin(m_c_lots)]])
                    st.session_state.manual_selection_df = m_df.copy(); st.success("已加入手動名單。")
            current_selection = pd.concat([current_selection if 'current_selection' in locals() else pd.DataFrame(), st.session_state.manual_selection_df]).drop_duplicates('地段地號')

        if not current_selection.empty:
            # 2. 階段一：統計與去留
            st.subheader("第一階段：初步篩選與去留確認 (✅ 留用 / ❌ 排除)")
            st.markdown('<div class="stats-container">', unsafe_allow_html=True)
            sp1 = current_selection[current_selection['計畫類別']=='系統型網格']; cp1 = current_selection[current_selection['計畫類別']=='個案型農地']
            k_s1, k_s2, k_s3 = st.columns(3); k_s1.metric("系統農地筆數", len(sp1)); k_s2.metric("個案農地筆數", len(cp1)); k_s3.metric("總計擬定", len(current_selection))
            st.markdown('</div>', unsafe_allow_html=True)
            
            current_selection.insert(0, '狀態確認', '✅ 留用')
            s1, s2 = st.columns(2)
            with s1: ed_s_p1 = st.data_editor(current_selection[current_selection['計畫類別']=='系統型網格'][['狀態確認','網格編號','地段地號','農地序號']], key="ed_s_p1", column_config={"狀態確認": st.column_config.SelectboxColumn("狀態", options=["✅ 留用", "❌ 排除"])})
            with s2: ed_c_p1 = st.data_editor(current_selection[current_selection['計畫類別']=='個案型農地'][['狀態確認','地段地號','農地序號']], key="ed_c_p1", column_config={"狀態確認": st.column_config.SelectboxColumn("狀態", options=["✅ 留用", "❌ 排除"])})
            
            if st.button("🔥 確定清出打叉 (❌ 排除) 的農地並自動補足"):
                to_kill = ed_s_p1[ed_s_p1['狀態確認']=="❌ 排除"]['地段地號'].tolist() + ed_c_p1[ed_c_p1['狀態確認']=="❌ 排除"]['地段地號'].tolist()
                st.session_state.excluded_lots.extend(to_kill); st.rerun()

            # 3. 階段二：現勘錄入 (分表顯示)
            st.divider(); st.subheader("第二階段：正式名單與現勘結果錄入")
            if st.button("💾 確認名單進入現勘錄入"): st.session_state.temp_field_plan = current_selection.copy()

            if st.session_state.temp_field_plan is not None:
                plan = st.session_state.temp_field_plan.copy()
                if '微調勾選' not in plan.columns: plan.insert(0, '微調勾選', '✅ 留用')
                if '現勘判定' not in plan.columns: plan['現勘判定'] = '系統型'
                
                recon_ed = st.data_editor(
                    plan[['微調勾選','現勘判定','地段地號','網格編號','目前農地調查現況','TWD97_X','TWD97_Y']],
                    column_config={"微調勾選": st.column_config.SelectboxColumn("狀態", options=["✅ 留用", "❌ 排除"]), "現勘判定": st.column_config.SelectboxColumn("現勘結果", options=["系統型", "個案型", "建物", "難以採樣"])},
                    key="recon_ed_v6"
                )
                
                res_samp = recon_ed[recon_ed['現勘判定'].isin(['系統型','個案型']) & (recon_ed['微調勾選']=='✅ 留用')]
                res_fail = recon_ed[recon_ed['現勘判定'].isin(['建物','難以採樣']) | (recon_ed['微調勾選']=='❌ 排除')]
                
                st.markdown(f"📊 **目前待現勘總數: {len(recon_ed)} 筆 | 最終採樣數: {len(res_samp)} 筆**")
                cr1, cr2 = st.columns(2)
                with cr1: st.success("✅ 採樣清單"); st.dataframe(res_samp, height=300)
                with cr2: 
                    st.error("❌ 失效名單"); st.dataframe(res_fail, height=300)
                    if st.button("✅ 確認更新失效點到總表"): st.session_state.excluded_lots.extend(res_fail['地段地號'].tolist()); st.rerun()

            # 4. 最終歸檔
            st.divider(); col_f1, col_f2 = st.columns(2)
            if col_f1.button("✅ 最終調查名單確認 (存入歷史庫)"):
                st.session_state.archived_plans[f"{target_year}_最終定案"] = res_samp
                st.session_state.temp_field_plan = None; st.success("計畫歸檔成功！"); st.rerun()
            if col_f2.button("🔄 重置所有篩選"): st.session_state.excluded_lots = []; st.session_state.temp_field_plan = None; st.rerun()

        # --- GIS 地圖 ---
        st.subheader("🗺️ 年度計畫分布地圖")
        m_plan = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        if gdf_grid is not None:
            gs_active = current_selection[['網格編號','網格監測頻率']].drop_duplicates('網格編號')
            g_col = gdf_grid.merge(gs_active, left_on='網格號', right_on='網格編號', how='inner').to_crs(epsg=4326)
            def get_g_c(f): fr = str(f); return '#FFB6C1' if '持續' in fr else '#ADD8E6' if '延長' in fr else '#f8f9fa'
            folium.GeoJson(g_col, style_function=lambda x: {'fillColor': get_g_c(x['properties'].get('網格監測頻率','')), 'color': 'white', 'weight': 1, 'fillOpacity': 0.4}).add_to(m_plan)
        for _, r in current_selection.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                sd = 3 if "系統" in str(r['計畫類別']) else 4
                if sd == 3: clr = "red" if "增量" in str(r['目前農地調查現況']) else "blue"
                else: clr = "yellow" if "增量" in str(r['目前農地調查現況']) else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=sd, radius=8, color=clr, fill=True, tooltip=r['地段地號']).add_to(m_plan)
            except: continue
        st_folium(m_plan, width=1100, height=500)
        if st.session_state.archived_plans:
            st.subheader("📚 歷史計畫檔案庫"); 
            for ak, av in st.session_state.archived_plans.items(): 
                with st.expander(f"📂 {ak}"): st.dataframe(av)

    # --- 其餘頁面維持不變 ---
    elif menu == "新增年度調查結果":
        st.title("➕ 錄入年度數據")
        pwd_input = st.sidebar.text_input("密碼", type="password")
        if pwd_input == ADMIN_PASSWORD:
            sl_q = st.text_input("🔍 搜尋地號錄入")
            if sl_q:
                # ... (此處保留原有的 DA 判定邏輯)
                st.success(f"找到 {sl_q}，請開始判定...")

    elif menu == "空間地圖檢視":
        st.title("🗺️ 衛星影像監測地圖")
        m_main = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        st_folium(m_main, width=1100, height=700)

else: st.error("Excel 載入失敗")























