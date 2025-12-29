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
# 1. 系統初始化與基礎函式 (必須置頂防止 NameError)
# ==========================================

# 初始化所有會話狀態
if 'excluded_lots' not in st.session_state: st.session_state.excluded_lots = []
if 'temp_field_plan' not in st.session_state: st.session_state.temp_field_plan = None
if 'archived_plans' not in st.session_state: st.session_state.archived_plans = {}
if 'manual_selection_list' not in st.session_state: st.session_state.manual_selection_list = pd.DataFrame()
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
# 2. 系統設定與專業美化 CSS
# ==========================================

st.set_page_config(page_title="彰化農地智慧管理系統", layout="wide", page_icon="🌾")
ADMIN_PASSWORD = "ET23597010"
EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"
LOG_PATH = "edit_log.csv"

transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ["汞", "砷", "銅", "鉻", "鎘", "鉛", "鋅", "鎳"]

st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; }
    .stMetric { background-color: #ffffff; padding: 10px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .status-header { padding: 10px; border-radius: 10px; text-align: center; font-weight: bold; font-size: 16px; border: 1px solid #ddd; margin-bottom: 5px; }
    .bg-p { background-color: #FFB6C1; color: #721c24; }
    .bg-l { background-color: #ADD8E6; color: #004085; }
    .bg-e { background-color: #90EE90; color: #155724; }
    .stats-container { background-color: #e8f5e9; padding: 20px; border-radius: 15px; margin-bottom: 20px; border: 1px solid #c8e6c9; }
    .filter-card { background-color: #f8f9fa; padding: 15px; border-radius: 10px; border-left: 5px solid #2e7d32; margin-bottom: 15px; }
    .legend-box { padding: 10px; background: white; border: 1px solid gray; line-height: 1.5em; border-radius: 5px; font-size: 12px; }
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
        if '延長頻率' not in df_m.columns: df_m['延長頻率'] = 10
        df_h = pd.read_excel(xl, sheet_name=get_s("歷年調查紀錄"))
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

def find_grid_by_coords(x, y, gdf):
    if gdf is None: return "未知"
    p = Point(x, y)
    match = gdf[gdf.contains(p)]
    return str(match.iloc[0]['網格號']) if not match.empty else "範圍外"

# ==========================================
# 4. 側邊欄與導覽
# ==========================================

st.sidebar.title("🌿 系統導覽")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

if df_master is not None:
    # 全局數據預處理
    abs_total = len(df_master)
    sampling_pts_count = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])
    
    grid_sys_all = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
    grid_uniq = grid_sys_all.drop_duplicates('網格編號').copy()
    grid_uniq['網格監測頻率'] = grid_uniq['網格監測頻率'].fillna('無狀態')
    
    case_data_master = df_master[~df_master['調查方式'].str.contains('系統', na=False)].copy()
    c_map_ref = {"增量":"持續", "延長":"延長", "正常":"退場", "管制":"管制", "建物":"建物", "難以採樣":"難以採樣"}
    case_data_master['對應狀態'] = case_data_master['目前農地調查現況'].map(c_map_ref).fillna(case_data_master['目前農地調查現況'])

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
        g_c = len(grid_uniq[grid_uniq['網格監測頻率'] == '持續'])
        g_l = len(grid_uniq[grid_uniq['網格監測頻率'] == '延長'])
        g_e = len(grid_uniq[grid_uniq['網格監測頻率'] == '退場'])
        g_sum = g_c + g_l + g_e; g_none = len(grid_uniq) - g_sum
        gc = st.columns(5); gc[0].metric("持續網格", g_c); gc[1].metric("延長網格", g_l); gc[2].metric("退場網格", g_e); gc[3].metric("有效網格合計", g_sum); gc[4].metric("無網格狀態", g_none)
        
        st.divider()
        st.subheader("📦 個案型農地統計")
        cc_m = case_data_master['對應狀態'].value_counts()
        cc = st.columns(6)
        for i, lab in enumerate(["持續", "延長", "退場", "管制", "難以採樣", "建物"]):
            cc[i].metric(lab, cc_m.get(lab, 0))
            
        st.divider()
        st.subheader("🔍 網格查詢系統")
        qs = st.text_input("輸入網格號碼 (如: G2405)")
        if qs:
            res = df_master[df_master['網格編號'] == clean_id(qs)]
            if not res.empty: st.dataframe(res.style.apply(lambda x: ['background-color: #FFFFCC' if x.代表性=='代表點' else '' for _ in x], axis=1))
            else: st.warning("查無此網格編號")
            
        st.divider()
        st.subheader("📊 近三年調查分佈 (樹狀圖)")
        tree_cols = st.columns(3); cy = get_minguo_year()
        for i, y in enumerate([cy-2, cy-1, cy]):
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
        st.title("📂 數據查詢中心")
        admin_mode = False
        with st.sidebar.expander("🔐 管理員修正權限"):
            if st.text_input("輸入修正密碼", type="password") == ADMIN_PASSWORD: admin_mode = True; st.success("編輯模式開啟")
        
        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單", "📜 修改紀錄"])
        
        with tabs[0]:
            sq = st.text_input("🔍 搜尋地號/序號/網格", key="m_s_q")
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號')+1; cols.insert(idx, cols.pop(cols.index('代表性顯示'))); df_p = df_p[cols]
            if sq: df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(sq)).any(axis=1)]
            st.dataframe(df_p, height=800, use_container_width=True)
            towrite = io.BytesIO(); df_master.to_excel(towrite, index=False, engine='xlsxwriter'); st.download_button("📥 下載 Excel", data=towrite.getvalue(), file_name="農地總表.xlsx")
            
        with tabs[1]:
            y_avail_list = sorted(df_history['調查年度'].unique(), reverse=True); y_sel = st.selectbox("選擇年度", y_avail_list if y_avail_list else [113])
            y_res = df_history[df_history['調查年度'] == int(y_sel)].copy()
            if not y_res.empty: st.dataframe(y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left'), use_container_width=True)
            else: st.warning("無數據")
            
        with tabs[2]:
            st.subheader("🏠 坵塊群組管理")
            bq = st.text_input("🔍 搜尋地號確認群組"); 
            if bq and bq in df_block['農地地段地號'].values:
                gid = df_block[df_block['農地地段地號']==bq].iloc[0]['農地群組編號']; st.dataframe(df_block[df_block['農地群組編號']==gid])
            with st.expander("➕ 批次新增同坵塊關聯"):
                try: 
                    last_id = df_block['農地群組編號'].str.extract('(\d+)').dropna().astype(int).max()[0]
                    nid = f"BLOCK_{str(last_id+1).zfill(3)}"
                except: nid = "BLOCK_001"
                with st.form("nb"):
                    st.write(f"下一組編號: **{nid}**"); li = st.text_area("地號清單(每行一筆)"); ri = st.text_input("指定代表點"); 
                    if st.form_submit_button("建立"): st.info("已加入排程")
            st.write("**現有對照表：**"); st.dataframe(df_block, height=400)

        with tabs[3]:
            st.subheader("🌐 系統型分類看板")
            s_q_gt = st.text_input("🔍 全域網格號搜尋", key="s_gt_t")
            sc1, sc2, sc3 = st.columns(3)
            for i, f in enumerate(['持續','延長','退場']):
                with sc1 if i==0 else sc2 if i==1 else sc3:
                    st.markdown(f'<div class="status-header bg-{"p" if i==0 else "l" if i==1 else "e"}">{f}網格</div>', unsafe_allow_html=True)
                    ids = grid_uniq[grid_uniq['網格監測頻率']==f]['網格編號'].tolist()
                    sel = st.selectbox(f"{f}名單", ["未選"]+ids, key=f"s_box_{f}")
                    if sel != "未選":
                        g_data_v = df_master[df_master['網格編號']==sel].copy()
                        st.write(f"筆數: {len(g_data_v)} | 採樣代表: {len(g_data_v[g_data_v['代表性']=='代表點'])}")
                        g_data_v['圖示'] = g_data_v.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
                        st.dataframe(g_data_v[['圖示','農地序號','地段地號','農地監測狀態','目前農地調查現況']])

        with tabs[4]:
            st.subheader("📦 個案型監測看板")
            cq = st.text_input("🔍 搜尋個案地號/序號")
            c_cols_tab = st.columns(6); c_labels = ["持續", "延長", "退場", "管制", "難以採樣", "建物"]
            for i, lab in enumerate(c_labels):
                sub = case_data_master[case_data_master['對應狀態']==lab]
                c_cols_tab[i].metric(lab, len(sub))
            st.write("**個案型總清單：**"); st.dataframe(case_data_master, height=400)

        with tabs[5]:
            st.subheader("📜 系統修改紀錄日誌")
            if os.path.exists(LOG_PATH): st.dataframe(pd.read_csv(LOG_PATH).sort_values(by="修改時間", ascending=False))
            else: st.info("無紀錄")

    # --- C. [重點整合] 新年度調查點篩選名單 ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 調查計畫決策與自動補位系統")
        
        # 模式與設定區
        col_m1, col_m2 = st.columns([2, 1])
        plan_mode = col_m1.radio("選擇模式", ["自動系統產生", "手動挑選名單"], horizontal=True)
        auto_backfill = col_m2.checkbox("開啟自動補位功能", value=True)

        c1, c2 = st.columns(2)
        target_year = c1.number_input("設定目標年度 (民國)", value=get_minguo_year()+1)
        quota = c2.number_input("設定年度預算筆數", value=500, step=10)

        df_calc = df_master.copy()
        
        if plan_mode == "自動系統產生":
            # 演算法：排除 T-1，持續 2年 / 延長 10年
            sys_pool = []
            grids = df_calc[df_calc['調查方式'].str.contains('系統', na=False)]['網格編號'].unique()
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
            
            case_active = df_calc[~df_calc['調查方式'].str.contains('系統', na=False)].copy()
            case_active = case_active[~case_active['農地監測狀態'].isin(['管制','建物','正常'])]
            def c_algo_prio(r):
                if r['最後調查年分'] == target_year - 1: return 99
                if str(r['目前農地調查現況']) == '增量' and (target_year-r['最後調查年分'] >= 2): return 1
                if str(r['目前農地調查現況']) == '延長' and (target_year-r['最後調查年分'] >= 10): return 3
                return 99
            case_active['優先權重'] = case_active.apply(c_algo_prio, axis=1)
            case_pool = case_active[case_active['優先權重'] < 99].copy()
            case_pool['計畫類別'], case_pool['篩選備註'] = '個案型農地', '獨立判定'

            full_pool = pd.concat([pd.concat(sys_pool) if sys_pool else pd.DataFrame(), case_pool]).sort_values(['優先權重', '網格編號'])
            eligible = full_pool[~full_pool['地段地號'].isin(st.session_state.excluded_lots)]
            current_selection = eligible.head(int(quota)).copy() if auto_backfill else eligible.copy()
        
        else: # 手動自選名單 (還原代碼)
            col_s1, col_s2 = st.columns(2)
            with col_s1: st.session_state.manual_sys_grids = st.multiselect("選取網格", df_master['網格編號'].unique(), default=st.session_state.manual_sys_grids)
            with col_s2: st.session_state.manual_case_lots = st.multiselect("選取地號", df_master[~df_master['調查方式'].str.contains('系統', na=False)]['地段地號'].unique(), default=st.session_state.manual_case_lots)
            m_s = df_master[df_master['網格編號'].isin(st.session_state.manual_sys_grids) & (df_master['代表性']=='代表點')].copy()
            m_c = df_master[df_master['地段地號'].isin(st.session_state.manual_case_lots)].copy()
            current_selection = pd.concat([m_s, m_c])
            current_selection = current_selection[~current_selection['地段地號'].isin(st.session_state.excluded_lots)].copy()

        if not current_selection.empty:
            # 統計看板
            st.markdown('<div class="stats-container">', unsafe_allow_html=True)
            sp = current_selection[current_selection['計畫類別']=='系統型網格']; cp = current_selection[current_selection['計畫類別']=='個案型農地']
            k_s1, k_s2, k_s3 = st.columns(3); k_s1.metric("系統農地筆數", len(sp)); k_s2.metric("個案農地筆數", len(cp)); k_s3.metric("總計擬定", len(current_selection))
            st.markdown('</div>', unsafe_allow_html=True)

            # 階段一：名單確認
            st.subheader("第一階段：初步篩選與去留確認 (✅ 留用 / ❌ 排除)")
            current_selection.insert(0, '狀態確認', '✅ 留用')
            
            s1, s2 = st.columns(2)
            with s1: ed_s_p1 = st.data_editor(current_selection[current_selection['計畫類別']=='系統型網格'][['狀態確認','網格編號','地段地號','農地序號','目前農地調查現況']], key="ed_s_p1", column_config={"狀態確認": st.column_config.SelectboxColumn("狀態", options=["✅ 留用", "❌ 排除"], required=True)})
            with s2: ed_c_p1 = st.data_editor(current_selection[current_selection['計畫類別']=='個案型農地'][['狀態確認','地段地號','農地序號','目前農地調查現況']], key="ed_c_p1", column_config={"狀態確認": st.column_config.SelectboxColumn("狀態", options=["✅ 留用", "❌ 排除"], required=True)})

            if st.button("🔥 確定清出打叉 (❌ 排除) 的農地並補位"):
                to_kill = ed_s_p1[ed_s_p1['狀態確認']=="❌ 排除"]['地段地號'].tolist() + ed_c_p1[ed_c_p1['狀態確認']=="❌ 排除"]['地段地號'].tolist()
                st.session_state.excluded_lots.extend(to_kill); st.rerun()

            # 階段二：現勘錄入 (分表顯示)
            st.divider(); st.subheader("第二階段：正式名單與現勘結果錄入")
            if st.button("💾 確認名單進入現勘錄入"): st.session_state.temp_field_plan = current_selection.copy()

            if st.session_state.temp_field_plan is not None:
                plan = st.session_state.temp_field_plan.copy()
                if '微調勾選' not in plan.columns: plan.insert(0, '微調勾選', '✅ 留用')
                if '現勘判定' not in plan.columns: plan['現勘判定'] = '系統型'
                
                recon_ed = st.data_editor(
                    plan[['微調勾選','現勘判定','地段地號','網格編號','目前農地調查現況','TWD97_X','TWD97_Y']],
                    column_config={"微調勾選": st.column_config.SelectboxColumn("狀態", options=["✅ 留用", "❌ 排除"]), "現勘判定": st.column_config.SelectboxColumn("現場結果", options=["系統型", "個案型", "建物", "難以採樣"])},
                    key="recon_ed_v2"
                )
                
                # 分表邏輯
                col_r1, col_r2 = st.columns(2)
                res_samp = recon_ed[recon_ed['現勘判定'].isin(['系統型','個案型']) & (recon_ed['微調勾選']=='✅ 留用')]
                res_fail = recon_ed[recon_ed['現勘判定'].isin(['建物','難以採樣']) | (recon_ed['微調勾選']=='❌ 排除')]
                
                with col_r1: st.markdown('<div class="bg-e status-header">✅ 最終採樣清單</div>', unsafe_allow_html=True); st.dataframe(res_samp, height=300)
                with col_r2:
                    st.markdown('<div class="bg-p status-header">❌ 失效點位 (建物/難以採樣)</div>', unsafe_allow_html=True); st.dataframe(res_fail, height=300)
                    if st.button("✅ 確認更新失效點到總表"): st.session_state.excluded_lots.extend(res_fail['地段地號'].tolist()); st.success("已移除失效點，請執行上方補位。"); st.rerun()

                # 異動分析窗
                if st.button("🔄 分析現勘後對大資料庫的影響"):
                    st.session_state.recon_log = [f"🚩 網格 {r['網格編號']}: 代表點 {r['地段地號']} 已失效。" for _, r in res_fail.iterrows()]
                    st.info("📋 異動分析窗已產出。")
                if st.session_state.recon_log:
                    for log in st.session_state.recon_log: st.write(log)

            # 最終歸檔
            st.divider(); col_f1, col_f2 = st.columns(2)
            if col_f1.button("✅ 最終調查名單確認 (移入歷史庫)"):
                arc_id = f"{target_year}_最終定案_{datetime.now().strftime('%H%M')}"
                st.session_state.archived_plans[arc_id] = res_samp.copy()
                st.session_state.temp_field_plan = None; st.success("計畫歸檔成功！工作區已清空。")
            if col_f2.button("🔄 重置篩選"): st.session_state.excluded_lots = []; st.session_state.temp_field_plan = None; st.rerun()

        # --- GIS 地圖與網格 (移動到最底) ---
        st.divider(); st.subheader("🗺️ 年度擬定計畫分布地圖")
        m_plan = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        
        if gdf_grid is not None:
            gs_active = current_selection[['網格編號','網格監測頻率']].drop_duplicates('網格編號')
            # 只有有狀態的網格才著色
            g_col_gdf = gdf_grid.merge(gs_active, left_on='網格號', right_on='網格編號', how='inner').to_crs(epsg=4326)
            def get_g_c(f): f_s = str(f); return '#FFB6C1' if '持續' in f_s else '#ADD8E6' if '延長' in f_s else '#f8f9fa'
            folium.GeoJson(g_col_gdf, style_function=lambda x: {'fillColor': get_g_c(x['properties'].get('網格監測頻率','')), 'color': 'white', 'weight': 1, 'fillOpacity': 0.4}).add_to(m_plan)
        
        for _, r in current_selection.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                sd = 3 if "系統" in str(r['計畫類別']) else 4
                if sd == 3: clr = "red" if "增量" in str(r['目前農地調查現況']) else "blue"
                else: clr = "yellow" if "增量" in str(r['目前農地調查現況']) else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=sd, radius=8, color=clr, fill=True, tooltip=r['地段地號']).add_to(m_plan)
            except: continue
        st_folium(m_plan, width=1100, height=550)

        # 歷史存檔區
        if st.session_state.archived_plans:
            st.divider(); st.subheader("📚 歷史計畫資料庫")
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
        st.title("🗺️ 衛星影像與網格著色圖")
        m = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        if gdf_grid is not None:
            gs = df_master.drop_duplicates('網格編號')[['網格編號', '網格監測頻率']]
            merged = gdf_grid.to_crs(epsg=4326).merge(gs, left_on='網格號', right_on='網格編號', how='left')
            def get_c(f): return '#FFB6C1' if '持續' in str(f) else '#ADD8E6' if '延長' in str(f) else '#90EE90' if '退場' in str(f) else '#F8F8F8'
            folium.GeoJson(merged, style_function=lambda x: {'fillColor': get_c(x['properties'].get('網格監測頻率')), 'color': 'white', 'weight': 1, 'fillOpacity': 0.4}).add_to(m)
        st_folium(m, width=1100, height=700)
else:
    st.error("❌ Excel 載入失敗")




















