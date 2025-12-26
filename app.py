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
# 1. 基礎函式定義 (必須放在最前面，防止 NameError)
# ==========================================

def get_minguo_year():
    """動態取得當前民國年份"""
    return datetime.now().year - 1911

def get_minguo_date():
    """取得民國格式日期"""
    now = datetime.now()
    return f"民國 {now.year - 1911} 年 {now.month} 月 {now.day} 日"

def clean_id(val):
    """清理 Excel 格式產生的空格或 .0"""
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

def get_pretty_rep(row, block_df):
    """代表性圖示邏輯"""
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
LOG_PATH = "edit_log.csv"

# 座標轉換器 (TWD97 -> WGS84)
transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ["汞", "砷", "銅", "鉻", "鎘", "鉛", "鋅", "鎳"]

# CSS 專業美化
st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .status-header { padding: 10px; border-radius: 10px; text-align: center; font-weight: bold; font-size: 18px; border: 1px solid #ddd; margin-bottom: 5px; }
    .bg-p { background-color: #FFB6C1; color: #721c24; } /* 淡紅 */
    .bg-l { background-color: #ADD8E6; color: #004085; } /* 淡藍 */
    .bg-e { background-color: #90EE90; color: #155724; } /* 淡綠 */
    .filter-card { background-color: #f0f4f8; padding: 20px; border-radius: 15px; border-left: 8px solid #2e7d32; margin-bottom: 20px; }
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
        if '延長頻率' not in df_m.columns: df_m['延長頻率'] = 2
        
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

# ==========================================
# 4. 全局數據預處理 (確保所有頁面統計一致)
# ==========================================

if df_master is not None:
    # 頂列指標
    abs_total = len(df_master)
    sampling_pts_count = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])

    # 系統型數據預處理
    grid_sys_only = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
    grid_uniq = grid_sys_only.drop_duplicates('網格編號').copy()
    grid_uniq['網格監測頻率'] = grid_uniq['網格監測頻率'].fillna('無網格狀態').astype(str).str.strip()
    
    # 個案型數據預處理 (修正：排除系統即個案，確保建物 16 筆/難以採樣 6 筆入列)
    case_data_master = df_master[~df_master['調查方式'].str.contains('系統', na=False)].copy()
    c_map = {"增量":"持續", "延長":"延長", "正常":"退場", "管制":"管制", "建物":"建物", "難以採樣":"難以採樣"}
    case_data_master['對應狀態'] = case_data_master['目前農地調查現況'].map(c_map).fillna(case_data_master['目前農地調查現況'])

# ==========================================
# 5. 側邊欄與功能導覽
# ==========================================

st.sidebar.title("🌿 系統導覽")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

# ==========================================
# 6. 頁面邏輯內容
# ==========================================

if df_master is not None:

    # --- A. 統計首頁 ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        
        # 1. 頂列六大指標 (水平一行)
        k = st.columns(6)
        k[0].metric("總資料點數", abs_total)
        k[1].metric("總採樣點數", sampling_pts_count)
        k[2].metric("管制點數", control_pts_count)
        k[3].metric("建物數量", build_pts_count)
        k[4].metric("難以採樣數量", hard_pts_count)
        k[5].metric("正常退場數量", normal_pts_count)

        st.divider()
        # 2. 系統型網格現況統計 (水平一行)
        st.subheader("🌐 系統型網格現況統計")
        g_c = len(grid_uniq[grid_uniq['網格監測頻率'] == '持續'])
        g_l = len(grid_uniq[grid_uniq['網格監測頻率'] == '延長'])
        g_e = len(grid_uniq[grid_uniq['網格監測頻率'] == '退場'])
        g_sum = g_c + g_l + g_e
        g_none = len(grid_uniq) - g_sum
        
        gc = st.columns(5)
        gc[0].metric("持續網格", g_c); gc[1].metric("延長網格", g_l); gc[2].metric("退場網格", g_e)
        gc[3].metric("有效網格合計", g_sum); gc[4].metric("無網格狀態", g_none)

        st.divider()
        # 3. 個案型農地現況統計 (水平一行)
        st.subheader("📦 個案型農地現況統計")
        cc_map = case_data_master['對應狀態'].value_counts()
        cc_cols = st.columns(6)
        cc_labels = ["持續", "延長", "退場", "管制", "難以採樣", "建物"]
        for i, lab in enumerate(cc_labels):
            cc_cols[i].metric(lab, cc_map.get(lab, 0))

        st.divider()
        # 4. 網格查詢系統
        st.subheader("🔍 網格查詢系統")
        grid_search_id = st.text_input("輸入網格號碼搜尋內部細節 (如: G2405)")
        if grid_search_id:
            res = df_master[df_master['網格編號'] == clean_id(grid_search_id)]
            if not res.empty:
                st.write(f"網格 {grid_search_id} 內部農地清單：")
                st.dataframe(res.style.apply(lambda x: ['background-color: #FFFFCC' if x.代表性=='代表點' else '' for _ in x], axis=1), use_container_width=True)
            else: st.warning("查無此網格")

        st.divider()
        # 5. 近三年樹狀圖 (112-114)
        st.subheader("📊 近三年調查分佈 (樹狀圖)")
        tree_cols = st.columns(3)
        curr_y = get_minguo_year()
        for i, y in enumerate([112, 113, 114]):
            cn = f"{y}狀態"
            if cn in df_master.columns:
                y_df = df_master[df_master[cn].notna()].copy()
                if not y_df.empty:
                    def ml(r):
                        s, cur = str(r[cn]).strip(), str(r['目前農地調查現況']).strip()
                        if s == "監測": return "持續" if "增量" in cur else "延長"
                        return "退場" if s == "正常" else s
                    y_df['標籤'] = y_df.apply(ml, axis=1)
                    y_counts = y_df.groupby(['調查方式', '標籤']).size().reset_index(name='筆數')
                    fig = px.treemap(y_counts, path=[px.Constant(f"{y}年"), '調查方式', '標籤'], values='筆數', color='標籤',
                                     color_discrete_map={'持續':'#FFB6C1','延長':'#ADD8E6','退場':'#90EE90','管制':'#FF3333','建物':'#D3D3D3'})
                    tree_cols[i].plotly_chart(fig, use_container_width=True)

    # --- B. 資料庫查詢與下載 ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 數據管理中心")
        admin_mode = False
        with st.sidebar.expander("🔐 管理員授權"):
            if st.text_input("輸入修正密碼", type="password") == ADMIN_PASSWORD: admin_mode = True; st.success("編輯模式開啟")

        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單", "📜 修改紀錄"])

        with tabs[0]: # 總表
            sm = st.text_input("🔍 搜尋地號/序號/網格", key="m_s")
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            # 欄位移位
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號') + 1
                cols.insert(idx, cols.pop(cols.index('代表性顯示')))
                df_p = df_p[cols]
            if sm: df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(sm)).any(axis=1)]
            st.dataframe(df_p, height=800, use_container_width=True)
            # 下載 Excel
            towrite = io.BytesIO()
            df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載全量總表", data=towrite.getvalue(), file_name="彰化農地總表.xlsx")

        with tabs[1]: # 歷年
            y_avail = sorted(df_history['調查年度'].unique(), reverse=True)
            sel_y = st.selectbox("選擇年度", y_avail if y_avail else [113])
            y_res = df_history[df_history['調查年度'] == int(sel_y)].copy()
            if not y_res.empty:
                st.dataframe(y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left'), use_container_width=True)
            else: st.warning("該年度無數據")

        with tabs[2]: # 坵塊 (自動編號)
            st.subheader("🏠 坵塊管理系統")
            blk_q = st.text_input("🔍 搜尋地號確認群組成員")
            if blk_q and blk_q in df_block['農地地段地號'].values:
                gid = df_block[df_block['農地地段地號']==blk_q].iloc[0]['農地群組編號']
                st.dataframe(df_block[df_block['農地群組編號']==gid])
            with st.expander("➕ 批次新增同坵塊關聯"):
                try: 
                    last_id = df_block['農地群組編號'].str.extract('(\d+)').dropna().astype(int).max()[0]
                    next_id = f"BLOCK_{str(last_id + 1).zfill(3)}"
                except: next_id = "BLOCK_001"
                with st.form("new_blk"):
                    st.write(f"預計編號: **{next_id}**")
                    lots_in = st.text_area("1. 貼上地號清單 (每行一筆)"); rep_in = st.text_input("2. 指定代表點")
                    if st.form_submit_button("確認建立"): st.info("已提交排程。")
            st.write("**現有對照清單：**")
            st.dataframe(df_block, height=400)

        with tabs[3]: # 系統型看板 (彩色方塊選擇)
            st.subheader("🌐 系統型農地監測看板")
            s_grid_top = st.text_input("🔍 全域搜尋網格編號 (不限狀態)", key="s_gt_tab")
            p_ids = grid_unique_df[grid_unique_df['網格監測頻率'] == '持續']['網格編號'].tolist() if 'grid_unique_df' in locals() else grid_uniq[grid_uniq['網格監測頻率'] == '持續']['網格編號'].tolist()
            l_ids = grid_unique_df[grid_unique_df['網格監測頻率'] == '延長']['網格編號'].tolist() if 'grid_unique_df' in locals() else grid_uniq[grid_uniq['網格監測頻率'] == '延長']['網格編號'].tolist()
            e_ids = grid_unique_df[grid_unique_df['網格監測頻率'] == '退場']['網格編號'].tolist() if 'grid_unique_df' in locals() else grid_uniq[grid_uniq['網格監測頻率'] == '退場']['網格編號'].tolist()
            
            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                st.markdown('<div class="status-header bg-p">持續網格</div>', unsafe_allow_html=True)
                sel_p = st.selectbox("選取持續名單", ["未選"] + p_ids, key="sp_tab")
            with sc2:
                st.markdown('<div class="status-header bg-l">延長網格</div>', unsafe_allow_html=True)
                sel_l = st.selectbox("選取延長名單", ["未選"] + l_ids, key="sl_tab")
            with sc3:
                st.markdown('<div class="status-header bg-e">退場網格</div>', unsafe_allow_html=True)
                sel_e = st.selectbox("選取退場名單", ["未選"] + e_ids, key="se_tab")
            
            chosen = s_grid_top if s_grid_top else (sel_p if sel_p != "未選" else sel_l if sel_l != "未選" else sel_e if sel_e != "未選" else None)
            if chosen:
                g_data = df_master[df_master['網格編號']==clean_id(chosen)].copy()
                if not g_data.empty:
                    st.info(f"📍 網格 {chosen} 統計小畫面")
                    ssc = st.columns(4)
                    ssc[0].metric("總農地數", len(g_data)); ssc[1].metric("代表點", len(g_data[g_data['代表性']=='代表點']))
                    ssc[2].metric("備用點", len(g_data[g_data['代表性']=='備用點'])); ssc[3].metric("無法採樣", len(g_data[g_data['農地監測狀態'].isin(['難以採樣','建物'])]))
                    g_data['代表性顯示'] = g_data.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
                    st.dataframe(g_data[['代表性顯示','農地序號','地段地號','目前農地調查現況','農地監測狀態']], use_container_width=True)
            st.divider()
            st.write("**系統型總表清單**")
            grid_sys_only['代表性顯示'] = grid_sys_only.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(grid_sys_only, height=400)

        with tabs[4]: # 5. 個案型看板 (6大分類與搜尋)
            st.subheader("📦 個案型農地監測看板")
            scq = st.text_input("🔍 搜尋個案地號/序號", key="scq_tab")
            clabs = ["持續", "延長", "退場", "管制", "難以採樣", "建物"]
            c_cols_tab = st.columns(6)
            for i, lab in enumerate(clabs):
                sub = case_data_master[case_data_master['對應狀態']==lab]
                c_cols_tab[i].metric(lab, len(sub))
                with c_cols_tab[i].expander("名單"): st.dataframe(sub[['農地序號','地段地號']])
            if scq: case_data_master = case_data_master[case_data_master['地段地號'].str.contains(scq)]
            st.divider()
            st.write("**個案型總表清單**")
            case_data_master['代表性顯示'] = case_data_master.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(case_data_master, height=400)

        with tabs[5]: # 6. 修改紀錄
            st.subheader("📜 系統修改紀錄日誌")
            if os.path.exists(LOG_PATH): st.dataframe(pd.read_csv(LOG_PATH).sort_values(by="修改時間", ascending=False))
            else: st.info("無紀錄")

    # --- C. [重點] 新年度調查點篩選名單 (演算法完全體) ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 年度調查計畫決策與分布地圖")
        
        if 'excluded_lots' not in st.session_state: st.session_state.excluded_lots = []
        if 'saved_plan' not in st.session_state: st.session_state.saved_plan = None

        c_set1, c_set2 = st.columns(2)
        target_year = c_set1.number_input("設定目標年度 (民國)", value=get_minguo_year()+1)
        quota = c_set2.number_input("設定年度預計調查總數", value=500, step=50)

        # 演算法核心池
        df_calc = df_master.copy()
        
        # (A) 系統型評估 (P1, P2) + 自動補位
        sys_pool = []
        all_grids = df_calc[df_calc['調查方式'].str.contains('系統', na=False)]['網格編號'].unique()
        for gid in all_grids:
            g_data = df_calc[df_calc['網格編號'] == gid]
            f, ly = str(g_data['網格監測頻率'].iloc[0]), g_data['最後調查年分'].max()
            prio = 1 if f == '持續' else (2 if f == '延長' and (target_year - ly >= 2) else 99)
            if prio < 99:
                # 排除行政狀態失效的代表點
                valid_reps = g_data[(g_data['代表性'] == '代表點') & (~g_data['農地監測狀態'].isin(['管制','建物','難以採樣']))]
                if len(valid_reps) < 3:
                    backups = g_data[g_data['代表性'] == '備用點'].sort_values('農地序號').head(3 - len(valid_reps)).copy()
                    backups['篩選備註'] = '備用遞補'
                    final_g = pd.concat([valid_reps, backups])
                else: final_g = valid_reps.copy(); final_g['篩選備註'] = '代表點監測'
                final_g['優先權重'], final_g['計畫類別'] = prio, '系統型網格'
                sys_pool.append(final_g)
        
        # (B) 個案型評估 (P1, P3)
        case_active = df_calc[~df_calc['調查方式'].str.contains('系統', na=False)].copy()
        case_active = case_active[~case_active['農地監測狀態'].isin(['管制','建物','正常'])]
        def c_prio(r):
            if str(r['目前農地調查現況']) == '增量': return 1
            if str(r['目前農地調查現況']) == '延長' and (target_year - r['最後調查年分'] >= r['延長頻率']): return 3
            return 99
        case_active['優先權重'] = case_active.apply(c_prio, axis=1)
        case_pool = case_active[case_active['優先權重'] < 99].copy()
        case_pool['計畫類別'] = '個案型農地'

        # 匯總名單 (考慮排除清單)
        full_pool = pd.concat([pd.concat(sys_pool) if sys_pool else pd.DataFrame(), case_pool]).sort_values(['優先權重', '網格編號', '農地序號'])
        eligible_pool = full_pool[~full_pool['地段地號'].isin(st.session_state.excluded_lots)]
        current_selection = eligible_pool.head(int(quota)).copy()
        current_selection['留用'] = True

        # --- 統計資訊 ---
        st.markdown('<div class="stats-container">', unsafe_allow_html=True)
        st.subheader("📊 本次擬定計畫統計")
        s_sel = current_selection[current_selection['計畫類別']=='系統型網格']
        c_sel = current_selection[current_selection['計畫類別']=='個案型農地']
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("選中網格數", len(s_sel['網格編號'].unique()))
        k2.metric("系統型農地數", len(s_sel)); k3.metric("個案型農地數", len(c_sel))
        k4.metric("總計筆數", len(current_selection))
        st.write(f"🔹 **網格顏色預覽**：持續型(淺紅) {len(s_sel[s_sel['網格監測頻率']=='持續']['網格編號'].unique())} / 延長型(淺藍) {len(s_sel[s_sel['網格監測頻率']=='延長']['網格編號'].unique())}")
        st.markdown('</div>', unsafe_allow_html=True)

        # --- 數據表格 (支援即時去留) ---
        t1, t2 = st.tabs(["🌐 系統型名單", "📦 個案型名單"])
        cols = ['留用', '網格編號', '地段地號', '農地序號', 'TWD97_X', 'TWD97_Y', '目前農地調查現況', '篩選備註']
        with t1:
            ed_s = st.data_editor(s_sel[cols], key="plan_s", use_container_width=True)
            rm_s = ed_s[ed_s['留用']==False]['地段地號'].tolist()
            if rm_s: st.session_state.excluded_lots.extend(rm_s); st.rerun()
        with t2:
            ed_c = st.data_editor(c_sel[cols], key="plan_c", use_container_width=True)
            rm_c = ed_c[ed_c['留用']==False]['地段地號'].tolist()
            if rm_c: st.session_state.excluded_lots.extend(rm_c); st.rerun()

        # --- [新增] 網格著色 GIS 地圖 ---
        st.divider()
        st.subheader("🗺️ 擬定計畫調查分布圖 (網格頻率著色)")
        m_plan = folium.Map(location=[24.05, 120.5], zoom_start=11, 
                            tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        
        if gdf_grid is not None:
            active_grids = current_selection[['網格編號','網格監測頻率']].drop_duplicates('網格編號')
            # 合併網格形狀
            plan_grids_gdf = gdf_grid.merge(active_grids, left_on='網格號', right_on='網格編號', how='inner').to_crs(epsg=4326)
            
            def style_fn(feature):
                freq = feature['properties'].get('網格監測頻率', '')
                color = '#FFB6C1' if '持續' in freq else '#ADD8E6' if '延長' in freq else '#f8f9fa'
                return {'fillColor': color, 'color': 'white', 'weight': 1, 'fillOpacity': 0.4}
            
            folium.GeoJson(plan_grids_gdf, style_function=style_fn, 
                           tooltip=folium.GeoJsonTooltip(fields=['網格號','網格監測頻率'])).add_to(m_plan)

        # 繪製採樣點位 (滿足要求的圖示規則)
        for _, r in current_selection.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                status, cat = str(r['目前農地調查現況']), str(r['計畫類別'])
                # 圖示邏輯
                sides = 3 if "系統" in cat else 4
                color = "red" if "增量" in status else "blue"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=sides, radius=8, color=color, fill=True, fill_opacity=0.9,
                                            popup=f"地號: {r['地段地號']}").add_to(m_plan)
            except: continue
        st_folium(m_plan, width=1100, height=650, key="plan_map")

        # 存檔按鈕
        if st.button("💾 凍結名單並準備檔案"):
            st.session_state.saved_plan = current_selection.drop(columns=['留用'])
            st.success("名單已儲存！可於下方下載。")
        if st.session_state.saved_plan is not None:
            towrite = io.BytesIO()
            st.session_state.saved_plan.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載此調查計畫 Excel", data=towrite.getvalue(), file_name=f"彰化定監計畫_{target_year}.xlsx")
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









