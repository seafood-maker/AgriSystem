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
    """取得當前民國年份"""
    return datetime.now().year - 1911

def get_minguo_date():
    """取得民國格式日期"""
    now = datetime.now()
    return f"民國 {now.year - 1911} 年 {now.month} 月 {now.day} 日"

def clean_id(val):
    """清理 Excel 格式產生的空格或 .0"""
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

# ==========================================
# 2. 系統設定與美化
# ==========================================

st.set_page_config(page_title="彰化農地智慧管理系統 26.0", layout="wide", page_icon="🌾")

ADMIN_PASSWORD = "ET23597010"
EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"
LOG_PATH = "edit_log.csv"

# 座標轉換器 (TWD97 -> WGS84)
transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ["汞", "砷", "銅", "鉻", "鎘", "鉛", "鋅", "鎳"]

# CSS 美化
st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; }
    .stMetric { background-color: #ffffff; padding: 10px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .status-header { padding: 10px; border-radius: 10px; text-align: center; font-weight: bold; font-size: 18px; border: 1px solid #ddd; margin-bottom: 5px; }
    .bg-p { background-color: #FFB6C1; color: #721c24; } /* 淡紅 */
    .bg-l { background-color: #ADD8E6; color: #004085; } /* 淡藍 */
    .bg-e { background-color: #90EE90; color: #155724; } /* 淡綠 */
    .filter-card { background-color: #f0f4f8; padding: 20px; border-radius: 15px; border-left: 8px solid #2e7d32; margin-bottom: 20px; }
    </style>
    """, unsafe_allow_html=True)

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

df_master, df_history, df_block, df_settings = load_all_data()

# ==========================================
# 4. 主選單與數據預算 (解決 NameError)
# ==========================================

st.sidebar.title("🌿 系統導覽")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

if df_master is not None:
    # --- 全域統計數據 ---
    abs_total = len(df_master)
    sampling_pts_count = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])

    # --- A. 統計首頁 ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        k = st.columns(6)
        k[0].metric("總資料點數", abs_total); k[1].metric("總採樣點數", sampling_pts_count)
        k[2].metric("管制點數", control_count); k[3].metric("建物數量", build_count)
        k[4].metric("難以採樣數量", hard_count); k[5].metric("正常退場數量", normal_count)
        
        st.divider()
        st.subheader("📊 近三年調查分佈 (樹狀圖)")
        tree_cols = st.columns(3)
        cy = get_minguo_year()
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
        st.title("📂 數據管理中心")
        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單", "📜 修改紀錄"])

        # 1. 總表
        with tabs[0]:
            search_m = st.text_input("🔍 搜尋地號/序號/網格", key="m_search")
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號') + 1
                cols.insert(idx, cols.pop(cols.index('代表性顯示')))
                df_p = df_p[cols]
            if search_m: df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(search_m)).any(axis=1)]
            st.dataframe(df_p, height=800, use_container_width=True)

        # 2. 歷史
        with tabs[1]:
            y_avail = sorted(df_history['調查年度'].unique(), reverse=True)
            sel_y = st.selectbox("選擇查詢年度", y_avail if y_avail else [113])
            y_res = df_history[df_history['調查年度'] == int(sel_y)].copy()
            if not y_res.empty:
                st.dataframe(y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left'), use_container_width=True)

        # 3. 坵塊
        with tabs[2]:
            st.subheader("🏠 坵塊搜尋")
            blk_q = st.text_input("🔍 搜尋地號找群組夥伴")
            if blk_q and blk_q in df_block['農地地段地號'].values:
                gid = df_block[df_block['農地地段地號']==blk_q].iloc[0]['農地群組編號']
                st.dataframe(df_block[df_block['農地群組編號']==gid])
            st.write("**現有對照清單：**")
            st.dataframe(df_block, height=400, use_container_width=True)

        # 4. 系統型農地清單 (滿足方塊與全部顯示要求)
        with tabs[3]:
            st.subheader("🌐 系統型農地監測看板")
            s_grid_q = st.text_input("🔍 快速搜尋網格編號 (不限狀態)", key="s_g_q")
            sys_df = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
            grid_uniq = sys_df.drop_duplicates('網格編號')
            
            p_ids = grid_uniq[grid_uniq['網格監測頻率'] == '持續']['網格編號'].tolist()
            l_ids = grid_uniq[grid_uniq['網格監測頻率'] == '延長']['網格編號'].tolist()
            e_ids = grid_uniq[grid_uniq['網格監測頻率'] == '退場']['網格編號'].tolist()

            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f'<div class="status-header bg-p">持續網格 ({len(p_ids)})</div>', unsafe_allow_html=True)
                sel_p = st.selectbox("選取持續網格", ["未選擇"] + p_ids, key="sel_p")
            with c2:
                st.markdown(f'<div class="status-header bg-l">延長網格 ({len(l_ids)})</div>', unsafe_allow_html=True)
                sel_l = st.selectbox("選取延長網格", ["未選擇"] + l_ids, key="sel_l")
            with c3:
                st.markdown(f'<div class="status-header bg-e">退場網格 ({len(e_ids)})</div>', unsafe_allow_html=True)
                sel_e = st.selectbox("選取退場網格", ["未選擇"] + e_ids, key="sel_e")

            # 判定顯示哪個網格
            chosen_g = s_grid_q if s_grid_q else (sel_p if sel_p != "未選擇" else sel_l if sel_l != "未選擇" else sel_e if sel_e != "未選擇" else None)
            
            if chosen_g:
                g_data = df_master[df_master['網格編號'] == clean_id(chosen_g)].copy()
                if not g_data.empty:
                    st.info(f"📍 網格 {chosen_g} 詳情：總數 {len(g_data)} | 採樣 {len(g_data[g_data['代表性']=='代表點'])} | 備用 {len(g_data[g_data['代表性']=='備用點'])}")
                    g_data['代表性顯示'] = g_data.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
                    st.dataframe(g_data[['代表性顯示','農地序號','地段地號','目前農地調查現況','農地監測狀態']], use_container_width=True)

            st.divider()
            st.write("**系統型農地總清單**")
            sys_df['代表性顯示'] = sys_df.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(sys_df, height=400, use_container_width=True)

        # 5. 個案型農地清單 (補齊看板與搜尋)
        with tabs[4]:
            st.subheader("📦 個案型農地監測看板")
            s_case_q = st.text_input("🔍 搜尋個案地號/序號", key="s_c_q")
            case_df = df_master[~df_master['調查方式'].str.contains('系統', na=False)].copy()
            c_mapping = {'持續':'增量', '延長':'延長', '退場':'正常', '管制':'管制', '難以採樣':'難以採樣', '建物':'建物'}
            c_met = st.columns(6)
            for i, (lab, val) in enumerate(c_mapping.items()):
                sub = case_df[case_df['目前農地調查現況']==val]
                c_met[i].metric(lab, len(sub))
                with c_met[i].expander("名單"): st.dataframe(sub[['農地序號','地段地號']])
            
            if s_case_q: case_df = case_df[case_df.astype(str).apply(lambda x: x.str.contains(s_case_q)).any(axis=1)]
            st.divider()
            st.write("**個案型農地總清單**")
            case_df['代表性顯示'] = case_df.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(case_df, height=400, use_container_width=True)

    # --- C. [重點功能] 新年度調查點篩選名單 ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 年度調查計畫決策系統")
        c1, c2 = st.columns(2)
        target_year = c1.number_input("1. 設定計畫目標年度 (民國)", value=get_minguo_year()+1)
        quota = c2.number_input("2. 設定本年度預計調查配額 (筆數)", value=500, step=50)

        st.markdown(f"""<div class="filter-card"><b>💡 篩選邏輯 (Priority)：</b><br>
        1. <b>P1 (最優先)：</b>持續型系統網格代表點 + 個案型『增量』農地。<br>
        2. <b>P2 (次優先)：</b>延長型系統網格到期點 (距上次調查 >= 2年)。<br>
        3. <b>P3 (最後)：</b>延長型個案到期農地。<br>
        4. <b>遞補機制：</b>網格代表點若失效，自動依序號由備用點補齊至 3 筆。</div>""", unsafe_allow_html=True)

        # 演算法執行
        # (A) 系統型評估
        sys_final = []
        grids = df_master[df_master['調查方式'].str.contains('系統', na=False)]['網格編號'].unique()
        for gid in grids:
            g_data = df_master[df_master['網格編號'] == gid]
            freq, last_y = str(g_data['網格監測頻率'].iloc[0]), g_data['最後調查年分'].max()
            prio = 1 if freq == '持續' else (2 if freq == '延長' and (target_year - last_y >= 2) else 99)
            if prio < 99:
                active_reps = g_data[(g_data['代表性'] == '代表點') & (~g_data['農地監測狀態'].isin(['管制','建物','難以採樣']))]
                if len(active_reps) < 3:
                    backups = g_data[g_data['代表性'] == '備用點'].sort_values('農地序號').head(3 - len(active_reps)).copy()
                    backups['篩選備註'] = '備用點遞補'
                    final_g = pd.concat([active_reps, backups])
                else: final_g = active_reps.copy(); final_g['篩選備註'] = '代表點監測'
                final_g['優先權重'], final_g['計畫類別'] = prio, f'網格({freq})'
                sys_final.append(final_g)

        # (B) 個案型評估
        case_active = df_master[~df_master['調查方式'].str.contains('系統', na=False)].copy()
        case_active = case_active[~case_active['農地監測狀態'].isin(['管制','建物','正常'])]
        def c_algo(r):
            if str(r['目前農地調查現況']) == '增量': return 1
            if str(r['目前農地調查現況']) == '延長' and (target_year - r['最後調查年分'] >= r['延長頻率']): return 3
            return 99
        case_active['優先權重'] = case_active.apply(c_algo, axis=1)
        case_plan = case_active[case_active['優先權重'] < 99].copy()
        case_plan['計畫類別'], case_plan['篩選備註'] = '個案型', '獨立判定'

        full_pool = pd.concat([pd.concat(sys_final) if sys_final else pd.DataFrame(), case_plan]).sort_values('優先權重')
        final_list = full_pool.head(int(quota))
        st.subheader(f"🎯 {target_year} 年度建議清單 (共 {len(final_list)} 筆)")
        st.dataframe(final_list[['優先權重','網格編號','地段地號','農地序號','計畫類別','目前農地調查現況','篩選備註']], use_container_width=True)

    # --- D. 新增結果 ---
    elif menu == "新增年度調查結果":
        st.title("➕ 錄入年度數據")
        pwd = st.sidebar.text_input("後台密碼", type="password")
        if pwd == ADMIN_PASSWORD: st.success("解鎖成功")
        else: st.warning("密碼鎖定中")

    # --- E. 空間地圖 ---
    elif menu == "空間地圖檢視":
        st.title("🗺️ 衛星影像監測圖")
        m = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        st_folium(m, width=1100, height=700)
else:
    st.error("❌ Excel 載入失敗")





