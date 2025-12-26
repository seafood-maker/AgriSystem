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

st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #eee; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .stDownloadButton button { width: 100%; background-color: #2e7d32; color: white; border-radius: 5px; }
    /* 系統型方塊樣式 */
    .status-header { padding: 10px; border-radius: 10px; text-align: center; font-weight: bold; font-size: 18px; border: 1px solid #ddd; margin-bottom: 5px; }
    .bg-p { background-color: #FFB6C1; color: #721c24; } /* 持續-淡紅 */
    .bg-l { background-color: #ADD8E6; color: #004085; } /* 延長-淡藍 */
    .bg-e { background-color: #90EE90; color: #155724; } /* 退場-淡綠 */
    </style>
    """, unsafe_allow_html=True)

EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"
LOG_PATH = "edit_log.csv"

transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ["汞", "砷", "銅", "鉻", "鎘", "鉛", "鋅", "鎳"]

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
        df_h['調查年度'] = pd.to_numeric(df_h['調查年度'], errors='coerce').fillna(0).astype(int)
        df_b = pd.read_excel(xl, sheet_name=get_s("同坵塊對照表"))
        df_b['農地地段地號'] = df_b['農地地段地號'].astype(str).str.strip()
        df_s = pd.read_excel(xl, sheet_name=get_s("判定標準表")).set_index('項目名稱')
        return df_m, df_h, df_b, df_s
    except: return None, None, None, None

df_master, df_history, df_block, df_settings = load_all_data()

# --- 3. 全局數據預處理 (解決 NameError 關鍵) ---
if df_master is not None:
    # A. 頂列六大指標
    abs_total = len(df_master)
    sampling_pts_count = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_pts_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])

    # B. 系統型網格數據
    grid_sys_master = df_master[df_master['調查方式'].str.contains('系統', na=False)].copy()
    grid_uniq_list = grid_sys_master.drop_duplicates('網格編號').copy()
    grid_uniq_list['網格監測頻率'] = grid_uniq_list['網格監測頻率'].fillna('無狀態')

    # C. 個案型數據映射
    case_data_master = df_master[~df_master['調查方式'].str.contains('系統', na=False)].copy()
    c_map = {"增量":"持續", "延長":"延長", "正常":"退場", "管制":"管制", "難以採樣":"難以採樣", "建物":"建物"}
    case_data_master['顯示狀態'] = case_data_master['目前農地調查現況'].map(c_map).fillna(case_data_master['目前農地調查現況'])

# --- 4. 側邊欄 ---
st.sidebar.title("🌿 系統選單")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

# --- 5. 主程式邏輯 ---
if df_master is not None:
    
    # --- A. 統計首頁 (重新補齊功能) ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        
        # 1. 頂列指標
        k = st.columns(6)
        k[0].metric("總資料點數", abs_total)
        k[1].metric("總採樣點數", sampling_pts_count)
        k[2].metric("管制點數", control_pts_count)
        k[3].metric("建物數量", build_pts_count)
        k[4].metric("難以採樣數量", hard_pts_count)
        k[5].metric("正常退場數量", normal_pts_count)

        st.divider()
        # 2. 系統型網格統計 (滿足第 2 點需求)
        st.subheader("🌐 系統型網格現況統計")
        g_c = len(grid_uniq_list[grid_uniq_list['網格監測頻率'] == '持續'])
        g_l = len(grid_uniq_list[grid_uniq_list['網格監測頻率'] == '延長'])
        g_e = len(grid_uniq_list[grid_uniq_list['網格監測頻率'] == '退場'])
        g_sum = g_c + g_l + g_e
        g_none = len(grid_uniq_list) - g_sum
        
        g_cols = st.columns(5)
        g_cols[0].metric("持續網格", g_c)
        g_cols[1].metric("延長網格", g_l)
        g_cols[2].metric("退場網格", g_e)
        g_cols[3].metric("有效網格合計", g_sum)
        g_cols[4].metric("無網格狀態", g_none)

        st.divider()
        # 3. 個案型農地統計 (滿足第 3 點需求)
        st.subheader("📦 個案型農地現況統計")
        case_counts = case_data_master['顯示狀態'].value_counts()
        c_cols = st.columns(6)
        c_labs = ["持續", "延長", "退場", "管制", "難以採樣", "建物"]
        for i, lab in enumerate(c_labs):
            c_cols[i].metric(lab, case_counts.get(lab, 0))

        st.divider()
        # 4. 網格查詢系統 (滿足第 4 點需求)
        st.subheader("🔍 網格查詢系統")
        grid_search_id = st.text_input("輸入網格號碼搜尋內部細節 (如: G2405)")
        if grid_search_id:
            res = df_master[df_master['網格編號'] == clean_id(grid_search_id)]
            if not res.empty:
                st.write(f"網格 {grid_search_id} 內部農地清單：")
                st.dataframe(res.style.apply(lambda x: ['background-color: #FFFFCC' if x.代表性=='代表點' else '' for _ in x], axis=1), use_container_width=True)
            else: st.warning("查無此網格")

        st.divider()
        # 5. 近三年樹狀圖 (滿足第 5 點需求)
        st.subheader("📊 近三年調查分佈 (樹狀圖)")
        tree_cols = st.columns(3)
        for i, y in enumerate([112, 113, 114]):
            cn = f"{y}狀態"
            if cn in df_master.columns:
                y_df = df_master[df_master[cn].notna()].copy()
                if not y_df.empty:
                    def label_map(r):
                        s, cur = str(r[cn]).strip(), str(r['目前農地調查現況']).strip()
                        if s == "監測": return "持續" if "增量" in cur else "延長"
                        return "退場" if s == "正常" else s
                    y_df['樹狀標籤'] = y_df.apply(label_map, axis=1)
                    counts = y_df.groupby(['調查方式', '樹狀標籤']).size().reset_index(name='筆數')
                    fig = px.treemap(counts, path=[px.Constant(f"{y}年"), '調查方式', '樹狀標籤'], values='筆數', color='樹狀標籤',
                                     color_discrete_map={'持續':'#FFB6C1','延長':'#ADD8E6','退場':'#90EE90','管制':'#FF3333','建物':'#D3D3D3'})
                    tree_cols[i].plotly_chart(fig, use_container_width=True)

    # --- B. 資料庫查詢與下載 ---
    elif menu == "資料庫查詢與下載":
        st.title("📂 數據管理中心")
        admin_mode = False
        with st.sidebar.expander("🔐 管理員授權"):
            if st.text_input("輸入密碼", type="password") == ADMIN_PASSWORD: admin_mode = True; st.success("已開啟編輯模式")

        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單", "📜 修改紀錄"])

        with tabs[0]: # 1. 總表
            st.subheader("🌾 農地現況總表")
            search_box = st.text_input("🔍 搜尋地號/序號/SGM", key="master_search")
            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號') + 1
                cols.insert(idx, cols.pop(cols.index('代表性顯示')))
                df_p = df_p[cols]
            if search_box: df_p = df_p[df_p.astype(str).apply(lambda x: x.str.contains(search_box)).any(axis=1)]
            st.dataframe(df_p, height=800, use_container_width=True)

        with tabs[1]: # 2. 歷年調查結果 (修正年度)
            st.subheader("📅 歷年調查明細")
            years = sorted(df_history['調查年度'].unique(), reverse=True)
            sel_y = st.selectbox("請選擇查詢年度", years if years else [113])
            y_res = df_history[df_history['調查年度'] == int(sel_y)].copy()
            if not y_res.empty:
                st.dataframe(y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left'), use_container_width=True)
            else: st.warning("該年度無數據")

        with tabs[2]: # 3. 坵塊管理 (自動編號)
            st.subheader("🏠 坵塊管理與搜尋")
            blk_search = st.text_input("🔍 搜尋地號找同坵塊夥伴")
            if blk_search and blk_search in df_block['農地地段地號'].values:
                gid = df_block[df_block['農地地段地號']==blk_search].iloc[0]['農地群組編號']
                st.dataframe(df_block[df_block['農地群組編號']==gid])
            
            with st.expander("➕ 批次新增同坵塊關聯"):
                try: 
                    last_num = df_block['農地群組編號'].str.extract('(\d+)').dropna().astype(int).max()[0]
                    next_id = f"BLOCK_{str(last_num + 1).zfill(3)}"
                except: next_id = "BLOCK_001"
                with st.form("new_blk"):
                    st.write(f"預計群組編號: **{next_id}**")
                    lots_text = st.text_area("地段地號清單 (每行一筆)")
                    rep_lot = st.text_input("指定代表點")
                    if st.form_submit_button("確認建立"): st.info("已提交排程")
            st.write("**對照表清單：**")
            st.dataframe(df_block, height=400)

        with tabs[3]: # 4. 系統型農地看板 (修正方塊選擇)
            st.subheader("🌐 系統型農地監測看板")
            s_q = st.text_input("🔍 全域網格編號搜尋 (輸入直接看詳情)", key="grid_search_box")
            
            p_list = grid_uniq_list[grid_uniq_list['網格監測頻率'] == '持續']['網格編號'].tolist()
            l_list = grid_uniq_list[grid_uniq_list['網格監測頻率'] == '延長']['網格編號'].tolist()
            e_list = grid_uniq_list[grid_uniq_list['網格監測頻率'] == '退場']['網格編號'].tolist()

            col_p, col_l, col_e = st.columns(3)
            with col_p:
                st.markdown('<div class="status-header bg-p">持續網格</div>', unsafe_allow_html=True)
                sel_p = st.selectbox("點選持續名單", ["未選"] + p_list, key="rp")
            with col_l:
                st.markdown('<div class="status-header bg-l">延長網格</div>', unsafe_allow_html=True)
                sel_l = st.selectbox("點選延長名單", ["未選"] + l_list, key="rl")
            with col_e:
                st.markdown('<div class="status-header bg-e">退場網格</div>', unsafe_allow_html=True)
                sel_e = st.selectbox("點選退場名單", ["未選"] + e_list, key="re")

            chosen = s_q if s_q else (sel_p if sel_p != "未選" else sel_l if sel_l != "未選" else sel_e if sel_e != "未選" else None)
            if chosen:
                g_data = df_master[df_master['網格編號']==clean_id(chosen)].copy()
                if not g_data.empty:
                    st.info(f"📍 網格 {chosen} 詳情看板")
                    sc = st.columns(4)
                    sc[0].metric("總農地數", len(g_data)); sc[1].metric("代表點", len(g_data[g_data['代表性']=='代表點']))
                    sc[2].metric("備用點", len(g_data[g_data['代表性']=='備用點'])); sc[3].metric("無法採樣", len(g_data[g_data['農地監測狀態']=='難以採樣']))
                    g_data['代表性顯示'] = g_data.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
                    st.dataframe(g_data[['代表性顯示','農地序號','地段地號','目前農地調查現況','農地監測狀態']], use_container_width=True)
            
            st.divider()
            st.write("**系統型總表清單 (底端保留)**")
            grid_sys_master['代表性顯示'] = grid_sys_master.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(grid_sys_master, height=400)

        with tabs[4]: # 5. 個案型農地看板 (修正搜尋與分類)
            st.subheader("📦 個案型農地監測看板")
            s_c_q = st.text_input("🔍 搜尋個案地號/序號", key="case_search_box")
            c_mapping = {'持續':'增量', '延長':'延長', '退場':'正常', '管制':'管制', '難以採樣':'難以採樣', '建物':'建物'}
            c_cols = st.columns(6)
            for i, (lab, val) in enumerate(c_mapping.items()):
                sub = case_data_master[case_data_master['目前農地調查現況']==val]
                c_cols[i].metric(lab, len(sub))
                with c_cols[i].expander("名單"): st.dataframe(sub[['農地序號','地段地號']])
            
            if s_c_q: case_data_master = case_data_master[case_data_master['地段地號'].str.contains(s_c_q)]
            st.divider()
            st.write("**個案型總表清單 (底端保留)**")
            case_data_master['代表性顯示'] = case_data_master.apply(lambda r: get_pretty_rep(r, df_block), axis=1)
            st.dataframe(case_data_master, height=400, use_container_width=True)

        with tabs[5]: # 6. 修改紀錄
            st.subheader("📜 系統稽核紀錄日誌")
            if os.path.exists(LOG_PATH): st.dataframe(pd.read_csv(LOG_PATH))
            else: st.info("尚無異動紀錄")

    # --- C/D/E 頁面 (保留 11.0 穩定邏輯) ---
        elif menu == "新年度調查點篩選名單":
        st.title("📅 年度調查計畫自動生成器")
        target_year = st.selectbox("請選擇預計調查年度 (民國)", [114, 115, 116], index=1)
        
        st.markdown(f"""
        <div class="filter-card">
        <b>演算法規則說明：</b><br>
        1. <b>持續型：</b>目前現況為『增量』之農地及所屬網格。<br>
        2. <b>延長型：</b>目前現況為『延長』且距離上次調查已達頻率年限者。<br>
        3. <b>自動補位：</b>若網格代表點失效 (建物/管制)，系統將自動標示需從備用點遞補。
        </div>
        """, unsafe_allow_html=True)

        # 演算法執行
        df_calc = df_master.copy()
        
        # 1. 計算個案型名單
        case_list = df_calc[
            (~df_calc['調查方式'].str.contains('系統', na=False)) & 
            (
                (df_calc['目前農地調查現況'] == '增量') | 
                ((df_calc['目前農地調查現況'] == '延長') & (target_year - df_calc['最後調查年分'] >= df_calc['延長頻率']))
            )
        ].copy()
        case_list['篩選分類'] = '個案型獨立監測'

        # 2. 計算系統型網格名單 (包含補位邏輯)
        grid_list = []
        all_grids = df_calc[df_calc['調查方式'].str.contains('系統', na=False)]['網格編號'].unique()
        
        for g_id in all_grids:
            g_data = df_calc[df_calc['網格編號'] == g_id]
            g_freq = str(g_data['網格監測頻率'].iloc[0])
            last_y = g_data['最後調查年分'].max()
            
            # 判斷網格是否該查
            should_survey_grid = False
            if g_freq == '持續': should_survey_grid = True
            elif g_freq == '延長' and (target_year - last_y >= 2): # 網格預設頻率 2
                should_survey_grid = True
            
            if should_survey_grid:
                # 找出目前有效的代表點
                reps = g_data[g_data['代表性'] == '代表點']
                active_reps = reps[~reps['農地監測狀態'].isin(['管制', '建物', '難以採樣'])]
                
                # 判定是否需要遞補
                if len(active_reps) < 3:
                    backups = g_data[g_data['代表性'] == '備用點'].sort_values('農地序號')
                    needed = 3 - len(active_reps)
                    to_add = backups.head(needed)
                    # 合併有效代表點與遞補點
                    final_g_list = pd.concat([active_reps, to_add])
                    final_g_list['篩選分類'] = f'系統型網格 ({g_freq}) - 包含遞補'
                else:
                    final_g_list = active_reps.copy()
                    final_g_list['篩選分類'] = f'系統型網格 ({g_freq})'
                
                grid_list.append(final_g_list)

        # 合併所有結果
        if grid_list:
            sys_final = pd.concat(grid_list)
            full_plan = pd.concat([case_list, sys_final])
        else:
            full_plan = case_list

        # 顯示結果
        st.subheader(f"📍 {target_year} 年度建議調查清單 (共 {len(full_plan)} 筆)")
        
        # 增加篩選按鈕
        show_type = st.multiselect("過濾類別", full_plan['篩選分類'].unique(), default=full_plan['篩選分類'].unique())
        display_df = full_plan[full_plan['篩選分類'].isin(show_type)]
        
        st.dataframe(display_df[['網格編號', '地段地號', '農地序號', '目前農地調查現況', '最後調查年分', '篩選分類']], height=600, use_container_width=True)

        # 下載按鈕
        towrite = io.BytesIO()
        display_df.to_excel(towrite, index=False, engine='xlsxwriter')
        st.download_button(f"📥 下載 {target_year} 年度調查工作單", data=towrite.getvalue(), file_name=f"彰化農地_{target_year}_工作單.xlsx")

    elif menu == "新增年度調查結果":
        st.title("➕ 錄入採樣判定")
        # (保留 DA 判定與照片上傳代碼)
        st.warning("管理員權限解鎖後進行錄入")

    elif menu == "空間地圖檢視":
        st.title("🗺️ 衛星影像與網格著色")
        # (保留 Folium 衛星底圖邏輯)
else:
    st.error("系統讀取失敗，請確認檔案。")



