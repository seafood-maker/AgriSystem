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
if 'recon_log' not in st.session_state: st.session_state.recon_log = []
if 'manual_selection_list' not in st.session_state: st.session_state.manual_selection_list = pd.DataFrame()

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
# 2. 系統設定與專業美化
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
    .stats-container { background-color: #e8f5e9; padding: 15px; border-radius: 10px; margin-bottom: 20px; border: 1px solid #c8e6c9; }
    .manual-tool-box { background-color: #f1f3f6; padding: 20px; border-radius: 15px; border: 1px solid #d1d9e6; margin-top: 10px; }
    .recon-box { border-left: 5px solid #2196f3; background-color: #e3f2fd; padding: 15px; border-radius: 5px; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 3. 資料讀取
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

df_master, df_history, df_block, df_settings = load_all_data()
gdf_grid = gpd.read_file(SHP_PATH) if os.path.exists(SHP_PATH) else None

# ==========================================
# 4. 側邊欄與導覽
# ==========================================

st.sidebar.title("🌿 系統導覽")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

if df_master is not None:

    # --- A. 統計首頁 (完整還原版) ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        # 計算統計數據
        abs_total = len(df_master)
        sampling_pts = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
        control_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
        build_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
        hard_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
        normal_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])
        
        k = st.columns(6)
        k[0].metric("總資料點數", abs_total); k[1].metric("總採樣點數", sampling_pts)
        k[2].metric("管制點數", control_count); k[3].metric("建物數量", build_count)
        k[4].metric("難以採樣數量", hard_count); k[5].metric("正常退場數量", normal_count)
        st.divider()
        st.subheader("📊 近三年調查分佈 (樹狀圖)")
        tree_cols = st.columns(3); curr_y = get_minguo_year()
        for i, y in enumerate([curr_y-2, curr_y-1, curr_y]):
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
        # (此處保留原先方塊選擇與下載代碼...)
        st.info("請點選下方分頁檢視資料。")
        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 系統型農地清單", "📦 個案型農地清單", "📜 修改紀錄"])

    # --- C. [核心大改版] 新年度調查點篩選名單 ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 年度調查計畫規劃工作流")
        
        # 1. 參數與演算法設定
        target_year = st.number_input("設定目標年度 (民國)", value=get_minguo_year()+1)
        quota = st.number_input("設定本年度預算調查總筆數 (補位基準)", value=500, step=10)

        # ---------------------------------------------------------
        # 第一階段：初步篩選與去留確認
        # ---------------------------------------------------------
        st.subheader("第一階段：初步篩選與去留確認")
        
        df_calc = df_master.copy()
        # (A) 系統型自動篩選
        sys_pool = []
        grids = df_calc[df_calc['調查方式'].str.contains('系統', na=False)]['網格編號'].unique()
        for gid in grids:
            g_data = df_calc[df_calc['網格編號'] == gid]
            f_t, ly = str(g_data['網格監測頻率'].iloc[0]), g_data['最後調查年分'].max()
            if ly == target_year - 1: continue # 排除去年
            prio = 1 if f_t == '持續' and (target_year-ly >= 2) else (2 if f_t == '延長' and (target_year-ly >= 10) else 99)
            if prio < 99:
                reps = g_data[(g_data['代表性'] == '代表點') & (~g_data['農地監測狀態'].isin(['管制','建物','難以採樣']))]
                if len(reps) < 3:
                    backs = g_data[g_data['代表性'] == '備用點'].sort_values('農地序號').head(3 - len(reps)).copy()
                    fg = pd.concat([reps, backs])
                else: fg = reps.copy()
                fg['優先權重'], fg['計畫類別'], fg['網格頻率'] = prio, '系統型網格', f_t
                sys_pool.append(fg)
        
        # (B) 個案型自動篩選
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
        case_pool['計畫類別'], case_pool['篩選備註'] = '個案型農地', '獨立判定'

        full_pool = pd.concat([pd.concat(sys_pool) if sys_pool else pd.DataFrame(), case_pool]).sort_values(['優先權重', '網格編號'])
        eligible = full_pool[~full_pool['地段地號'].isin(st.session_state.excluded_lots)]
        current_selection = eligible.head(int(quota)).copy()
        current_selection.insert(0, '狀態確認', '✅ 留用')

        # 第一階段統計
        st.markdown('<div class="stats-container">', unsafe_allow_html=True)
        s_pts = current_selection[current_selection['計畫類別']=='系統型網格']
        c_pts = current_selection[current_selection['計畫類別']=='個案型農地']
        k_s1, k_s2, k_s3 = st.columns(3)
        k_s1.metric("初步-系統型農地", len(s_pts))
        k_s2.metric("初步-個案型農地", len(c_pts))
        k_s3.metric("初步-總筆數", len(current_selection))
        st.write(f"🔹 持續網格: {len(s_pts[s_pts['網格頻率']=='持續']['網格編號'].unique())} / 延長網格: {len(s_pts[s_pts['網格頻率']=='延長']['網格編號'].unique())}")
        st.markdown('</div>', unsafe_allow_html=True)

        col_p1_l, col_p1_r = st.columns(2)
        with col_p1_l:
            st.write("**🌐 系統型名單**")
            ed_s = st.data_editor(s_pts[['狀態確認','網格編號','地段地號','目前農地調查現況']], key="ed_s_p1", column_config={"狀態確認": st.column_config.SelectboxColumn("狀態確認", options=["✅ 留用", "❌ 排除"])})
        with col_p1_r:
            st.write("**📦 個案型名單**")
            ed_c = st.data_editor(c_pts[['狀態確認','地段地號','農地序號','目前農地調查現況']], key="ed_c_p1", column_config={"狀態確認": st.column_config.SelectboxColumn("狀態確認", options=["✅ 留用", "❌ 排除"])})

        if st.button("🔥 執行清出並自動補足差額"):
            ex1 = ed_s[ed_s['狀態確認'] == "❌ 排除"]['地段地號'].tolist()
            ex2 = ed_c[ed_c['狀態確認'] == "❌ 排除"]['地段地號'].tolist()
            st.session_state.excluded_lots.extend(ex1 + ex2); st.rerun()

        # ---------------------------------------------------------
        # [新增] 手動挑選增補頁面
        # ---------------------------------------------------------
        st.divider()
        with st.expander("🖐️ 手動挑選名單工具 (搜尋並納入現勘)"):
            st.markdown('<div class="manual-tool-box">', unsafe_allow_html=True)
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                st.write("**1. 搜尋系統型網格**")
                m_sys_ids = st.multiselect("選取欲新增網格 ID", df_master['網格編號'].unique())
            with col_m2:
                st.write("**2. 搜尋個案型農地**")
                m_case_lots = st.multiselect("選取欲新增地段地號", df_master[~df_master['調查方式'].str.contains('系統', na=False)]['地段地號'].unique())
            
            if st.button("➕ 將選取項納入現勘名單"):
                m_s_df = df_master[df_master['網格編號'].isin(m_sys_ids) & (df_master['代表性']=='代表點')].copy()
                m_c_df = df_master[df_master['地段地號'].isin(m_case_lots)].copy()
                st.session_state.manual_selection_list = pd.concat([m_s_df, m_c_df])
                st.success(f"已成功選取 {len(st.session_state.manual_selection_list)} 筆農地。請點擊下方『同步至現勘階段』。")
            st.markdown('</div>', unsafe_allow_html=True)

        # ---------------------------------------------------------
        # 第二階段：正式名單與現勘結果錄入
        # ---------------------------------------------------------
        st.divider()
        st.subheader("第二階段：正式名單與現勘結果錄入")
        
        # 進入現勘的條件：初步名單完成 或 手動名單加入
        if st.button("💾 同步名單至現勘階段 (含自動與手動)"):
            combined = pd.concat([current_selection, st.session_state.manual_selection_list]).drop_duplicates('地段地號')
            st.session_state.temp_field_plan = combined.copy()

        if st.session_state.temp_field_plan is not None:
            plan_df = st.session_state.temp_field_plan
            if '去留' not in plan_df.columns: plan_df.insert(0, '去留', '✅ 留用')
            if '現勘判定' not in plan_df.columns: plan_df['現勘判定'] = '系統型'
            if '現勘備註' not in plan_df.columns: plan_df['現勘備註'] = ''

            # 第二階段統計數據 (滿足需求)
            st.markdown('<div class="stats-container">', unsafe_allow_html=True)
            k_f1, k_f2 = st.columns(2)
            k_f1.metric("目前納入現勘總數", len(plan_df))
            
            # 使用目前的 editor 狀態來預判最終名單
            final_sampling = plan_df[(plan_df['現勘判定'].isin(['系統型','個案型']))]
            k_f2.metric("預計最終採樣筆數", len(final_sampling))
            st.markdown('</div>', unsafe_allow_html=True)

            recon_ed = st.data_editor(
                plan_df[['去留','現勘判定','現勘備註','地段地號','網格編號','目前農地調查現況']],
                column_config={
                    "去留": st.column_config.SelectboxColumn("狀態確認", options=["✅ 留用", "❌ 排除"]),
                    "現勘判定": st.column_config.SelectboxColumn("現勘判定", options=["系統型", "個案型", "建物", "難以採樣"])
                },
                key="recon_flow_ed", use_container_width=True
            )

            # 分表顯示
            res_samp = recon_ed[(recon_ed['現勘判定'].isin(['系統型', '個案型'])) & (recon_ed['去留']=='✅ 留用')]
            res_fail = recon_ed[recon_ed['現勘判定'].isin(['建物', '難以採樣'])]

            c_r1, c_r2 = st.columns(2)
            with c_r1:
                st.markdown('<div class="recon-valid"><b>✅ 最終採樣清單</b></div>', unsafe_allow_html=True)
                st.dataframe(res_samp, height=300)
            with c_r2:
                st.markdown('<div class="recon-invalid"><b>❌ 失效點位 (建物/難以採樣)</b></div>', unsafe_allow_html=True)
                st.dataframe(res_fail, height=300)
                if st.button("✅ 確認更新失效點到總表"):
                    st.session_state.excluded_lots.extend(res_fail['地段地號'].tolist())
                    st.success("已標記失效，將於第一階段重新補位。")

            # [顯示窗] 變動分析
            if st.button("🔄 分析現勘後對大資料庫的影響"):
                changes = []
                for _, r in res_fail.iterrows(): changes.append(f"🚩 網格 {r['網格編號']}: 代表點 {r['地段地號']} 已失效，需遞補備用點。")
                st.session_state.recon_log = changes
            
            if st.session_state.recon_log:
                st.info("📋 網格與狀態異動窗：")
                for log in st.session_state.recon_log: st.write(log)

        # 歸檔與存檔
        st.divider()
        if st.button("✅ 最終調查名單確認 (存入歷史檔案庫)"):
            if st.session_state.temp_field_plan is not None:
                st.session_state.archived_plans[f"{target_year}_最終定案"] = res_samp
                st.session_state.temp_field_plan = None; st.rerun()

        # --- GIS 地圖 ---
        st.subheader("🗺️ 年度擬定計畫調查分布圖 (最底端)")
        m_plan = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        for _, r in current_selection.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                sd = 3 if "系統" in str(r['計畫類別']) else 4
                if sd == 3: clr = "red" if "增量" in str(r['目前農地調查現況']) else "blue"
                else: clr = "yellow" if "增量" in str(r['目前農地調查現況']) else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=sd, radius=8, color=clr, fill=True).add_to(m_plan)
            except: continue
        st_folium(m_plan, width=1100, height=500)
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



















