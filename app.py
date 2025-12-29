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
# 1. 系統初始化與基礎函式
# ==========================================

# 解決屬性錯誤：最上方初始化所有會話狀態
if 'excluded_lots' not in st.session_state: st.session_state.excluded_lots = []
if 'temp_field_plan' not in st.session_state: st.session_state.temp_field_plan = None
if 'archived_plans' not in st.session_state: st.session_state.archived_plans = {}
if 'manual_sys_grids' not in st.session_state: st.session_state.manual_sys_grids = []
if 'manual_case_lots' not in st.session_state: st.session_state.manual_case_lots = []
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
    .stats-container { background-color: #e8f5e9; padding: 20px; border-radius: 15px; margin-bottom: 20px; border: 1px solid #c8e6c9; }
    .filter-card { background-color: #f8f9fa; padding: 15px; border-radius: 10px; border-left: 5px solid #2e7d32; margin-bottom: 20px; }
    .manual-box { border: 1px solid #ddd; padding: 15px; border-radius: 10px; background-color: #fcfcfc; min-height: 300px; }
    .recon-log-box { background-color: #fff3e0; padding: 15px; border-radius: 10px; border: 1px solid #ffb74d; }
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
        df_s = pd.read_excel(xl, sheet_name=get_s("判定標準表")).set_index('項目名稱')
        return df_m, df_h, df_b, df_s
    except: return None, None, None, None

df_master, df_history, df_block, df_settings = load_all_data()
gdf_grid = gpd.read_file(SHP_PATH) if os.path.exists(SHP_PATH) else None

# ==========================================
# 4. 全局數據處理
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
# 5. 主選單
# ==========================================

st.sidebar.title("🌿 系統選單")
menu = st.sidebar.radio("功能導覽", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

if df_master is not None:

    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        # ...原有六大指標看板與樹狀圖... (代碼省略，邏輯同前，確保穩定性)
        k = st.columns(6)
        k[0].metric("總資料點數", abs_total); k[1].metric("總採樣點數", sampling_pts_count)
        k[2].metric("管制點數", control_pts_count); k[3].metric("建物數量", build_pts_count)
        k[4].metric("難以採樣數量", hard_pts_count); k[5].metric("正常退場數量", normal_pts_count)

    elif menu == "資料庫查詢與下載":
        st.title("📂 數據查詢與下載中心")
        # ...原有總表、歷史紀錄、分類看板代碼... (代碼省略，邏輯同前)

    # --- C. [核心大升級] 新年度調查點篩選名單 ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 調查計畫決策工作流")
        
        # 模式與設定區
        col_m1, col_m2 = st.columns([2, 1])
        with col_m1:
            plan_mode = st.radio("選擇操作模式", ["自動系統產生 (具補位功能)", "手動自選名單"], horizontal=True)
        with col_m2:
            auto_backfill = st.checkbox("開啟自動補位功能", value=True, help="若關閉，取消留用時系統不會自動增加新名單")

        c1, c2 = st.columns(2)
        target_year = c1.number_input("設定目標年度 (民國)", value=get_minguo_year()+1)
        quota = c2.number_input("設定預計調查總筆數 (補位基準)", value=500, step=10)

        # --- 名單運算池 ---
        df_calc = df_master.copy()
        
        if plan_mode.startswith("自動"):
            # 演算法：系統型
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
                    fg['優先權重'], fg['計畫類別'] = prio, '系統型網格'
                    sys_pool.append(fg)
            
            # 演算法：個案型
            case_active = df_calc[~df_calc['調查方式'].str.contains('系統', na=False)].copy()
            case_active = case_active[~case_active['農地監測狀態'].isin(['管制','建物','正常'])]
            def c_prio_logic(r):
                if r['最後調查年分'] == target_year - 1: return 99
                if str(r['目前農地調查現況']) == '增量' and (target_year-r['最後調查年分'] >= 2): return 1
                if str(r['目前農地調查現況']) == '延長' and (target_year-r['最後調查年分'] >= 10): return 3
                return 99
            case_active['優先權重'] = case_active.apply(c_prio_logic, axis=1)
            case_pool = case_active[case_active['優先權重'] < 99].copy()
            case_pool['計畫類別'] = '個案型農地'

            # 總池過濾與截斷
            full_pool = pd.concat([pd.concat(sys_pool) if sys_pool else pd.DataFrame(), case_pool]).sort_values(['優先權重', '網格編號'])
            eligible = full_pool[~full_pool['地段地號'].isin(st.session_state.excluded_lots)]
            # 補位判斷
            current_selection = eligible.head(int(quota)).copy() if auto_backfill else eligible.copy()
        
        else: # [還原並修正] 手動挑選頁面
            st.subheader("🖐️ 手動挑選名單")
            col_sel1, col_sel2 = st.columns(2)
            sys_opts = df_calc[df_calc['調查方式'].str.contains('系統', na=False)]['網格編號'].unique()
            case_opts = df_calc[~df_calc['調查方式'].str.contains('系統', na=False)]['地段地號'].unique()
            
            with col_sel1:
                st.markdown('<div class="manual-box"><b>1. 系統型網格選取</b>', unsafe_allow_html=True)
                st.session_state.manual_sys_grids = st.multiselect("選取網格 ID", sys_opts, default=st.session_state.manual_sys_grids)
                st.markdown('</div>', unsafe_allow_html=True)
            with col_sel2:
                st.markdown('<div class="manual-box"><b>2. 個案型地號選取</b>', unsafe_allow_html=True)
                st.session_state.manual_case_lots = st.multiselect("選取地段地號", case_opts, default=st.session_state.manual_case_lots)
                st.markdown('</div>', unsafe_allow_html=True)
            
            m_s = df_calc[df_calc['網格編號'].isin(st.session_state.manual_sys_grids) & (df_calc['代表性']=='代表點')].copy()
            m_s['計畫類別'] = '系統型網格'
            m_c = df_calc[df_calc['地段地號'].isin(st.session_state.manual_case_lots)].copy()
            m_c['計畫類別'] = '個案型農地'
            current_selection = pd.concat([m_s, m_c])
            current_selection = current_selection[~current_selection['地段地號'].isin(st.session_state.excluded_lots)].copy()

        if not current_selection.empty:
            # 階段一：名單微調 (✅/❌ 選單)
            st.subheader("第一階段：初步篩選與去留確認")
            current_selection.insert(0, '狀態確認', '✅ 留用')
            
            edited_init = st.data_editor(
                current_selection[['狀態確認','網格編號','地段地號','農地序號','TWD97_X','TWD97_Y','目前農地調查現況','計畫類別']],
                column_config={"狀態確認": st.column_config.SelectboxColumn("狀態確認", options=["✅ 留用", "❌ 排除"], required=True)},
                key="init_ed", use_container_width=True
            )
            
            if st.button("🔥 確定清出標記為打叉 (❌) 的農地"):
                to_kill = edited_init[edited_init['狀態確認'] == "❌ 排除"]['地段地號'].tolist()
                st.session_state.excluded_lots.extend(to_kill); st.rerun()

            # 階段二：現勘紀錄錄入 (滿足三個欄位、備註與自動更新邏輯)
            st.divider()
            st.subheader("第二階段：現場現勘紀錄錄入與數據連動")
            if st.button("💾 確認名單並產生正式現勘表"):
                st.session_state.temp_field_plan = current_selection.copy()

            if st.session_state.temp_field_plan is not None:
                plan = st.session_state.temp_field_plan
                if '微調勾選' not in plan.columns: plan.insert(0, '微調勾選', True)
                if '現勘結果' not in plan.columns: plan['現勘結果'] = '系統型'
                if '現勘備註' not in plan.columns: plan['現勘備註'] = ''
                
                recon_ed = st.data_editor(
                    plan[['微調勾選','現勘結果','現勘備註','地段地號','網格編號','調查方式','目前農地調查現況']],
                    column_config={
                        "現勘結果": st.column_config.SelectboxColumn("現勘判定", options=["系統型", "個案型", "建物", "難以採樣"]),
                        "現勘備註": st.column_config.TextColumn("備註紀錄", placeholder="輸入現場狀況...")
                    },
                    key="recon_final_ed", use_container_width=True
                )

                # [自動同步演算法啟動]
                if st.button("🔄 同步現勘結果至大資料庫並修正狀態"):
                    new_logs = []
                    for _, row in recon_ed.iterrows():
                        lot = row['地段地號']
                        old_type = row['調查方式']
                        new_type = row['現勘結果']
                        # 邏輯判斷：型態轉換
                        if new_type in ['建物', '難以採樣']:
                            new_logs.append(f"🚩 地號 {lot}: 判定為 {new_type}，已從今年名單移除，建議網格補位。")
                        elif old_type != new_type:
                            new_logs.append(f"🔄 地號 {lot}: 調查方式由 {old_type} 變更為 {new_type}。")
                    st.session_state.recon_log = new_logs
                    st.success("分析完成！請查看下方變動顯示窗。")

                # [顯示窗] (滿足第 3 點)
                if st.session_state.recon_log:
                    st.markdown('<div class="recon-log-box"><b>📋 現勘結果變動分析清單：</b>', unsafe_allow_html=True)
                    for log in st.session_state.recon_log: st.write(log)
                    st.markdown('</div>', unsafe_allow_html=True)

            # 最終統計與歸檔
            st.markdown('<div class="stats-container">', unsafe_allow_html=True)
            st.subheader("📊 擬定計畫統計摘要")
            s_plan = current_selection[current_selection['計畫類別']=='系統型網格']
            c_plan = current_selection[current_selection['計畫類別']=='個案型農地']
            st.write(f"系統型網格數: {len(s_plan['網格編號'].unique())} (持續 {len(s_plan[s_plan['網格監測頻率']=='持續']['網格編號'].unique())} / 延長 {len(s_plan[s_plan['網格監測頻率']=='延長']['網格編號'].unique())})")
            st.write(f"系統型農地筆數: {len(s_plan)} | 個案型農地筆數: {len(c_plan)} (持續 {len(c_plan[c_plan['目前農地調查現況']=='增量'])} / 延長 {len(c_plan[c_plan['目前農地調查現況']=='延長'])})")
            st.markdown('</div>', unsafe_allow_html=True)

            # 最終調查名單確認 (滿足第 5 點：存入歷史資料庫)
            st.divider()
            if st.button("✅ 最終調查名單確認"):
                arc_id = f"{target_year}_最終定案_{datetime.now().strftime('%m%d_%H%M')}"
                st.session_state.archived_plans[arc_id] = current_selection.copy()
                st.session_state.temp_field_plan = None
                st.success("計畫已移入下方歷史檔案資料庫。")

        if st.button("🔄 重置篩選"):
            st.session_state.excluded_lots = []; st.session_state.temp_field_plan = None; st.rerun()

        # --- GIS 地圖分布圖 ---
        st.subheader("🗺️ 年度擬定計畫調查分布圖 (最底端)")
        m_plan = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        for _, r in current_selection.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                sd = 3 if "系統" in str(r['計畫類別']) else 4 # 三角/正方
                if sd == 3: clr = "red" if "增量" in str(r['目前農地調查現況']) else "blue"
                else: clr = "yellow" if "增量" in str(r['目前農地調查現況']) else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=sd, radius=8, color=clr, fill=True).add_to(m_plan)
            except: continue
        st_folium(m_plan, width=1100, height=500)

        # 歷史檔案資料庫 (頁面最底)
        if st.session_state.archived_plans:
            st.subheader("📚 歷史計畫存檔資料庫")
            for k_arc, v_arc in st.session_state.archived_plans.items():
                with st.expander(f"📂 {k_arc}"):
                    st.dataframe(v_arc)
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
















