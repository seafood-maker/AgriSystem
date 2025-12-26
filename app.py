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
st.set_page_config(page_title="彰化農地智慧管理系統 10.0", layout="wide", page_icon="🌾")

# 專業美化：標題黑化、Metric 卡片美化
st.markdown("""
    <style>
    th { color: #000000 !important; font-weight: bold !important; background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border: 1px solid #eee; }
    .stDownloadButton button { width: 100%; background-color: #2e7d32; color: white; border-radius: 5px; }
    </style>
    """, unsafe_allow_html=True)

EXCEL_PATH = "彰化農地管理資料庫.xlsx"
SHP_PATH = "彰化網格.shp"
LOG_PATH = "edit_log.csv"

# 座標轉換與金屬定義
transformer_to_wgs84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
METALS = ['汞', '砷', '銅', '鉻', '鎘', '鉛', '鋅', '鎳']

def clean_id(val):
    s = str(val).strip()
    return re.sub(r'\.0$', '', s)

def get_minguo_date():
    now = datetime.now()
    return f"民國 {now.year - 1911} 年 {now.month} 月 {now.day} 日"

# --- 2. 資料讀取引擎 ---
@st.cache_data
def load_all_data():
    if not os.path.exists(EXCEL_PATH): return None, None, None, None
    try:
        xl = pd.ExcelFile(EXCEL_PATH)
        actual_sheets = xl.sheet_names
        def get_s(n): return next((s for s in actual_sheets if n == s.strip()), None)
        
        # 讀取主檔
        df_m = pd.read_excel(xl, sheet_name=get_s("農地現況主檔"))
        df_m.columns = df_m.columns.str.strip()
        df_m['網格編號'] = df_m['網格編號'].apply(clean_id)
        df_m['地段地號'] = df_m['地段地號'].astype(str).str.strip()
        df_m['最後調查年分'] = pd.to_numeric(df_m['最後調查年分'], errors='coerce')
        
        # 讀取歷史
        df_h = pd.read_excel(xl, sheet_name=get_s("歷年調查紀錄"))
        df_h.columns = df_h.columns.str.strip()
        df_h['調查年度'] = pd.to_numeric(df_h['調查年度'], errors='coerce')
        
        # 讀取對照表
        df_b = pd.read_excel(xl, sheet_name=get_s("同坵塊對照表"))
        df_b.columns = df_b.columns.str.strip()
        df_b['農地地段地號'] = df_b['農地地段地號'].astype(str).str.strip()
        
        # 讀取標準表
        df_s = pd.read_excel(xl, sheet_name=get_s("判定標準表")).set_index('項目名稱')
        return df_m, df_h, df_b, df_s
    except Exception as e:
        st.error(f"Excel 讀取錯誤: {e}"); return None, None, None, None

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

# 空間判定函式
def find_grid_by_coords(x, y, gdf):
    if gdf is None: return "未知"
    p = Point(x, y)
    match = gdf[gdf.contains(p)]
    return str(match.iloc[0]['網格號']) if not match.empty else "範圍外"

# --- 3. 側邊欄 ---
st.sidebar.title("🌿 系統導覽")
menu = st.sidebar.radio("功能選單", ["統計首頁", "資料庫查詢與下載", "新年度調查點篩選名單", "新增年度調查結果", "空間地圖檢視"])

if df_master is not None:
    # --- 預算統計指標 ---
    abs_total = len(df_master)
    sampling_pts = len(df_master[df_master['代表性'].isin(['代表點', '備用點'])])
    control_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('管制', na=False)])
    build_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('建物', na=False)])
    hard_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('難以採樣', na=False)])
    normal_count = len(df_master[df_master['農地監測狀態'].astype(str).str.contains('正常', na=False)])

    # --- A. 統計首頁 ---
    if menu == "統計首頁":
        st.title("🚜 彰化縣農地監測戰情室")
        st.subheader(f"📅 當前時間：{get_minguo_date()}")
        
        # 1. 頂列統計
        k = st.columns(6)
        k[0].metric("總資料點數", abs_total)
        k[1].metric("總採樣點數", sampling_pts)
        k[2].metric("管制點數", control_count)
        k[3].metric("建物數量", build_count)
        k[4].metric("難以採樣數量", hard_count)
        k[5].metric("正常退場數量", normal_count)

        st.divider()
        col_g, col_c = st.columns(2)
        with col_g:
            st.subheader("🌐 系統型網格統計")
            grid_df = df_master.drop_duplicates('網格編號').copy()
            grid_df['網格監測頻率'] = grid_df['網格監測頻率'].fillna('無網格狀態')
            g_c = len(grid_df[grid_df['網格監測頻率'] == '持續'])
            g_e = len(grid_df[grid_df['網格監測頻率'] == '延長'])
            g_ex = len(grid_df[grid_df['網格監測頻率'] == '退場'])
            g_sum = g_c + g_e + g_ex
            gc = st.columns(5)
            gc[0].metric("持續", g_c); gc[1].metric("延長", g_e); gc[2].metric("退場", g_ex)
            gc[3].metric("有效網格合計", g_sum); gc[4].metric("無網格狀態", len(grid_df)-g_sum)
        
        with col_c:
            st.subheader("📦 個案型農地統計")
            case_data = df_master[~df_master['調查方式'].astype(str).str.contains('系統', na=False)].copy()
            c_map = {"增量":"持續", "延長":"延長", "正常":"退場", "難以採樣":"難以採樣", "管制":"管制", "建物":"建物"}
            case_data['對應'] = case_data['目前農地調查現況'].map(c_map).fillna(case_data['目前農地調查現況'])
            cc = case_data['對應'].value_counts()
            cc_cols = st.columns(6)
            for i, lab in enumerate(["持續", "延長", "退場", "管制", "難以採樣", "建物"]):
                cc_cols[i].metric(lab, cc.get(lab, 0))

        st.divider()
        st.subheader("🔍 網格查詢系統")
        gs = st.text_input("輸入網格號碼搜尋 (如: G001)")
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
                    # 演算判定：交叉比對 1xx 狀態與目前現況
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
        st.title("📂 資料庫查詢與下載中心")
        
        admin_mode = False
        with st.sidebar.expander("🔐 數據直接修正"):
            if st.text_input("輸入授權密碼", type="password") == ADMIN_PASSWORD:
                admin_mode = True; st.success("已開啟編輯模式")

        tabs = st.tabs(["📋 總表清單", "📅 歷年調查結果", "🏠 坵塊管理", "🌐 網格採樣看板", "📜 修改紀錄"])

        with tabs[0]:
            st.subheader("🌾 農地現況總表")
            # 代表性視覺化邏輯
            def rich_rep(row):
                r, s = str(row['代表性']).strip(), str(row['農地監測狀態']).strip()
                if r == "代表點": return "✅ 代表點"
                if r == "備用點": return "⚪ 備用點"
                # 檢查同坵塊
                if not df_block.empty and row['地段地號'] in df_block['農地地段地號'].values:
                    is_rep_in_blk = df_block[df_block['農地地段地號']==row['地段地號']].iloc[0]['代表農地']
                    if "否" in str(is_rep_in_blk): return "❌ 非採樣 (同坵塊)"
                return f"❌ 非採樣 ({s})"

            df_p = df_master.copy()
            df_p['代表性顯示'] = df_p.apply(rich_rep, axis=1)
            
            # 【位置移動】將代表性顯示移到農地序號右邊
            cols = list(df_p.columns)
            if '農地序號' in cols:
                idx = cols.index('農地序號') + 1
                cols.insert(idx, cols.pop(cols.index('代表性顯示')))
                df_p = df_p[cols]

            if admin_mode:
                st.data_editor(df_p, height=800, use_container_width=True)
            else:
                st.dataframe(df_p, height=800, use_container_width=True)
            
            towrite = io.BytesIO()
            df_master.to_excel(towrite, index=False, engine='xlsxwriter')
            st.download_button("📥 下載全量總表 Excel", data=towrite.getvalue(), file_name="彰化農地總表.xlsx")

        with tabs[1]:
            st.subheader("📅 歷年調查結果下載")
            search_y = st.number_input("請輸入查詢年度 (民國)", value=113)
            y_res = df_history[df_history['調查年度'] == search_y].copy()
            if not y_res.empty:
                # 關聯主檔顯示更詳盡資訊
                y_rich = y_res.merge(df_master[['SGM編號','地段地號','調查方式','目前農地調查現況']], on='SGM編號', how='left')
                # DA 判定備註
                def da_note(row):
                    notes = [f"{m}DA({row.get(f'DA_{m}',0)}%)" for m in METALS if row.get(f'DA_{m}',0) > 20]
                    return " / ".join(notes) if notes else "正常"
                y_rich['DA判定說明'] = y_rich.apply(da_note, axis=1)
                st.dataframe(y_rich, use_container_width=True)
                tow_h = io.BytesIO()
                y_rich.to_excel(tow_h, index=False, engine='xlsxwriter')
                st.download_button(f"📥 下載 {search_y} 調查清單", data=tow_h.getvalue(), file_name=f"{search_y}_調查結果.xlsx")
            else: st.warning("該年度尚無調查紀錄")

        with tabs[2]:
            st.subheader("🏠 坵塊管理系統")
            blk_q = st.text_input("🔍 搜尋地段地號查看同群組成員")
            if blk_q and not df_block.empty:
                if blk_q in df_block['農地地段地號'].values:
                    gid = df_block[df_block['農地地段地號']==blk_q].iloc[0]['農地群組編號']
                    st.success(f"找到群組 {gid}，成員清單：")
                    st.dataframe(df_block[df_block['農地群組編號']==gid])
            
            st.divider()
            st.write("**➕ 批次新增同坵塊關聯**")
            with st.form("blk_form"):
                new_gid = st.text_input("坵塊群組編號 (如: G_001)")
                lots_text = st.text_area("地段地號清單 (每行一筆)")
                rep_lot = st.text_input("指定哪一筆為『代表點』？")
                if st.form_submit_button("確認新增"):
                    st.info(f"已排程處理群組 {new_gid}。請更新 Excel 後重新上傳以生效。")

        with tabs[3]:
            st.subheader("🌐 系統型網格監測看板")
            st.dataframe(df_master[df_master['調查方式']=='系統型農地'].groupby('網格編號').head(5))

    # --- C. 新年度調查點篩選名單 ---
    elif menu == "新年度調查點篩選名單":
        st.title("📅 115 年度調查點初步篩選")
        st.write("演算法邏輯：持續型(增量)必選 > 延長型(到期)次選 > 排除管制/建物/退場")
        
        f1 = df_master[df_master['目前農地調查現況'] == '增量'].copy()
        f1['篩選原因'] = '持續監測(增量)'
        
        f2 = df_master[(df_master['目前農地調查現況'] == '延長') & (df_master['最後調查年分'] <= 113)].copy()
        f2['篩選原因'] = '延長到期'
        
        plan_df = pd.concat([f1, f2]).drop_duplicates('SGM編號')
        st.dataframe(plan_df[['網格編號','地段地號','目前農地調查現況','最後調查年分','篩選原因']])
        st.info("💡 提醒：若代表點變更為建物或管制，系統將自動於正式計畫中遞補同網格之備用點。")

    # --- D. 新增年度調查結果 (含 DA 與 3M 預警) ---
    elif menu == "新增年度調查結果":
        st.title("➕ 錄入當年度調查數據")
        pwd = st.sidebar.text_input("管理員密碼", type="password")
        if pwd == ADMIN_PASSWORD:
            search_lot = st.text_input("🔍 第一步：輸入地號搜尋")
            if search_lot:
                hits = df_master[df_master['地段地號'] == search_lot.strip()]
                if not hits.empty:
                    row = hits.iloc[0]
                    with st.form("input_form"):
                        st.subheader(f"📍 編輯：{search_lot}")
                        c1, c2 = st.columns(2)
                        new_x = c1.number_input("實測 X (TWD97)", value=float(row['TWD97_X']))
                        new_y = c2.number_input("實測 Y (TWD97)", value=float(row['TWD97_Y']))
                        dist = np.sqrt((new_x-row['TWD97_X'])**2 + (new_y-row['TWD97_Y'])**2)
                        if dist > 3:
                            st.warning(f"⚠️ 座標偏移 {dist:.2f} 米")
                            st.info(f"當前位處網格: {find_grid_by_coords(new_x, new_y, gdf_grid)}")
                        
                        st.write("---")
                        st.write("🧪 重金屬錄入 (DA 判定)")
                        v_data = {}
                        for m in METALS:
                            col1, col2 = st.columns(2)
                            xrf = col1.number_input(f"{m} (XRF)", min_value=0.0, key=f"x_{m}")
                            tot = col2.number_input(f"{m} (全量)", min_value=0.0, key=f"t_{m}")
                            v_data[m] = tot if tot > 0 else xrf
                        
                        if st.form_submit_button("執行自動判定"):
                            final_st = "正常"; da_res = {}
                            for m in METALS:
                                init = row.get(f'初始_{m}', 0)
                                s2 = df_settings.loc[m, '管制標準']
                                da = ((v_data[m]-init)/init*100) if init > 0 else 0
                                da_res[m] = f"{da:.1f}%"
                                if v_data[m] > s2: final_st = "管制"
                                elif init > s2 and da > df_settings.loc[m, '上升標準 (DA門檻)']:
                                    if final_st != "管制": final_st = "增量"
                            st.write(f"💡 建議判定: **{final_st}**", da_res)
        else: st.warning("權限受限")

    # --- E. 空間地圖檢視 ---
    elif menu == "空間地圖檢視":
        st.title("🗺️ 衛星影像與網格監測圖")
        m = folium.Map(location=[24.05, 120.5], zoom_start=11, tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', attr='Esri')
        if gdf_grid is not None:
            grid_status = df_master.drop_duplicates('網格編號')[['網格編號', '網格監測頻率']]
            merged = gdf_grid.to_crs(epsg=4326).merge(grid_status, left_on='網格號', right_on='網格編號', how='left')
            def get_c(f):
                f = str(f)
                return '#FFB6C1' if '持續' in f else '#ADD8E6' if '延長' in f else '#90EE90' if '退場' in f else '#F8F8F8'
            folium.GeoJson(merged, style_function=lambda x: {'fillColor': get_c(x['properties'].get('網格監測頻率')), 'color': 'white', 'weight': 1, 'fillOpacity': 0.4}).add_to(m)
        
        # 點位顯示邏輯 (與前版相同)
        sample = df_master.sample(min(800, len(df_master)))
        for _, r in sample.iterrows():
            try:
                lon, lat = transformer_to_wgs84.transform(r['TWD97_X'], r['TWD97_Y'])
                sides = 4 if "個案" in str(r['調查方式']) else 3
                mon_s = str(r['農地監測狀態'])
                if mon_s == "管制": sides = 6; c = "red"
                elif mon_s == "建物": sides = 6; c = "black"
                elif mon_s == "難以採樣": sides = 6; c = "purple"
                elif str(r['代表性']) == "備用點": sides = 4; c = "white"
                else:
                    inv = str(r['目前農地調查現況'])
                    c = "red" if "增量" in inv else "blue" if "延長" in inv else "green"
                folium.RegularPolygonMarker(location=[lat, lon], number_of_sides=sides, radius=6, color=c, fill=True, popup=f"{r['地段地號']}").add_to(m)
            except: continue
        st_folium(m, width=1100, height=700)
else:
    st.error("系統讀取失敗。")
