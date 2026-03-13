import streamlit as st
import pandas as pd
import numpy as np
import datetime
import re
import os
import io
import json

# ==========================================
# 页面基础配置 (必须放在最前面)
# ==========================================
st.set_page_config(page_title="天猫电商数据中台", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

HISTORY_FILE = 'dashboard_history.csv'
CONFIG_FILE = 'app_config.json'

# ==========================================
# 配置记忆系统 (读取与保存)
# ==========================================
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except: pass
    return {}

def save_config(config_data):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config_data, f, ensure_ascii=False, indent=4)

config = load_config()

# ==========================================
# 核心读取与清洗函数 (公用)
# ==========================================
def parse_money(s):
    if pd.isna(s) or s == '': return 0.0
    try: return float(str(s).replace('¥','').replace(',','').replace('%','').replace('元','').strip())
    except: return 0.0

def clean_id(raw_id):
    if pd.isna(raw_id): return ""
    return str(raw_id).strip().replace(' ', '').split('.')[0]

def get_category_by_mapping(item_id, title, mapping_dict):
    c_id = clean_id(item_id)
    if mapping_dict and c_id in mapping_dict:
        cat_cn = str(mapping_dict[c_id]).strip()
        if '灯' in cat_cn: return 'Lighting'
        if '家具' in cat_cn: return 'Furniture'
        return 'ACC' 
    t = str(title).lower()
    if any(x in t for x in['灯', 'light', 'lamp', 'portable', 'shade', 'pendant', 'bulb', 'sconce', 'apex']): return 'Lighting'
    if any(x in t for x in['沙发', '柜', '床', 'sofa', 'cabinet', 'cabine', 'bed', 'desk', 'chair', '椅']): return 'Furniture'
    return 'ACC'

@st.cache_data
def read_excel_smart(file_obj, keyword):
    if file_obj is None: return None
    file_obj.seek(0)
    content = file_obj.getvalue()
    try: text = content.decode('utf-8')
    except: text = content.decode('gbk', errors='ignore')
    if '<html' in text.lower() or '<table' in text.lower():
        try:
            dfs = pd.read_html(io.StringIO(text))
            raw_df = dfs[0]
            if keyword in str(raw_df.columns): df = raw_df
            else:
                header_idx = -1
                for idx, row in raw_df.iterrows():
                    if keyword in str(row.values):
                        header_idx = idx
                        break
                if header_idx == -1: return None
                df = raw_df.iloc[header_idx+1:].copy()
                df.columns = raw_df.iloc[header_idx]
            df.columns =[str(c).strip() for c in df.columns]
            return df
        except: pass
    file_obj.seek(0)
    ext = os.path.splitext(file_obj.name)[-1].lower()
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'
    try:
        raw_df = pd.read_excel(file_obj, header=None, nrows=20, engine=engine)
        header_idx = -1
        for idx, row in raw_df.iterrows():
            if keyword in str(row.values):
                header_idx = idx
                break
        if header_idx == -1: return None
        file_obj.seek(0)
        df = pd.read_excel(file_obj, header=header_idx, engine=engine)
        df.columns =[str(c).strip() for c in df.columns]
        return df
    except: return None

def extract_date_from_excel(file_obj):
    if file_obj is None: return None
    file_obj.seek(0)
    content = file_obj.getvalue()
    try: text = content.decode('utf-8')
    except: text = content.decode('gbk', errors='ignore')
    if '<html' in text.lower() or '<table' in text.lower():
        try:
            dfs = pd.read_html(io.StringIO(text))
            raw_df = dfs[0]
            match = re.search(r'(202\d-\d{2}-\d{2})', str(raw_df.columns))
            if match: return match.group(1)
            for _, row in raw_df.head(10).iterrows():
                row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                match = re.search(r'(202\d-\d{2}-\d{2})', row_text)
                if match: return match.group(1)
        except: pass
    else:
        try:
            file_obj.seek(0)
            ext = os.path.splitext(file_obj.name)[-1].lower()
            engine = 'xlrd' if ext == '.xls' else 'openpyxl'
            raw_df = pd.read_excel(file_obj, header=None, nrows=10, engine=engine)
            for _, row in raw_df.iterrows():
                row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                match = re.search(r'(202\d-\d{2}-\d{2})', row_text)
                if match: return match.group(1)
        except: pass
    return None

def get_col_val(df, keywords):
    if df is None or df.empty: return 0.0
    for col in df.columns:
        col_str = str(col).replace(' ', '').replace('\n', '').lower()
        for kw in keywords:
            if kw.lower() in col_str:
                return parse_money(df[col].iloc[0])
    return 0.0

def calc_achi(actual, target):
    if pd.isna(target) or target is None or target == 0 or pd.isna(actual) or actual is None:
        return None
    return actual / target

# ==========================================
# 侧边栏：顶级导航菜单
# ==========================================
st.sidebar.title("🧭 导航菜单")
app_mode = st.sidebar.radio("请选择功能模块：", ["🌞 每日看板 (Daily Dashboard)", "📅 月度排行 (HAY Ranking)"])
st.sidebar.markdown("---")

# =======================================================================================================
# ========================================== 模块一：每日看板 ==========================================
# =======================================================================================================
if app_mode == "🌞 每日看板 (Daily Dashboard)":
    
    st.title("📊 Tmall Store Daily Dashboard")
    st.markdown("💡 **绝对严谨版**：严格提取表格数据，拒绝任何瞎推算，缺失数据优雅画横杠(-)。")

    # --- 侧边栏：业务参数配置中心 ---
    with st.sidebar:
        st.header("⚙️ 业务参数配置中心")
        st.info("💡 填好后点击最下方【保存】，下次打开自动恢复！")
        
        st.subheader("1. 月度目标 (Monthly Target)")
        tgt_gmv_month = st.number_input("全月 GMV 目标", value=config.get('tgt_gmv_month', 2492647.0), step=1000.0)
        tgt_net_month = st.number_input("全月 Net 目标", value=config.get('tgt_net_month', 1500000.0), step=1000.0)
        
        st.subheader("2. 阶段目标 (MTD Target)")
        tgt_gmv_mtd = st.number_input("MTD GMV 目标", value=config.get('tgt_gmv_mtd', 1160000.0), step=1000.0)
        tgt_net_mtd = st.number_input("MTD Net 目标", value=config.get('tgt_net_mtd', 700000.0), step=1000.0)
        
        st.subheader("3. 业务预估 (Estimated Rest)")
        est_rest_gmv = st.number_input("预估剩余 GMV", value=config.get('est_rest_gmv', 1640000.0), step=1000.0)
        est_rest_net = st.number_input("预估剩余 Net", value=config.get('est_rest_net', 940000.0), step=1000.0)
        
        st.subheader("4. 历史对比基数 (LY Actual)")
        ly_whole_gmv = st.number_input("去年全月 GMV", value=config.get('ly_whole_gmv', 2243769.0), step=1000.0)
        ly_whole_net = st.number_input("去年全月 Net", value=config.get('ly_whole_net', 1349223.0), step=1000.0)

        st.subheader("5. 每日需更新的累计数据")
        mtd_refund_actual = st.number_input("🔙 本月累计退款 (必填!)", value=config.get('mtd_refund_actual', 369286.84), step=1000.0)
        manual_mtd_gmv = st.number_input("📌 本月累计GMV (防单品误差)", value=config.get('manual_mtd_gmv', 1061347.0), step=1000.0)

        if st.button("💾 保存以上配置", type="primary", use_container_width=True):
            new_config = {
                'tgt_gmv_month': tgt_gmv_month, 'tgt_net_month': tgt_net_month,
                'tgt_gmv_mtd': tgt_gmv_mtd, 'tgt_net_mtd': tgt_net_mtd,
                'est_rest_gmv': est_rest_gmv, 'est_rest_net': est_rest_net,
                'ly_whole_gmv': ly_whole_gmv, 'ly_whole_net': ly_whole_net,
                'mtd_refund_actual': mtd_refund_actual, 'manual_mtd_gmv': manual_mtd_gmv
            }
            save_config(new_config)
            st.success("配置已永久保存！")

    st.markdown("### 🗂️ 请上传数据源")
    col1, col2, col3 = st.columns(3)
    with col1: 
        file_order = st.file_uploader("1. 📥 今日订单底表", type=['xlsx', 'xls', 'csv'], key="d1")
        file_ly = st.file_uploader("4. 🕰️ 去年当日单品", type=['xlsx', 'xls', 'csv'], key="d4")
    with col2: 
        file_item = st.file_uploader("2. 📥 今日单品生参", type=['xlsx', 'xls', 'csv'], key="d2")
        file_mapping = st.file_uploader("5. 🏷️ 商品分类映射表", type=['xlsx', 'xls', 'csv'], key="d5")
    with col3: 
        file_store = st.file_uploader("3. 📥 店铺大盘表", type=['xlsx', 'xls', 'csv'], key="d3")
        file_history = st.file_uploader("6. 💾 历史记录表(选传)", type=['csv'], key="d6")

    st.divider()

    if st.button("⚡ 严谨生成日报", type="primary", use_container_width=True):
        if not file_order and not file_item and not file_store:
            st.warning("⚠️ 至少需要上传前 3 个文件中的一个！")
            st.stop()

        with st.spinner("🔄 AI 正在严格校验数据，应用商品主数据映射..."):
            try:
                auto_date_str = extract_date_from_excel(file_store)
                if not auto_date_str: auto_date_str = extract_date_from_excel(file_item)
                if not auto_date_str: auto_date_str = extract_date_from_excel(file_order)

                if auto_date_str:
                    d = datetime.datetime.strptime(auto_date_str, '%Y-%m-%d')
                else:
                    d = datetime.date.today() - datetime.timedelta(days=1)

                DATE_STR = f"{d.month}/{d.day}"
                DAYS_PASSED = d.day
                MONTH_NAME = d.strftime('%b') 
                WEEK_NUM = d.isocalendar()[1] 
                ly_year_str = str(d.year - 1)[-2:]

                id_to_cat = {}
                if file_mapping:
                    mapping_df = read_excel_smart(file_mapping, '一级')
                    if mapping_df is None: 
                        mapping_df = read_excel_smart(file_mapping, '商品ID')
                    if mapping_df is not None and '商品ID' in mapping_df.columns and '一级' in mapping_df.columns:
                        mapping_df['Clean_ID'] = mapping_df['商品ID'].apply(clean_id)
                        id_to_cat = dict(zip(mapping_df['Clean_ID'], mapping_df['一级']))

                orders = read_excel_smart(file_order, '商品标题')
                if orders is None: orders = pd.DataFrame()
                for c in['商品价格', '购买数量', '买家应付货款', '买家实付金额', '商品ID', '商家编码', '商品属性', '商品标题']:
                    if c not in orders.columns: orders[c] = 0.0 if '金额' in c or '数量' in c else ''

                sycm_item = read_excel_smart(file_item, '支付金额')
                if sycm_item is None: sycm_item = pd.DataFrame()
                for c in['商品名称', '支付金额', '支付件数', '月累计支付金额', '月累计支付件数', '商品ID']:
                    if c not in sycm_item.columns: sycm_item[c] = 0.0 if '金额' in c or '件数' in c else ''

                sycm_store = read_excel_smart(file_store, '支付金额')
                ly_sycm = read_excel_smart(file_ly, '支付金额') if file_ly else None

                store_gmv = get_col_val(sycm_store,['支付金额', '成交额'])
                store_demand = get_col_val(sycm_store,['下单金额', 'Gross Sales Demand']) 
                store_refund = get_col_val(sycm_store,['成功退款金额', '退款金额（完结时间）'])
                store_traffic = get_col_val(sycm_store,['访客数'])
                store_buyers = get_col_val(sycm_store,['支付买家数'])
                store_cr = get_col_val(sycm_store,['支付转化率'])
                if store_cr > 1: store_cr /= 100 
                new_followers = get_col_val(sycm_store,['新增粉丝数', '关注店铺人数'])
                acc_followers = get_col_val(sycm_store,['累计粉丝数', '总粉丝数'])

                LY_STORE_GMV = get_col_val(sycm_store,['去年今日GMV', '去年当日GMV'])
                LY_STORE_TRAFFIC = get_col_val(sycm_store,['去年今日访客', '去年当日访客'])
                LY_STORE_BUYERS = get_col_val(sycm_store,['去年今日买家', '去年当日买家'])
                LY_STORE_UNITS = get_col_val(sycm_store,['去年今日件数', '去年当日件数'])
                LY_STORE_REFUND = get_col_val(sycm_store,['去年今日退款', '去年当日退款'])
                LY_STORE_DEMAND = get_col_val(sycm_store,['去年今日下单金额'])

                TARGET_GMV_MONTH = tgt_gmv_month
                TARGET_NET_MONTH = tgt_net_month
                TARGET_GMV_MTD = tgt_gmv_mtd
                TARGET_NET_MTD = tgt_net_mtd
                EST_REST_MONTH_GMV = est_rest_gmv
                EST_REST_MONTH_NET = est_rest_net
                LY_WHOLE_MONTH_ACTUAL_GMV = ly_whole_gmv
                LY_WHOLE_MONTH_ACTUAL_NET = ly_whole_net

                if ly_sycm is not None and not ly_sycm.empty:
                    ly_sycm['商品ID'] = ly_sycm.get('商品ID', pd.Series([])).astype(str).str.replace(r'\.0$', '', regex=True)
                    ly_sycm['Category'] = ly_sycm.apply(lambda x: get_category_by_mapping(x.get('商品ID'), x.get('商品名称'), id_to_cat), axis=1)
                    ly_item_gmv = ly_sycm['支付金额'].apply(parse_money).sum()
                    ly_item_units = ly_sycm['支付件数'].apply(parse_money).sum()
                    ly_cat = ly_sycm.groupby('Category')['支付金额'].apply(lambda x: x.apply(parse_money).sum())
                    ly_acc = ly_cat.get('ACC', 0)
                    ly_furn = ly_cat.get('Furniture', 0)
                    ly_light = ly_cat.get('Lighting', 0)
                else:
                    ly_item_gmv = ly_item_units = ly_acc = ly_furn = ly_light = 0

                mtd_gmv_item = sycm_item['月累计支付金额'].apply(parse_money).sum() if not sycm_item.empty else 0
                mtd_gmv = manual_mtd_gmv if manual_mtd_gmv > 0 else mtd_gmv_item
                mtd_net_sales = (mtd_gmv - mtd_refund_actual) / 1.13 if mtd_refund_actual > 0 else mtd_gmv / 1.13

                est_whole_month_gmv = mtd_gmv + EST_REST_MONTH_GMV if EST_REST_MONTH_GMV else None
                est_whole_month_net = mtd_net_sales + EST_REST_MONTH_NET if EST_REST_MONTH_NET else None

                yoy_est_gmv = None
                if est_whole_month_gmv and LY_WHOLE_MONTH_ACTUAL_GMV:
                    yoy_est_gmv = (est_whole_month_gmv - LY_WHOLE_MONTH_ACTUAL_GMV) / LY_WHOLE_MONTH_ACTUAL_GMV
                yoy_est_net = None
                if est_whole_month_net and LY_WHOLE_MONTH_ACTUAL_NET:
                    yoy_est_net = (est_whole_month_net - LY_WHOLE_MONTH_ACTUAL_NET) / LY_WHOLE_MONTH_ACTUAL_NET

                df_p1 = pd.DataFrame({
                    f'Updated {DATE_STR}':['GMV', 'Net sales(Tax excluded )'],
                    'MTD Actual':[mtd_gmv, mtd_net_sales],
                    'MTD Target':[TARGET_GMV_MTD, TARGET_NET_MTD],
                    'MTD Achi%':[calc_achi(mtd_gmv, TARGET_GMV_MTD), calc_achi(mtd_net_sales, TARGET_NET_MTD)],
                    'estimated sales of the rest of Month':[EST_REST_MONTH_GMV if EST_REST_MONTH_GMV>0 else None, EST_REST_MONTH_NET if EST_REST_MONTH_NET>0 else None],
                    'estimated sales of the whole month':[est_whole_month_gmv, est_whole_month_net],
                    f'Monthly target {MONTH_NAME}':[TARGET_GMV_MONTH, TARGET_NET_MONTH],
                    'estimated Achi% of the whole month':[calc_achi(est_whole_month_gmv, TARGET_GMV_MONTH), calc_achi(est_whole_month_net, TARGET_NET_MONTH)],
                    f'estimated whole month sales vs Y{ly_year_str} {MONTH_NAME}':[yoy_est_gmv, yoy_est_net] 
                })

                sycm_item['商品ID'] = sycm_item.get('商品ID', pd.Series([])).astype(str).str.replace(r'\.0$', '', regex=True)
                sycm_item['Category'] = sycm_item.apply(lambda x: get_category_by_mapping(x.get('商品ID'), x.get('商品名称'), id_to_cat), axis=1)
                sycm_item['月累计金额'] = sycm_item['月累计支付金额'].apply(parse_money)
                sycm_item['月累计件数'] = sycm_item['月累计支付件数'].apply(parse_money)
                sycm_item['当日金额'] = sycm_item['支付金额'].apply(parse_money)

                cat_group = sycm_item.groupby('Category').agg(
                    Daily_GMV=('当日金额', 'sum'),
                    MTD_Actual=('月累计金额', 'sum'),
                    MTD_Units=('月累计件数', 'sum')
                ).reset_index()
                
                TARGET_LIGHTING_MTD = TARGET_GMV_MTD * 0.10 if TARGET_GMV_MTD else None
                TARGET_FURNITURE_MTD = TARGET_GMV_MTD * 0.35 if TARGET_GMV_MTD else None
                TARGET_ACC_MTD = TARGET_GMV_MTD * 0.55 if TARGET_GMV_MTD else None
                target_dict = {'Lighting': TARGET_LIGHTING_MTD, 'Furniture': TARGET_FURNITURE_MTD, 'ACC': TARGET_ACC_MTD}
                
                if cat_group.empty:
                    cat_group = pd.DataFrame({'Category':['Lighting', 'Furniture', 'ACC'], 'Daily_GMV':[None,None,None], 'MTD_Actual':[None,None,None], 'MTD_Units':[None,None,None]})
                
                cat_group['MTD Target'] = cat_group['Category'].map(target_dict)
                cat_group['MTD Achi%'] = cat_group.apply(lambda r: calc_achi(r['MTD_Actual'], r['MTD Target']), axis=1)
                cat_group['MTD contribution%'] = cat_group['MTD_Actual'] / mtd_gmv if mtd_gmv else None
                cat_group['MTD AUV'] = cat_group.apply(lambda r: r['MTD_Actual']/r['MTD_Units'] if r['MTD_Units'] and r['MTD_Units']>0 else None, axis=1)

                df_p2 = cat_group[['Category', 'Daily_GMV', 'MTD_Actual', 'MTD Target', 'MTD Achi%', 'MTD contribution%', 'MTD AUV', 'MTD_Units']].copy()
                df_p2['Category'] = pd.Categorical(df_p2['Category'], categories=['Lighting', 'Furniture', 'ACC'], ordered=True)
                df_p2 = df_p2.sort_values('Category')
                df_p2.rename(columns={'Daily_GMV': f'GMV {DATE_STR}'}, inplace=True)

                total_row = pd.DataFrame([{
                    'Category': 'Total',
                    f'GMV {DATE_STR}': store_gmv if store_gmv > 0 else df_p2[f'GMV {DATE_STR}'].sum(),
                    'MTD_Actual': mtd_gmv,
                    'MTD Target': df_p2['MTD Target'].sum() if not df_p2['MTD Target'].isna().all() else None,
                    'MTD Achi%': calc_achi(mtd_gmv, df_p2['MTD Target'].sum()) if not df_p2['MTD Target'].isna().all() else None,
                    'MTD contribution%': 1.0 if mtd_gmv > 0 else None,
                    'MTD AUV': mtd_gmv / df_p2['MTD_Units'].sum() if df_p2['MTD_Units'].sum() else None,
                    'MTD_Units': df_p2['MTD_Units'].sum()
                }])
                df_p2 = pd.concat([df_p2, total_row], ignore_index=True)

                df_p3 = pd.DataFrame({
                    f'Updated {DATE_STR}':['New Followers', 'Accumulated Followers'],
                    'Data':[new_followers if new_followers else None, acc_followers if acc_followers else None]
                })

                orders['购买数量'] = orders['购买数量'].apply(parse_money)
                orders['买家实付金额'] = orders['买家实付金额'].apply(parse_money)
                total_units = orders['购买数量'].sum()
                
                gross_demand = store_demand if store_demand > 0 else orders['买家实付金额'].sum()
                today_gmv_fallback = store_gmv if store_gmv > 0 else sycm_item['当日金额'].sum()
                today_auv = today_gmv_fallback / total_units if total_units else 0

                net_sales_tax_inc = today_gmv_fallback - store_refund
                net_sales_tax_excl = net_sales_tax_inc / 1.13

                today_record = {
                    'Date': DATE_STR,
                    'Traffic': store_traffic if store_traffic else None,
                    'CR%': store_cr if store_cr else None,
                    'Buyers': store_buyers if store_buyers else None,
                    'ATV 客单价': today_gmv_fallback / store_buyers if store_buyers else None,
                    'UPT 客单件': total_units / store_buyers if store_buyers else None,
                    'AUV 件单价': today_auv,
                    'Units Sold 件数': total_units,
                    'Gross Sales Demand 下单金额': gross_demand,
                    'GMV （ 成交额）': today_gmv_fallback,
                    'ACC': df_p2.loc[df_p2['Category']=='ACC', f'GMV {DATE_STR}'].sum(),
                    'Furniture': df_p2.loc[df_p2['Category']=='Furniture', f'GMV {DATE_STR}'].sum(),
                    'Lighting': df_p2.loc[df_p2['Category']=='Lighting', f'GMV {DATE_STR}'].sum(),
                    'Returns  退款': store_refund if store_refund else None,
                    'Net sales（含税）': net_sales_tax_inc,
                    'Net sales（去税）': net_sales_tax_excl
                }
                df_today = pd.DataFrame([today_record])

                if file_history is not None:
                    history_df = pd.read_csv(file_history)
                elif os.path.exists(HISTORY_FILE):
                    history_df = pd.read_csv(HISTORY_FILE)
                else:
                    history_df = pd.DataFrame()

                if not history_df.empty:
                    history_df = history_df[history_df['Date'] != DATE_STR] 
                    history_df = pd.concat([history_df, df_today], ignore_index=True)
                else:
                    history_df = df_today
                    
                try:
                    history_df['SortDate'] = pd.to_datetime(history_df['Date'] + f'/{d.year}', format='%m/%d/%Y')
                    history_df = history_df.sort_values('SortDate').drop('SortDate', axis=1)
                except: pass
                
                history_df.to_csv(HISTORY_FILE, index=False)

                recent_7 = history_df.tail(7).copy()
                wtd_row = recent_7.sum(numeric_only=True)
                wtd_row['CR%'] = wtd_row['Buyers'] / wtd_row['Traffic'] if wtd_row.get('Traffic', 0) else None
                wtd_row['ATV 客单价'] = wtd_row['GMV （ 成交额）'] / wtd_row['Buyers'] if wtd_row.get('Buyers', 0) else None
                wtd_row['UPT 客单件'] = wtd_row['Units Sold 件数'] / wtd_row['Buyers'] if wtd_row.get('Buyers', 0) else None
                wtd_row['AUV 件单价'] = wtd_row['GMV （ 成交额）'] / wtd_row['Units Sold 件数'] if wtd_row.get('Units Sold 件数', 0) else None

                recent_7.loc['WTD'] = wtd_row
                recent_7.at['WTD', 'Date'] = 'Week to Date'

                def get_yoy(today_val, ly_val):
                    if ly_val and ly_val > 0 and pd.notna(today_val): 
                        return (today_val - ly_val) / ly_val
                    return None

                ly_row = pd.Series(dtype=object)
                ly_row['Date'] = 'Today vs LY'
                
                ly_gmv_base = LY_STORE_GMV if LY_STORE_GMV > 0 else ly_item_gmv
                today_gmv_base = today_gmv_fallback if LY_STORE_GMV > 0 else (sycm_item['当日金额'].sum() if not sycm_item.empty else 0)
                
                ly_units_base = LY_STORE_UNITS if LY_STORE_UNITS > 0 else ly_item_units
                today_units_base = total_units if LY_STORE_UNITS > 0 else (sycm_item['支付件数'].apply(parse_money).sum() if not sycm_item.empty and '支付件数' in sycm_item.columns else 0)

                ly_row['Traffic'] = get_yoy(store_traffic, LY_STORE_TRAFFIC) if LY_STORE_TRAFFIC > 0 else None
                
                if LY_STORE_TRAFFIC > 0 and store_buyers:
                    ly_cr = LY_STORE_BUYERS / LY_STORE_TRAFFIC
                    ly_row['CR%'] = (store_cr - ly_cr) / ly_cr if ly_cr > 0 else None
                else: ly_row['CR%'] = None
                    
                ly_row['Buyers'] = get_yoy(store_buyers, LY_STORE_BUYERS) if LY_STORE_BUYERS > 0 else None
                
                if LY_STORE_BUYERS > 0 and LY_STORE_GMV > 0 and today_record['ATV 客单价']:
                    ly_row['ATV 客单价'] = get_yoy(today_record['ATV 客单价'], LY_STORE_GMV / LY_STORE_BUYERS)
                    ly_row['UPT 客单件'] = get_yoy(today_record['UPT 客单件'], LY_STORE_UNITS / LY_STORE_BUYERS)
                else:
                    ly_row['ATV 客单价'] = None
                    ly_row['UPT 客单件'] = None
                    
                ly_row['AUV 件单价'] = get_yoy(today_auv, ly_gmv_base/ly_units_base if ly_units_base else 0)
                ly_row['Units Sold 件数'] = get_yoy(today_units_base, ly_units_base)
                ly_row['Gross Sales Demand 下单金额'] = get_yoy(gross_demand, LY_STORE_DEMAND) if LY_STORE_DEMAND > 0 else None
                ly_row['GMV （ 成交额）'] = get_yoy(today_gmv_fallback, ly_gmv_base)
                ly_row['ACC'] = get_yoy(today_record['ACC'], ly_acc)
                ly_row['Furniture'] = get_yoy(today_record['Furniture'], ly_furn)
                ly_row['Lighting'] = get_yoy(today_record['Lighting'], ly_light)
                
                ly_row['Returns  退款'] = get_yoy(store_refund, LY_STORE_REFUND) if LY_STORE_REFUND > 0 else None
                
                if LY_STORE_REFUND > 0 and LY_STORE_GMV > 0:
                    ly_net_inc = LY_STORE_GMV - LY_STORE_REFUND
                    ly_row['Net sales（含税）'] = get_yoy(net_sales_tax_inc, ly_net_inc)
                    ly_row['Net sales（去税）'] = get_yoy(net_sales_tax_excl, ly_net_inc / 1.13)
                else:
                    ly_row['Net sales（含税）'] = None 
                    ly_row['Net sales（去税）'] = None 

                recent_7.loc['LY'] = ly_row
                df_p4 = recent_7.set_index('Date').T.reset_index()
                df_p4.rename(columns={'index': 'Daily KPIs'}, inplace=True)

                orders['商品ID'] = orders['商品ID'].astype(str).str.replace(r'\.0$', '', regex=True)

                def extract_spu(title):
                    if pd.isna(title): return ''
                    match = re.findall(r'[A-Za-z][A-Za-z0-9\s\-_/]{3,}', str(title))
                    if match:
                        longest = max([m.strip() for m in match if m.strip()], key=len)
                        cleaned = re.sub(r'\bHAY\b', '', longest, flags=re.IGNORECASE).strip()
                        return cleaned if cleaned else str(title)
                    return str(title)
                orders['Description'] = orders['商品标题'].apply(extract_spu)

                def join_unique(x):
                    return '\n'.join(sorted(set([str(i).strip() for i in x if pd.notna(i) and str(i).strip() != 'nan'])))
                def clean_color(text):
                    if pd.isna(text): return ''
                    s = str(text).replace('：', ':').replace('；', ';')
                    parts = s.split(';')
                    cleaned =[p.split(':', 1)[1].strip() if ':' in p else p.strip() for p in parts if p.strip()]
                    return ' '.join(cleaned)
                orders['商品属性'] = orders['商品属性'].apply(clean_color)

                orders['Category'] = orders.apply(lambda x: get_category_by_mapping(x.get('商品ID'), x.get('商品标题'), id_to_cat), axis=1)

                bestsellers = orders.groupby('商品ID').agg(
                    Order_Gross_Sales=('买家实付金额', 'sum'),
                    Units=('购买数量', 'sum'),
                    SKU=('商家编码', join_unique),
                    Colour=('商品属性', join_unique),
                    Description=('Description', 'first')
                ).reset_index()

                if not sycm_item.empty and sycm_item['当日金额'].sum() > 0:
                    sycm_item['商品ID'] = sycm_item['商品ID'].astype(str).str.replace(r'\.0$', '', regex=True)
                    sycm_item_gmv = sycm_item.groupby('商品ID')['当日金额'].sum().reset_index()
                    sycm_item_gmv.rename(columns={'当日金额': 'SYCM_GMV'}, inplace=True)

                    bestsellers = pd.merge(bestsellers, sycm_item_gmv, on='商品ID', how='left')
                    bestsellers['Gross_Sales'] = bestsellers['SYCM_GMV'].fillna(0)
                else:
                    bestsellers['Gross_Sales'] = bestsellers['Order_Gross_Sales']

                bestsellers = bestsellers.sort_values('Gross_Sales', ascending=False).head(15).reset_index(drop=True)
                bestsellers['Contribution%'] = bestsellers['Gross_Sales'] / today_gmv_fallback if today_gmv_fallback > 0 else None
                bestsellers['No.'] = bestsellers.index + 1
                bestsellers['Pictures'] = ''
                df_p5 = bestsellers[['No.', 'SKU', 'Description', 'Colour', 'Pictures', 'Gross_Sales', 'Units', 'Contribution%']]

                output_excel = io.BytesIO()
                with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    worksheet = workbook.add_worksheet('Daily Dashboard')
                    
                    fmt_title = workbook.add_format({'bold': True, 'bg_color': '#000000', 'font_color': 'white', 'align': 'center_across', 'valign': 'vcenter', 'font_size': 12})
                    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'center'})
                    fmt_num = workbook.add_format({'num_format': '#,##0', 'border': 1, 'align': 'center'})
                    fmt_float = workbook.add_format({'num_format': '0.00', 'border': 1, 'align': 'center'}) 
                    fmt_pct_int = workbook.add_format({'num_format': '0%', 'border': 1, 'align': 'center'})    
                    fmt_pct_cr = workbook.add_format({'num_format': '0.00%', 'border': 1, 'align': 'center'}) 
                    fmt_text = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})

                    current_row = 0
                    def write_block(title, df, start_row):
                        for col_num in range(len(df.columns)):
                            if col_num == 0: worksheet.write(start_row, col_num, title, fmt_title)
                            else: worksheet.write(start_row, col_num, "", fmt_title) 
                                
                        for col_num, value in enumerate(df.columns):
                            worksheet.write(start_row + 1, col_num, value, fmt_header)
                        for r_idx, row in df.iterrows():
                            metric_name = str(df.iloc[r_idx, 0])
                            for c_idx, val in enumerate(row):
                                if c_idx == 0:
                                    worksheet.write(start_row + 2 + r_idx, c_idx, val, fmt_text)
                                    continue
                                col_name = str(df.columns[c_idx])
                                if pd.isna(val) or val is None or val == '':
                                    worksheet.write(start_row + 2 + r_idx, c_idx, '-', fmt_text) 
                                elif 'vs Y' in col_name or 'vs LY' in col_name:
                                    worksheet.write(start_row + 2 + r_idx, c_idx, float(val), fmt_pct_int) 
                                elif 'CR' in metric_name:
                                    worksheet.write(start_row + 2 + r_idx, c_idx, val, fmt_pct_cr) 
                                elif '%' in metric_name or '%' in col_name:
                                    worksheet.write(start_row + 2 + r_idx, c_idx, val, fmt_pct_int) 
                                elif 'UPT' in metric_name:
                                    worksheet.write(start_row + 2 + r_idx, c_idx, val, fmt_float) 
                                elif isinstance(val, (int, float)):
                                    worksheet.write(start_row + 2 + r_idx, c_idx, val, fmt_num)
                                else:
                                    worksheet.write(start_row + 2 + r_idx, c_idx, val, fmt_text)
                        return start_row + len(df) + 3

                    current_row = write_block("Tmall Store Sales performance- MTD", df_p1, current_row)
                    current_row = write_block("Tmall Store Sales by categories - MTD", df_p2, current_row)
                    current_row = write_block("Tmall Store Followers performance", df_p3, current_row)
                    current_row = write_block(f"Tmall Store Week {WEEK_NUM} Daily Dashboard", df_p4, current_row)
                    current_row = write_block("Tmall Store Bestsellers - Daily (TOP 15)", df_p5, current_row)

                    worksheet.set_column('A:A', 25)
                    worksheet.set_column('B:D', 20)
                    worksheet.set_column('E:Z', 15)

                st.success("🎉 日报数据处理完成！")
                
                dl_col1, dl_col2 = st.columns(2)
                with dl_col1:
                    st.download_button(
                        label="📊 1. 下载电商日报 (Excel)",
                        data=output_excel.getvalue(),
                        file_name=f"Tmall_Daily_Dashboard_{DATE_STR.replace('/','')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                with dl_col2:
                    history_csv = history_df.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        label="💾 2. 下载历史库 (明后天上传用)",
                        data=history_csv,
                        file_name=f"dashboard_history.csv",
                        mime="text/csv"
                    )

            except Exception as e:
                st.error(f"❌ 程序发生错误: {e}")


# =======================================================================================================
# ========================================== 模块二：月度排行 (双核心版) ===============================
# =======================================================================================================
elif app_mode == "📅 月度排行 (HAY Ranking)":
    
    st.title("📊 生意参谋双年份排行 - HAY Ranking")
    st.info("💡 只要你把26年和25年的表一起传上来，系统会自动给你抽出【两套独立报表】：一套26年的(带同比)，一套25年的(干干净净无同比)！")
    
    st.sidebar.header("📂 数据上传")
    file_curr = st.sidebar.file_uploader("1. 上传【今年当月】生意参谋数据 (如 2026年)", type=["xlsx", "xls", "csv"])
    file_last = st.sidebar.file_uploader("2. 上传【去年当月】生意参谋数据 (如 2025年)", type=["xlsx", "xls", "csv"])
    file_map = st.sidebar.file_uploader("3. 上传【分类映射表】", type=["xlsx", "xls", "csv"])
    
    # --- 辅助数据清洗 ---
    def load_data(file):
        if file.name.endswith('.csv'):
            try:
                df = pd.read_csv(file, encoding='gbk')
            except:
                file.seek(0)
                df = pd.read_csv(file, encoding='utf-8')
        else:
            df = pd.read_excel(file)
    
        df.columns = df.columns.astype(str).str.strip()
        if '商品ID' not in df.columns:
            for i in range(min(20, len(df))):
                row_values = df.iloc[i].astype(str).str.strip().values
                if '商品ID' in row_values:
                    df.columns = row_values
                    df = df.iloc[i+1:].reset_index(drop=True)
                    break
        df.columns = df.columns.astype(str).str.strip()
        return df
    
    def clean_id(df):
        if '商品ID' in df.columns:
            df['商品ID'] = df['商品ID'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        return df
    
    def to_numeric_col(series):
        cleaned = series.astype(str).str.replace(',', '').str.replace(' ', '')
        return pd.to_numeric(cleaned, errors='coerce').fillna(0)
        
    # --- 排版画板 ---
    def generate_excel_dashboard(df_ttl, df_fav, dict_cats, df_return, display_cat_names):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            title_fmt = workbook.add_format({'bold': True, 'font_size': 12, 'bg_color': '#CDE9F5', 'valign': 'vcenter'})
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#CDE9F5', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
            cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
            num_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0'})
            pct_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0%'})
            
            for cat_name in display_cat_names:
                sheet_name = f"HAY Ranking - {cat_name}"[:31]
                worksheet = workbook.add_worksheet(sheet_name)
                worksheet.write('A1', 'HAY Ranking', title_fmt)
                worksheet.set_row(0, 20)
                
                def write_table_to_excel(df, start_row, start_col):
                    for col_idx, col_name in enumerate(df.columns):
                        worksheet.write(start_row, start_col + col_idx, col_name, header_fmt)
                    for row_idx, row in enumerate(df.values):
                        for col_idx, val in enumerate(row):
                            col_name = df.columns[col_idx]
                            fmt = cell_fmt
                            if isinstance(val, (int, float)):
                                if 'Share%' in col_name or 'YOY' in col_name or '率' in col_name:
                                    fmt = pct_fmt
                                elif 'Value' in col_name or 'QTY' in col_name or '人数' in col_name:
                                    fmt = num_fmt
                            if pd.isna(val) or val == "":
                                worksheet.write(start_row + 1 + row_idx, start_col + col_idx, "-", cell_fmt)
                            else:
                                worksheet.write(start_row + 1 + row_idx, start_col + col_idx, val, fmt)
                
                write_table_to_excel(df_ttl, start_row=2, start_col=0)
                write_table_to_excel(df_fav, start_row=2, start_col=8)
                
                cat_df = dict_cats[cat_name].copy()
                cat_df.rename(columns={'Rank': f'{cat_name}\nRank', 'Share% of Category': f'Share% of\n{cat_name}'}, inplace=True)
                write_table_to_excel(cat_df, start_row=21, start_col=0)
                write_table_to_excel(df_return, start_row=21, start_col=8)
                
                worksheet.set_column('A:A', 8)
                worksheet.set_column('B:B', 30) 
                worksheet.set_column('C:C', 8)
                worksheet.set_column('D:G', 12)
                worksheet.set_column('H:H', 2)
                worksheet.set_column('I:I', 8)
                worksheet.set_column('J:J', 30) 
                worksheet.set_column('K:K', 8)
                worksheet.set_column('L:M', 12)
        return output.getvalue()

    # --- 核心：万能抽取函数 (给它一份数据，它吐出整套 Top15) ---
    def get_ranking_dfs(df_main_raw, df_last_raw, df_map_raw, is_ly_only=False):
        df_main = df_main_raw.copy()
        if '商品名称' in df_main.columns:
            df_main = df_main.sort_values(by='商品名称', na_position='last').drop_duplicates(subset=['商品ID'], keep='first')
        else:
            df_main = df_main.drop_duplicates(subset=['商品ID'], keep='first')

        if '一级' in df_map_raw.columns:
            df_map_unique = df_map_raw.drop_duplicates(subset=['商品ID'], keep='first')[['商品ID', '一级']]
        else:
            df_map_unique = pd.DataFrame(columns=['商品ID', '一级'])

        df_merged = pd.merge(df_main, df_map_unique, on='商品ID', how='left')
        df_merged['一级'] = df_merged['一级'].fillna('未分类')

        has_yoy = False
        if not is_ly_only and df_last_raw is not None and not df_last_raw.empty:
            df_last = df_last_raw.copy()
            if '支付金额' in df_last.columns:
                df_last['去年支付金额'] = to_numeric_col(df_last['支付金额'])
                df_last_sales = df_last.groupby('商品ID', as_index=False)['去年支付金额'].sum()
                df_merged = pd.merge(df_merged, df_last_sales, on='商品ID', how='left')
                has_yoy = True

        if '去年支付金额' not in df_merged.columns:
            df_merged['去年支付金额'] = 0.0

        numeric_cols =['支付金额', '去年支付金额', '支付件数', '商品收藏人数', '商品加购人数', '商品访客数', '成功退款金额']
        for col in numeric_cols:
            if col in df_merged.columns:
                df_merged[col] = to_numeric_col(df_merged[col])
            else:
                df_merged[col] = 0.0

        total_store_value = df_merged['支付金额'].sum()
        total_store_refund = df_merged['成功退款金额'].sum() if df_merged['成功退款金额'].sum() > 0 else 1

        df_merged['Value'] = df_merged['支付金额']
        df_merged['QTY'] = df_merged['支付件数']
        df_merged['Share% of TTL'] = np.where(total_store_value > 0, df_merged['Value'] / total_store_value, 0)

        if has_yoy:
            df_merged['YOY'] = np.where(df_merged['去年支付金额'] > 0, 
                                       (df_merged['Value'] - df_merged['去年支付金额']) / df_merged['去年支付金额'], 
                                       np.nan)

        df_merged['收加人数'] = df_merged['商品收藏人数'] + df_merged['商品加购人数']
        df_merged['收加率%'] = np.where(df_merged['商品访客数'] > 0, df_merged['收加人数'] / df_merged['商品访客数'], 0)
        df_merged['Picture'] = ""

        if '商品名称' in df_merged.columns:
            df_merged['Product'] = df_merged['商品名称'].fillna("未命名_ID:" + df_merged['商品ID'])
        else:
            df_merged['Product'] = df_merged['商品ID']

        df_merged['Return Value'] = df_merged['成功退款金额']
        df_merged['Return Share%'] = df_merged['Return Value'] / total_store_refund

        # ==========================================
        # 🚀 固定抽 Top 15，按需动态移除 YOY 列
        # ==========================================
        ttl_cols = ['Product', 'Picture', 'Value', 'QTY', 'Share% of TTL']
        cat_cols = ['Product', 'Picture', 'Value', 'QTY', 'Share% of Category']
        if has_yoy:
            ttl_cols.append('YOY')
            cat_cols.append('YOY')

        raw_ttl = df_merged.sort_values(by='Value', ascending=False).head(15)[ttl_cols].copy()
        raw_ttl.insert(0, 'TTL Rank', range(1, len(raw_ttl) + 1))

        raw_fav = df_merged.sort_values(by='收加人数', ascending=False).head(15)[['Product', 'Picture', '收加人数', '收加率%']].copy()
        raw_fav.insert(0, 'Rank', range(1, len(raw_fav) + 1))

        category_sales = df_merged.groupby('一级')['Value'].sum().sort_values(ascending=False)
        top_3_categories = [cat for cat in category_sales.index if cat != '未分类'][:3]

        raw_cats = {}
        for cat in top_3_categories:
            c_df = df_merged[df_merged['一级'] == cat].copy()
            c_total = c_df['Value'].sum()
            c_df['Share% of Category'] = np.where(c_total > 0, c_df['Value'] / c_total, 0)
            c_top15 = c_df.sort_values(by='Value', ascending=False).head(15)[cat_cols].copy()
            c_top15.insert(0, 'Rank', range(1, len(c_top15) + 1))
            raw_cats[cat] = c_top15

        # 植入全店总榜
        total_cat_df = raw_ttl.copy()
        total_cat_df.rename(columns={'TTL Rank': 'Rank', 'Share% of TTL': 'Share% of Category'}, inplace=True)
        raw_cats['全店总榜'] = total_cat_df
        display_cat_names = ['全店总榜'] + top_3_categories

        raw_return = df_merged.sort_values(by='Return Value', ascending=False).head(15)[['Product', 'Picture', 'Return Value', 'Return Share%']].copy()
        raw_return.rename(columns={'Return Value': 'Returned Value', 'Return Share%': 'Share% of TTL'}, inplace=True)
        raw_return.insert(0, 'HAY Rank', range(1, len(raw_return) + 1))

        return raw_ttl, raw_fav, raw_cats, raw_return, display_cat_names


    # --- 统一网页展现函数 ---
    def render_dashboard_ui(ttl, fav, cats, ret, display_names):
        def fmt_display(df):
            res = df.copy()
            for col in res.columns:
                if 'Value' in col or 'QTY' in col or '人数' in col:
                    res[col] = res[col].apply(lambda x: f"{x:,.0f}" if pd.notnull(x) else "-")
                elif 'Share%' in col or 'YOY' in col or '率' in col:
                    res[col] = res[col].apply(lambda x: f"{x:.0%}" if pd.notnull(x) else "-")
            return res.set_index(res.columns[0])

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("🏆 全店销售 Top 15 (TTL Rank)")
            st.dataframe(fmt_display(ttl), use_container_width=True)
        with col2:
            st.subheader("❤️ 收藏加购 Top 15")
            st.dataframe(fmt_display(fav), use_container_width=True)
        st.markdown("---")
        col3, col4 = st.columns(2)
        with col3:
            st.subheader("📦 各类目及全店总盘 Top 15")
            tabs = st.tabs([f"{cat} Rank" for cat in display_names])
            for i, cat in enumerate(display_names):
                with tabs[i]:
                    st.dataframe(fmt_display(cats[cat]), use_container_width=True)
        with col4:
            st.subheader("↩️ 退货 Top 15 (按退款金额)")
            st.dataframe(fmt_display(ret), use_container_width=True)

    # ==========================
    # 执行主逻辑
    # ==========================
    if file_curr and file_map:
        with st.spinner('正在执行强力去重与精密计算 (生成双年份 TOP 15)...'):
            
            df_curr_raw = clean_id(load_data(file_curr))
            df_map_raw = clean_id(load_data(file_map))
            
            df_last_raw = None
            if file_last:
                df_last_raw = clean_id(load_data(file_last))
                
            # 1️⃣ 计算【今年 (比如2026)】的数据看板 (有去年的参照，带YOY)
            curr_ttl, curr_fav, curr_cats, curr_ret, curr_names = get_ranking_dfs(df_curr_raw, df_last_raw, df_map_raw, is_ly_only=False)
            excel_curr = generate_excel_dashboard(curr_ttl, curr_fav, curr_cats, curr_ret, curr_names)
            
            # 2️⃣ 如果传了去年文件，系统顺手计算【去年 (比如2025)】的独立看板 (无参照，不带YOY)
            if df_last_raw is not None:
                ly_ttl, ly_fav, ly_cats, ly_ret, ly_names = get_ranking_dfs(df_last_raw, None, df_map_raw, is_ly_only=True)
                excel_ly = generate_excel_dashboard(ly_ttl, ly_fav, ly_cats, ly_ret, ly_names)

            # --- 显示界面 ---
            if df_last_raw is not None:
                st.success("✅ 双年份数据处理完成！已为你独立生成两份报表。")
                
                # 顶端双下载按钮
                dl_col1, dl_col2 = st.columns(2)
                with dl_col1:
                    st.download_button(label="📥 下载【今年当月】榜单 Excel (含YOY)", data=excel_curr, file_name="HAY_Ranking_CurrentYear.xlsx", type="primary")
                with dl_col2:
                    st.download_button(label="📥 下载【去年当月】独立榜单 Excel (无YOY)", data=excel_ly, file_name="HAY_Ranking_LastYear.xlsx")
                    
                # 网页展示用双标签切分
                tab_curr, tab_ly = st.tabs(["🔥 1. 今年当月排行 (对比去年，含YOY)", "⏪ 2. 去年当月排行 (仅抽Top15，无YOY)"])
                with tab_curr:
                    render_dashboard_ui(curr_ttl, curr_fav, curr_cats, curr_ret, curr_names)
                with tab_ly:
                    st.info("💡 这里的排名是用你去年的表格，独立跑出的一套 Top 15 总榜和细分榜。")
                    render_dashboard_ui(ly_ttl, ly_fav, ly_cats, ly_ret, ly_names)
                    
            else:
                # 只传了今年，没传去年
                st.success("✅ 单年份数据处理完成！(由于没有上传去年表格，已自动屏蔽YOY列)")
                st.download_button(label="📥 下载独立榜单 Excel (无YOY)", data=excel_curr, file_name="HAY_Ranking.xlsx", type="primary")
                render_dashboard_ui(curr_ttl, curr_fav, curr_cats, curr_ret, curr_names)

    else:
        st.info("👈 请在左侧依次上传：1.今年数据 2.去年数据 3.分类映射表")
