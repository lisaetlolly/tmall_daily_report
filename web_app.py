import streamlit as st
import pandas as pd
import numpy as np
import datetime
import re
import os
import io

# ==========================================
# 页面基础配置
# ==========================================
st.set_page_config(page_title="天猫电商日报生成器", page_icon="📊", layout="wide")
st.title("📊 Tmall Store Daily Dashboard")
st.markdown("💡 **新版逻辑**：表格自动提取单品数据 + 网页手工填写大盘数据。互不干涉，极致清爽。")

HISTORY_FILE = 'dashboard_history.csv'

# ==========================================
# 核心读取与清洗函数 
# ==========================================
def parse_money(s):
    if pd.isna(s) or s == '': return 0.0
    try: return float(str(s).replace('¥','').replace(',','').replace('%','').replace('元','').strip())
    except: return 0.0

def categorize_item(title):
    t = str(title).lower()
    if any(x in t for x in['灯', 'light', 'lamp', 'portable', 'shade', 'pendant', 'bulb', 'sconce', 'apex']): return 'Lighting'
    if any(x in t for x in['椅', '柜', '沙发', 'stool', 'table', 'cabinet', 'cabine', 'chair', 'bench', 'desk', 'sofa', 'pouf', 'bed', 'rack', 'shelv']): return 'Furniture'
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
    except Exception as e:
        st.error(f"❌ 读取失败: {e}")
        return None

def find_value_in_df(df, keyword):
    if df is None or df.empty: return 0.0
    for r in range(len(df)):
        for c in range(len(df.columns)):
            if keyword in str(df.iloc[r, c]):
                if c + 1 < len(df.columns) and pd.notna(df.iloc[r, c+1]) and str(df.iloc[r, c+1]).strip() != '':
                    return parse_money(df.iloc[r, c+1])
                if r + 1 < len(df) and pd.notna(df.iloc[r+1, c]) and str(df.iloc[r+1, c]).strip() != '':
                    return parse_money(df.iloc[r+1, c])
    return 0.0

# ==========================================
# UI 布局: 模块化分离
# ==========================================

# 📅 1. 日期选择
st.markdown("### 📅 第一步：确认报表日期")
report_date = st.date_input("选择你要生成哪一天的报表？", datetime.date.today() - datetime.timedelta(days=1))

# 🗂️ 2. 文件上传区
st.markdown("### 🗂️ 第二步：上传表格文件")
col1, col2, col3, col4 = st.columns(4)
with col1: file_order = st.file_uploader("1. 今日订单底表 (必传)", type=['xlsx', 'xls', 'csv'])
with col2: file_item = st.file_uploader("2. 今日单品生参 (必传)", type=['xlsx', 'xls', 'csv'])
with col3: file_target = st.file_uploader("3. 月度规划表 (每月传)", type=['xlsx', 'xls', 'csv'])
with col4: file_ly = st.file_uploader("4. 去年当日单品 (选传)", type=['xlsx', 'xls', 'csv'])

# ✍️ 3. 每日手工填报区
st.markdown("### ✍️ 第三步：填写今日大盘核心数据")
st.info("直接照着生意参谋的概览页面抄进来，免去导大盘表的烦恼！")

with st.container():
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown("**【今日成交漏斗】**")
        in_gmv = st.number_input("今日 全店支付金额 (GMV)", value=0.0, step=100.0)
        in_demand = st.number_input("今日 下单金额 (Demand)", value=0.0, step=100.0)
        in_refund = st.number_input("今日 成功退款金额", value=0.0, step=100.0)
    with c2:
        st.markdown("**【今日流量与用户】**")
        in_traffic = st.number_input("今日 访客数", value=0, step=100)
        in_buyers = st.number_input("今日 支付买家数", value=0, step=10)
        in_cr = st.number_input("今日 转化率 (%) 填数字", value=0.00, step=0.1, help="如 2.58% 填 2.58")
    with c3:
        st.markdown("**【粉丝与累计退款】**")
        in_new_fans = st.number_input("今日 新增粉丝数", value=0, step=10)
        in_acc_fans = st.number_input("累计 总粉丝数", value=0, step=100)
        in_mtd_refund = st.number_input("🔙 本月累计退款 (必填)", value=0.0, step=1000.0, help="用于计算本月累计净销售额")
    with c4:
        st.markdown("**【去年今日精准对比】**")
        in_ly_gmv = st.number_input("去年今日 GMV", value=0.0, step=100.0)
        in_ly_traffic = st.number_input("去年今日 访客数", value=0, step=100)
        in_ly_buyers = st.number_input("去年今日 买家数", value=0, step=10)
        in_ly_units = st.number_input("去年今日 件数", value=0, step=10)

st.divider()

# ==========================================
# 执行生成逻辑
# ==========================================
if st.button("⚡ 一键生成智能日报", type="primary", use_container_width=True):
    if not file_order and not file_item:
        st.warning("⚠️ 订单底表和单品生参至少要传一个！")
        st.stop()

    with st.spinner("🔄 AI 正在深度融合您的数据..."):
        try:
            # --- 日期变量构建 ---
            DATE_STR = f"{report_date.month}/{report_date.day}"
            DAYS_PASSED = report_date.day
            MONTH_NAME = report_date.strftime('%b') 
            WEEK_NUM = report_date.isocalendar()[1] 
            TOTAL_DAYS = 31 if report_date.month in[1,3,5,7,8,10,12] else (28 if report_date.month == 2 else 30)

            # --- 安全读取底层数据 ---
            orders = read_excel_smart(file_order, '商品标题')
            if orders is None: orders = pd.DataFrame()
            for c in['商品价格', '购买数量', '买家应付货款', '买家实付金额', '商品ID', '商家编码', '商品属性', '商品标题']:
                if c not in orders.columns: orders[c] = 0.0 if '金额' in c or '数量' in c else ''

            sycm_item = read_excel_smart(file_item, '支付金额')
            if sycm_item is None: sycm_item = pd.DataFrame()
            for c in['商品名称', '支付金额', '支付件数', '月累计支付金额', '月累计支付件数', '商品ID']:
                if c not in sycm_item.columns: sycm_item[c] = 0.0 if '金额' in c or '件数' in c else ''

            ly_sycm = read_excel_smart(file_ly, '支付金额') if file_ly else None

            # --- 获取前端手工参数 ---
            store_gmv = float(in_gmv)
            store_demand = float(in_demand)
            store_refund = float(in_refund)
            store_traffic = int(in_traffic)
            store_buyers = int(in_buyers)
            store_cr = float(in_cr) / 100.0  # 转为小数
            new_followers = int(in_new_fans)
            acc_followers = int(in_acc_fans)
            MTD_REFUND_ACTUAL = float(in_mtd_refund)

            LY_STORE_GMV = float(in_ly_gmv)
            LY_STORE_TRAFFIC = int(in_ly_traffic)
            LY_STORE_BUYERS = int(in_ly_buyers)
            LY_STORE_UNITS = int(in_ly_units)

            # ==============================================================
            # 🎯 全自动解析《月度目标规划表》
            # ==============================================================
            TARGET_GMV_MONTH, TARGET_NET_MONTH = 0.0, 0.0
            TARGET_GMV_MTD, TARGET_NET_MTD = 0.0, 0.0
            EST_REST_MONTH_GMV, EST_REST_MONTH_NET = 0.0, 0.0
            LY_WHOLE_MONTH_ACTUAL_GMV, LY_WHOLE_MONTH_ACTUAL_NET = 0.0, 0.0
            
            if file_target:
                try:
                    file_target.seek(0)
                    t_ext = os.path.splitext(file_target.name)[-1].lower()
                    t_engine = 'xlrd' if t_ext == '.xls' else 'openpyxl'
                    target_df = pd.read_excel(file_target, header=None, engine=t_engine)
                    
                    # 1. 抓取 MTD 目标
                    total_col_idx = -1
                    for r in range(min(10, len(target_df))):
                        for c in range(len(target_df.columns)):
                            if '合计' in str(target_df.iloc[r, c]):
                                total_col_idx = c
                                break
                        if total_col_idx != -1: break
                    
                    gmv_row_idx = -1
                    net_row_idx = -1
                    for r in range(len(target_df)):
                        row_vals =[str(x).strip().lower() for x in target_df.iloc[r].values]
                        if 'gmv' in row_vals and gmv_row_idx == -1: gmv_row_idx = r
                        if any('不含税ns' in x for x in row_vals) and net_row_idx == -1: net_row_idx = r
                    
                    if total_col_idx != -1:
                        if gmv_row_idx != -1:
                            TARGET_GMV_MONTH = parse_money(target_df.iloc[gmv_row_idx, total_col_idx])
                            TARGET_GMV_MTD = sum([parse_money(target_df.iloc[gmv_row_idx, total_col_idx + i]) for i in range(1, DAYS_PASSED + 1)])
                        if net_row_idx != -1:
                            TARGET_NET_MONTH = parse_money(target_df.iloc[net_row_idx, total_col_idx])
                            TARGET_NET_MTD = sum([parse_money(target_df.iloc[net_row_idx, total_col_idx + i]) for i in range(1, DAYS_PASSED + 1)])

                    # 2. 抓取人工写的预估与历史数据 (从目标表里挖)
                    EST_REST_MONTH_GMV = find_value_in_df(target_df, '预估剩余GMV')
                    EST_REST_MONTH_NET = find_value_in_df(target_df, '预估剩余Net')
                    LY_WHOLE_MONTH_ACTUAL_GMV = find_value_in_df(target_df, '去年全月GMV')
                    LY_WHOLE_MONTH_ACTUAL_NET = find_value_in_df(target_df, '去年全月Net')
                except Exception as e:
                    st.warning(f"规划表解析异常，使用了默认0值: {e}")

            # --- 提取去年单品同比 ---
            if ly_sycm is not None and not ly_sycm.empty:
                ly_sycm['Category'] = ly_sycm['商品名称'].apply(categorize_item)
                ly_item_gmv = ly_sycm['支付金额'].apply(parse_money).sum()
                ly_item_units = ly_sycm['支付件数'].apply(parse_money).sum()
                ly_cat = ly_sycm.groupby('Category')['支付金额'].apply(lambda x: x.apply(parse_money).sum())
                ly_acc = ly_cat.get('ACC', 0)
                ly_furn = ly_cat.get('Furniture', 0)
                ly_light = ly_cat.get('Lighting', 0)
            else:
                ly_item_gmv = ly_item_units = ly_acc = ly_furn = ly_light = 0

            # ==========================================
            # Part 1: MTD
            # ==========================================
            mtd_gmv = sycm_item['月累计支付金额'].apply(parse_money).sum() if not sycm_item.empty else 0
            mtd_net_sales = (mtd_gmv - MTD_REFUND_ACTUAL) / 1.13 if MTD_REFUND_ACTUAL > 0 else mtd_gmv / 1.13

            est_whole_month_gmv = mtd_gmv + EST_REST_MONTH_GMV if EST_REST_MONTH_GMV > 0 else (mtd_gmv / DAYS_PASSED) * TOTAL_DAYS
            est_whole_month_net = mtd_net_sales + EST_REST_MONTH_NET if EST_REST_MONTH_NET > 0 else (mtd_net_sales / DAYS_PASSED) * TOTAL_DAYS

            yoy_est_gmv = (est_whole_month_gmv - LY_WHOLE_MONTH_ACTUAL_GMV) / LY_WHOLE_MONTH_ACTUAL_GMV if LY_WHOLE_MONTH_ACTUAL_GMV > 0 else None
            yoy_est_net = (est_whole_month_net - LY_WHOLE_MONTH_ACTUAL_NET) / LY_WHOLE_MONTH_ACTUAL_NET if LY_WHOLE_MONTH_ACTUAL_NET > 0 else None

            df_p1 = pd.DataFrame({
                f'Updated {DATE_STR}':['GMV', 'Net sales(Tax excluded )'],
                'MTD Actual':[mtd_gmv, mtd_net_sales],
                'MTD Target':[TARGET_GMV_MTD, TARGET_NET_MTD],
                'MTD Achi%':[mtd_gmv/TARGET_GMV_MTD if TARGET_GMV_MTD else 0, mtd_net_sales/TARGET_NET_MTD if TARGET_NET_MTD else 0],
                'estimated sales of the rest of Month':[EST_REST_MONTH_GMV, EST_REST_MONTH_NET],
                'estimated sales of the whole month':[est_whole_month_gmv, est_whole_month_net],
                f'Monthly target {MONTH_NAME}':[TARGET_GMV_MONTH, TARGET_NET_MONTH],
                'estimated Achi% of the whole month':[est_whole_month_gmv/TARGET_GMV_MONTH if TARGET_GMV_MONTH else 0, est_whole_month_net/TARGET_NET_MONTH if TARGET_NET_MONTH else 0],
                f'estimated whole month sales vs Y25 {MONTH_NAME}':[yoy_est_gmv, yoy_est_net] 
            })

            # ==========================================
            # Part 2: Categories MTD
            # ==========================================
            sycm_item['Category'] = sycm_item['商品名称'].apply(categorize_item)
            sycm_item['月累计金额'] = sycm_item['月累计支付金额'].apply(parse_money)
            sycm_item['月累计件数'] = sycm_item['月累计支付件数'].apply(parse_money)
            sycm_item['当日金额'] = sycm_item['支付金额'].apply(parse_money)

            cat_group = sycm_item.groupby('Category').agg(
                Daily_GMV=('当日金额', 'sum'),
                MTD_Actual=('月累计金额', 'sum'),
                MTD_Units=('月累计件数', 'sum')
            ).reset_index()
            
            # 各品类目标通常按照整体进度同比例切割（如果没有具体填写）
            target_ratio = {'Lighting': 0.1, 'Furniture': 0.35, 'ACC': 0.55}
            cat_group['MTD Target'] = cat_group['Category'].map(lambda x: TARGET_GMV_MTD * target_ratio.get(x, 0))
            
            if cat_group.empty:
                cat_group = pd.DataFrame({'Category':['Lighting', 'Furniture', 'ACC'], 'Daily_GMV': [0,0,0], 'MTD_Actual':[0,0,0], 'MTD_Units':[0,0,0], 'MTD Target':[0,0,0]})
            
            cat_group['MTD Achi%'] = cat_group['MTD_Actual'] / cat_group['MTD Target'].replace(0, 1)
            cat_group['MTD contribution%'] = cat_group['MTD_Actual'] / mtd_gmv if mtd_gmv else 0
            cat_group['MTD AUV'] = cat_group['MTD_Actual'] / cat_group['MTD_Units'].replace(0, 1)

            df_p2 = cat_group[['Category', 'Daily_GMV', 'MTD_Actual', 'MTD Target', 'MTD Achi%', 'MTD contribution%', 'MTD AUV', 'MTD_Units']].copy()
            df_p2['Category'] = pd.Categorical(df_p2['Category'], categories=['Lighting', 'Furniture', 'ACC'], ordered=True)
            df_p2 = df_p2.sort_values('Category')
            df_p2.rename(columns={'Daily_GMV': f'GMV {DATE_STR}'}, inplace=True)

            total_row = pd.DataFrame([{
                'Category': 'Total',
                f'GMV {DATE_STR}': df_p2[f'GMV {DATE_STR}'].sum(),
                'MTD_Actual': df_p2['MTD_Actual'].sum(),
                'MTD Target': df_p2['MTD Target'].sum(),
                'MTD Achi%': df_p2['MTD_Actual'].sum() / df_p2['MTD Target'].sum() if df_p2['MTD Target'].sum() else 0,
                'MTD contribution%': 1.0 if df_p2['MTD_Actual'].sum() > 0 else 0,
                'MTD AUV': df_p2['MTD_Actual'].sum() / df_p2['MTD_Units'].sum() if df_p2['MTD_Units'].sum() else 0,
                'MTD_Units': df_p2['MTD_Units'].sum()
            }])
            df_p2 = pd.concat([df_p2, total_row], ignore_index=True)

            # ==========================================
            # Part 3: Followers Performance
            # ==========================================
            df_p3 = pd.DataFrame({
                f'Updated {DATE_STR}':['New Followers', 'Accumulated Followers'],
                'Data':[new_followers, acc_followers]
            })

            # ==========================================
            # Part 4: Daily Dashboard
            # ==========================================
            orders['购买数量'] = orders['购买数量'].apply(parse_money)
            orders['买家实付金额'] = orders['买家实付金额'].apply(parse_money)
            total_units = orders['购买数量'].sum()
            
            gross_demand = store_demand if store_demand > 0 else orders['买家实付金额'].sum()
            today_gmv_fallback = store_gmv if store_gmv > 0 else sycm_item['当日金额'].sum()
            today_auv = today_gmv_fallback / total_units if total_units else 0

            net_sales_tax_inc = today_gmv_fallback - store_refund
            net_sales_tax_excl = net_sales_tax_inc / 1.13

            today_item_gmv = sycm_item['当日金额'].sum() if not sycm_item.empty else 0
            today_item_units = sycm_item['支付件数'].apply(parse_money).sum() if '支付件数' in sycm_item.columns else 0

            today_record = {
                'Date': DATE_STR,
                'Traffic': store_traffic,
                'CR%': store_cr,
                'Buyers': store_buyers,
                'ATV 客单价': today_gmv_fallback / store_buyers if store_buyers else 0,
                'UPT 客单件': total_units / store_buyers if store_buyers else 0,
                'AUV 件单价': today_auv,
                'Units Sold 件数': total_units,
                'Gross Sales Demand 下单金额': gross_demand,
                'GMV （ 成交额）': today_gmv_fallback,
                'ACC': df_p2.loc[df_p2['Category']=='ACC', f'GMV {DATE_STR}'].sum(),
                'Furniture': df_p2.loc[df_p2['Category']=='Furniture', f'GMV {DATE_STR}'].sum(),
                'Lighting': df_p2.loc[df_p2['Category']=='Lighting', f'GMV {DATE_STR}'].sum(),
                'Returns  退款': store_refund,
                'Net sales（含税）': net_sales_tax_inc,
                'Net sales（去税）': net_sales_tax_excl
            }
            df_today = pd.DataFrame([today_record])

            if os.path.exists(HISTORY_FILE):
                history_df = pd.read_csv(HISTORY_FILE)
                history_df = history_df[history_df['Date'] != DATE_STR] 
                history_df = pd.concat([history_df, df_today], ignore_index=True)
            else:
                history_df = df_today
            history_df.to_csv(HISTORY_FILE, index=False)

            recent_7 = history_df.tail(7).copy()
            wtd_row = recent_7.sum(numeric_only=True)
            wtd_row['CR%'] = wtd_row['Buyers'] / wtd_row['Traffic'] if wtd_row.get('Traffic', 0) else 0
            wtd_row['ATV 客单价'] = wtd_row['GMV （ 成交额）'] / wtd_row['Buyers'] if wtd_row.get('Buyers', 0) else 0
            wtd_row['UPT 客单件'] = wtd_row['Units Sold 件数'] / wtd_row['Buyers'] if wtd_row.get('Buyers', 0) else 0
            wtd_row['AUV 件单价'] = wtd_row['GMV （ 成交额）'] / wtd_row['Units Sold 件数'] if wtd_row.get('Units Sold 件数', 0) else 0

            recent_7.loc['WTD'] = wtd_row
            recent_7.at['WTD', 'Date'] = 'Week to Date'

            def get_yoy(today_val, ly_val):
                if ly_val and ly_val > 0: return (today_val - ly_val) / ly_val
                return None

            ly_row = pd.Series(dtype=object)
            ly_row['Date'] = 'Today vs LY'
            
            ly_gmv_base = LY_STORE_GMV if LY_STORE_GMV > 0 else ly_item_gmv
            today_gmv_base = today_gmv_fallback if LY_STORE_GMV > 0 else today_item_gmv
            
            ly_units_base = LY_STORE_UNITS if LY_STORE_UNITS > 0 else ly_item_units
            today_units_base = total_units if LY_STORE_UNITS > 0 else today_item_units

            ly_row['Traffic'] = get_yoy(store_traffic, LY_STORE_TRAFFIC) if LY_STORE_TRAFFIC > 0 else None
            
            if LY_STORE_TRAFFIC > 0 and LY_STORE_BUYERS > 0:
                ly_cr = LY_STORE_BUYERS / LY_STORE_TRAFFIC
                ly_row['CR%'] = (store_cr - ly_cr) / ly_cr if ly_cr > 0 else None
            else: ly_row['CR%'] = None
                
            ly_row['Buyers'] = get_yoy(store_buyers, LY_STORE_BUYERS) if LY_STORE_BUYERS > 0 else None
            
            if LY_STORE_BUYERS > 0 and LY_STORE_GMV > 0:
                ly_row['ATV 客单价'] = get_yoy(today_record['ATV 客单价'], LY_STORE_GMV / LY_STORE_BUYERS)
                ly_row['UPT 客单件'] = get_yoy(today_record['UPT 客单件'], LY_STORE_UNITS / LY_STORE_BUYERS)
            else:
                ly_row['ATV 客单价'] = None
                ly_row['UPT 客单件'] = None
                
            ly_row['AUV 件单价'] = get_yoy(today_auv, ly_gmv_base/ly_units_base if ly_units_base else 0)
            ly_row['Units Sold 件数'] = get_yoy(today_units_base, ly_units_base)
            ly_row['Gross Sales Demand 下单金额'] = None 
            ly_row['GMV （ 成交额）'] = get_yoy(today_gmv_base, ly_gmv_base)
            ly_row['ACC'] = get_yoy(today_record['ACC'], ly_acc)
            ly_row['Furniture'] = get_yoy(today_record['Furniture'], ly_furn)
            ly_row['Lighting'] = get_yoy(today_record['Lighting'], ly_light)
            
            # 手填了去年退款才能算退款和净销同比
            in_ly_refund = float(st.session_state.get('in_ly_refund', 0.0)) if 'in_ly_refund' in st.session_state else 0.0
            ly_row['Returns  退款'] = get_yoy(store_refund, in_ly_refund) if in_ly_refund > 0 else None
            
            if in_ly_refund > 0 and LY_STORE_GMV > 0:
                ly_net_inc = LY_STORE_GMV - in_ly_refund
                ly_row['Net sales（含税）'] = get_yoy(net_sales_tax_inc, ly_net_inc)
                ly_row['Net sales（去税）'] = get_yoy(net_sales_tax_excl, ly_net_inc / 1.13)
            else:
                ly_row['Net sales（含税）'] = None 
                ly_row['Net sales（去税）'] = None 

            recent_7.loc['LY'] = ly_row
            df_p4 = recent_7.set_index('Date').T.reset_index()
            df_p4.rename(columns={'index': 'Daily KPIs'}, inplace=True)

            # ==========================================
            # Part 5: Top 15 
            # ==========================================
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
            bestsellers['Contribution%'] = bestsellers['Gross_Sales'] / today_gmv_fallback if today_gmv_fallback > 0 else 0
            bestsellers['No.'] = bestsellers.index + 1
            bestsellers['Pictures'] = ''
            df_p5 = bestsellers[['No.', 'SKU', 'Description', 'Colour', 'Pictures', 'Gross_Sales', 'Units', 'Contribution%']]

            # ==========================================
            # 导出排版
            # ==========================================
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
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
                            if pd.isna(val) or val == '':
                                worksheet.write(start_row + 2 + r_idx, c_idx, '-', fmt_text) 
                            elif 'vs Y25' in col_name or 'vs LY' in col_name:
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
                current_row = write_block("Tmall Store Bestsellers - Daily", df_p5, current_row)

                worksheet.set_column('A:A', 25)
                worksheet.set_column('B:D', 20)
                worksheet.set_column('E:Z', 15)

            st.success("🎉 报表生成成功！点击下方按钮下载。")
            output.seek(0)
            
            st.download_button(
                label="📥 下载 Tmall_Daily_Dashboard.xlsx",
                data=output,
                file_name=f"Tmall_Daily_Dashboard_{DATE_STR.replace('/','')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        except Exception as e:
            st.error(f"❌ 程序发生错误: {e}")
