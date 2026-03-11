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
st.markdown("上传表格 + 补全人工参数，生成 100% 完美对齐的业务大盘！")

HISTORY_FILE = 'dashboard_history.csv'

# ==========================================
# 核心函数 
# ==========================================
def parse_money(s):
    if pd.isna(s) or s == '': return 0.0
    try: return float(str(s).replace('¥','').replace(',','').replace('%','').replace('元','').strip())
    except: return 0.0

def categorize_item(title):
    t = str(title).lower()
    if any(x in t for x in['灯', 'light', 'lamp', 'portable', 'shade', 'pendant', 'bulb', 'sconce', 'apex']): return 'Lighting'
    # 剔除 table, rug 等容易混淆的词，严格锁定核心大件家具
    if any(x in t for x in['椅', '柜', '沙发', 'stool', 'cabinet', 'cabine', 'chair', 'bench', 'desk', 'sofa', 'pouf', 'bed', 'rack', 'shelv']): return 'Furniture'
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
        st.error(f"❌ 读取 {file_obj.name} 失败: {e}")
        return None

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

# ==========================================
# UI 布局: 表格上传区 + 手工参数补全区
# ==========================================
col1, col2 = st.columns(2)
with col1:
    file_order = st.file_uploader("1. 📥 今日订单底表", type=['xlsx', 'xls', 'csv'])
    file_store = st.file_uploader("3. 📥 店铺大盘表", type=['xlsx', 'xls', 'csv'])
    file_target = st.file_uploader("5. 🎯 月度目标规划表", type=['xlsx', 'xls', 'csv'])
with col2:
    file_item = st.file_uploader("2. 📥 今日单品生参", type=['xlsx', 'xls', 'csv'])
    file_ly = st.file_uploader("4. 🕰️ 去年当日单品", type=['xlsx', 'xls', 'csv'])

st.divider()

# 🔥 核心升级：让你直接手填那些无法导出的参数！
with st.expander("🛠️ 点击展开：手工业务参数补全 (填入后同比和预估 100% 准确)"):
    st.markdown("*(以下数据若有则填，不填则系统自动估算或不显示同比)*")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**【预估与月度数据】**")
        manual_mtd_refund = st.number_input("💸 本月累计退款 (用于算MTD Net)", value=0)
        manual_est_rest_gmv = st.number_input("🔮 预估剩余全月 GMV", value=0)
        manual_est_rest_net = st.number_input("🔮 预估剩余全月 Net", value=0)
    with c2:
        st.markdown("**【去年今日大盘(算同比)】**")
        manual_ly_traffic = st.number_input("👥 去年今日 访客数", value=0)
        manual_ly_buyers = st.number_input("🛒 去年今日 买家数", value=0)
        manual_ly_refund = st.number_input("🔙 去年今日 退款金额", value=0)
    with c3:
        st.markdown("**【其他】**")
        manual_ly_demand = st.number_input("💰 去年今日 下单金额", value=0)
        manual_ly_whole_gmv = st.number_input("📅 去年同月 全月GMV", value=0)
        manual_ly_whole_net = st.number_input("📅 去年同月 全月Net", value=0)

# ==========================================
# 执行生成逻辑
# ==========================================
if st.button("⚡ 一键生成智能日报", type="primary", use_container_width=True):
    if not file_order and not file_item and not file_store:
        st.warning("⚠️ 请至少上传前 3 个文件中的 1 个！")
        st.stop()

    with st.spinner("🔄 AI 正在深度融合数据与人工参数..."):
        try:
            # --- 智能日期识别 ---
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
            TOTAL_DAYS = 31 if d.month in[1,3,5,7,8,10,12] else (28 if d.month == 2 else 30)

            # --- 安全读取底层数据 ---
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

            # --- 提取动态大盘数据 ---
            # 支持多种可能的导出列名
            def get_val(df, keys):
                if df is None: return 0.0
                for k in keys:
                    if k in df.columns: return parse_money(df[k].iloc[0])
                return 0.0

            store_gmv = get_val(sycm_store,['支付金额', '成交额'])
            store_demand = get_val(sycm_store,['下单金额', 'Gross Sales Demand', '访客价值金额']) # 增加宽容度
            store_refund = get_val(sycm_store, ['成功退款金额', '退款金额（完结时间）'])
            store_traffic = get_val(sycm_store, ['访客数'])
            store_buyers = get_val(sycm_store, ['支付买家数'])
            store_cr = get_val(sycm_store, ['支付转化率'])
            if store_cr > 1: store_cr /= 100 
            new_followers = get_val(sycm_store, ['新增粉丝数', '关注店铺人数'])
            acc_followers = get_val(sycm_store, ['累计粉丝数', '总粉丝数'])

            # ==============================================================
            # 解析《月度目标规划表》
            # ==============================================================
            TARGET_GMV_MONTH, TARGET_NET_MONTH = 0.0, 0.0
            TARGET_GMV_MTD, TARGET_NET_MTD = 0.0, 0.0
            
            if file_target:
                try:
                    file_target.seek(0)
                    t_ext = os.path.splitext(file_target.name)[-1].lower()
                    t_engine = 'xlrd' if t_ext == '.xls' else 'openpyxl'
                    target_df = pd.read_excel(file_target, header=None, engine=t_engine)
                    
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
                except Exception as e:
                    pass

            # ==========================================
            # 整合 LY 同比数据
            # ==========================================
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
            # 使用人工填写的月累计退款，最精准
            mtd_net_sales = (mtd_gmv - manual_mtd_refund) / 1.13 if manual_mtd_refund > 0 else mtd_gmv / 1.13

            est_whole_month_gmv = mtd_gmv + manual_est_rest_gmv if manual_est_rest_gmv > 0 else (mtd_gmv / DAYS_PASSED) * TOTAL_DAYS
            est_whole_month_net = mtd_net_sales + manual_est_rest_net if manual_est_rest_net > 0 else (mtd_net_sales / DAYS_PASSED) * TOTAL_DAYS

            yoy_est_gmv = (est_whole_month_gmv - manual_ly_whole_gmv) / manual_ly_whole_gmv if manual_ly_whole_gmv > 0 else None
            yoy_est_net = (est_whole_month_net - manual_ly_whole_net) / manual_ly_whole_net if manual_ly_whole_net > 0 else None

            df_p1 = pd.DataFrame({
                f'Updated {DATE_STR}':['GMV', 'Net sales(Tax excluded )'],
                'MTD Actual':[mtd_gmv, mtd_net_sales],
                'MTD Target':[TARGET_GMV_MTD, TARGET_NET_MTD],
                'MTD Achi%':[mtd_gmv/TARGET_GMV_MTD if TARGET_GMV_MTD else 0, mtd_net_sales/TARGET_NET_MTD if TARGET_NET_MTD else 0],
                'estimated sales of the rest of Month':[manual_est_rest_gmv, manual_est_rest_net],
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

            # 提取大盘表里的品类目标 (如果配了的话)
            TARGET_LIGHTING_MTD = get_val(sycm_store,['MTD_Lighting目标'])
            TARGET_FURNITURE_MTD = get_val(sycm_store,['MTD_Furniture目标'])
            TARGET_ACC_MTD = get_val(sycm_store, ['MTD_ACC目标'])
            target_dict = {'Lighting': TARGET_LIGHTING_MTD, 'Furniture': TARGET_FURNITURE_MTD, 'ACC': TARGET_ACC_MTD}
            
            if cat_group.empty:
                cat_group = pd.DataFrame({'Category':['Lighting', 'Furniture', 'ACC'], 'Daily_GMV': [0,0,0], 'MTD_Actual':[0,0,0], 'MTD_Units':[0,0,0]})
            
            cat_group['MTD Target'] = cat_group['Category'].map(target_dict).fillna(0)
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
            # Part 3: Followers
            # ==========================================
            df_p3 = pd.DataFrame({
                f'Updated {DATE_STR}':['New Followers', 'Accumulated Followers'],
                'Data': [new_followers, acc_followers]
            })

            # ==========================================
            # Part 4: Daily Dashboard
            # ==========================================
            orders['购买数量'] = orders['购买数量'].apply(parse_money)
            orders['买家实付金额'] = orders['买家实付金额'].apply(parse_money)
            total_units = orders['购买数量'].sum()
            
            # 安全兜底：如果没有大盘下单金额，用底表实付补齐
            gross_demand = store_demand if store_demand > 0 else orders['买家实付金额'].sum()
            today_gmv_fallback = store_gmv if store_gmv > 0 else sycm_item['当日金额'].sum()
            today_auv = today_gmv_fallback / total_units if total_units else 0

            net_sales_tax_inc = today_gmv_fallback - store_refund
            net_sales_tax_excl = net_sales_tax_inc / 1.13

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

            # 管理历史
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

            # 🛠️ 终极同比计算 (融合手工参数)
            def get_yoy(today_val, ly_val):
                if ly_val and ly_val > 0: return (today_val - ly_val) / ly_val
                return None

            ly_row = pd.Series(dtype=object)
            ly_row['Date'] = 'Today vs LY'
            
            # 若手填了LY流量则算，否则为 -
            ly_row['Traffic'] = get_yoy(store_traffic, manual_ly_traffic) if manual_ly_traffic > 0 else None
            
            # CR% 同比
            if manual_ly_traffic > 0 and manual_ly_buyers > 0:
                ly_cr = manual_ly_buyers / manual_ly_traffic
                ly_row['CR%'] = (store_cr - ly_cr) / ly_cr if ly_cr > 0 else None
            else:
                ly_row['CR%'] = None
                
            ly_row['Buyers'] = get_yoy(store_buyers, manual_ly_buyers) if manual_ly_buyers > 0 else None
            
            # ATV & UPT 同比
            if manual_ly_buyers > 0 and ly_item_gmv > 0:
                ly_row['ATV 客单价'] = get_yoy(today_record['ATV 客单价'], ly_item_gmv / manual_ly_buyers)
                ly_row['UPT 客单件'] = get_yoy(today_record['UPT 客单件'], ly_item_units / manual_ly_buyers)
            else:
                ly_row['ATV 客单价'] = None
                ly_row['UPT 客单件'] = None
                
            ly_row['AUV 件单价'] = get_yoy(today_auv, ly_item_gmv/ly_item_units if ly_item_units else 0)
            ly_row['Units Sold 件数'] = get_yoy(total_units, ly_item_units)
            ly_row['Gross Sales Demand 下单金额'] = get_yoy(gross_demand, manual_ly_demand) if manual_ly_demand > 0 else None
            ly_row['GMV （ 成交额）'] = get_yoy(today_gmv_fallback, ly_item_gmv)
            ly_row['ACC'] = get_yoy(today_record['ACC'], ly_acc)
            ly_row['Furniture'] = get_yoy(today_record['Furniture'], ly_furn)
            ly_row['Lighting'] = get_yoy(today_record['Lighting'], ly_light)
            ly_row['Returns  退款'] = get_yoy(store_refund, manual_ly_refund) if manual_ly_refund > 0 else None
            
            # 净销售同比
            if manual_ly_refund > 0:
                ly_net_inc = ly_item_gmv - manual_ly_refund
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

            st.success("🎉 报表生成成功！")
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
