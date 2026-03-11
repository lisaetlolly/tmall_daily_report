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
st.set_page_config(page_title="天猫电商日报生成器", page_icon="📊", layout="centered")
st.title("📊 Tmall Store Daily Dashboard")
st.markdown("请在下方拖拽或点击上传 Excel 数据源。**（支持 xlsx / xls / 网页伪装xls，不强制全传）**")

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
    if any(x in t for x in['灯', 'light', 'lamp', 'portable']): return 'Lighting'
    if any(x in t for x in['椅', '桌', '柜', '沙发', 'stool', 'table', 'cabinet', 'cabine', 'chair', 'bench', 'desk', 'sofa']): return 'Furniture'
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
            if keyword in str(raw_df.columns):
                df = raw_df
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

def safe_get(df, col_names):
    if df is None or df.empty: return 0.0
    for c in col_names:
        if c in df.columns:
            return parse_money(df[c].iloc[0])
    return 0.0

# ==========================================
# UI: 文件上传区
# ==========================================
col1, col2 = st.columns(2)
with col1:
    file_order = st.file_uploader("1. 📥 今日订单底表", type=['xlsx', 'xls', 'csv'])
    file_store = st.file_uploader("3. 📥 店铺大盘表 (中央控制台)", type=['xlsx', 'xls', 'csv'])
with col2:
    file_item = st.file_uploader("2. 📥 今日单品生参", type=['xlsx', 'xls', 'csv'])
    file_ly = st.file_uploader("4. 📥 去年当日单品[选填]", type=['xlsx', 'xls', 'csv'])

# ==========================================
# 执行生成逻辑
# ==========================================
if st.button("⚡ 一键生成日报", type="primary", use_container_width=True):
    if not any([file_order, file_item, file_store]):
        st.warning("⚠️ 请至少上传一个文件！")
        st.stop()

    with st.spinner("🔄 正在飞速计算数据并校验同比逻辑..."):
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
            MONTH_NAME = d.strftime('%b') 
            WEEK_NUM = d.isocalendar()[1] 

            # --- 安全读取数据 ---
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

            # --- 提取大盘与配置数据 ---
            store_gmv = safe_get(sycm_store, ['支付金额'])
            store_demand = safe_get(sycm_store, ['下单金额'])
            store_refund = safe_get(sycm_store,['成功退款金额', '退款金额（完结时间）'])
            store_traffic = safe_get(sycm_store, ['访客数'])
            store_buyers = safe_get(sycm_store,['支付买家数'])
            store_cr = safe_get(sycm_store,['支付转化率'])
            if store_cr > 1: store_cr /= 100 
            new_followers = safe_get(sycm_store,['新增粉丝数', '关注店铺人数'])
            acc_followers = safe_get(sycm_store,['累计粉丝数'])

            TARGET_GMV_MONTH = safe_get(sycm_store,['全月GMV目标'])
            TARGET_NET_MONTH = safe_get(sycm_store,['全月Net目标'])
            TARGET_GMV_MTD = safe_get(sycm_store,['MTD_GMV目标'])
            TARGET_NET_MTD = safe_get(sycm_store,['MTD_Net目标'])
            TARGET_LIGHTING_MTD = safe_get(sycm_store,['MTD_Lighting目标'])
            TARGET_FURNITURE_MTD = safe_get(sycm_store,['MTD_Furniture目标'])
            TARGET_ACC_MTD = safe_get(sycm_store,['MTD_ACC目标'])
            EST_REST_MONTH_GMV = safe_get(sycm_store,['预估剩余GMV'])
            EST_REST_MONTH_NET = safe_get(sycm_store, ['预估剩余Net'])
            LY_WHOLE_MONTH_ACTUAL_GMV = safe_get(sycm_store, ['去年全月GMV'])
            LY_WHOLE_MONTH_ACTUAL_NET = safe_get(sycm_store,['去年全月Net'])
            MTD_REFUND_ACTUAL = safe_get(sycm_store,['本月累计退款'])

            # 🌟 新增：手动填入的"去年今日"真实大盘数据（用于终极精准同比）
            LY_STORE_GMV = safe_get(sycm_store,['去年今日GMV', '去年当日GMV'])
            LY_STORE_TRAFFIC = safe_get(sycm_store,['去年今日访客', '去年当日访客'])
            LY_STORE_BUYERS = safe_get(sycm_store, ['去年今日买家', '去年当日买家'])
            LY_STORE_UNITS = safe_get(sycm_store, ['去年今日件数', '去年当日件数'])

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
            mtd_net_sales = (mtd_gmv - MTD_REFUND_ACTUAL) / 1.13

            est_whole_month_gmv = mtd_gmv + EST_REST_MONTH_GMV
            est_whole_month_net = mtd_net_sales + EST_REST_MONTH_NET

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
            # Part 3: Followers Performance
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
            gross_demand = store_demand if store_demand > 0 else orders['买家实付金额'].sum()
            
            today_gmv_fallback = store_gmv if store_gmv > 0 else sycm_item['当日金额'].sum()
            today_auv = today_gmv_fallback / total_units if total_units else 0

            net_sales_tax_inc = today_gmv_fallback - store_refund
            net_sales_tax_excl = net_sales_tax_inc / 1.13

            # 单品聚合的品类 GMV (用于同比对齐)
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
            
            # 🌟 同比修复核心逻辑：苹果对苹果！
            # 如果人工配置了去年真实大盘，优先用；如果没配置，就用今年单品总和 vs 去年单品总和
            ly_gmv_base = LY_STORE_GMV if LY_STORE_GMV > 0 else ly_item_gmv
            today_gmv_base = today_gmv_fallback if LY_STORE_GMV > 0 else today_item_gmv
            
            ly_units_base = LY_STORE_UNITS if LY_STORE_UNITS > 0 else ly_item_units
            today_units_base = total_units if LY_STORE_UNITS > 0 else today_item_units
            
            ly_auv_base = ly_gmv_base / ly_units_base if ly_units_base > 0 else 0
            today_auv_base = today_gmv_base / today_units_base if today_units_base > 0 else 0

            ly_row['Traffic'] = get_yoy(store_traffic, LY_STORE_TRAFFIC) if LY_STORE_TRAFFIC > 0 else None
            ly_row['CR%'] = None  # 转化率同比意义不大，留空
            ly_row['Buyers'] = get_yoy(store_buyers, LY_STORE_BUYERS) if LY_STORE_BUYERS > 0 else None
            ly_row['ATV 客单价'] = None         
            ly_row['UPT 客单件'] = None         
            ly_row['AUV 件单价'] = get_yoy(today_auv_base, ly_auv_base)
            ly_row['Units Sold 件数'] = get_yoy(today_units_base, ly_units_base)
            ly_row['Gross Sales Demand 下单金额'] = None 
            ly_row['GMV （ 成交额）'] = get_yoy(today_gmv_base, ly_gmv_base)
            ly_row['ACC'] = get_yoy(today_record['ACC'], ly_acc)
            ly_row['Furniture'] = get_yoy(today_record['Furniture'], ly_furn)
            ly_row['Lighting'] = get_yoy(today_record['Lighting'], ly_light)
            ly_row['Returns  退款'] = None     
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
            # 在内存中写入 Excel
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
                        if col_num == 0:
                            worksheet.write(start_row, col_num, title, fmt_title)
                        else:
                            worksheet.write(start_row, col_num, "", fmt_title) 
                            
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
