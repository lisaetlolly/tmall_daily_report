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
st.markdown("请在下方拖拽或点击上传今天的 4 个 Excel 数据源。无需更改文件名。")

# 本地历史记录文件路径
HISTORY_FILE = 'dashboard_history.csv'

# ==========================================
# 核心函数 (适配网页流格式)
# ==========================================
def parse_money(s):
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
    raw_df = pd.read_excel(file_obj, header=None, nrows=20)
    header_idx = -1
    for idx, row in raw_df.iterrows():
        if keyword in str(row.values):
            header_idx = idx
            break
    if header_idx == -1: return None
    file_obj.seek(0)
    df = pd.read_excel(file_obj, header=header_idx)
    df.columns =[str(c).strip() for c in df.columns]
    return df

def extract_date_from_excel(file_obj):
    if file_obj is None: return None
    try:
        file_obj.seek(0)
        raw_df = pd.read_excel(file_obj, header=None, nrows=10)
        for _, row in raw_df.iterrows():
            text = ' '.join([str(v) for v in row.values if pd.notna(v)])
            match = re.search(r'(202\d-\d{2}-\d{2})', text)
            if match: return match.group(1)
    except: pass
    return None

# ==========================================
# UI: 文件上传区
# ==========================================
col1, col2 = st.columns(2)
with col1:
    file_order = st.file_uploader("1. 📥 今日订单底表", type=['xlsx', 'xls'])
    file_store = st.file_uploader("3. 📥 店铺大盘表 (含目标预估)", type=['xlsx', 'xls'])
with col2:
    file_item = st.file_uploader("2. 📥 今日单品生参", type=['xlsx', 'xls'])
    file_ly = st.file_uploader("4. 📥 去年当日单品 (选填)", type=['xlsx', 'xls'])

# ==========================================
# 执行生成逻辑
# ==========================================
if st.button("⚡ 一键生成日报", type="primary", use_container_width=True):
    if not file_order or not file_item or not file_store:
        st.warning("⚠️ 警告：前 3 个文件必须上传！")
        st.stop()

    with st.spinner("🔄 正在飞速计算数据，请稍候..."):
        try:
            # --- 日期识别 ---
            auto_date_str = extract_date_from_excel(file_store)
            if not auto_date_str: auto_date_str = extract_date_from_excel(file_item)

            if auto_date_str:
                d = datetime.datetime.strptime(auto_date_str, '%Y-%m-%d')
                st.success(f"✅ 成功识别数据日期: {auto_date_str}")
            else:
                d = datetime.date.today() - datetime.timedelta(days=1)
                st.info("⚠️ 无法提取日期，默认按昨天计算。")

            DATE_STR = f"{d.month}/{d.day}"
            MONTH_NAME = d.strftime('%b') 
            WEEK_NUM = d.isocalendar()[1] 

            # --- 读取数据 ---
            orders = read_excel_smart(file_order, '商品标题')
            sycm_item = read_excel_smart(file_item, '支付金额')
            sycm_store = read_excel_smart(file_store, '支付金额')
            ly_sycm = read_excel_smart(file_ly, '支付金额') if file_ly else None

            if orders is None or sycm_item is None or sycm_store is None:
                st.error("❌ 读取失败：表格格式不对，请检查表头。")
                st.stop()

            # --- 提取动态大盘数据 ---
            store_gmv = parse_money(sycm_store.get('支付金额', pd.Series([0]))[0])
            store_demand = parse_money(sycm_store.get('下单金额', pd.Series([0]))[0]) 
            store_refund = parse_money(sycm_store.get('成功退款金额', pd.Series([0]))[0])
            store_traffic = parse_money(sycm_store.get('访客数', pd.Series([0]))[0])
            store_buyers = parse_money(sycm_store.get('支付买家数', pd.Series([0]))[0])
            store_cr = parse_money(sycm_store.get('支付转化率', pd.Series([0]))[0])
            if store_cr > 1: store_cr = store_cr / 100 
            new_followers = parse_money(sycm_store.get('新增粉丝数', pd.Series([0]))[0])
            acc_followers = parse_money(sycm_store.get('累计粉丝数', pd.Series([0]))[0])

            # --- 提取人工配置目标 (从 sycm_store) ---
            TARGET_GMV_MONTH = parse_money(sycm_store.get('全月GMV目标', pd.Series([0]))[0])
            TARGET_NET_MONTH = parse_money(sycm_store.get('全月Net目标', pd.Series([0]))[0])
            TARGET_GMV_MTD = parse_money(sycm_store.get('MTD_GMV目标', pd.Series([0]))[0])
            TARGET_NET_MTD = parse_money(sycm_store.get('MTD_Net目标', pd.Series([0]))[0])
            TARGET_LIGHTING_MTD = parse_money(sycm_store.get('MTD_Lighting目标', pd.Series([0]))[0])
            TARGET_FURNITURE_MTD = parse_money(sycm_store.get('MTD_Furniture目标', pd.Series([0]))[0])
            TARGET_ACC_MTD = parse_money(sycm_store.get('MTD_ACC目标', pd.Series([0]))[0])
            EST_REST_MONTH_GMV = parse_money(sycm_store.get('预估剩余GMV', pd.Series([0]))[0])
            EST_REST_MONTH_NET = parse_money(sycm_store.get('预估剩余Net', pd.Series([0]))[0])
            LY_WHOLE_MONTH_ACTUAL_GMV = parse_money(sycm_store.get('去年全月GMV', pd.Series([0]))[0])
            LY_WHOLE_MONTH_ACTUAL_NET = parse_money(sycm_store.get('去年全月Net', pd.Series([0]))[0])
            MTD_REFUND_ACTUAL = parse_money(sycm_store.get('本月累计退款', pd.Series([0]))[0])

            # --- 提取去年同比 ---
            if ly_sycm is not None:
                ly_sycm['Category'] = ly_sycm['商品名称'].apply(categorize_item)
                ly_gmv = ly_sycm.get('支付金额', pd.Series([0])).apply(parse_money).sum()
                ly_units = ly_sycm.get('支付件数', pd.Series([0])).apply(parse_money).sum()
                ly_cat = ly_sycm.groupby('Category')['支付金额'].apply(lambda x: x.apply(parse_money).sum())
                ly_acc = ly_cat.get('ACC', 0)
                ly_furn = ly_cat.get('Furniture', 0)
                ly_light = ly_cat.get('Lighting', 0)
            else:
                ly_gmv = ly_units = ly_acc = ly_furn = ly_light = 0

            # ==========================================
            # Part 1: MTD
            # ==========================================
            mtd_gmv = sycm_item.get('月累计支付金额', sycm_item.get('支付金额')).apply(parse_money).sum()
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
            sycm_item['月累计金额'] = sycm_item.get('月累计支付金额', sycm_item.get('支付金额')).apply(parse_money)
            sycm_item['月累计件数'] = sycm_item.get('月累计支付件数', sycm_item.get('支付件数')).apply(parse_money)
            sycm_item['当日金额'] = sycm_item.get('支付金额').apply(parse_money)

            cat_group = sycm_item.groupby('Category').agg(
                Daily_GMV=('当日金额', 'sum'),
                MTD_Actual=('月累计金额', 'sum'),
                MTD_Units=('月累计件数', 'sum')
            ).reset_index()

            target_dict = {'Lighting': TARGET_LIGHTING_MTD, 'Furniture': TARGET_FURNITURE_MTD, 'ACC': TARGET_ACC_MTD}
            cat_group['MTD Target'] = cat_group['Category'].map(target_dict)
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
                'MTD contribution%': 1.0,
                'MTD AUV': df_p2['MTD_Actual'].sum() / df_p2['MTD_Units'].sum() if df_p2['MTD_Units'].sum() else 0,
                'MTD_Units': df_p2['MTD_Units'].sum()
            }])
            df_p2 = pd.concat([df_p2, total_row], ignore_index=True)

            # ==========================================
            # Part 3: Followers Performance
            # ==========================================
            df_p3 = pd.DataFrame({
                f'Updated {DATE_STR}': ['New Followers', 'Accumulated Followers'],
                'Data': [new_followers, acc_followers]
            })

            # ==========================================
            # Part 4: Daily Dashboard
            # ==========================================
            orders['购买数量'] = orders.get('购买数量', pd.Series([0])).apply(parse_money)
            total_units = orders['购买数量'].sum()
            gross_demand = store_demand if store_demand > 0 else orders.get('买家实付金额', pd.Series([0])).apply(parse_money).sum()
            today_auv = store_gmv / total_units if total_units else 0

            net_sales_tax_inc = store_gmv - store_refund
            net_sales_tax_excl = net_sales_tax_inc / 1.13

            today_record = {
                'Date': DATE_STR,
                'Traffic': store_traffic,
                'CR%': store_cr,
                'Buyers': store_buyers,
                'ATV 客单价': store_gmv / store_buyers if store_buyers else 0,
                'UPT 客单件': total_units / store_buyers if store_buyers else 0,
                'AUV 件单价': today_auv,
                'Units Sold 件数': total_units,
                'Gross Sales Demand 下单金额': gross_demand,
                'GMV （ 成交额）': store_gmv,
                'ACC': df_p2.loc[df_p2['Category']=='ACC', f'GMV {DATE_STR}'].sum(),
                'Furniture': df_p2.loc[df_p2['Category']=='Furniture', f'GMV {DATE_STR}'].sum(),
                'Lighting': df_p2.loc[df_p2['Category']=='Lighting', f'GMV {DATE_STR}'].sum(),
                'Returns  退款': store_refund,
                'Net sales（含税）': net_sales_tax_inc,
                'Net sales（去税）': net_sales_tax_excl
            }
            df_today = pd.DataFrame([today_record])

            # 处理历史数据保存 (写入本地)
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
            ly_row['Traffic'] = None          
            ly_row['CR%'] = None              
            ly_row['Buyers'] = None           
            ly_row['ATV 客单价'] = None         
            ly_row['UPT 客单件'] = None         
            ly_row['AUV 件单价'] = get_yoy(today_auv, ly_gmv/ly_units if ly_units else 0)
            ly_row['Units Sold 件数'] = get_yoy(total_units, ly_units)
            ly_row['Gross Sales Demand 下单金额'] = None 
            ly_row['GMV （ 成交额）'] = get_yoy(store_gmv, ly_gmv)
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
            orders['商品ID'] = orders.get('商品ID', pd.Series([])).astype(str).str.replace(r'\.0$', '', regex=True)

            def extract_spu(title):
                if pd.isna(title): return ''
                match = re.findall(r'[A-Za-z][A-Za-z0-9\s\-_/]{3,}', str(title))
                if match:
                    longest = max([m.strip() for m in match if m.strip()], key=len)
                    cleaned = re.sub(r'\bHAY\b', '', longest, flags=re.IGNORECASE).strip()
                    return cleaned if cleaned else str(title)
                return str(title)
            orders['Description'] = orders.get('商品标题', pd.Series([''])).apply(extract_spu)

            def join_unique(x):
                return '\n'.join(sorted(set([str(i).strip() for i in x if pd.notna(i) and str(i).strip() != 'nan'])))
            def clean_color(text):
                if pd.isna(text): return ''
                s = str(text).replace('：', ':').replace('；', ';')
                parts = s.split(';')
                cleaned =[p.split(':', 1)[1].strip() if ':' in p else p.strip() for p in parts if p.strip()]
                return ' '.join(cleaned)
            orders['商品属性'] = orders.get('商品属性', pd.Series([''])).apply(clean_color)

            bestsellers = orders.groupby('商品ID').agg(
                Gross_Sales=('买家实付金额', 'sum'),
                Units=('购买数量', 'sum'),
                SKU=('商家编码', join_unique),
                Colour=('商品属性', join_unique),
                Description=('Description', 'first')
            ).reset_index()

            bestsellers = bestsellers.sort_values('Gross_Sales', ascending=False).head(15).reset_index(drop=True)
            bestsellers['Contribution%'] = bestsellers['Gross_Sales'] / store_gmv if store_gmv > 0 else 0
            bestsellers['No.'] = bestsellers.index + 1
            bestsellers['Pictures'] = ''
            df_p5 = bestsellers[['No.', 'SKU', 'Description', 'Colour', 'Pictures', 'Gross_Sales', 'Units', 'Contribution%']]

            # ==========================================
            # 在内存中写入 Excel (用于网页下载)
            # ==========================================
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet('Daily Dashboard')
                
                fmt_title = workbook.add_format({'bold': True, 'bg_color': '#000000', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
                fmt_header = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'center'})
                fmt_num = workbook.add_format({'num_format': '#,##0', 'border': 1, 'align': 'center'})
                fmt_float = workbook.add_format({'num_format': '0.00', 'border': 1, 'align': 'center'}) 
                fmt_pct_int = workbook.add_format({'num_format': '0%', 'border': 1, 'align': 'center'})    
                fmt_pct_cr = workbook.add_format({'num_format': '0.00%', 'border': 1, 'align': 'center'}) 
                fmt_text = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})

                current_row = 0

                def write_block(title, df, start_row):
                    worksheet.merge_range(start_row, 0, start_row, len(df.columns)-1, title, fmt_title)
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
            
            # 提供下载按钮
            st.download_button(
                label="📥 下载 Tmall_Daily_Dashboard.xlsx",
                data=output,
                file_name=f"Tmall_Daily_Dashboard_{DATE_STR.replace('/','')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        except Exception as e:
            st.error(f"❌ 程序发生错误: {e}")
            
