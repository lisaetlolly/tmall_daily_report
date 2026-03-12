import streamlit as st
import pandas as pd
import numpy as np
import io

# 设置页面布局为宽屏
st.set_page_config(page_title="生意参谋看板 - HAY Ranking", layout="wide")

st.title("📊 生意参谋商品数据看板 - HAY Ranking")

# 1. 侧边栏：文件上传区
st.sidebar.header("📂 数据上传")
file_curr = st.sidebar.file_uploader("1. 上传【今年当月】生意参谋数据 (Excel/CSV)", type=["xlsx", "xls", "csv"])
file_last = st.sidebar.file_uploader("2. 上传【去年当月】生意参谋数据 (Excel/CSV)", type=["xlsx", "xls", "csv"])
file_map = st.sidebar.file_uploader("3. 上传【分类映射表】 (Excel/CSV)", type=["xlsx", "xls", "csv"])

# --- 核心数据清洗辅助函数 ---
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

# --- Excel 完美还原排版生成函数 ---
def generate_excel_dashboard(df_ttl, df_fav, dict_cats, df_return, cat_names):
    output = io.BytesIO()
    # 使用 xlsxwriter 引擎进行精细化排版
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # 定义各种单元格样式 (复刻图片风格)
        title_fmt = workbook.add_format({'bold': True, 'font_size': 12, 'bg_color': '#CDE9F5', 'valign': 'vcenter'}) # 浅蓝色 HAY Ranking
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#CDE9F5', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        num_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0'})
        pct_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0%'})
        
        # 为每个三大类目生成一个单独的Sheet页面，完美展示四宫格
        for cat_name in cat_names:
            sheet_name = f"HAY Ranking - {cat_name}"[:31] # Excel要求sheet名不超过31字
            worksheet = workbook.add_worksheet(sheet_name)
            
            # 写入左上角大标题
            worksheet.write('A1', 'HAY Ranking', title_fmt)
            worksheet.set_row(0, 20) # 标题行高
            
            # --- 辅助写入单表函数 ---
            def write_table_to_excel(df, start_row, start_col):
                # 写入表头
                for col_idx, col_name in enumerate(df.columns):
                    worksheet.write(start_row, start_col + col_idx, col_name, header_fmt)
                # 写入数据
                for row_idx, row in enumerate(df.values):
                    for col_idx, val in enumerate(row):
                        col_name = df.columns[col_idx]
                        fmt = cell_fmt
                        
                        # 动态判断数字格式
                        if isinstance(val, (int, float)):
                            if 'Share%' in col_name or 'YOY' in col_name or '率' in col_name:
                                fmt = pct_fmt
                            elif 'Value' in col_name or 'QTY' in col_name or '人数' in col_name:
                                fmt = num_fmt
                                
                        if pd.isna(val):
                            worksheet.write(start_row + 1 + row_idx, start_col + col_idx, "-", cell_fmt)
                        else:
                            worksheet.write(start_row + 1 + row_idx, start_col + col_idx, val, fmt)
            
            # 1. 左上：全店 TTL Rank
            write_table_to_excel(df_ttl, start_row=2, start_col=0) # 从 A3 开始
            
            # 2. 右上：收藏加购 Rank
            write_table_to_excel(df_fav, start_row=2, start_col=8) # 从 I3 开始
            
            # 3. 左下：类目 Rank (动态替换列名以匹配图片)
            cat_df = dict_cats[cat_name].copy()
            cat_df.rename(columns={'Rank': f'{cat_name}\nRank', 'Share% of Category': f'Share% of\n{cat_name}'}, inplace=True)
            write_table_to_excel(cat_df, start_row=16, start_col=0) # 从 A17 开始
            
            # 4. 右下：退款 Rank
            write_table_to_excel(df_return, start_row=16, start_col=8) # 从 I17 开始
            
            # --- 调整列宽让排版更美观 ---
            worksheet.set_column('A:A', 8)   # Rank
            worksheet.set_column('B:B', 20)  # Product
            worksheet.set_column('C:C', 8)   # Picture
            worksheet.set_column('D:G', 12)  # Value, QTY 等
            worksheet.set_column('H:H', 2)   # 中间空白隔离带
            worksheet.set_column('I:I', 8)   # 右侧 Rank
            worksheet.set_column('J:J', 20)  # 右侧 Product
            worksheet.set_column('K:K', 8)   # 右侧 Picture
            worksheet.set_column('L:M', 12)  # 人数等
            
    return output.getvalue()

# ----------------------------

if file_curr and file_last and file_map:
    with st.spinner('正在进行精密计算与排版...'):
        
        # 数据读取与合并
        df_curr = clean_id(load_data(file_curr))
        df_last = clean_id(load_data(file_last))
        df_map = clean_id(load_data(file_map))

        if '支付金额' in df_last.columns:
            df_last_sales = df_last[['商品ID', '支付金额']].rename(columns={'支付金额': '去年支付金额'})
        else:
            df_last_sales = pd.DataFrame(columns=['商品ID', '去年支付金额'])

        df_merged = pd.merge(df_curr, df_last_sales, on='商品ID', how='left')
        
        if '一级' in df_map.columns:
            df_merged = pd.merge(df_merged, df_map[['商品ID', '一级']], on='商品ID', how='left')
        else:
            df_merged['一级'] = '未分类'
        df_merged['一级'] = df_merged['一级'].fillna('未分类')

        numeric_columns =['支付金额', '去年支付金额', '支付件数', '商品收藏人数', '商品加购人数', '商品访客数', '成功退款金额']
        for col in numeric_columns:
            if col in df_merged.columns:
                df_merged[col] = to_numeric_col(df_merged[col])
            else:
                df_merged[col] = 0.0 

        # 核心指标计算
        total_store_value = df_merged['支付金额'].sum()
        total_store_refund = df_merged['成功退款金额'].sum() if df_merged['成功退款金额'].sum() > 0 else 1

        df_merged['Value'] = df_merged['支付金额']
        df_merged['QTY'] = df_merged['支付件数']
        df_merged['Share% of TTL'] = np.where(total_store_value > 0, df_merged['Value'] / total_store_value, 0)
        df_merged['YOY'] = np.where(df_merged['去年支付金额'] > 0, 
                                   (df_merged['Value'] - df_merged['去年支付金额']) / df_merged['去年支付金额'], 
                                   np.nan)
        df_merged['收加人数'] = df_merged['商品收藏人数'] + df_merged['商品加购人数']
        df_merged['收加率%'] = np.where(df_merged['商品访客数'] > 0, df_merged['收加人数'] / df_merged['商品访客数'], 0)
        df_merged['Picture'] = ""
        df_merged['Product'] = df_merged['商品名称'] if '商品名称' in df_merged.columns else df_merged['商品ID']
        df_merged['Return Value'] = df_merged['成功退款金额']
        df_merged['Return Share%'] = df_merged['Return Value'] / total_store_refund

        # --- 为 Excel 准备纯净数据 (保持纯数字类型) ---
        # 1. TTL
        raw_ttl = df_merged.sort_values(by='Value', ascending=False).head(10)[['Product', 'Picture', 'Value', 'QTY', 'Share% of TTL', 'YOY']].copy()
        raw_ttl.insert(0, 'TTL Rank', range(1, len(raw_ttl) + 1))
        
        # 2. Fav
        raw_fav = df_merged.sort_values(by='收加人数', ascending=False).head(10)[['Product', 'Picture', '收加人数', '收加率%']].copy()
        raw_fav.insert(0, 'Rank', range(1, len(raw_fav) + 1))
        
        # 3. Category (存入字典)
        category_sales = df_merged.groupby('一级')['Value'].sum().sort_values(ascending=False)
        top_3_categories = [cat for cat in category_sales.index if cat != '未分类'][:3]
        
        raw_cats = {}
        for cat in top_3_categories:
            c_df = df_merged[df_merged['一级'] == cat].copy()
            c_total = c_df['Value'].sum()
            c_df['Share% of Category'] = np.where(c_total > 0, c_df['Value'] / c_total, 0)
            c_top10 = c_df.sort_values(by='Value', ascending=False).head(10)[['Product', 'Picture', 'Value', 'QTY', 'Share% of Category', 'YOY']].copy()
            c_top10.insert(0, 'Rank', range(1, len(c_top10) + 1))
            raw_cats[cat] = c_top10
            
        # 4. Return
        raw_return = df_merged.sort_values(by='Return Value', ascending=False).head(10)[['Product', 'Picture', 'Return Value', 'Return Share%']].copy()
        raw_return.rename(columns={'Return Value': 'Returned Value', 'Return Share%': 'Share% of TTL'}, inplace=True)
        raw_return.insert(0, 'HAY Rank', range(1, len(raw_return) + 1))

        # 🚀 触发 Excel 生成
        excel_data = generate_excel_dashboard(raw_ttl, raw_fav, raw_cats, raw_return, top_3_categories)

        # ---------------- 页面展示 ----------------
        
        # 顶部增加显眼的下载按钮
        st.success("✅ 数据计算及 Excel 原版排版生成完成！")
        st.download_button(
            label="📥 一键下载原版 Excel 报表",
            data=excel_data,
            file_name="HAY_Ranking_Dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary" # 让按钮变红色/主题色，显眼一点
        )
        st.markdown("*(⬇️ 下方为网页预览，下载的 Excel 已经为您排版成 2x2 四宫格样式)*")
        st.markdown("---")

        # --- 网页格式化展现 (带逗号和百分号) ---
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
            st.subheader("🏆 全店销售 Top 10 (TTL Rank)")
            st.dataframe(fmt_display(raw_ttl), use_container_width=True)
        with col2:
            st.subheader("❤️ 收藏加购 Top 10")
            st.dataframe(fmt_display(raw_fav), use_container_width=True)

        st.markdown("---")
        col3, col4 = st.columns(2)
        with col3:
            st.subheader("📦 三大核心类目 Top 10")
            if len(top_3_categories) > 0:
                tabs = st.tabs([f"{cat} Rank" for cat in top_3_categories])
                for i, cat in enumerate(top_3_categories):
                    with tabs[i]:
                        st.dataframe(fmt_display(raw_cats[cat]), use_container_width=True)
        with col4:
            st.subheader("↩️ 退货 Top 10 (按退款金额)")
            st.dataframe(fmt_display(raw_return), use_container_width=True)

else:
    st.info("👈 请在左侧依次上传：1.今年当月数据 2.去年当月数据 3.分类映射表")
