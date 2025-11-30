import streamlit as st
import pandas as pd
import io
import xlsxwriter
import os
from datetime import datetime

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(
    page_title="财务自动核对系统 (旗舰点击版)", 
    layout="wide", 
    page_icon="🖱️",
    initial_sidebar_state="expanded"
)

st.title("🖱️ 销售折让 vs ERP - 智能核对系统")
st.markdown("### ✨ 特性：点击穿透 | 数据完整性监控 | 差异筛选 | 自动匹配")
st.markdown("---")

# ==========================================
# 2. 侧边栏
# ==========================================
st.sidebar.header("1. 任务设置")
SCENARIO_OPTIONS = [
    "商务一级", "商务二级", "其他折让", 
    "大健康新零售", "大健康商超", "大健康海外", 
    "澳诺", "OTC医疗备案", "自定义"
]
selected_scenario = st.sidebar.selectbox("📂 核对场景", SCENARIO_OPTIONS)

if selected_scenario == "商务二级":
    st.sidebar.warning("ℹ️ 逻辑：基于【名称】匹配")
else:
    st.sidebar.info("ℹ️ 逻辑：基于【编码】匹配")

current_month_str = datetime.now().strftime("%Y-%m")
match_month = st.sidebar.text_input("📅 核对月份", value=current_month_str)

st.sidebar.markdown("---")
st.sidebar.header("2. 数据上传")

DEFAULT_MATCH_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "匹配表.xlsx")
match_file_source = None
uploaded_match_file = st.sidebar.file_uploader("① 上传匹配表 (可选)", type=["xlsx"])

if uploaded_match_file:
    match_file_source = uploaded_match_file
    st.sidebar.success("✅ 使用上传的匹配表")
elif os.path.exists(DEFAULT_MATCH_FILE):
    match_file_source = DEFAULT_MATCH_FILE
    st.sidebar.success(f"✅ 自动加载本地: 匹配表.xlsx")
else:
    st.sidebar.warning(f"⚠️ 未找到本地匹配表，请上传。")

provision_file = st.sidebar.file_uploader("② 上传折让暂估台账", type=["xlsx", "csv"])
erp_file = st.sidebar.file_uploader("③ 上传ERP导出表", type=["xlsx", "csv"])

# ==========================================
# 3. 工具函数
# ==========================================

def clean_str(val):
    s = str(val).strip()
    if s == 'nan' or s == 'None': return ''
    if s.endswith('.0'): s = s[:-2]
    return s

def normalize_brackets(val):
    s = clean_str(val)
    return s.replace('（', '(').replace('）', ')')

def clean_amount(series):
    return pd.to_numeric(series, errors='coerce').fillna(0)

def strip_suffix(code):
    code = clean_str(code)
    if '-' in code:
        return code.split('-')[0].strip()
    return code

def add_total_row(df):
    if df.empty: return df
    df_out = df.copy()
    sum_row = df_out.sum(numeric_only=True)
    df_out.loc['=== 总计 ==='] = sum_row
    return df_out.fillna('')

@st.cache_data
def load_mappings(file_path_or_buffer):
    try:
        df_bus = pd.read_excel(file_path_or_buffer, sheet_name='业务线', header=None)
        bus_map = dict(zip(df_bus.iloc[:, 0].apply(clean_str), df_bus.iloc[:, 1].apply(clean_str)))
        
        df_rel = pd.read_excel(file_path_or_buffer, sheet_name='关联方')
        df_rel.columns = df_rel.columns.astype(str).str.strip()
        
        col_cust_code = None
        col_cust_name = None
        
        for c in df_rel.columns:
            if '客户' in c and '编码' in c: col_cust_code = c
            if '名称' in c: col_cust_name = c 
            
        if not col_cust_code:
            st.error("❌ 关联方表头识别失败！")
            return None, None, None, None
            
        valid_codes = set(df_rel[col_cust_code].apply(strip_suffix).unique())
        valid_names = set()
        if col_cust_name:
            valid_names = set(df_rel[col_cust_name].apply(normalize_brackets).unique())
        
        return bus_map, valid_codes, valid_names, col_cust_code
    except Exception as e:
        st.error(f"匹配表异常: {e}")
        return None, None, None, None

# ==========================================
# 4. 数据处理逻辑
# ==========================================

def process_provision(df, valid_codes, valid_names, scenario):
    df.columns = df.columns.astype(str).str.strip()
    col_code = next((c for c in df.columns if '一级客户编码' in c), None)
    col_name = next((c for c in df.columns if '一级客户名称' in c), None)
    
    if not col_code: 
        st.error("❌ 未找到【一级客户编码】")
        return pd.DataFrame()

    df['原始编码'] = df[col_code].apply(clean_str)
    df['原始名称'] = df[col_name].apply(clean_str) if col_name else ''
    if '业务线' not in df.columns: df['业务线'] = ''
    df['业务线'] = df['业务线'].apply(clean_str)
    
    amt_col = '传ERP金额'
    if amt_col not in df.columns:
        amt_col = next((c for c in df.columns if 'ERP' in c and '金额' in c), None)
    if not amt_col: 
        st.error("❌ 未找到金额列")
        return pd.DataFrame()
    df['传ERP金额'] = clean_amount(df[amt_col])

    if scenario == "商务二级":
        df['标准名称'] = df['原始名称'].apply(normalize_brackets)
        df['透视Key'] = df['标准名称'] + df['业务线']
        df['是否关联方'] = df['标准名称'].apply(lambda x: x in valid_names)
    else:
        df['Code_Clean'] = df['原始编码'].apply(strip_suffix)
        df['透视Key'] = df['Code_Clean'] + df['业务线']
        df['是否关联方'] = df['Code_Clean'].apply(lambda x: x in valid_codes)

    df['金额_不含税'] = (df['传ERP金额'] / 1.13).round(2)
    df['税额'] = (df['传ERP金额'] / 1.13 * 0.13).round(2)
    return df

def process_erp(df, bus_map, valid_codes, valid_names, scenario):
    df.columns = df.columns.astype(str).str.strip()
    
    if '交易对象编码' not in df.columns: st.error("ERP缺少 '交易对象编码'"); return pd.DataFrame()
    
    def clean_prefix(t):
        t = clean_str(t)
        if ':' in t: return t.split(':')[0] if len(t.split(':'))==1 else t.split(':')[-1].strip()
        return t

    df['原始交易编码'] = df['交易对象编码'].apply(clean_prefix)
    df['Code_Clean'] = df['原始交易编码'].apply(strip_suffix)
    if '交易对象名称' in df.columns:
        df['原始交易名称'] = df['交易对象名称'].apply(clean_str)
    else:
        df['原始交易名称'] = ''
    
    df['帐户'] = df['帐户'].astype(str).str.strip()
    df['金额_借贷'] = clean_amount(df['本位币借方']) + clean_amount(df['本位币贷方'])
    
    def extract_bus(acc):
        if not acc: return None
        parts = acc.split('.')
        return next((p for p in parts if p.startswith(('A','B')) and len(p)>1), None)

    df['提取_业务线Code'] = df['帐户'].apply(extract_bus)
    df['业务线'] = df['提取_业务线Code'].apply(clean_str).map(bus_map) if bus_map else None
    
    if scenario == "商务二级":
        df['标准名称'] = df['原始交易名称'].apply(normalize_brackets)
        df['透视Key'] = df.apply(lambda x: x['标准名称'] + x['业务线'] if pd.notna(x['业务线']) else None, axis=1)
        df['是否关联方'] = df['标准名称'].apply(lambda x: x in valid_names)
    else:
        df['透视Key'] = df.apply(lambda x: x['Code_Clean'] + x['业务线'] if pd.notna(x['业务线']) else None, axis=1)
        if valid_codes:
            df['是否关联方'] = df['Code_Clean'].apply(lambda x: x in valid_codes)
        else:
            df['是否关联方'] = False
        
    return df

def perform_reconciliation(df_p, df_e, filter_related=False):
    if filter_related:
        df_p = df_p[df_p['是否关联方'] == True]
        df_e = df_e[df_e['是否关联方'] == True]
        
    key_col = '透视Key'
    
    p_agg = df_p.dropna(subset=[key_col]).groupby(key_col).agg({
        '传ERP金额':'sum', '金额_不含税':'sum', '税额':'sum'
    }).rename(columns={'传ERP金额':'折让_价税合计', '金额_不含税':'折让_金额', '税额':'折让_税额'})
    
    targets = ['应收账款-应收账款（总账专用）', '主营业务收入-商品收入-贸易类', '应交税费-待转销项税额']
    e_data = df_e.dropna(subset=[key_col])
    if '会计科目' in e_data.columns:
        e_data = e_data[e_data['会计科目'].isin(targets)]
        e_pivot = e_data.pivot_table(index=key_col, columns='会计科目', values='金额_借贷', aggfunc='sum', fill_value=0)
        for c in targets: 
            if c not in e_pivot.columns: e_pivot[c] = 0.0
    else:
        e_pivot = pd.DataFrame(columns=targets)
            
    col_map = {
        '应收账款-应收账款（总账专用）': 'ERP_应收账款',
        '主营业务收入-商品收入-贸易类': 'ERP_主营收入',
        '应交税费-待转销项税额': 'ERP_销项税'
    }
    e_pivot.rename(columns=col_map, inplace=True)
    
    merged = pd.merge(p_agg, e_pivot, left_index=True, right_index=True, how='outer').fillna(0)
    
    merged['核对_应收(0)'] = merged['折让_价税合计'] + merged['ERP_应收账款']
    merged['核对_收入(0)'] = merged['折让_金额'] + merged['ERP_主营收入']
    merged['核对_税额(0)'] = merged['折让_税额'] + merged['ERP_销项税']
    
    cols = [
        '折让_价税合计', 'ERP_应收账款', '核对_应收(0)',
        '折让_金额', 'ERP_主营收入', '核对_收入(0)',
        '折让_税额', 'ERP_销项税', '核对_税额(0)'
    ]
    return merged[[c for c in cols if c in merged.columns]]

def apply_styles(df):
    def hl(val): 
        if isinstance(val, (int, float)) and abs(val) > 0.01:
            return 'background-color: #ffcccc; color: red'
        return ''
    chk = [c for c in df.columns if '核对' in c]
    return df.style.map(hl, subset=chk).format("{:,.2f}")

# ==========================================
# 5. 主程序执行入口
# ==========================================
if match_file_source and provision_file and erp_file:
    bus_map, valid_codes, valid_names, col_cust_name = load_mappings(match_file_source)
    
    if bus_map:
        try:
            prov_raw = pd.read_csv(provision_file) if provision_file.name.endswith('.csv') else pd.read_excel(provision_file)
            h_row = 3
            erp_raw = pd.read_excel(erp_file, header=h_row) if not erp_file.name.endswith('.csv') else pd.read_csv(erp_file, header=h_row)
            
            st.info(f"📊 数据清洗监控 | 当前场景: **{selected_scenario}**")
            
            df_p = process_provision(prov_raw, valid_codes, valid_names, selected_scenario)
            df_e = process_erp(erp_raw, bus_map, valid_codes, valid_names, selected_scenario)
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("折让原始", len(prov_raw))
            c2.metric("折让清洗", len(df_p))
            c3.metric("ERP原始", len(erp_raw))
            c4.metric("ERP清洗", len(df_e))
            
            t1, t2, t3 = st.tabs(["👥 客户对账", "🏢 关联方对账", "📥 结果导出"])
            
            # === 通用渲染函数 (恢复点击交互) ===
            def render_click_tab(df_main, source_p, source_e, key_prefix):
                # A. 筛选
                col_filt, _ = st.columns([1, 4])
                show_diff = col_filt.checkbox("🧨 只看差异", key=f"chk_{key_prefix}")
                
                df_view = df_main.copy()
                if show_diff:
                    # 只要任意一列核对值不为0，就保留
                    chk_cols = [c for c in df_view.columns if '核对' in c]
                    condition = df_view[chk_cols].apply(lambda x: x.abs() > 0.01).any(axis=1)
                    df_view = df_view[condition]
                
                # B. 合计
                df_total = add_total_row(df_view)
                
                # C. 点击表格
                st.write("👉 **点击** 下方表格的任意行，查看明细：")
                selection = st.dataframe(
                    apply_styles(df_total), 
                    use_container_width=True, 
                    height=500,
                    on_select="rerun",  # 恢复点击功能
                    selection_mode="single-row",
                    key=f"grid_{key_prefix}"
                )
                
                # D. 穿透展示
                if selection.selection["rows"]:
                    idx = selection.selection["rows"][0]
                    sel_key = df_total.index[idx]
                    
                    if sel_key != "=== 总计 ===":
                        st.markdown(f"### 👇 明细数据: `{sel_key}`")
                        d1, d2 = st.columns(2)
                        
                        dp = source_p[source_p['透视Key'] == sel_key]
                        de = source_e[source_e['透视Key'] == sel_key]
                        
                        with d1:
                            st.caption("📘 折让系统")
                            st.dataframe(dp, use_container_width=True)
                        with d2:
                            st.caption("📙 ERP系统")
                            st.dataframe(de, use_container_width=True)
                    else:
                        st.info("合计行无法穿透。")

            # --- Tab 1 ---
            with t1:
                res_cust = perform_reconciliation(df_p, df_e, False)
                render_click_tab(res_cust, df_p, df_e, "cust")
                
            # --- Tab 2 ---
            with t2:
                res_rel = perform_reconciliation(df_p, df_e, True)
                if res_rel.empty:
                    st.warning("⚠️ 无关联方数据")
                else:
                    render_click_tab(res_rel, df_p, df_e, "rel")
            
            # --- Tab 3 ---
            with t3:
                fname = f"{selected_scenario}_{match_month}_核对结果.xlsx"
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                    res_cust_final = add_total_row(res_cust)
                    res_cust_final.to_excel(w, sheet_name='客户对账')
                    
                    if not res_rel.empty:
                        res_rel_final = add_total_row(res_rel)
                        res_rel_final.to_excel(w, sheet_name='关联方对账')
                    
                    df_p.to_excel(w, sheet_name='折让明细_清洗后', index=False)
                    df_e.to_excel(w, sheet_name='ERP明细_清洗后', index=False)
                st.download_button("📥 下载完整 Excel (含合计行)", out.getvalue(), fname, mime="application/vnd.ms-excel")
                
        except Exception as e:
            st.error(f"处理错误: {e}")
            st.exception(e)
else:
    st.info("👈 请上传文件以开始")