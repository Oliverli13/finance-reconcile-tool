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
    page_title="财务智能核对系统 (最终版)", 
    layout="wide", 
    page_icon="🧬",
    initial_sidebar_state="expanded"
)

st.title("🧬 销售折让 vs ERP - 智能核对系统")
st.markdown("### ✨ 特性：冲销相减核对 | 智能场景匹配 | 包含清洗后明细导出")
st.markdown("---")

# ==========================================
# 2. 侧边栏：全局设置
# ==========================================
st.sidebar.header("1. 全局设置")

TASK_MODE = st.sidebar.radio("🛠️ 选择任务模式", ["暂估核对 (Provision)", "冲销核对 (Write-off)"])

SCENARIO_OPTIONS = [
    "商务一级", 
    "商务二级", 
    "其他折让", 
    "大健康新零售", 
    "大健康商超", 
    "大健康海外", 
    "澳诺", 
    "OTC医疗备案", 
    "自定义"
]
selected_scenario = st.sidebar.selectbox("📂 业务场景 / 筛选维度", SCENARIO_OPTIONS)

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
uploaded_match_file = st.sidebar.file_uploader("① 上传匹配表 (通用)", type=["xlsx"])

if uploaded_match_file:
    match_file_source = uploaded_match_file
    st.sidebar.success("✅ 使用上传的匹配表")
elif os.path.exists(DEFAULT_MATCH_FILE):
    match_file_source = DEFAULT_MATCH_FILE
    st.sidebar.success(f"✅ 自动加载本地: 匹配表.xlsx")
else:
    st.sidebar.warning(f"⚠️ 未找到本地匹配表，请上传。")

if TASK_MODE == "暂估核对 (Provision)":
    file_label_1 = "② 上传【折让暂估台账】"
    file_label_2 = "③ 上传【ERP导出表】"
else:
    file_label_1 = "② 上传【折让冲销总表】(包含所有场景)"
    file_label_2 = "③ 上传【ERP导出表】(对应当前场景)"

file_left = st.sidebar.file_uploader(file_label_1, type=["xlsx", "csv"])
file_right = st.sidebar.file_uploader(file_label_2, type=["xlsx", "csv"])

# ==========================================
# 3. 智能场景映射
# ==========================================
def get_search_keyword(scenario):
    MAPPING = {
        "商务一级": ["商务一级", "商务一级备案"],
        "商务二级": ["商务二级", "商务二级备案"],
        "其他折让": ["其他折扣", "其他折让"],
        "大健康新零售": ["大健康新零售", "大健康-新零售"],
        "大健康商超":   ["大健康-商超", "大健康商超"],
        "大健康海外":   ["大健康-海外", "大健康海外"],
        "OTC医疗备案": ["OTC-医疗备案", "OTC医疗备案", "OTC备案"],
        "澳诺":       ["OTX-澳诺备案", "澳诺", "OTX澳诺"]
    }
    return MAPPING.get(scenario, [scenario])

# ==========================================
# 4. 核心工具函数
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
        
        col_cust_code = next((c for c in df_rel.columns if '客户' in c and '编码' in c), None)
        col_cust_name = next((c for c in df_rel.columns if '名称' in c), None)
        
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
# 5. 数据处理逻辑
# ==========================================

def process_provision_data(df, valid_codes, valid_names, scenario):
    """暂估处理"""
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
    
    amt_col = next((c for c in df.columns if '传ERP金额' in c or ('ERP' in c and '金额' in c)), '传ERP金额')
    df['传ERP金额'] = clean_amount(df[amt_col]) if amt_col in df.columns else 0

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

def process_writeoff_discount(df, target_scenario):
    """冲销处理"""
    df.columns = df.columns.astype(str).str.strip()
    
    col_code = next((c for c in df.columns if '客户' in c and ('号' in c or '编码' in c)), None) 
    col_biz = next((c for c in df.columns if '业务线' in c), None) 
    col_amt = next((c for c in df.columns if '汇总' in c or '金额' in c), None) 
    col_type = next((c for c in df.columns if '类型' in c), None) 
    
    if not col_code or not col_biz or not col_amt or not col_type:
        st.error(f"❌ 冲销表列识别失败。\n需包含：客户号、业务线、汇总、类型更新。\n读取到: {list(df.columns)}")
        return pd.DataFrame()

    df['Code_Raw'] = df[col_code].apply(clean_str)
    df['Code_Clean'] = df['Code_Raw'].apply(strip_suffix) 
    df['业务线'] = df[col_biz].apply(clean_str)
    df['金额'] = clean_amount(df[col_amt])
    df['类型'] = df[col_type].apply(clean_str) 
    
    if target_scenario != "自定义":
        keywords = get_search_keyword(target_scenario)
        pattern = "|".join([k.replace('-', r'\-') for k in keywords])
        
        filtered_df = df[df['类型'].str.contains(pattern, na=False, case=False)]
        
        if filtered_df.empty:
            st.error(f"❌ 筛选失败！在冲销表的『{col_type}』列中，未找到包含以下关键词的数据：{keywords}")
            st.info(f"💡 当前『{col_type}』列包含的值有：")
            st.write(df['类型'].unique())
            return filtered_df
        else:
            df = filtered_df
            
    df['透视Key'] = df['Code_Clean'] + df['业务线']
    return df

def process_erp_generic(df, bus_map, valid_codes, valid_names, scenario, mode):
    df.columns = df.columns.astype(str).str.strip()
    
    if '交易对象编码' not in df.columns: st.error("ERP缺少 '交易对象编码'"); return pd.DataFrame()
    
    def clean_prefix(t):
        t = clean_str(t)
        if ':' in t: return t.split(':')[0] if len(t.split(':'))==1 else t.split(':')[-1].strip()
        return t

    df['原始交易编码'] = df['交易对象编码'].apply(clean_prefix)
    df['Code_Clean'] = df['原始交易编码'].apply(strip_suffix)
    df['原始交易名称'] = df['交易对象名称'].apply(clean_str) if '交易对象名称' in df.columns else ''
    
    df['帐户'] = df['帐户'].astype(str).str.strip()
    df['金额_借贷'] = clean_amount(df['本位币借方']) + clean_amount(df['本位币贷方'])
    
    def extract_bus(acc):
        if not acc: return None
        parts = acc.split('.')
        return next((p for p in parts if p.startswith(('A','B')) and len(p)>1), None)

    df['提取_业务线Code'] = df['帐户'].apply(extract_bus)
    df['业务线'] = df['提取_业务线Code'].apply(clean_str).map(bus_map) if bus_map else None
    
    if mode == "PROVISION" and scenario == "商务二级":
        df['标准名称'] = df['原始交易名称'].apply(normalize_brackets)
        df['透视Key'] = df['标准名称'] + df['业务线']
        df['是否关联方'] = df['标准名称'].apply(lambda x: x in valid_names)
    else:
        df['透视Key'] = df.apply(lambda x: x['Code_Clean'] + x['业务线'] if pd.notna(x['业务线']) else None, axis=1)
        if valid_codes:
            df['是否关联方'] = df['Code_Clean'].apply(lambda x: x in valid_codes)
        else:
            df['是否关联方'] = False
        
    return df

# ==========================================
# 6. 核对执行
# ==========================================

def perform_reconciliation(df_p, df_e, mode):
    key_col = '透视Key'
    
    if mode == "PROVISION":
        p_agg = df_p.dropna(subset=[key_col]).groupby(key_col).agg({
            '传ERP金额':'sum', '金额_不含税':'sum', '税额':'sum'
        }).rename(columns={'传ERP金额':'折让_总额', '金额_不含税':'折让_金额', '税额':'折让_税额'})
    else:
        p_data = df_p.dropna(subset=[key_col])
        if p_data.empty: return pd.DataFrame()
        p_agg = p_data.pivot_table(index=key_col, columns='类型', values='金额', aggfunc='sum', fill_value=0)
        p_agg['折让_汇总总计'] = p_agg.sum(axis=1)
        
    e_data = df_e.dropna(subset=[key_col])
    
    if mode == "PROVISION":
        targets = ['应收账款-应收账款（总账专用）', '主营业务收入-商品收入-贸易类', '应交税费-待转销项税额']
    else:
        targets = ['应收账款-应收账款（总账专用）']
    
    if '会计科目' in e_data.columns:
        e_data = e_data[e_data['会计科目'].isin(targets)]
        e_pivot = e_data.pivot_table(index=key_col, columns='会计科目', values='金额_借贷', aggfunc='sum', fill_value=0)
        for c in targets: 
            if c not in e_pivot.columns: e_pivot[c] = 0.0
    else:
        e_pivot = pd.DataFrame(columns=targets)
        
    if mode == "PROVISION":
        col_map = {
            '应收账款-应收账款（总账专用）': 'ERP_应收账款',
            '主营业务收入-商品收入-贸易类': 'ERP_主营收入',
            '应交税费-待转销项税额': 'ERP_销项税'
        }
    else:
        col_map = {
            '应收账款-应收账款（总账专用）': 'ERP_应收账款(总账)'
        }
    e_pivot.rename(columns=col_map, inplace=True)
    
    merged = pd.merge(p_agg, e_pivot, left_index=True, right_index=True, how='outer').fillna(0)
    
    if mode == "PROVISION":
        merged['核对_应收(0)'] = merged['折让_总额'] + merged['ERP_应收账款']
        merged['核对_收入(0)'] = merged['折让_金额'] + merged['ERP_主营收入']
        merged['核对_税额(0)'] = merged['折让_税额'] + merged['ERP_销项税']
        cols = ['折让_总额', 'ERP_应收账款', '核对_应收(0)', '折让_金额', 'ERP_主营收入', '核对_收入(0)', '折让_税额', 'ERP_销项税', '核对_税额(0)']
        return merged[[c for c in cols if c in merged.columns]]
    else:
        # 相减逻辑
        merged['核对_差额(0)'] = merged['折让_汇总总计'] - merged['ERP_应收账款(总账)']
        fixed_cols = ['折让_汇总总计', 'ERP_应收账款(总账)', '核对_差额(0)']
        other_cols = [c for c in merged.columns if c not in fixed_cols and 'ERP' not in c]
        return merged[fixed_cols + other_cols]

def apply_styles(df):
    def hl(val): 
        if isinstance(val, (int, float)) and abs(val) > 0.01:
            return 'background-color: #ffcccc; color: red'
        return ''
    chk = [c for c in df.columns if '核对' in c]
    return df.style.map(hl, subset=chk).format("{:,.2f}")

# ==========================================
# 7. 主程序执行入口
# ==========================================
if match_file_source and file_left and file_right:
    bus_map, valid_codes, valid_names, col_cust_name = load_mappings(match_file_source)
    
    if bus_map:
        try:
            df_l = pd.read_csv(file_left) if file_left.name.endswith('.csv') else pd.read_excel(file_left)
            h_row = 3
            df_r = pd.read_excel(file_right, header=h_row) if not file_right.name.endswith('.csv') else pd.read_csv(file_right, header=h_row)
            
            st.info(f"🚀 正在执行：{TASK_MODE} | 场景：{selected_scenario}")
            
            def render_safe_tab(df_main, source_p, source_e, key_prefix):
                col_f, _ = st.columns([1,4])
                show_diff = col_f.checkbox("🧨 只看差异", key=f"chk_{key_prefix}")
                
                df_view = df_main.copy()
                if show_diff:
                    chk = [c for c in df_view.columns if '核对' in c]
                    cond = df_view[chk].apply(lambda x: x.abs()>0.01).any(axis=1)
                    df_view = df_view[cond]
                
                df_t = add_total_row(df_view)
                
                valid_opts = [i for i in df_t.index if i != "=== 总计 ==="]
                c_sel, _ = st.columns([2,3])
                selected_key = c_sel.selectbox("🔍 选择查看明细:", ["(请选择)"] + list(valid_opts), key=f"sel_{key_prefix}")
                
                st.dataframe(apply_styles(df_t), use_container_width=True, height=500)
                
                if selected_key and selected_key != "(请选择)":
                    st.markdown(f"### 👇 明细: `{selected_key}`")
                    d1, d2 = st.columns(2)
                    with d1: st.caption("📘 折让系统"); st.dataframe(source_p[source_p['透视Key']==selected_key], use_container_width=True)
                    with d2: st.caption("📙 ERP系统"); st.dataframe(source_e[source_e['透视Key']==selected_key], use_container_width=True)

            if TASK_MODE == "暂估核对 (Provision)":
                df_p = process_provision_data(df_l, valid_codes, valid_names, selected_scenario)
                df_e = process_erp_generic(df_r, bus_map, valid_codes, valid_names, selected_scenario, "PROVISION")
                
                t1, t2, t3 = st.tabs(["👥 客户对账", "🏢 关联方对账", "📥 结果导出"])
                
                with t1:
                    res = perform_reconciliation(df_p, df_e, "PROVISION")
                    render_safe_tab(res, df_p, df_e, "cust")
                with t2:
                    df_p_r = df_p[df_p['是否关联方']==True]
                    df_e_r = df_e[df_e['是否关联方']==True]
                    res_rel = perform_reconciliation(df_p_r, df_e_r, "PROVISION")
                    if res_rel.empty: st.warning("无关联方数据")
                    else: render_safe_tab(res_rel, df_p, df_e, "rel")
                with t3:
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                        add_total_row(res).to_excel(w, sheet_name='客户对账')
                        if not res_rel.empty: add_total_row(res_rel).to_excel(w, sheet_name='关联方对账')
                        # 导出清洗后明细
                        df_p.to_excel(w, sheet_name='折让明细_清洗后', index=False)
                        df_e.to_excel(w, sheet_name='ERP明细_清洗后', index=False)
                    st.download_button("📥 下载暂估核对 (含清洗后明细)", out.getvalue(), "暂估核对.xlsx")

            else:
                # 冲销模式
                df_p = process_writeoff_discount(df_l, selected_scenario)
                if df_p.empty: st.stop()
                
                df_e = process_erp_generic(df_r, bus_map, valid_codes, None, selected_scenario, "WRITEOFF")
                res_wo = perform_reconciliation(df_p, df_e, "WRITEOFF")
                
                st.write(f"📊 数据行数: 折让 {len(df_p)} | ERP {len(df_e)}")
                render_safe_tab(res_wo, df_p, df_e, "wo")
                
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                    add_total_row(res_wo).to_excel(w, sheet_name='冲销核对')
                    # 导出清洗后明细 (统一名称)
                    df_p.to_excel(w, sheet_name='折让明细_清洗后', index=False)
                    df_e.to_excel(w, sheet_name='ERP明细_清洗后', index=False)
                st.download_button("📥 下载冲销核对 (含清洗后明细)", out.getvalue(), "冲销核对.xlsx")

        except Exception as e:
            st.error(f"处理错误: {e}")
            st.exception(e)
else:
    st.info("👈 请上传文件以开始")