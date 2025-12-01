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
    page_title="财务智能核对系统 (双模式版)", 
    layout="wide", 
    page_icon="💹",
    initial_sidebar_state="expanded"
)

st.title("💹 销售折让 vs ERP - 智能核对系统")

# ==========================================
# 2. 侧边栏：全局设置
# ==========================================
st.sidebar.header("1. 全局设置")

# 【新增】任务模式切换
TASK_MODE = st.sidebar.radio("🛠️ 选择任务模式", ["暂估核对 (Provision)", "冲销核对 (Write-off)"])

# 场景选择 (通用)
# 注意：冲销核对时，这个选项用于筛选折让总表
SCENARIO_OPTIONS = [
    "商务一级", "商务二级", "其他折让", 
    "大健康新零售", "大健康商超", "大健康海外", 
    "澳诺", "OTC医疗备案", "自定义",
    "OTC-医疗备案", "OTX-澳诺备案", "商务二级备案", "商务一级备案", "其他折扣" # 补充冲销场景
]
selected_scenario = st.sidebar.selectbox("📂 业务场景 / 筛选维度", SCENARIO_OPTIONS)

# 自动月份
current_month_str = datetime.now().strftime("%Y-%m")
match_month = st.sidebar.text_input("📅 核对月份", value=current_month_str)

st.sidebar.markdown("---")
st.sidebar.header("2. 数据上传")

# 匹配表 (通用)
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

# 根据模式显示不同的上传框
if TASK_MODE == "暂估核对 (Provision)":
    file_label_1 = "② 上传【折让暂估台账】"
    file_label_2 = "③ 上传【ERP导出表】"
else:
    file_label_1 = "② 上传【折让冲销总表】(包含所有场景)"
    file_label_2 = "③ 上传【ERP导出表】(对应当前场景)"

file_left = st.sidebar.file_uploader(file_label_1, type=["xlsx", "csv"])
file_right = st.sidebar.file_uploader(file_label_2, type=["xlsx", "csv"])

# ==========================================
# 3. 核心工具函数
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
    """
    核心：去除 -00 后缀
    解决 A0686929-00OTC 与 A0686929OTC 不匹配的问题
    """
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
        # 业务线
        df_bus = pd.read_excel(file_path_or_buffer, sheet_name='业务线', header=None)
        bus_map = dict(zip(df_bus.iloc[:, 0].apply(clean_str), df_bus.iloc[:, 1].apply(clean_str)))
        
        # 关联方
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
        
        return bus_map, valid_codes, valid_names
    except Exception as e:
        st.error(f"匹配表异常: {e}")
        return None, None, None

# ==========================================
# 4. 数据处理逻辑 - 暂估模式 (Provision)
# ==========================================

def process_provision_data(df, valid_codes, valid_names, scenario):
    """暂估-折让数据处理"""
    df.columns = df.columns.astype(str).str.strip()
    col_code = next((c for c in df.columns if '一级客户编码' in c), None)
    col_name = next((c for c in df.columns if '一级客户名称' in c), None)
    
    if not col_code: return pd.DataFrame()

    df['原始编码'] = df[col_code].apply(clean_str)
    df['原始名称'] = df[col_name].apply(clean_str) if col_name else ''
    if '业务线' not in df.columns: df['业务线'] = ''
    df['业务线'] = df['业务线'].apply(clean_str)
    
    amt_col = next((c for c in df.columns if '传ERP金额' in c or ('ERP' in c and '金额' in c)), '传ERP金额')
    df['传ERP金额'] = clean_amount(df[amt_col]) if amt_col in df.columns else 0

    # 场景分流
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

# ==========================================
# 5. 数据处理逻辑 - 冲销模式 (Write-off)
# ==========================================

def process_writeoff_discount(df, target_scenario):
    """
    冲销-折让数据处理
    逻辑：读取总表 -> 筛选场景 -> 强制去后缀匹配
    字段：A客户号, B公司, C业务线, D汇总, E类型更新
    """
    # 假设没有表头，或者用户上传的是标准格式，这里尝试按列位置或名称读取
    # 为了稳健，我们先清洗列名
    df.columns = df.columns.astype(str).str.strip()
    
    # 映射列 (根据描述)
    # 如果是无表头或标准表头，尝试智能识别
    # 这里假设用户上传的表有表头，且包含关键列
    # 如果是纯数据无表头，可能需要按 index 读取。这里假设有表头。
    
    # 尝试寻找列
    col_code = next((c for c in df.columns if '客户' in c and ('号' in c or '编码' in c)), None) # A列
    col_biz = next((c for c in df.columns if '业务线' in c), None) # C列
    col_amt = next((c for c in df.columns if '汇总' in c or '金额' in c), None) # D列
    col_type = next((c for c in df.columns if '类型' in c), None) # E列
    
    if not col_code or not col_biz or not col_amt:
        st.error(f"❌ 冲销表列识别失败。需包含：客户号、业务线、汇总(金额)。\n读取到: {list(df.columns)}")
        return pd.DataFrame()

    # 1. 筛选场景 (业务线)
    # 模糊匹配：比如选择 "OTC-医疗备案"，只要业务线里包含 "OTC" 且包含 "医疗" 即可，或者完全匹配
    # 根据描述，这里先不做太严格的筛选，或者直接全量处理，最后在核对时筛选？
    # 描述说：“生成透视表后，需要筛选业务场景”。那我们先处理全量。
    
    # 2. 清洗
    df['Code_Raw'] = df[col_code].apply(clean_str)
    df['Code_Clean'] = df['Code_Raw'].apply(strip_suffix) # 强制去后缀！解决 A0686929-00 问题
    
    df['业务线'] = df[col_biz].apply(clean_str)
    df['金额'] = clean_amount(df[col_amt])
    df['类型'] = df[col_type].apply(clean_str) if col_type else '默认'
    
    # 3. 筛选场景 (根据侧边栏)
    # 如果用户选了 "自定义"，则不筛选
    if target_scenario != "自定义":
        # 简单包含逻辑
        df = df[df['业务线'].str.contains(target_scenario, na=False, case=False)]
        if df.empty:
            st.warning(f"⚠️ 在冲销表中未找到业务线包含 '{target_scenario}' 的数据。")
            
    # 4. 生成 Key
    df['透视Key'] = df['Code_Clean'] + df['业务线']
    
    return df

# ==========================================
# 6. 通用 ERP 处理 (支持两种模式)
# ==========================================

def process_erp_generic(df, bus_map, valid_codes, valid_names, scenario, mode):
    """
    ERP处理通用函数
    mode: "PROVISION" or "WRITEOFF"
    """
    df.columns = df.columns.astype(str).str.strip()
    
    if '交易对象编码' not in df.columns: 
        st.error("ERP缺少 '交易对象编码'"); return pd.DataFrame()
    
    # 1. 基础清洗
    def clean_prefix(t):
        t = clean_str(t)
        if ':' in t: return t.split(':')[0] if len(t.split(':'))==1 else t.split(':')[-1].strip()
        return t

    df['原始交易编码'] = df['交易对象编码'].apply(clean_prefix)
    df['Code_Clean'] = df['原始交易编码'].apply(strip_suffix) # 强制去后缀
    
    df['帐户'] = df['帐户'].astype(str).str.strip()
    df['金额_借贷'] = clean_amount(df['本位币借方']) + clean_amount(df['本位币贷方'])
    
    # 2. 解析业务线
    def extract_bus(acc):
        if not acc: return None
        parts = acc.split('.')
        return next((p for p in parts if p.startswith(('A','B')) and len(p)>1), None)

    df['提取_业务线Code'] = df['帐户'].apply(extract_bus)
    df['业务线'] = df['提取_业务线Code'].apply(clean_str).map(bus_map) if bus_map else None
    
    # 3. Key 生成逻辑
    if mode == "PROVISION" and scenario == "商务二级":
        # 暂估-商务二级：特殊用名称
        if '交易对象名称' in df.columns:
            df['标准名称'] = df['交易对象名称'].apply(normalize_brackets)
            df['透视Key'] = df['标准名称'] + df['业务线']
            df['是否关联方'] = df['标准名称'].apply(lambda x: x in valid_names)
        else:
            st.error("ERP缺少 '交易对象名称' 列 (商务二级必须)")
            return pd.DataFrame()
    else:
        # 其他所有情况 (暂估其他场景 & 冲销所有场景)：都用编码
        # 冲销核对要求：二级也用编码处理 (Requirement 1)
        df['透视Key'] = df['Code_Clean'] + df['业务线']
        
        # 关联方判断 (暂估才需要，冲销主要是全量核对，但保留逻辑无妨)
        if valid_codes:
            df['是否关联方'] = df['Code_Clean'].apply(lambda x: x in valid_codes)
        else:
            df['是否关联方'] = False
            
    return df

# ==========================================
# 7. 核对执行函数
# ==========================================

def perform_reconciliation(df_p, df_e, mode):
    """
    执行核对
    mode: "PROVISION" (暂估) or "WRITEOFF" (冲销)
    """
    key_col = '透视Key'
    
    # --- 左边 (折让) ---
    if mode == "PROVISION":
        # 暂估：按金额、税额拆分
        p_agg = df_p.dropna(subset=[key_col]).groupby(key_col).agg({
            '传ERP金额':'sum', '金额_不含税':'sum', '税额':'sum'
        }).rename(columns={'传ERP金额':'折让_总额', '金额_不含税':'折让_金额', '税额':'折让_税额'})
    else:
        # 冲销：按类型透视 (Requirement: 类型更新作为列)
        # 冲销数据里，'金额' 是汇总值
        # 这里我们需要做一个 Pivot Table：行=Key, 列=类型, 值=金额
        p_data = df_p.dropna(subset=[key_col])
        p_agg = p_data.pivot_table(index=key_col, columns='类型', values='金额', aggfunc='sum', fill_value=0)
        # 计算一个行总计，方便和ERP核对
        p_agg['折让_汇总总计'] = p_agg.sum(axis=1)
        
    # --- 右边 (ERP) ---
    e_data = df_e.dropna(subset=[key_col])
    
    if mode == "PROVISION":
        # 暂估：筛选3个科目
        targets = ['应收账款-应收账款（总账专用）', '主营业务收入-商品收入-贸易类', '应交税费-待转销项税额']
    else:
        # 冲销：只筛选应收账款 (Requirement 2)
        targets = ['应收账款-应收账款（总账专用）']
    
    if '会计科目' in e_data.columns:
        e_data = e_data[e_data['会计科目'].isin(targets)]
        e_pivot = e_data.pivot_table(index=key_col, columns='会计科目', values='金额_借贷', aggfunc='sum', fill_value=0)
        # 补全列
        for c in targets: 
            if c not in e_pivot.columns: e_pivot[c] = 0.0
    else:
        e_pivot = pd.DataFrame(columns=targets)
        
    # 重命名 ERP 列
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
    
    # --- 合并 ---
    merged = pd.merge(p_agg, e_pivot, left_index=True, right_index=True, how='outer').fillna(0)
    
    # --- 计算差异 ---
    if mode == "PROVISION":
        merged['核对_应收(0)'] = merged['折让_总额'] + merged['ERP_应收账款']
        merged['核对_收入(0)'] = merged['折让_金额'] + merged['ERP_主营收入']
        merged['核对_税额(0)'] = merged['折让_税额'] + merged['ERP_销项税']
        # 列排序
        cols = ['折让_总额', 'ERP_应收账款', '核对_应收(0)', '折让_金额', 'ERP_主营收入', '核对_收入(0)', '折让_税额', 'ERP_销项税', '核对_税额(0)']
        return merged[[c for c in cols if c in merged.columns]]
    else:
        # 冲销核对：折让汇总 vs ERP应收
        # 注意：这里假设方向是相反的，所以相加为0。如果方向相同，可能需要相减。
        # 通常冲销是减少应收，所以可能和暂估方向相反。如果相加不为0，请尝试相减。
        # 暂定逻辑：A + B = 0
        merged['核对_差额(0)'] = merged['折让_汇总总计'] + merged['ERP_应收账款(总账)']
        
        # 把折让的透视列也放进去展示
        first_cols = ['折让_汇总总计', 'ERP_应收账款(总账)', '核对_差额(0)']
        other_cols = [c for c in merged.columns if c not in first_cols]
        return merged[first_cols + other_cols]

def apply_styles(df):
    def hl(val): 
        if isinstance(val, (int, float)) and abs(val) > 0.01:
            return 'background-color: #ffcccc; color: red'
        return ''
    chk = [c for c in df.columns if '核对' in c]
    return df.style.map(hl, subset=chk).format("{:,.2f}")

# ==========================================
# 8. 主程序
# ==========================================
if match_file_source and file_left and file_right:
    bus_map, valid_codes, valid_names = load_mappings(match_file_source)
    
    if bus_map:
        try:
            # 读取
            df_l = pd.read_csv(file_left) if file_left.name.endswith('.csv') else pd.read_excel(file_left)
            h_row = 3 # ERP Header
            df_r = pd.read_excel(file_right, header=h_row) if not file_right.name.endswith('.csv') else pd.read_csv(file_right, header=h_row)
            
            st.info(f"🚀 正在执行：{TASK_MODE} | 场景：{selected_scenario}")
            
            # === 分流处理 ===
            if TASK_MODE == "暂估核对 (Provision)":
                # 1. 暂估处理
                df_p = process_provision_data(df_l, valid_codes, valid_names, selected_scenario)
                df_e = process_erp_generic(df_r, bus_map, valid_codes, valid_names, selected_scenario, "PROVISION")
                
                # 2. 暂估核对 (分客户/关联方)
                t1, t2, t3 = st.tabs(["👥 客户对账", "🏢 关联方对账", "📥 结果导出"])
                
                with t1:
                    res = perform_reconciliation(df_p, df_e, "PROVISION")
                    res_view = add_total_row(res)
                    st.dataframe(apply_styles(res_view), use_container_width=True, height=500)
                    
                with t2:
                    # 关联方筛选
                    df_p_rel = df_p[df_p['是否关联方']==True]
                    df_e_rel = df_e[df_e['是否关联方']==True]
                    res_rel = perform_reconciliation(df_p_rel, df_e_rel, "PROVISION")
                    st.dataframe(apply_styles(add_total_row(res_rel)), use_container_width=True, height=500)
                    
                with t3:
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                        res.to_excel(w, sheet_name='客户对账')
                        res_rel.to_excel(w, sheet_name='关联方对账')
                    st.download_button("下载暂估核对结果", out.getvalue(), "暂估核对.xlsx")

            else:
                # === 冲销核对 (Write-off) ===
                # 1. 冲销处理
                # 折让：读取、去后缀、筛选场景、透视
                df_p = process_writeoff_discount(df_l, selected_scenario)
                
                # ERP：通用处理 (强制用Code模式)、筛选科目(在核对步)
                # 注意：冲销模式下，商务二级也强制用 Code (valid_names传空即可或在函数内控制)
                df_e = process_erp_generic(df_r, bus_map, valid_codes, None, selected_scenario, "WRITEOFF")
                
                # 2. 冲销核对 (只有一张大表)
                st.write(f"📊 冲销数据预览: 折让行数 {len(df_p)} | ERP行数 {len(df_e)}")
                
                res_wo = perform_reconciliation(df_p, df_e, "WRITEOFF")
                res_wo_final = add_total_row(res_wo)
                
                # 筛选差异
                chk_cols = [c for c in res_wo_final.columns if '核对' in c]
                diff_val = res_wo_final.loc['=== 总计 ===', chk_cols[0]] if not res_wo_final.empty else 0
                
                c1, c2 = st.columns(2)
                c1.metric("总行数", len(res_wo))
                c2.metric("总差异", f"{diff_val:,.2f}", delta_color="inverse")
                
                # 交互筛选
                show_diff = st.checkbox("🧨 只看差异行")
                if show_diff:
                    # 排除合计行进行筛选
                    data_only = res_wo_final.drop(index='=== 总计 ===', errors='ignore')
                    cond = data_only[chk_cols].apply(lambda x: x.abs() > 0.01).any(axis=1)
                    st.dataframe(apply_styles(add_total_row(data_only[cond])), use_container_width=True)
                else:
                    st.dataframe(apply_styles(res_wo_final), use_container_width=True)
                
                # 下载
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                    res_wo_final.to_excel(w, sheet_name='冲销核对')
                    df_p.to_excel(w, sheet_name='折让明细', index=False)
                    df_e.to_excel(w, sheet_name='ERP明细', index=False)
                st.download_button("下载冲销核对结果", out.getvalue(), "冲销核对.xlsx")

        except Exception as e:
            st.error(f"运行出错: {e}")
            st.exception(e)
else:
    st.info("👈 请先上传所需文件")