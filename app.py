import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from io import BytesIO
import uuid
import sqlite3
import os
import zipfile
import re

# ========== 数据库路径 ==========
DB_PATH = os.path.join(os.path.dirname(__file__), "app_data.db")

# ========== 初始化数据库 ==========
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS companies (
        id TEXT PRIMARY KEY, company_name TEXT, province TEXT, city TEXT, district TEXT, tax_id TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS templates (
        id TEXT PRIMARY KEY, province TEXT, city TEXT, district TEXT, report_type TEXT,
        template_name TEXT, template_version TEXT, source_url TEXT, source_authority TEXT,
        publish_date TEXT, required_fields TEXT, status TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS rules (
        id TEXT PRIMARY KEY, city TEXT, province TEXT, unit_social REAL, personal_social REAL,
        unit_fund REAL, personal_fund REAL, social_min REAL, social_max REAL,
        fund_min REAL, fund_max REAL, source_quote TEXT, is_default BOOLEAN DEFAULT 0
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS export_history (
        id TEXT PRIMARY KEY, company_id TEXT, template_id TEXT, company_name TEXT,
        city TEXT, province TEXT, report_type TEXT, period_type TEXT, generated_at TEXT,
        review_status TEXT, reviewer TEXT, reviewed_at TEXT, file_name TEXT, file_data BLOB,
        data_source TEXT, month_used TEXT, year_used TEXT
    )''')
    conn.commit()
    conn.close()

init_db()

# ========== 数据库迁移 ==========
def migrate_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("PRAGMA table_info(rules)")
    columns_rules = [col[1] for col in c.fetchall()]
    if 'province' not in columns_rules:
        c.execute("ALTER TABLE rules ADD COLUMN province TEXT")
    c.execute("PRAGMA table_info(export_history)")
    columns_export = [col[1] for col in c.fetchall()]
    for col in ['data_source', 'month_used', 'year_used']:
        if col not in columns_export:
            c.execute(f"ALTER TABLE export_history ADD COLUMN {col} TEXT")
    conn.commit()
    conn.close()

migrate_db()

# ========== 数据操作函数 ==========
def dict_fetchall(cursor):
    columns = [col[0] for col in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]

def safe_execute_query(query, params=None):
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        if params:
            c.execute(query, params)
        else:
            c.execute(query)
        rows = dict_fetchall(c)
        conn.close()
        return rows
    except sqlite3.OperationalError as e:
        if "no such table" in str(e):
            return []
        else:
            raise e

def load_companies():
    return safe_execute_query("SELECT * FROM companies")

def load_templates():
    return safe_execute_query("SELECT * FROM templates WHERE status='active'")

def load_rules():
    return safe_execute_query("SELECT * FROM rules ORDER BY province, city")

def load_export_history():
    return safe_execute_query("SELECT * FROM export_history ORDER BY generated_at DESC")

def save_companies(companies):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM companies")
    for comp in companies:
        c.execute('''INSERT INTO companies (id, company_name, province, city, district, tax_id)
            VALUES (?,?,?,?,?,?)''',
            (comp.get('id', str(uuid.uuid4())[:8]), comp['company_name'], comp['province'],
             comp['city'], comp.get('district',''), comp.get('tax_id','')))
    conn.commit()
    conn.close()

def save_template(template):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO templates 
        (id, province, city, district, report_type, template_name, template_version,
         source_url, source_authority, publish_date, required_fields, status)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)''',
        (template['id'], template['province'], template['city'], template.get('district',''),
         template['report_type'], template['template_name'], template['template_version'],
         template.get('source_url',''), template.get('source_authority',''),
         template.get('publish_date',''), template.get('required_fields',''), template.get('status','active')))
    conn.commit()
    conn.close()

def save_rules(rules):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM rules")
    for r in rules:
        c.execute('''INSERT INTO rules 
            (id, city, province, unit_social, personal_social, unit_fund, personal_fund,
             social_min, social_max, fund_min, fund_max, source_quote, is_default)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (r['id'], r['city'], r.get('province',''), r['unit_social'], r['personal_social'],
             r['unit_fund'], r['personal_fund'], r.get('social_min',0), r.get('social_max',999999),
             r.get('fund_min',0), r.get('fund_max',999999), r.get('source_quote',''),
             r.get('is_default',0)))
    conn.commit()
    conn.close()

def save_export(record):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO export_history 
        (id, company_id, template_id, company_name, city, province, report_type, period_type,
         generated_at, review_status, reviewer, reviewed_at, file_name, file_data,
         data_source, month_used, year_used)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (record['id'], record.get('company_id',''), record.get('template_id',''),
         record['company_name'], record.get('city',''), record.get('province',''),
         record.get('report_type',''), record.get('period_type',''), record['generated_at'],
         record.get('review_status','pending'), record.get('reviewer',''), record.get('reviewed_at',''),
         record.get('file_name',''), record.get('file_data', None),
         record.get('data_source',''), record.get('month_used',''), record.get('year_used','')))
    conn.commit()
    conn.close()

def update_export_status(export_id, status, reviewer):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''UPDATE export_history 
        SET review_status=?, reviewer=?, reviewed_at=?
        WHERE id=?''',
        (status, reviewer, datetime.now().isoformat(), export_id))
    conn.commit()
    conn.close()

# ========== 全国省份及城市规则 ==========
# （与之前相同，此处省略以节省篇幅，实际代码中包含完整的PROVINCE_DEFAULT_RULES）

# ========== 全国官方模板库 ==========
# （与之前相同，此处省略，实际代码中包含完整的generate_all_templates和DEFAULT_TEMPLATES）

# ========== 标准化匹配函数 ==========
def normalize_name(name):
    if not name:
        return name
    # 去除常见后缀
    for suffix in ['省', '市', '区', '县', '自治区', '特别行政区']:
        if name.endswith(suffix):
            name = name[:-len(suffix)]
    return name.strip()

def match_template_with_details(province, city, district, report_type):
    templates = load_templates()
    if not templates:
        return None, None, []
    
    norm_prov = normalize_name(province)
    norm_city = normalize_name(city)
    norm_dist = normalize_name(district) if district else ''
    
    matched = None
    match_level = None
    
    # 1. 区级匹配
    for t in templates:
        if normalize_name(t['province']) == norm_prov and \
           normalize_name(t['city']) == norm_city and \
           normalize_name(t.get('district', '')) == norm_dist and \
           t['report_type'] == report_type:
            matched = t
            match_level = "区级模板"
            break
    # 2. 市级匹配
    if not matched:
        for t in templates:
            if normalize_name(t['province']) == norm_prov and \
               normalize_name(t['city']) == norm_city and \
               t['report_type'] == report_type:
                matched = t
                match_level = "市级模板"
                break
    # 3. 省级匹配
    if not matched:
        for t in templates:
            if normalize_name(t['province']) == norm_prov and \
               t['report_type'] == report_type:
                matched = t
                match_level = "省级模板"
                break
    
    # 候选模板（用于手动选择）
    candidates = [t for t in templates if normalize_name(t['province']) == norm_prov and t['report_type'] == report_type]
    return matched, match_level, candidates

# ========== 解析上传的Excel ==========
def parse_uploaded_excel(file):
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    all_companies = []
    for sheet in sheets:
        try:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            header_row = None
            for i, row in df.iterrows():
                row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                if '所属城市' in row_text or '城市' in row_text or '分公司' in row_text:
                    header_row = i
                    break
            if header_row is not None:
                df = pd.read_excel(file, sheet_name=sheet, skiprows=header_row)
                df.columns = [str(c).strip() for c in df.columns]
                city_col = None
                company_col = None
                district_col = None
                for col in df.columns:
                    if '所属城市' in col or '城市' in col:
                        city_col = col
                    elif '分公司' in col or '公司' in col:
                        company_col = col
                    elif '区县' in col or '区' in col:
                        district_col = col
                if city_col and company_col:
                    for _, row in df.iterrows():
                        city = str(row[city_col]) if pd.notna(row[city_col]) else ''
                        company = str(row[company_col]) if pd.notna(row[company_col]) else ''
                        district = str(row[district_col]) if district_col and pd.notna(row[district_col]) else ''
                        if city and company:
                            province = city
                            for r in PROVINCE_DEFAULT_RULES:
                                if r['city'] == city and r.get('province'):
                                    province = r['province']
                                    break
                            all_companies.append({
                                'company_name': company,
                                'province': province,
                                'city': city,
                                'district': district,
                                'tax_id': ''
                            })
        except:
            continue
    unique = []
    seen = set()
    for c in all_companies:
        key = (c['company_name'], c['city'])
        if key not in seen:
            seen.add(key)
            unique.append(c)
    return unique

def get_rule_for_city(city):
    rules = load_rules()
    for r in rules:
        if r['city'] == city:
            return r
    for r in rules:
        if r.get('province') == city:
            return r
    return None

def get_data_source_info(df):
    info = {}
    if df is not None and not df.empty:
        for col in df.columns:
            col_lower = str(col).lower()
            if '年份' in col_lower or '年度' in col_lower:
                info['year'] = df[col].iloc[0] if not df[col].empty else '2025'
            if '月份' in col_lower or '月' in col_lower:
                if '统计月份' in col_lower or '月份' in col_lower:
                    info['month'] = df[col].iloc[0] if not df[col].empty else '12'
    return info

# ========== 初始化默认数据 ==========
PROVINCE_DEFAULT_RULES = [
    # （完整规则列表，由于篇幅省略，实际代码中包含所有规则）
]
DEFAULT_TEMPLATES = generate_all_templates()

def init_default_data():
    if not load_rules():
        all_rules = []
        for r in PROVINCE_DEFAULT_RULES:
            all_rules.append({
                'id': str(uuid.uuid4())[:8],
                'city': r['city'],
                'province': r.get('province', r['city']),
                'unit_social': r['unit_social'],
                'personal_social': r['personal_social'],
                'unit_fund': r['unit_fund'],
                'personal_fund': r['personal_fund'],
                'social_min': r.get('social_min', 0),
                'social_max': r.get('social_max', 999999),
                'fund_min': r.get('fund_min', 0),
                'fund_max': r.get('fund_max', 999999),
                'source_quote': r.get('source_quote', '省份默认'),
                'is_default': 1 if r['city'] == r.get('province', r['city']) else 0
            })
        save_rules(all_rules)
    
    if not load_templates():
        for t in DEFAULT_TEMPLATES:
            save_template(t)

init_default_data()

# ========== Streamlit 页面 ==========
st.set_page_config(page_title="官方模板匹配器", layout="wide")
st.title("📋 官方模板匹配器（优化版）")
st.markdown("**上传Excel → 自动提取城市/公司 → 选择模板和统计口径 → 生成待复核版Excel**")

template_count = len(load_templates())
rule_count = len(load_rules())
st.success(f"✅ 已内置 {rule_count} 个城市的规则，以及 {template_count} 个官方模板")

# ===== 侧边栏 =====
with st.sidebar:
    st.header("📤 上传数据Excel")
    uploaded_file = st.file_uploader("选择Excel文件（.xlsx）", type=["xlsx"])
    
    if uploaded_file:
        with st.spinner("正在解析Excel..."):
            companies = parse_uploaded_excel(uploaded_file)
            if companies:
                save_companies(companies)
                st.success(f"成功提取 {len(companies)} 家公司")
                try:
                    xls = pd.ExcelFile(uploaded_file)
                    data_sheet = None
                    for s in xls.sheet_names:
                        if '明细' in s or '月度' in s or '数据' in s:
                            data_sheet = s
                            break
                    if data_sheet:
                        df_data = pd.read_excel(uploaded_file, sheet_name=data_sheet)
                        st.session_state['imported_df'] = df_data
                        st.session_state['data_sheet_name'] = data_sheet
                        st.success(f"已读取数据Sheet「{data_sheet}」，共{len(df_data)}行")
                except:
                    pass
            else:
                st.warning("未识别到公司数据，请确认Excel包含「城市」和「公司」列")
    
    with st.sidebar.expander("🏢 当前公司列表"):
        companies = load_companies()
        if companies:
            st.dataframe(pd.DataFrame(companies))
            st.caption(f"共 {len(companies)} 家公司")
        else:
            st.info("暂无数据")
    
    with st.sidebar.expander("📚 查看所有模板"):
        templates = load_templates()
        if templates:
            df_temp = pd.DataFrame(templates)
            st.dataframe(df_temp[['province', 'city', 'report_type', 'template_name', 'template_version']])
            st.caption(f"共 {len(templates)} 个模板")
        else:
            st.info("暂无模板")

# ===== 主体 =====
st.subheader("📊 导入数据预览")
data_source_info = ""
if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
    df_preview = st.session_state['imported_df']
    st.dataframe(df_preview.head(10))
    sheet_name = st.session_state.get('data_sheet_name', '未知Sheet')
    info = get_data_source_info(df_preview)
    year = info.get('year', '')
    month = info.get('month', '')
    data_source_info = f"数据来源：{sheet_name}"
    if year:
        data_source_info += f"，年份：{year}"
    if month:
        data_source_info += f"，月份：{month}"
    st.caption(f"共 {len(df_preview)} 行数据 | {data_source_info}")
else:
    st.info("上传Excel后，此处将显示数据预览")

companies = load_companies()
if not companies:
    st.info("👈 请先在侧边栏上传包含公司/城市数据的Excel")
    st.stop()

all_provinces = sorted(set(c['province'] for c in companies if c['province']))

col1, col2, col3 = st.columns(3)
with col1:
    province = st.selectbox("省份", [""] + all_provinces)
    cities = sorted(set(c['city'] for c in companies if c['province'] == province)) if province else sorted(set(c['city'] for c in companies))
    city = st.selectbox("城市", [""] + cities)
with col2:
    districts = sorted(set(c['district'] for c in companies if c['province'] == province and c['city'] == city)) if province and city else []
    district = st.selectbox("区县", [""] + districts)
    company_list = [c for c in companies if c['province'] == province and c['city'] == city and (not district or c['district'] == district)]
    company_names = [c['company_name'] for c in company_list]
    selected_company_names = st.multiselect("公司（可多选）", company_names)
with col3:
    report_type = st.selectbox("报表类型", ["", "增值税", "社保", "公积金", "个人所得税", "企业所得税", "年度汇算清缴"])
    period_type = st.selectbox("统计口径", ["月度（12月单月）", "累计（1-12月）"])

selected_companies = [c for c in company_list if c['company_name'] in selected_company_names]

if selected_companies and report_type:
    st.markdown("---")
    st.subheader("🔍 匹配结果")
    
    matched, match_level, candidates = match_template_with_details(province, city, district, report_type)
    
    # ===== 模板选择逻辑 =====
    selected_template = None
    if matched:
        st.success(f"✅ 自动匹配到官方模板（{match_level}）")
        # 显示匹配的模板
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**📄 模板信息**")
            st.write(f"模板名称：{matched['template_name']}")
            st.write(f"版本：{matched['template_version']}")
            st.write(f"发布机构：{matched['source_authority']}")
            st.write(f"发布日期：{matched['publish_date']}")
            st.write(f"必填字段：{matched['required_fields']}")
        with col_b:
            st.markdown("**🔗 来源信息**")
            st.write(f"来源URL：[{matched['source_url']}]({matched['source_url']})")
            st.write(f"适用地区：{matched['province']} {matched['city']} {matched['district']}")
            st.write(f"报表类型：{matched['report_type']}")
        
        # 允许用户切换模板
        if candidates and len(candidates) > 1:
            with st.expander("🔄 切换其他模板（点击展开）"):
                template_options = {f"{t['template_name']}（{t['province']}{t['city']}{t.get('district','')}）": t for t in candidates}
                selected_key = st.selectbox("选择模板", list(template_options.keys()))
                if selected_key:
                    selected_template = template_options[selected_key]
                    st.info(f"已选择：{selected_template['template_name']}")
        else:
            selected_template = matched
    else:
        st.warning("⚠️ 自动匹配失败，请从下方候选模板中选择")
        if candidates:
            st.info(f"找到 {len(candidates)} 个可用模板，请选择")
            template_options = {f"{t['template_name']}（{t['province']}{t['city']}{t.get('district','')}）": t for t in candidates}
            selected_key = st.selectbox("选择模板", list(template_options.keys()))
            if selected_key:
                selected_template = template_options[selected_key]
                st.success(f"已选择：{selected_template['template_name']}")
        else:
            st.error("❌ 当前地区没有可用模板，请使用通用模板")
            selected_template = {
                'id': 'gen001',
                'template_name': f'{report_type}通用申报表',
                'template_version': 'v1.0',
                'source_authority': '系统通用',
                'publish_date': datetime.now().strftime('%Y-%m-%d'),
                'required_fields': '纳税人识别号,公司名称,申报金额',
                'source_url': '#'
            }
            match_level = "通用模板"
    
    # 使用选中的模板
    matched = selected_template if selected_template else matched
    
    # ===== 模板预览 =====
    st.subheader("📋 模板预览")
    if matched:
        fields = matched['required_fields'].split(',')
        st.markdown(f"**字段列表**：{', '.join(fields)}")
        
        sample_row = {}
        for f in fields:
            sample_values = {
                '纳税人识别号': '91310115MA1KXXXXX',
                '公司名称': selected_companies[0]['company_name'] if selected_companies else '示例公司',
                '销售额': '100,000.00',
                '进项税额': '13,000.00',
                '应纳税额': '0.00',
                '单位名称': selected_companies[0]['company_name'] if selected_companies else '示例公司',
                '社保登记号': 'SH123456',
                '基数': '8,000.00',
                '单位金额': '1,280.00',
                '个人金额': '640.00',
                '单位比例': '12.0%',
                '个人比例': '12.0%',
                '公积金账号': 'GJJ123456',
                '收入额': '100,000.00',
                '专项扣除': '0.00',
                '营业收入': '1,000,000.00',
                '营业成本': '600,000.00',
                '应纳税所得额': '100,000.00',
                '全年收入': '12,000,000.00',
                '全年成本': '7,200,000.00',
                '已预缴税额': '150,000.00',
                '应补退税额': '0.00',
                '申报金额': '100,000.00'
            }
            sample_row[f] = sample_values.get(f, f'<{f} 示例值>')
        
        preview_df = pd.DataFrame([{'字段名': f, '示例值': sample_row[f]} for f in fields])
        st.dataframe(preview_df, use_container_width=True)
        st.info(f"📌 当前使用模板：{matched['template_name']}（{match_level}）")
    
    # ===== 数据校验 =====
    st.subheader("📋 数据校验")
    data_source_text = "未知"
    if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
        data_source_text = st.session_state.get('data_sheet_name', '未知Sheet')
        info = get_data_source_info(st.session_state['imported_df'])
        if info.get('year'):
            data_source_text += f"（年份：{info.get('year')}"
        if info.get('month'):
            data_source_text += f"，月份：{info.get('month')}"
        if info.get('year') or info.get('month'):
            data_source_text += "）"
    
    st.info(f"📌 数据来源：{data_source_text}")
    
    missing_rules = []
    for comp in selected_companies:
        rule = get_rule_for_city(comp['city'])
        if rule is None:
            missing_rules.append(comp['city'])
    if missing_rules:
        st.warning(f"⚠️ 以下城市缺少规则，将使用默认值：{', '.join(set(missing_rules))}")
        st.info("💡 如需添加规则，请在「规则管理」中添加对应城市的缴费比例")
    else:
        st.success("✅ 所有城市已配置规则")
    
    # ===== 报表预览 =====
    st.subheader("📊 报表预览（生成前确认）")
    preview_data = []
    for comp in selected_companies:
        preview_data.append({
            '公司': comp['company_name'],
            '城市': comp['city'],
            '模板': matched['template_name'] if matched else '未选择',
            '匹配级别': match_level if matched else '无'
        })
    st.dataframe(pd.DataFrame(preview_data), use_container_width=True)
    
    reviewed = st.checkbox("✅ 我已人工复核确认数据无误", value=False)
    
    if st.button("📥 生成待复核版Excel", disabled=not reviewed):
        if not matched:
            st.error("请先选择模板")
        else:
            generated_files = []
            summary = []
            errors = []
            
            for comp in selected_companies:
                try:
                    rule = get_rule_for_city(comp['city'])
                    fields = matched['required_fields'].split(',')
                    
                    if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
                        df_data = st.session_state['imported_df']
                        company_col = None
                        for col in df_data.columns:
                            if '公司' in str(col) or '分公司' in str(col):
                                company_col = col
                                break
                        if company_col:
                            df_comp = df_data[df_data[company_col] == comp['company_name']]
                            if not df_comp.empty:
                                row_data = []
                                for f in fields:
                                    matched_col = None
                                    for col in df_data.columns:
                                        if f in str(col) or str(col) in f:
                                            matched_col = col
                                            break
                                    if matched_col:
                                        row_data.append(df_comp.iloc[0][matched_col])
                                    else:
                                        row_data.append('')
                            else:
                                row_data = [''] * len(fields)
                        else:
                            row_data = [''] * len(fields)
                    else:
                        sample_data = {
                            '纳税人识别号': comp.get('tax_id', ''),
                            '公司名称': comp['company_name'],
                            '销售额': '100,000.00',
                            '进项税额': '13,000.00',
                            '应纳税额': '0.00',
                            '单位名称': comp['company_name'],
                            '社保登记号': 'SH123456',
                            '基数': '8,000.00',
                            '单位金额': str(round(8000 * rule['unit_social'], 2)) if rule else '1,280.00',
                            '个人金额': str(round(8000 * rule['personal_social'], 2)) if rule else '640.00',
                            '单位比例': str(round(rule['unit_fund'] * 100, 1)) if rule else '12.0',
                            '个人比例': str(round(rule['personal_fund'] * 100, 1)) if rule else '12.0',
                            '公积金账号': 'GJJ123456',
                            '收入额': '100,000.00',
                            '专项扣除': '0.00',
                            '营业收入': '1,000,000.00',
                            '营业成本': '600,000.00',
                            '应纳税所得额': '100,000.00',
                            '申报金额': '100,000.00',
                            '全年收入': '12,000,000.00',
                            '全年成本': '7,200,000.00',
                            '已预缴税额': '150,000.00',
                            '应补退税额': '0.00'
                        }
                        row_data = [sample_data.get(f, '') for f in fields]
                    
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "申报表"
                    ws.append(fields)
                    ws.append(row_data)
                    
                    ws.insert_rows(1)
                    ws['A1'] = f'【系统生成 - 待复核版】统计口径：{period_type}'
                    ws['A1'].font = Font(color='FF0000', bold=True, size=14)
                    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(fields))
                    ws['A1'].alignment = Alignment(horizontal='center')
                    ws['A1'].fill = PatternFill(start_color='FFF9E6', end_color='FFF9E6', fill_type='solid')
                    
                    ws.insert_rows(2)
                    ws['A2'] = f'模板名称：{matched["template_name"]}  版本：{matched["template_version"]}  匹配级别：{match_level}'
                    ws['A2'].font = Font(color='666666', size=10)
                    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(fields))
                    
                    ws.insert_rows(3)
                    ws['A3'] = f'来源：{matched.get("source_authority","")}  发布日期：{matched.get("publish_date","")}'
                    ws['A3'].font = Font(color='666666', size=10)
                    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(fields))
                    
                    ws.insert_rows(4)
                    ws['A4'] = f'数据来源：{data_source_text}  统计口径：{period_type}'
                    ws['A4'].font = Font(color='666666', size=10)
                    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=len(fields))
                    
                    # 年检汇总
                    ws_annual = wb.create_sheet("年检汇总")
                    ws_annual.append(['年检汇总数据'])
                    ws_annual.merge_cells('A1:B1')
                    ws_annual['A1'].font = Font(bold=True, size=12)
                    
                    if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
                        df_all = st.session_state['imported_df']
                        social_col = None
                        fund_col = None
                        unit_col = None
                        personal_col = None
                        total_col = None
                        people_col = None
                        for col in df_all.columns:
                            col_str = str(col)
                            if '社保' in col_str and '合计' in col_str:
                                if '单位' in col_str:
                                    social_col = col
                                elif '个人' in col_str:
                                    personal_col = col
                            elif '公积金' in col_str and '合计' in col_str:
                                fund_col = col
                            elif '单位总费用' in col_str:
                                unit_col = col
                            elif '个人总费用' in col_str:
                                personal_col = col
                            elif '全部总费用' in col_str or '总金额' in col_str:
                                total_col = col
                            elif '参保人数' in col_str:
                                people_col = col
                        
                        if social_col:
                            social_total = df_all[social_col].sum()
                        else:
                            social_total = 0
                        if fund_col:
                            fund_total = df_all[fund_col].sum()
                        else:
                            fund_total = 0
                        if people_col:
                            total_people = df_all[people_col].sum()
                        else:
                            total_people = len(df_all)
                        
                        if unit_col and personal_col:
                            unit_total = df_all[unit_col].sum()
                            personal_total = df_all[personal_col].sum()
                            grand_total = df_all[total_col].sum() if total_col else (unit_total + personal_total)
                        else:
                            unit_total = social_total
                            personal_total = personal_col if personal_col else 0
                            grand_total = social_total + fund_total if fund_total else social_total
                    else:
                        total_people = 0
                        social_total = 0
                        fund_total = 0
                        unit_total = 0
                        personal_total = 0
                        grand_total = 0
                    
                    ws_annual.append(['公司名称', comp['company_name']])
                    ws_annual.append(['所属城市', comp['city']])
                    ws_annual.append(['统计口径', period_type])
                    ws_annual.append(['参保人数（全年）', int(total_people) if total_people else 0])
                    ws_annual.append(['全年社保缴费基数总额', round(social_total, 2)])
                    ws_annual.append(['全年公积金缴费基数总额', round(fund_total, 2)])
                    ws_annual.append(['单位全年缴费总额', round(unit_total, 2)])
                    ws_annual.append(['个人全年缴费总额', round(personal_total, 2)])
                    ws_annual.append(['全年总费用', round(grand_total, 2)])
                    ws_annual.append(['报告生成时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                    ws_annual.append(['数据来源', data_source_text])
                    
                    audit = wb.create_sheet("审计日志")
                    audit.append(['操作时间', '操作类型', '操作人', '详情'])
                    audit.append([datetime.now().isoformat(), 'GENERATED', '系统', f'公司:{comp["company_name"]}, 城市:{comp["city"]}, 模板:{matched["template_name"]}, 匹配级别:{match_level}, 数据来源:{data_source_text}'])
                    
                    output = BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    fname = f"{comp['company_name']}_{report_type}_{period_type.replace('（','_').replace('）','')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    generated_files.append((fname, output.getvalue()))
                    summary.append({
                        '公司': comp['company_name'], 
                        '城市': comp['city'], 
                        '模板': matched['template_name'],
                        '匹配级别': match_level,
                        '数据来源': data_source_text,
                        '状态': '待复核'
                    })
                    
                    save_export({
                        'id': str(uuid.uuid4())[:8],
                        'company_id': comp['id'],
                        'template_id': matched.get('id', 'gen001'),
                        'company_name': comp['company_name'],
                        'city': comp['city'],
                        'province': comp.get('province', ''),
                        'report_type': report_type,
                        'period_type': period_type,
                        'generated_at': datetime.now().isoformat(),
                        'review_status': 'pending',
                        'file_name': fname,
                        'file_data': output.getvalue(),
                        'data_source': data_source_text,
                        'month_used': period_type,
                        'year_used': datetime.now().strftime('%Y')
                    })
                except Exception as e:
                    errors.append(f"{comp['company_name']}: {str(e)}")
            
            if errors:
                for err in errors:
                    st.warning(err)
            if generated_files:
                st.success(f"✅ 成功生成 {len(generated_files)} 份报表")
                st.dataframe(pd.DataFrame(summary), use_container_width=True)
                if len(generated_files) > 1:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w') as zf:
                        for fname, data in generated_files:
                            zf.writestr(fname, data)
                    zip_buffer.seek(0)
                    st.download_button("📦 下载全部报表（ZIP）", data=zip_buffer, file_name=f"报表_{datetime.now().strftime('%Y%m%d')}.zip", mime="application/zip")
                else:
                    fname, data = generated_files[0]
                    st.download_button(f"📥 下载 {fname}", data=BytesIO(data), file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    if not selected_companies:
        st.info("👆 请先选择公司")
    elif not report_type:
        st.info("👆 请选择报表类型")

# ===== 导出历史 =====
with st.expander("📋 导出历史记录"):
    history = load_export_history()
    if history:
        df_hist = pd.DataFrame(history)
        st.dataframe(df_hist[['company_name', 'city', 'report_type', 'period_type', 'data_source', 'generated_at', 'review_status']], use_container_width=True)
        
        pending = [h for h in history if h['review_status'] == 'pending']
        if pending:
            st.subheader("✅ 复核待处理报表")
            opts = [f"{h['company_name']} - {h['city']} ({h['generated_at'][:10]})" for h in pending]
            sel_idx = st.selectbox("选择要复核的报表", range(len(opts)), format_func=lambda x: opts[x])
            selected = pending[sel_idx]
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ 通过复核"):
                    update_export_status(selected['id'], 'approved', '复核员')
                    st.success("已通过复核")
                    st.rerun()
            with col2:
                if st.button("❌ 驳回"):
                    update_export_status(selected['id'], 'rejected', '复核员')
                    st.warning("已驳回")
                    st.rerun()
    else:
        st.info("暂无导出记录")

# ===== 查看知识库 =====
with st.expander("📚 官方模板知识库（按省份查看）"):
    templates = load_templates()
    if templates:
        provinces_in_templates = sorted(set(t['province'] for t in templates))
        selected_province = st.selectbox("选择省份查看模板", [""] + provinces_in_templates)
        if selected_province:
            filtered = [t for t in templates if t['province'] == selected_province]
            st.dataframe(pd.DataFrame(filtered)[['city', 'district', 'report_type', 'template_name', 'template_version', 'source_authority']])
        else:
            st.dataframe(pd.DataFrame(templates)[['province', 'city', 'report_type', 'template_name', 'template_version']])
        st.caption(f"共 {len(templates)} 个官方模板，覆盖 {len(provinces_in_templates)} 个省份，6种报表类型")
    else:
        st.info("暂无模板")
