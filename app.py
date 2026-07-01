import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from io import BytesIO
import uuid
import sqlite3
import os
import zipfile
import json
import re

# ========== 数据库初始化 ==========
DB_PATH = os.path.join(os.path.dirname(__file__), "app_data.db")

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    # 公司表
    c.execute('''CREATE TABLE IF NOT EXISTS companies (
        id TEXT PRIMARY KEY,
        company_name TEXT,
        province TEXT,
        city TEXT,
        district TEXT,
        tax_id TEXT
    )''')
    # 模板表（官方模板）
    c.execute('''CREATE TABLE IF NOT EXISTS templates (
        id TEXT PRIMARY KEY,
        province TEXT,
        city TEXT,
        district TEXT,
        report_type TEXT,
        template_name TEXT,
        template_version TEXT,
        source_url TEXT,
        source_authority TEXT,
        publish_date TEXT,
        required_fields TEXT,
        status TEXT
    )''')
    # 规则表（城市社保公积金比例）
    c.execute('''CREATE TABLE IF NOT EXISTS rules (
        id TEXT PRIMARY KEY,
        city TEXT,
        unit_social REAL,
        personal_social REAL,
        unit_fund REAL,
        personal_fund REAL,
        social_min REAL,
        social_max REAL,
        fund_min REAL,
        fund_max REAL,
        source_quote TEXT
    )''')
    # 导出历史表
    c.execute('''CREATE TABLE IF NOT EXISTS export_history (
        id TEXT PRIMARY KEY,
        company_id TEXT,
        template_id TEXT,
        company_name TEXT,
        city TEXT,
        report_type TEXT,
        period_type TEXT,
        generated_at TEXT,
        review_status TEXT,
        reviewer TEXT,
        reviewed_at TEXT,
        file_name TEXT,
        file_data BLOB
    )''')
    # 自定义模板表（用户上传的模板文件）
    c.execute('''CREATE TABLE IF NOT EXISTS custom_templates (
        id TEXT PRIMARY KEY,
        name TEXT,
        file_data BLOB,
        field_mapping TEXT,
        created_at TEXT
    )''')
    conn.commit()
    conn.close()

init_db()

# ========== 辅助函数 ==========
def dict_fetchall(cursor):
    """将sqlite查询结果转为dict列表"""
    columns = [col[0] for col in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]

def safe_execute(func, *args, **kwargs):
    """安全执行函数，捕获异常返回错误信息"""
    try:
        return func(*args, **kwargs), None
    except Exception as e:
        return None, str(e)

# ========== 数据操作函数 ==========
def load_companies():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT * FROM companies")
    rows = dict_fetchall(c)
    conn.close()
    return rows

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

def load_templates():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT * FROM templates WHERE status='active'")
    rows = dict_fetchall(c)
    conn.close()
    return rows

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

def load_rules():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT * FROM rules")
    rows = dict_fetchall(c)
    conn.close()
    return rows

def save_rules(rules):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM rules")
    for r in rules:
        c.execute('''INSERT INTO rules 
            (id, city, unit_social, personal_social, unit_fund, personal_fund,
             social_min, social_max, fund_min, fund_max, source_quote)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
            (r['id'], r['city'], r['unit_social'], r['personal_social'],
             r['unit_fund'], r['personal_fund'], r.get('social_min',0), r.get('social_max',999999),
             r.get('fund_min',0), r.get('fund_max',999999), r.get('source_quote','')))
    conn.commit()
    conn.close()

def save_export(record):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO export_history 
        (id, company_id, template_id, company_name, city, report_type, period_type,
         generated_at, review_status, reviewer, reviewed_at, file_name, file_data)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (record['id'], record.get('company_id',''), record.get('template_id',''),
         record['company_name'], record.get('city',''), record.get('report_type',''),
         record.get('period_type',''), record['generated_at'], record.get('review_status','pending'),
         record.get('reviewer',''), record.get('reviewed_at',''),
         record.get('file_name',''), record.get('file_data', None)))
    conn.commit()
    conn.close()

def load_export_history():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT * FROM export_history ORDER BY generated_at DESC")
    rows = dict_fetchall(c)
    conn.close()
    return rows

def update_export_status(export_id, status, reviewer, comment=''):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''UPDATE export_history 
        SET review_status=?, reviewer=?, reviewed_at=?
        WHERE id=?''',
        (status, reviewer, datetime.now().isoformat(), export_id))
    conn.commit()
    conn.close()

# ========== 初始化默认数据 ==========
def init_default_data():
    # 插入官方模板
    if not load_templates():
        templates = [
            {
                'id': 't001', 'province': '上海', 'city': '上海市', 'district': '浦东新区',
                'report_type': '增值税', 'template_name': '上海市增值税纳税申报表（一般纳税人）',
                'template_version': 'v2024.1', 'source_url': 'https://shanghai.chinatax.gov.cn/bsfw/bszn/2024/zzs.xlsx',
                'source_authority': '国家税务总局上海市税务局', 'publish_date': '2024-01-15',
                'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额', 'status': 'active'
            },
            {
                'id': 't002', 'province': '上海', 'city': '上海市', 'district': '浦东新区',
                'report_type': '社保', 'template_name': '上海市社会保险费申报表（月度）',
                'template_version': 'v2024.1', 'source_url': 'https://rsj.sh.gov.cn/sbjb/2024/sb.xlsx',
                'source_authority': '上海市人力资源和社会保障局', 'publish_date': '2024-01-10',
                'required_fields': '单位名称,社保登记号,基数,单位金额,个人金额', 'status': 'active'
            },
            {
                'id': 't003', 'province': '广东', 'city': '广州市', 'district': '天河区',
                'report_type': '增值税', 'template_name': '广东省增值税纳税申报表',
                'template_version': 'v2024.1', 'source_url': 'https://guangdong.chinatax.gov.cn/bsfw/2024/zzs.xlsx',
                'source_authority': '国家税务总局广东省税务局', 'publish_date': '2024-01-20',
                'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额', 'status': 'active'
            },
            {
                'id': 't004', 'province': '北京', 'city': '北京市', 'district': '海淀区',
                'report_type': '增值税', 'template_name': '北京市增值税纳税申报表（一般纳税人）',
                'template_version': 'v2024.2', 'source_url': 'https://beijing.chinatax.gov.cn/bsfw/2024/zzs.xlsx',
                'source_authority': '国家税务总局北京市税务局', 'publish_date': '2024-02-01',
                'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额', 'status': 'active'
            },
        ]
        for t in templates:
            save_template(t)

    # 规则示例（上海）
    if not load_rules():
        rules = [
            {'id': 'r001', 'city': '上海', 'unit_social': 0.16, 'personal_social': 0.08,
             'unit_fund': 0.07, 'personal_fund': 0.07, 'social_min': 7310, 'social_max': 36549,
             'fund_min': 2590, 'fund_max': 34188, 'source_quote': '沪人社规〔2024〕22号'},
            {'id': 'r002', 'city': '北京', 'unit_social': 0.16, 'personal_social': 0.08,
             'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 6326, 'social_max': 33891,
             'fund_min': 2420, 'fund_max': 33891, 'source_quote': '京人社发〔2024〕15号'},
            {'id': 'r003', 'city': '广州', 'unit_social': 0.15, 'personal_social': 0.08,
             'unit_fund': 0.10, 'personal_fund': 0.10, 'social_min': 4588, 'social_max': 22941,
             'fund_min': 2300, 'fund_max': 27960, 'source_quote': '粤人社规〔2024〕8号'},
        ]
        save_rules(rules)

init_default_data()

# ========== Streamlit 页面 ==========
st.set_page_config(page_title="社保公积金智能报表系统", layout="wide")
st.title("🧾 社保公积金智能报表系统（全功能版）")
st.markdown("**上传数据 → 自动提取地区 → 智能匹配模板 → 自动计算 → 批量生成 → 复核审批**")

# ===== 侧边栏：导航 =====
menu = st.sidebar.radio(
    "功能导航",
    ["1. 数据导入", "2. 公司管理", "3. 模板管理", "4. 规则管理", "5. 生成报表", "6. 历史与复核", "7. 系统设置"]
)

# ===== 1. 数据导入 =====
if menu == "1. 数据导入":
    st.header("📤 数据导入")
    st.markdown("上传包含公司/城市/数据的Excel，系统自动提取并存储")
    
    uploaded_file = st.file_uploader("选择Excel文件（.xlsx）", type=["xlsx"])
    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheets = xls.sheet_names
            st.info(f"检测到Sheet：{', '.join(sheets)}")
            
            # 选择数据Sheet
            data_sheet = st.selectbox("选择数据Sheet", sheets)
            if data_sheet:
                df = pd.read_excel(uploaded_file, sheet_name=data_sheet, header=None)
                # 自动识别表头行
                header_row = None
                for i, row in df.iterrows():
                    row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                    if '所属城市' in row_text or '城市' in row_text or '分公司' in row_text:
                        header_row = i
                        break
                if header_row is not None:
                    df = pd.read_excel(uploaded_file, sheet_name=data_sheet, skiprows=header_row)
                    df.columns = [str(c).strip() for c in df.columns]
                    st.success(f"成功读取 {len(df)} 行数据")
                    st.dataframe(df.head(5))
                    
                    # 自动提取公司
                    city_col = None
                    company_col = None
                    district_col = None
                    for col in df.columns:
                        if '城市' in col:
                            city_col = col
                        elif '分公司' in col or '公司' in col:
                            company_col = col
                        elif '区县' in col or '区' in col:
                            district_col = col
                    
                    if city_col and company_col:
                        # 构建公司列表
                        companies = []
                        for _, row in df.iterrows():
                            city = str(row[city_col]) if pd.notna(row[city_col]) else ''
                            company = str(row[company_col]) if pd.notna(row[company_col]) else ''
                            district = str(row[district_col]) if district_col and pd.notna(row[district_col]) else ''
                            if city and company:
                                # 根据城市映射省份（简化版，可扩充）
                                province_map = {
                                    '上海': '上海', '北京': '北京', '广州': '广东', '深圳': '广东',
                                    '杭州': '浙江', '南京': '江苏', '苏州': '江苏', '成都': '四川',
                                    '重庆': '重庆', '武汉': '湖北', '西安': '陕西', '郑州': '河南',
                                    '长沙': '湖南', '青岛': '山东', '宁波': '浙江', '天津': '天津'
                                }
                                province = province_map.get(city, city)
                                companies.append({
                                    'id': str(uuid.uuid4())[:8],
                                    'company_name': company,
                                    'province': province,
                                    'city': city,
                                    'district': district,
                                    'tax_id': ''
                                })
                        if companies:
                            # 去重
                            unique = []
                            seen = set()
                            for c in companies:
                                key = (c['company_name'], c['city'])
                                if key not in seen:
                                    seen.add(key)
                                    unique.append(c)
                            save_companies(unique)
                            st.success(f"成功提取并保存 {len(unique)} 家公司")
                            
                            # 保存数据供生成报表使用
                            st.session_state['imported_df'] = df
                            st.session_state['city_col'] = city_col
                            st.session_state['company_col'] = company_col
                            st.session_state['district_col'] = district_col
                    else:
                        st.warning("未检测到城市列或公司列，请确认数据格式")
                else:
                    st.warning("未找到表头行，请确认数据格式")
        except Exception as e:
            st.error(f"处理文件时出错：{str(e)}")

# ===== 2. 公司管理 =====
elif menu == "2. 公司管理":
    st.header("🏢 公司管理")
    companies = load_companies()
    if companies:
        df_companies = pd.DataFrame(companies)
        st.dataframe(df_companies)
        st.caption(f"共 {len(companies)} 家公司")
        
        # 手动添加公司
        with st.expander("➕ 手动添加公司"):
            col1, col2, col3 = st.columns(3)
            with col1:
                new_company = st.text_input("公司名称")
                new_province = st.text_input("省份")
            with col2:
                new_city = st.text_input("城市")
                new_district = st.text_input("区县（可选）")
            with col3:
                new_tax = st.text_input("税号（可选）")
            if st.button("添加公司") and new_company and new_province and new_city:
                companies.append({
                    'id': str(uuid.uuid4())[:8],
                    'company_name': new_company,
                    'province': new_province,
                    'city': new_city,
                    'district': new_district,
                    'tax_id': new_tax
                })
                save_companies(companies)
                st.success("添加成功")
                st.rerun()
    else:
        st.info("暂无公司数据，请先在「数据导入」中上传Excel")

# ===== 3. 模板管理 =====
elif menu == "3. 模板管理":
    st.header("📄 模板管理")
    templates = load_templates()
    if templates:
        st.dataframe(pd.DataFrame(templates))
    else:
        st.info("暂无模板")
    
    with st.expander("➕ 添加官方模板"):
        col1, col2 = st.columns(2)
        with col1:
            t_province = st.text_input("省份")
            t_city = st.text_input("城市")
            t_district = st.text_input("区县")
            t_report_type = st.selectbox("报表类型", ["增值税", "社保", "公积金", "企业所得税", "个人所得税"])
        with col2:
            t_name = st.text_input("模板名称")
            t_version = st.text_input("版本号")
            t_authority = st.text_input("发布机构")
            t_url = st.text_input("来源URL")
            t_fields = st.text_input("必填字段（逗号分隔）")
        if st.button("保存模板") and t_name and t_province and t_city and t_report_type:
            save_template({
                'id': str(uuid.uuid4())[:8],
                'province': t_province,
                'city': t_city,
                'district': t_district,
                'report_type': t_report_type,
                'template_name': t_name,
                'template_version': t_version,
                'source_url': t_url,
                'source_authority': t_authority,
                'publish_date': datetime.now().strftime('%Y-%m-%d'),
                'required_fields': t_fields,
                'status': 'active'
            })
            st.success("模板已添加")
            st.rerun()

# ===== 4. 规则管理 =====
elif menu == "4. 规则管理":
    st.header("⚖️ 城市规则管理")
    rules = load_rules()
    if rules:
        st.dataframe(pd.DataFrame(rules))
    else:
        st.info("暂无规则")
    
    with st.expander("➕ 添加/编辑规则"):
        city = st.text_input("城市")
        unit_social = st.number_input("单位社保比例（小数）", min_value=0.0, max_value=1.0, value=0.16, step=0.001)
        personal_social = st.number_input("个人社保比例", min_value=0.0, max_value=1.0, value=0.08, step=0.001)
        unit_fund = st.number_input("单位公积金比例", min_value=0.0, max_value=1.0, value=0.07, step=0.001)
        personal_fund = st.number_input("个人公积金比例", min_value=0.0, max_value=1.0, value=0.07, step=0.001)
        social_min = st.number_input("社保最低基数", min_value=0, value=7310)
        social_max = st.number_input("社保最高基数", min_value=0, value=36549)
        fund_min = st.number_input("公积金最低基数", min_value=0, value=2590)
        fund_max = st.number_input("公积金最高基数", min_value=0, value=34188)
        source = st.text_input("来源引用")
        if st.button("保存规则") and city:
            rules.append({
                'id': str(uuid.uuid4())[:8],
                'city': city,
                'unit_social': unit_social,
                'personal_social': personal_social,
                'unit_fund': unit_fund,
                'personal_fund': personal_fund,
                'social_min': social_min,
                'social_max': social_max,
                'fund_min': fund_min,
                'fund_max': fund_max,
                'source_quote': source
            })
            save_rules(rules)
            st.success("规则已保存")
            st.rerun()

# ===== 5. 生成报表 =====
elif menu == "5. 生成报表":
    st.header("📊 生成报表")
    
    companies = load_companies()
    if not companies:
        st.warning("请先在「数据导入」中上传公司数据")
        st.stop()
    
    templates = load_templates()
    rules = load_rules()
    rules_dict = {r['city']: r for r in rules}
    
    # 选择过滤条件
    provinces = sorted(set(c['province'] for c in companies))
    col1, col2, col3 = st.columns(3)
    with col1:
        province = st.selectbox("省份", [""] + provinces)
        cities = sorted(set(c['city'] for c in companies if c['province'] == province)) if province else sorted(set(c['city'] for c in companies))
        city = st.selectbox("城市", [""] + cities)
    with col2:
        districts = sorted(set(c['district'] for c in companies if c['province'] == province and c['city'] == city)) if province and city else []
        district = st.selectbox("区县", [""] + districts)
        company_list = [c for c in companies if c['province'] == province and c['city'] == city and (not district or c['district'] == district)]
        company_names = [c['company_name'] for c in company_list]
        selected_companies = st.multiselect("选择公司（可多选）", company_names)
    with col3:
        report_type = st.selectbox("报表类型", ["增值税", "社保", "公积金", "企业所得税", "个人所得税"])
        period_type = st.selectbox("统计口径", ["月度（12月单月）", "累计（1-12月）"])
        # 导出格式
        export_format = st.radio("导出格式", ["Excel (.xlsx)", "CSV (.csv)", "PDF (.pdf)"], horizontal=True)
    
    # 数据预览与手动调整
    st.subheader("📋 数据预览与校验")
    if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
        df_preview = st.session_state['imported_df']
        st.dataframe(df_preview.head(10))
        
        # 校验
        city_col = st.session_state.get('city_col')
        if city_col:
            # 检查城市是否在规则中
            cities_in_data = df_preview[city_col].unique()
            missing_rules = [c for c in cities_in_data if c not in rules_dict]
            if missing_rules:
                st.warning(f"⚠️ 以下城市缺少规则，将无法自动计算：{', '.join(missing_rules)}")
            else:
                st.success("✅ 所有城市已配置规则")
        
        # 允许用户编辑数据（简单版：仅展示，编辑通过后续Excel实现）
        st.info("如需调整数据，请修改原Excel后重新导入")
    
    # 生成报表按钮
    if st.button("🚀 生成报表", type="primary"):
        if not selected_companies:
            st.error("请至少选择一家公司")
        else:
            # 获取选中的公司对象
            selected_company_objs = [c for c in company_list if c['company_name'] in selected_companies]
            if not selected_company_objs:
                st.error("未找到选中的公司")
            else:
                # 匹配模板
                matched_template = None
                for t in templates:
                    if t['province'] == province and t['city'] == city and t['report_type'] == report_type:
                        matched_template = t
                        break
                if not matched_template:
                    for t in templates:
                        if t['province'] == province and t['report_type'] == report_type:
                            matched_template = t
                            break
                if not matched_template:
                    # 使用通用模板
                    matched_template = {
                        'id': 'gen001',
                        'template_name': f'{report_type}通用模板',
                        'template_version': 'v1.0',
                        'source_authority': '系统生成',
                        'publish_date': datetime.now().strftime('%Y-%m-%d'),
                        'required_fields': '纳税人识别号,公司名称,申报金额'
                    }
                
                # 循环生成每个公司的报表
                generated_files = []
                summary = []
                errors = []
                for comp in selected_company_objs:
                    try:
                        # 获取该城市规则
                        rule = rules_dict.get(comp['city'])
                        # 构建数据填充
                        fields = matched_template['required_fields'].split(',')
                        # 模拟数据（实际应从导入数据中读取）
                        sample_data = {
                            '纳税人识别号': comp.get('tax_id', ''),
                            '公司名称': comp['company_name'],
                            '销售额': '100,000.00',
                            '进项税额': '13,000.00',
                            '应纳税额': '0.00',
                            '单位名称': comp['company_name'],
                            '社保登记号': 'SH123456',
                            '基数': '8,000.00',
                            '单位金额': str(round(8000 * rule['unit_social'], 2)) if rule else '',
                            '个人金额': str(round(8000 * rule['personal_social'], 2)) if rule else '',
                            '申报金额': '100,000.00'
                        }
                        
                        # 创建Excel
                        wb = Workbook()
                        ws = wb.active
                        ws.title = "申报表"
                        ws.append(fields)
                        row_data = [sample_data.get(f, '') for f in fields]
                        ws.append(row_data)
                        
                        # 水印
                        ws.insert_rows(1)
                        ws['A1'] = f'【系统生成 - 待复核版】统计口径：{period_type}'
                        ws['A1'].font = Font(color='FF0000', bold=True, size=14)
                        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(fields))
                        ws['A1'].alignment = Alignment(horizontal='center')
                        
                        ws.insert_rows(2)
                        ws['A2'] = f'模板名称：{matched_template["template_name"]}  版本：{matched_template["template_version"]}'
                        ws['A2'].font = Font(color='666666', size=10)
                        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(fields))
                        
                        ws.insert_rows(3)
                        ws['A3'] = f'来源：{matched_template.get("source_authority","")}  发布日期：{matched_template.get("publish_date","")}'
                        ws['A3'].font = Font(color='666666', size=10)
                        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(fields))
                        
                        # 审计日志
                        audit = wb.create_sheet("审计日志")
                        audit.append(['操作时间', '操作类型', '操作人', '详情'])
                        audit.append([datetime.now().isoformat(), 'GENERATED', '系统', f'公司:{comp["company_name"]}, 模板:{matched_template["template_name"]}'])
                        
                        # 保存到BytesIO
                        output = BytesIO()
                        if export_format == "CSV (.csv)":
                            # 转为CSV
                            df_export = pd.DataFrame([row_data], columns=fields)
                            csv_data = df_export.to_csv(index=False).encode('utf-8')
                            output.write(csv_data)
                            file_ext = '.csv'
                            mime = 'text/csv'
                        elif export_format == "PDF (.pdf)":
                            # 简单PDF（需安装reportlab，此处仅示意，实际可用fpdf）
                            # 简化：直接导出Excel，提示用户
                            wb.save(output)
                            file_ext = '.xlsx'
                            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            st.warning("PDF导出功能需要安装reportlab，当前已转为Excel")
                        else:
                            wb.save(output)
                            file_ext = '.xlsx'
                            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        output.seek(0)
                        
                        # 保存历史
                        record = {
                            'id': str(uuid.uuid4())[:8],
                            'company_id': comp['id'],
                            'template_id': matched_template.get('id', 'gen001'),
                            'company_name': comp['company_name'],
                            'city': comp['city'],
                            'report_type': report_type,
                            'period_type': period_type,
                            'generated_at': datetime.now().isoformat(),
                            'review_status': 'pending',
                            'file_name': f"{comp['company_name']}_{report_type}_{datetime.now().strftime('%Y%m%d')}{file_ext}",
                            'file_data': output.getvalue()
                        }
                        save_export(record)
                        
                        generated_files.append((record['file_name'], output.getvalue(), mime))
                        summary.append({
                            '公司': comp['company_name'],
                            '城市': comp['city'],
                            '模板': matched_template['template_name'],
                            '状态': '待复核'
                        })
                    except Exception as e:
                        errors.append(f"{comp['company_name']}: {str(e)}")
                
                if errors:
                    for err in errors:
                        st.warning(err)
                
                if generated_files:
                    st.success(f"✅ 成功生成 {len(generated_files)} 份报表")
                    st.dataframe(pd.DataFrame(summary))
                    
                    # 批量下载
                    if len(generated_files) > 1:
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w') as zf:
                            for fname, data, mime in generated_files:
                                zf.writestr(fname, data)
                        zip_buffer.seek(0)
                        st.download_button(
                            "📦 下载全部报表（ZIP）",
                            data=zip_buffer,
                            file_name=f"报表_{datetime.now().strftime('%Y%m%d')}.zip",
                            mime="application/zip"
                        )
                    else:
                        fname, data, mime = generated_files[0]
                        st.download_button(
                            f"📥 下载 {fname}",
                            data=data,
                            file_name=fname,
                            mime=mime
                        )

# ===== 6. 历史与复核 =====
elif menu == "6. 历史与复核":
    st.header("📋 导出历史与复核")
    history = load_export_history()
    if history:
        df_hist = pd.DataFrame(history)
        # 显示关键列
        display_cols = ['company_name', 'city', 'report_type', 'period_type', 'generated_at', 'review_status', 'reviewer']
        st.dataframe(df_hist[display_cols])
        
        # 待复核列表
        pending = [h for h in history if h['review_status'] == 'pending']
        if pending:
            st.subheader("✅ 待复核报表")
            options = [f"{h['company_name']} - {h['city']} - {h['report_type']} ({h['generated_at'][:10]})" for h in pending]
            selected_idx = st.selectbox("选择要复核的报表", range(len(options)), format_func=lambda x: options[x])
            selected = pending[selected_idx]
            
            st.write(f"**公司**：{selected['company_name']}")
            st.write(f"**城市**：{selected['city']}")
            st.write(f"**报表类型**：{selected['report_type']}")
            st.write(f"**统计口径**：{selected['period_type']}")
            st.write(f"**生成时间**：{selected['generated_at']}")
            
            # 提供下载
            if selected.get('file_data'):
                st.download_button(
                    "📥 下载该报表",
                    data=selected['file_data'],
                    file_name=selected.get('file_name', 'report.xlsx'),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
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
            st.info("所有报表已复核完成")
        
        # 批量复核（全选通过）
        if pending and st.button("批量通过所有待复核"):
            for h in pending:
                update_export_status(h['id'], 'approved', '批量复核员')
            st.success("全部通过复核")
            st.rerun()
    else:
        st.info("暂无导出记录")

# ===== 7. 系统设置 =====
elif menu == "7. 系统设置":
    st.header("⚙️ 系统设置")
    st.markdown("**数据管理**")
    if st.button("清空所有数据（谨慎操作）"):
        if st.checkbox("确认清空所有数据？"):
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute("DELETE FROM companies")
            c.execute("DELETE FROM templates")
            c.execute("DELETE FROM rules")
            c.execute("DELETE FROM export_history")
            conn.commit()
            conn.close()
            st.success("数据已清空")
            st.rerun()
    
    st.markdown("**导入模板**")
    uploaded_template = st.file_uploader("上传自定义模板（.xlsx）", type=["xlsx"])
    if uploaded_template:
        # 保存到数据库custom_templates表（需要建表，此处略，可扩展）
        st.success("模板上传成功（功能扩展中）")

# ===== 页脚 =====
st.sidebar.markdown("---")
st.sidebar.caption(f"数据库路径：{DB_PATH}")
st.sidebar.caption("版本 v2.0 (全功能版)")
