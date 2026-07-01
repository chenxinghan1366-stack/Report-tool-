import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO
import uuid
import sqlite3
import os
import zipfile

# ========== 数据库初始化 ==========
DB_PATH = os.path.join(os.path.dirname(__file__), "app_data.db")

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
        id TEXT PRIMARY KEY, city TEXT, unit_social REAL, personal_social REAL,
        unit_fund REAL, personal_fund REAL, social_min REAL, social_max REAL,
        fund_min REAL, fund_max REAL, source_quote TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS export_history (
        id TEXT PRIMARY KEY, company_id TEXT, template_id TEXT, company_name TEXT,
        city TEXT, report_type TEXT, period_type TEXT, generated_at TEXT,
        review_status TEXT, reviewer TEXT, reviewed_at TEXT, file_name TEXT, file_data BLOB
    )''')
    conn.commit()
    conn.close()

init_db()

# ========== 数据操作函数 ==========
def dict_fetchall(cursor):
    columns = [col[0] for col in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]

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

def update_export_status(export_id, status, reviewer):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''UPDATE export_history 
        SET review_status=?, reviewer=?, reviewed_at=?
        WHERE id=?''',
        (status, reviewer, datetime.now().isoformat(), export_id))
    conn.commit()
    conn.close()

# ========== 16个城市完整规则 ==========
ALL_CITY_RULES = [
    {'id': 'r001', 'city': '上海', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.07, 'personal_fund': 0.07, 'social_min': 7310, 'social_max': 36549,
     'fund_min': 2590, 'fund_max': 34188, 'source_quote': '沪人社规〔2024〕22号'},
    {'id': 'r002', 'city': '北京', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 6326, 'social_max': 33891,
     'fund_min': 2420, 'fund_max': 33891, 'source_quote': '京人社发〔2024〕15号'},
    {'id': 'r003', 'city': '广州', 'unit_social': 0.15, 'personal_social': 0.08,
     'unit_fund': 0.10, 'personal_fund': 0.10, 'social_min': 4588, 'social_max': 22941,
     'fund_min': 2300, 'fund_max': 27960, 'source_quote': '粤人社规〔2024〕8号'},
    {'id': 'r004', 'city': '深圳', 'unit_social': 0.15, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 2360, 'social_max': 22941,
     'fund_min': 2360, 'fund_max': 27927, 'source_quote': '深人社规〔2024〕3号'},
    {'id': 'r005', 'city': '杭州', 'unit_social': 0.15, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3957, 'social_max': 22941,
     'fund_min': 2280, 'fund_max': 27874, 'source_quote': '浙人社发〔2024〕7号'},
    {'id': 'r006', 'city': '南京', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.08, 'personal_fund': 0.08, 'social_min': 4250, 'social_max': 22470,
     'fund_min': 2280, 'fund_max': 27841, 'source_quote': '苏人社发〔2024〕6号'},
    {'id': 'r007', 'city': '成都', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 4071, 'social_max': 20355,
     'fund_min': 2100, 'fund_max': 25401, 'source_quote': '川人社发〔2024〕9号'},
    {'id': 'r008', 'city': '重庆', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3957, 'social_max': 19784,
     'fund_min': 2100, 'fund_max': 24595, 'source_quote': '渝人社发〔2024〕5号'},
    {'id': 'r009', 'city': '天津', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.11, 'personal_fund': 0.11, 'social_min': 4400, 'social_max': 22434,
     'fund_min': 2180, 'fund_max': 24240, 'source_quote': '津人社发〔2024〕4号'},
    {'id': 'r010', 'city': '苏州', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 4250, 'social_max': 22470,
     'fund_min': 2280, 'fund_max': 27874, 'source_quote': '苏人社发〔2024〕6号'},
    {'id': 'r011', 'city': '武汉', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 4077, 'social_max': 20385,
     'fund_min': 2010, 'fund_max': 24114, 'source_quote': '鄂人社发〔2024〕5号'},
    {'id': 'r012', 'city': '西安', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.10, 'personal_fund': 0.10, 'social_min': 3957, 'social_max': 19784,
     'fund_min': 1950, 'fund_max': 23556, 'source_quote': '陕人社发〔2024〕4号'},
    {'id': 'r013', 'city': '郑州', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.10, 'personal_fund': 0.10, 'social_min': 3409, 'social_max': 17043,
     'fund_min': 2000, 'fund_max': 22892, 'source_quote': '豫人社发〔2024〕3号'},
    {'id': 'r014', 'city': '长沙', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3604, 'social_max': 18018,
     'fund_min': 1930, 'fund_max': 22998, 'source_quote': '湘人社发〔2024〕5号'},
    {'id': 'r015', 'city': '青岛', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3746, 'social_max': 18726,
     'fund_min': 2010, 'fund_max': 23496, 'source_quote': '鲁人社发〔2024〕6号'},
    {'id': 'r016', 'city': '宁波', 'unit_social': 0.15, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3957, 'social_max': 22941,
     'fund_min': 2280, 'fund_max': 27874, 'source_quote': '浙人社发〔2024〕7号'},
]

# ========== 初始化默认数据 ==========
def init_default_data():
    # 预置16个城市的规则
    existing_rules = load_rules()
    if not existing_rules:
        save_rules(ALL_CITY_RULES)
    
    if not load_templates():
        templates = [
            {'id': 't001', 'province': '上海', 'city': '上海市', 'district': '浦东新区',
             'report_type': '增值税', 'template_name': '上海市增值税纳税申报表（一般纳税人）',
             'template_version': 'v2024.1', 'source_url': 'https://shanghai.chinatax.gov.cn/bsfw/bszn/2024/zzs.xlsx',
             'source_authority': '国家税务总局上海市税务局', 'publish_date': '2024-01-15',
             'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额', 'status': 'active'},
            {'id': 't002', 'province': '上海', 'city': '上海市', 'district': '浦东新区',
             'report_type': '社保', 'template_name': '上海市社会保险费申报表（月度）',
             'template_version': 'v2024.1', 'source_url': 'https://rsj.sh.gov.cn/sbjb/2024/sb.xlsx',
             'source_authority': '上海市人力资源和社会保障局', 'publish_date': '2024-01-10',
             'required_fields': '单位名称,社保登记号,基数,单位金额,个人金额', 'status': 'active'},
            {'id': 't003', 'province': '广东', 'city': '广州市', 'district': '天河区',
             'report_type': '增值税', 'template_name': '广东省增值税纳税申报表',
             'template_version': 'v2024.1', 'source_url': 'https://guangdong.chinatax.gov.cn/bsfw/2024/zzs.xlsx',
             'source_authority': '国家税务总局广东省税务局', 'publish_date': '2024-01-20',
             'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额', 'status': 'active'},
            {'id': 't004', 'province': '北京', 'city': '北京市', 'district': '海淀区',
             'report_type': '增值税', 'template_name': '北京市增值税纳税申报表（一般纳税人）',
             'template_version': 'v2024.2', 'source_url': 'https://beijing.chinatax.gov.cn/bsfw/2024/zzs.xlsx',
             'source_authority': '国家税务总局北京市税务局', 'publish_date': '2024-02-01',
             'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额', 'status': 'active'},
        ]
        for t in templates:
            save_template(t)

init_default_data()

# ========== 解析上传的Excel ==========
def parse_uploaded_excel(file):
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    all_companies = []
    city_to_province = {
        '上海': '上海', '北京': '北京', '广州': '广东', '深圳': '广东',
        '杭州': '浙江', '南京': '江苏', '苏州': '江苏', '成都': '四川',
        '重庆': '重庆', '武汉': '湖北', '西安': '陕西', '郑州': '河南',
        '长沙': '湖南', '青岛': '山东', '宁波': '浙江', '天津': '天津'
    }
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
                            province = city_to_province.get(city, city)
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

# ========== 自动导入规则（从基础配置表） ==========
def import_rules_from_excel(file):
    try:
        xls = pd.ExcelFile(file)
        if '基础配置表' not in xls.sheet_names:
            return None, "未找到「基础配置表」"
        
        # 读取基础配置表
        df_rules = pd.read_excel(file, sheet_name='基础配置表', header=None)
        
        # 找表头行
        header_row = None
        for i, row in df_rules.iterrows():
            row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
            if '所属城市' in row_text or '城市' in row_text:
                header_row = i
                break
        
        if header_row is None:
            return None, "未找到表头行"
        
        df_rules = pd.read_excel(file, sheet_name='基础配置表', skiprows=header_row)
        df_rules.columns = [str(c).strip() for c in df_rules.columns]
        
        # 列名映射
        col_city = None
        col_unit_social = None
        col_personal_social = None
        col_unit_fund = None
        col_personal_fund = None
        col_social_min = None
        col_social_max = None
        col_fund_min = None
        col_fund_max = None
        
        for col in df_rules.columns:
            if '城市' in col:
                col_city = col
            elif '养老保险-单位' in col or '单位养老' in col:
                col_unit_social = col
            elif '养老保险-个人' in col or '个人养老' in col:
                col_personal_social = col
            elif '公积金-单位' in col or '单位公积金' in col:
                col_unit_fund = col
            elif '公积金-个人' in col or '个人公积金' in col:
                col_personal_fund = col
            elif '缴费基数下限' in col:
                col_social_min = col
            elif '缴费基数上限' in col:
                col_social_max = col
        
        if col_city is None or col_unit_social is None:
            return None, "缺少必要列（城市、养老保险-单位比例）"
        
        rules = []
        for _, row in df_rules.iterrows():
            city = row[col_city]
            if pd.isna(city):
                continue
            try:
                unit_social = float(row[col_unit_social]) if col_unit_social and pd.notna(row[col_unit_social]) else 0.16
                personal_social = float(row[col_personal_social]) if col_personal_social and pd.notna(row[col_personal_social]) else 0.08
                unit_fund = float(row[col_unit_fund]) if col_unit_fund and pd.notna(row[col_unit_fund]) else 0.12
                personal_fund = float(row[col_personal_fund]) if col_personal_fund and pd.notna(row[col_personal_fund]) else 0.12
                social_min = float(row[col_social_min]) if col_social_min and pd.notna(row[col_social_min]) else 0
                social_max = float(row[col_social_max]) if col_social_max and pd.notna(row[col_social_max]) else 999999
            except:
                continue
            rules.append({
                'id': str(uuid.uuid4())[:8],
                'city': city,
                'unit_social': unit_social,
                'personal_social': personal_social,
                'unit_fund': unit_fund,
                'personal_fund': personal_fund,
                'social_min': social_min,
                'social_max': social_max,
                'fund_min': 0,
                'fund_max': 999999,
                'source_quote': '从基础配置表导入'
            })
        if rules:
            save_rules(rules)
            return len(rules), None
        return None, "未解析到有效数据"
    except Exception as e:
        return None, str(e)

# ========== Streamlit 页面 ==========
st.set_page_config(page_title="官方模板匹配器", layout="wide")
st.title("📋 官方模板匹配器（含自动识别与统计口径）")
st.markdown("**上传Excel → 自动提取城市/公司 → 选择模板和统计口径 → 生成待复核版Excel**")

# ===== 侧边栏 =====
with st.sidebar:
    st.header("📤 上传数据Excel")
    st.markdown("上传包含公司/城市信息的Excel，系统自动提取所有地区，并自动导入规则")
    uploaded_file = st.file_uploader("选择Excel文件（.xlsx）", type=["xlsx"])
    
    if uploaded_file:
        with st.spinner("正在解析Excel..."):
            # 1. 提取公司
            companies = parse_uploaded_excel(uploaded_file)
            if companies:
                save_companies(companies)
                st.success(f"成功提取 {len(companies)} 家公司")
                
                # 2. 导入规则
                rule_count, error = import_rules_from_excel(uploaded_file)
                if rule_count:
                    st.success(f"✅ 成功从「基础配置表」导入 {rule_count
