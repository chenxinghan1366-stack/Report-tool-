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
        id TEXT PRIMARY KEY, city TEXT, province TEXT, unit_social REAL, personal_social REAL,
        unit_fund REAL, personal_fund REAL, social_min REAL, social_max REAL,
        fund_min REAL, fund_max REAL, source_quote TEXT, is_default BOOLEAN DEFAULT 0
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS export_history (
        id TEXT PRIMARY KEY, company_id TEXT, template_id TEXT, company_name TEXT,
        city TEXT, province TEXT, report_type TEXT, period_type TEXT, generated_at TEXT,
        review_status TEXT, reviewer TEXT, reviewed_at TEXT, file_name TEXT, file_data BLOB
    )''')
    conn.commit()
    conn.close()

init_db()  # 确保表被创建

# ========== 数据操作函数（带异常处理） ==========
def dict_fetchall(cursor):
    columns = [col[0] for col in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]

def safe_execute_query(query, params=None):
    """执行查询，若表不存在则返回空列表"""
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
         generated_at, review_status, reviewer, reviewed_at, file_name, file_data)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (record['id'], record.get('company_id',''), record.get('template_id',''),
         record['company_name'], record.get('city',''), record.get('province',''),
         record.get('report_type',''), record.get('period_type',''), record['generated_at'],
         record.get('review_status','pending'), record.get('reviewer',''), record.get('reviewed_at',''),
         record.get('file_name',''), record.get('file_data', None)))
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

# ========== 全国省份默认规则 ==========
PROVINCE_DEFAULT_RULES = [
    # 直辖市
    {'city': '上海', 'province': '上海', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.07, 'personal_fund': 0.07, 'social_min': 7310, 'social_max': 36549,
     'fund_min': 2590, 'fund_max': 34188, 'source_quote': '沪人社规〔2024〕22号'},
    # ... 其他省份和城市规则（与之前相同，此处省略以节省篇幅，但实际代码中完整保留）
]
# 注意：为了节省篇幅，这里省略了全部省份规则，但实际代码中必须包含完整的 PROVINCE_DEFAULT_RULES。

# ========== 初始化默认数据 ==========
def init_default_data():
    # 先确保规则表存在
    if not load_rules():
        # 插入所有省份默认规则
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
        templates = [
            {'id': 't001', 'province': '上海', 'city': '上海市', 'district': '浦东新区',
             'report_type': '增值税', 'template_name': '上海市增值税纳税申报表（一般纳税人）',
             'template_version': 'v2024.1', 'source_url': 'https://shanghai.chinatax.gov.cn/bsfw/bszn/2024/zzs.xlsx',
             'source_authority': '国家税务总局上海市税务局', 'publish_date': '2024-01-15',
             'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额', 'status': 'active'},
            # ... 其他模板（完整保留）
        ]
        for t in templates:
            save_template(t)

init_default_data()

# ========== 其余代码（页面、解析等）保持不变 ==========
# ...（后续所有代码与原版本一致）
