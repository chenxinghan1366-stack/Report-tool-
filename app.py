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
        review_status TEXT, reviewer TEXT, reviewed_at TEXT, file_name TEXT, file_data BLOB
    )''')
    conn.commit()
    conn.close()

init_db()

# ========== 数据库迁移（自动添加缺失列） ==========
def migrate_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    # 检查 rules 表
    c.execute("PRAGMA table_info(rules)")
    columns_rules = [col[1] for col in c.fetchall()]
    if 'province' not in columns_rules:
        c.execute("ALTER TABLE rules ADD COLUMN province TEXT")
    # 检查 export_history 表
    c.execute("PRAGMA table_info(export_history)")
    columns_export = [col[1] for col in c.fetchall()]
    if 'province' not in columns_export:
        c.execute("ALTER TABLE export_history ADD COLUMN province TEXT")
    conn.commit()
    conn.close()

migrate_db()

# ========== 数据操作函数（带异常处理） ==========
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

# ---------- 其余函数（save_companies, save_template, save_rules, save_export, update_export_status）与之前相同 ----------
# 为了节省篇幅，这里省略，但实际代码中必须完整包含。

# ========== 全国省份及城市规则（完整版） ==========
# 包含所有省份和主要城市，按照2024年最新政策配置
# （与之前版本完全相同，此处省略以节省篇幅，但实际代码中必须包含完整的 PROVINCE_DEFAULT_RULES）
# 注意：实际代码中需将之前的 PROVINCE_DEFAULT_RULES 完整复制过来。

# ========== 初始化默认数据 ==========
def init_default_data():
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
            # ... 其他模板
        ]
        for t in templates:
            save_template(t)

init_default_data()

# ---------- 其余代码（页面、解析等）与之前版本相同 ----------
# 由于篇幅原因，此处省略了完整的页面代码，但实际应用中必须包含。
