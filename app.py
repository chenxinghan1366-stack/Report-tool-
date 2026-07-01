import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook, load_workbook
from io import BytesIO
import uuid
import sqlite3
import os
import zipfile
import re

# ---------- 数据库路径 ----------
DB_PATH = os.path.join(os.path.dirname(__file__), "data.db")

# ---------- 初始化数据库 ----------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS companies (
        id TEXT PRIMARY KEY, company_name TEXT, province TEXT, city TEXT, district TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS rules (
        id TEXT PRIMARY KEY, province TEXT, city TEXT, report_type TEXT,
        unit_social REAL, personal_social REAL, unit_fund REAL, personal_fund REAL,
        social_min REAL, social_max REAL, fund_min REAL, fund_max REAL, source_quote TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS export_history (
        id TEXT PRIMARY KEY, company TEXT, city TEXT, report_type TEXT,
        year INTEGER, month INTEGER, generated_at TEXT, status TEXT,
        source_quote TEXT, total_people INTEGER, total_cost REAL,
        reviewer TEXT, reviewed_at TEXT, review_comment TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS audit_logs (
        id TEXT PRIMARY KEY, timestamp TEXT, action TEXT, detail TEXT, report_id TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS custom_templates (
        id TEXT PRIMARY KEY, name TEXT, file_data BLOB, field_mapping TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS employee_data (
        id TEXT PRIMARY KEY, company TEXT, city TEXT, name TEXT,
        base_salary REAL, social_base REAL, fund_base REAL,
        last_updated TEXT
    )''')
    conn.commit()
    conn.close()

# ---------- 数据库辅助 ----------
def load_table(table_name):
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql(f"SELECT * FROM {table_name}", conn)
    conn.close()
    return df.to_dict('records') if not df.empty else []

def save_table(table_name, records):
    if not records: return
    conn = sqlite3.connect(DB_PATH)
    pd.DataFrame(records).to_sql(table_name, conn, if_exists='replace', index=False)
    conn.close()

def insert_or_update_export(record):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO export_history 
        (id, company, city, report_type, year, month, generated_at, status, source_quote, 
         total_people, total_cost, reviewer, reviewed_at, review_comment)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (record['id'], record['company'], record['city'], record['report_type'],
         record['year'], record['month'], record['generated_at'], record['status'],
         record.get('source_quote',''), record.get('total_people',0), record.get('total_cost',0.0),
         record.get('reviewer',''), record.get('reviewed_at',''), record.get('review_comment','')))
    conn.commit()
    conn.close()

def insert_audit_log(log):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT INTO audit_logs (id, timestamp, action, detail, report_id)
        VALUES (?,?,?,?,?)''',
        (log['id'], log['timestamp'], log['action'], log['detail'], log.get('report_id','')))
    conn.commit()
    conn.close()

def get_export_history():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql("SELECT * FROM export_history ORDER BY generated_at DESC", conn)
    conn.close()
    return df.to_dict('records') if not df.empty else []

def get_audit_logs():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql("SELECT * FROM audit_logs ORDER BY timestamp DESC", conn)
    conn.close()
    return df.to_dict('records') if not df.empty else []

# ---------- 智能列名映射 ----------
def smart_column_mapping(df, required_cols, user_mapping=None):
    df_cols = list(df.columns)
    synonyms = {
        '公司': ['公司', '企业', '单位名称', '公司名称', '企业名称', '单位名称', 'name', 'company', '公司名', '所属公司', '公司全称'],
        '城市': ['城市', '市', 'city', '地区', '所属城市', '所在地', '城市名', '城市名称'],
        '姓名': ['姓名', '名字', '员工', 'name', 'employee', '人员', '员工编号', '员工姓名', '员工id'],
        '工资基数': ['工资基数', '基数', '工资', '月工资', '基本工资', 'base', 'salary', '月应发工资', '应发工资', '月薪', '工资总额', '月度工资'],
        '社保基数': ['社保基数', '养老基数', '社保', 'social_security', '社保缴费基数', '社会保险基数'],
        '公积金基数': ['公积金基数', '公积金', 'fund', '公积金缴费基数', '住房公积金基数']
    }
    
    if user_mapping:
        mapping = {k: v for k, v in user_mapping.items() if v}
        for std_name in required_cols:
            if std_name not in mapping or not mapping[std_name]:
                for df_col in df_cols:
                    if df_col in mapping.values():
                        continue
                    if df_col == std_name:
                        mapping[std_name] = df_col
                        break
                    for syn in synonyms.get(std_name, []):
                        if syn.lower() in df_col.lower() or df_col.lower() in syn.lower():
                            mapping[std_name] = df_col
                            break
                    if std_name in mapping and mapping[std_name]:
                        break
        return mapping
    else:
        mapping = {}
        for std_name in required_cols:
            matched = False
            for df_col in df_cols:
                if df_col == std_name:
                    mapping[std_name] = df_col
                    matched = True
                    break
                for syn in synonyms.get(std_name, []):
                    if syn.lower() in df_col.lower() or df_col.lower() in syn.lower():
                        mapping[std_name] = df_col
                        matched = True
                        break
                if matched:
                    break
            if std_name not in mapping:
                mapping[std_name] = None
        return mapping

# ===== 自动导入规则（从Excel的"基础配置表"） =====
def import_rules_from_excel(xls):
    """
    从Excel的"基础配置表"Sheet导入城市规则
    返回导入的城市数量
    """
    # 支持多种可能的Sheet名称
    possible_sheets = ["基础配置表", "规则", "城市规则", "规则表", "配置表"]
    rule_sheet = None
    for sheet in possible_sheets:
        if sheet in xls.sheet_names:
            rule_sheet = sheet
            break
    
    if rule_sheet is None:
        return 0
    
    df_rules = pd.read_excel(xls, sheet_name=rule_sheet)
    # 寻找表头行：包含"城市"、"社保基数下限"等关键字
    header_row = 0
    for i, row in df_rules.iterrows():
        row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
        if '城市' in row_text and ('社保基数' in row_text or '基数' in row_text):
            header_row = i
            break
    # 重新读取并跳过前header_row行
    df_rules = pd.read_excel(xls, sheet_name=rule_sheet, skiprows=header_row)
    # 列名映射（支持多种列名变体）
    col_map = {}
    for col in df_rules.columns:
        col_lower = str(col).lower()
        if '城市' in col_lower:
            col_map['城市'] = col
        elif '社保基数下限' in col_lower or '社保最低基数' in col_lower or '下限' in col_lower:
            col_map['社保最低基数'] = col
        elif '社保基数上限' in col_lower or '社保最高基数' in col_lower or '上限' in col_lower:
            col_map['社保最高基数'] = col
        elif '养老单位比例' in col_lower or '单位养老' in col_lower or '单位社保' in col_lower:
            col_map['单位社保比例'] = col
        elif '养老个人比例' in col_lower or '个人养老' in col_lower or '个人社保' in col_lower:
            col_map['个人社保比例'] = col
        elif '公积金单位比例' in col_lower or '单位公积金' in col_lower:
            col_map['单位公积金比例'] = col
        elif '公积金个人比例' in col_lower or '个人公积金' in col_lower:
            col_map['个人公积金比例'] = col
        elif '公积金基数下限' in col_lower:
            col_map['公积金最低基数'] = col
        elif '公积金基数上限' in col_lower:
            col_map['公积金最高基数'] = col
    
    # 检查必要列是否都存在
    required = ['城市', '单位社保比例', '个人社保比例', '单位公积金比例', '个人公积金比例',
                '社保最低基数', '社保最高基数', '公积金最低基数', '公积金最高基数']
    # 检查是否所有必需列都已映射
    missing = [r for r in required if r not in col_map]
    if missing:
        return 0
    
    # 构建规则列表
    rules_list = []
    for _, row in df_rules.iterrows():
        city = row[col_map['城市']]
        if pd.isna(city):
            continue
        try:
            rules_list.append({
                "id": str(uuid.uuid4())[:8],
                "province": city,
                "city": city,
                "report_type": "月度申报",
                "unit_social": float(row[col_map['单位社保比例']]),
                "personal_social": float(row[col_map['个人社保比例']]),
                "unit_fund": float(row[col_map['单位公积金比例']]),
                "personal_fund": float(row[col_map['个人公积金比例']]),
                "social_min": float(row[col_map['社保最低基数']]),
                "social_max": float(row[col_map['社保最高基数']]),
                "fund_min": float(row[col_map['公积金最低基数']]),
                "fund_max": float(row[col_map['公积金最高基数']]),
                "source_quote": f"自动从{rule_sheet}导入"
            })
        except (ValueError, TypeError):
            continue
    
    if rules_list:
        save_table('rules', rules_list)
        return len(rules_list)
    return 0

# ---------- 初始化 ----------
init_db()
if not load_table('companies'):
    save_table('companies', [
        {"id": "c001", "company_name": "上海科技公司", "province": "上海", "city": "上海市", "district": "浦东新区"},
        {"id": "c002", "company_name": "深圳科技公司", "province": "广东", "city": "深圳市", "district": "南山区"}
    ])
if not load_table('rules'):
    save_table('rules', [])

# ---------- Streamlit 页面 ----------
st.set_page_config(page_title="本地社保报表系统", layout="wide")
st.title("📋 本地社保公积金报表系统（全自动版）")
st.markdown("🔒 **上传Excel后自动识别「基础配置表」导入规则，自动识别公司列表**")

# 仪表盘
history = get_export_history()
companies = load_table('companies')
rules = load_table('rules')
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("🏢 公司", len(companies))
with col2:
    st.metric("🏙️ 城市", len(set(c.get('city','') for c in companies)))
with col3:
    st.metric("📋 报表总数", len(history))
with col4:
    pending = len([h for h in history if h.get('status') == '待复核'])
    st.metric("🟡 待复核", pending)

tab1, tab2, tab3, tab4 = st.tabs(["📊 生成报表", "📋 报表历史与复核", "✏️ 数据管理", "📄 自定义模板"])

# ==================== 生成报表 ====================
with tab1:
    companies_data = load_table('companies')
    rules_data = load_table('rules')
    custom_templates = load_table('custom_templates')
    
    st.subheader("📤 1. 上传数据（自动导入公司 + 规则）")
    st.caption("上传Excel后，系统自动提取公司列表，并自动识别「基础配置表」导入城市规则")
    uploaded = st.file_uploader("选择文件（.xlsx, .xls, .csv, .txt）", type=["xlsx", "xls", "csv", "txt"])
    
    df_raw = None
    df_std = None
    sheets = []
    selected_sheet = None
    rules_imported = False
    
    if uploaded:
        try:
            file_ext = uploaded.name.split('.')[-1].lower()
            if file_ext in ['xlsx', 'xls']:
                xls = pd.ExcelFile(uploaded)
                sheets = xls.sheet_names
                selected_sheet = st.selectbox("选择Sheet", sheets, index=0)
                if selected_sheet:
                    df_raw = pd.read_excel(uploaded, sheet_name=selected_sheet)
                
                # ===== 自动导入规则（检测基础配置表） =====
                with st.spinner("正在自动识别并导入城市规则..."):
                    count = import_rules_from_excel(xls)
                    if count > 0:
                        st.success(f"✅ 成功从「基础配置表」导入 {count} 个城市的规则！")
                        rules_imported = True
                        rules_data = load_table('rules')
                    else:
                        st.info("未检测到有效的「基础配置表」，如需规则请手动在「管理数据」中导入。")
            elif file_ext == 'csv':
                df_raw = pd.read_csv(uploaded, encoding='utf-8-sig')
            elif file_ext == 'txt':
                df_raw = pd.read_csv(uploaded, delimiter='\t', encoding='utf-8-sig')
            
            if df_raw is not None:
                st.info(f"📊 读取到 {len(df_raw)} 行，{len(df_raw.columns)} 列")
                st.dataframe(df_raw.head(5))
                
                # 自动检测表头行
                header_row = 0
                for i, row in df_raw.iterrows():
                    row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                    if any(key in row_text for key in ['公司', '城市', '姓名', '工资', '基数', '员工', '企业']):
                        header_row = i
                        break
                
                if header_row == 0:
                    header_row = st.number_input("表头行号（从0开始）", min_value=0, max_value=len(df_raw)-1, value=0, step=1)
                else:
                    st.info(f"自动检测到表头行：第 {header_row} 行")
                    override = st.checkbox("手动调整表头行号")
                    if override:
                        header_row = st.number_input("表头行号（从0开始）", min_value=0, max_value=len(df_raw)-1, value=header_row, step=1)
                
                if header_row < len(df_raw):
                    new_header = df_raw.iloc[header_row]
                    df_raw = df_raw[header_row+1:]
                    df_raw.columns = new_header
                    df_raw = df_raw.reset_index(drop=True)
                    df_raw = df_raw.dropna(how='all')
                    st.success(f"成功提取数据，共 {len(df_raw)} 行")
                    st.dataframe(df_raw.head(5))
                    
                    # ===== 列名映射 =====
                    st.subheader("🔍 2. 列名映射（用于生成报表）")
                    required_cols = ['公司', '城市', '姓名', '工资基数']
                    col_mapping = smart_column_mapping(df_raw, required_cols)
                    
                    for std_name, actual_col in col_mapping.items():
                        if actual_col is None or actual_col == '':
                            options = [''] + list(df_raw.columns)
                            selected = st.selectbox(f"请选择 '{std_name}' 对应的列", options, key=f"map_{std_name}_{uploaded.name}")
                            if selected:
                                col_mapping[std_name] = selected
                    
                    if col_mapping.get('公司') and col_mapping.get('工资基数'):
                        st.success("✅ 列映射完成！")
                        std_data = {}
                        for std_name, actual_col in col_mapping.items():
                            if actual_col and actual_col in df_raw.columns:
                                std_data[std_name] = df_raw[actual_col]
                        df_std = pd.DataFrame(std_data)
                        
                        # 处理姓名
                        if '姓名' not in df_std.columns or df_std['姓名'].isnull().all():
                            for col in df_raw.columns:
                                if '编号' in col or 'ID' in col or '员工' in col:
                                    df_std['姓名'] = df_raw[col]
                                    break
                        if '姓名' not in df_std.columns or df_std['姓名'].isnull().all():
                            df_std['姓名'] = [f"员工{i+1}" for i in range(len(df_std))]
                        
                        # 处理工资基数
                        if '工资基数' in df_std.columns:
                            df_std['工资基数'] = pd.to_numeric(df_std['工资基数'], errors='coerce')
                            df_std = df_std.dropna(subset=['工资基数'])
                            st.success(f"数据准备完成，共 {len(df_std)} 条有效记录")
                            st.dataframe(df_std.head(5))
                        else:
                            st.error("未找到工资基数列")
                            df_std = None
                    else:
                        st.warning("列映射未完成，请确保'公司'和'工资基数'已映射")
        except Exception as e:
            st.error(f"❌ 读取文件失败：{str(e)}")
    
    # ===== 选择公司和生成报表（直接从数据中提取公司列表） =====
    st.subheader("📤 3. 选择公司并生成报表")
    
    if df_std is not None and not df_std.empty:
        # 从数据中提取所有公司及对应城市
        company_city = df_std[['公司', '城市']].drop_duplicates()
        company_city['label'] = company_city['公司'] + ' (' + company_city['城市'] + ')'
        company_options = company_city['label'].tolist()
        company_map = {row['label']: (row['公司'], row['城市']) for _, row in company_city.iterrows()}
        
        selected_labels = st.multiselect("选择公司（可多选）", company_options)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            report_type = st.selectbox("报表类型", ["月度申报", "年度汇算"])
        with col2:
            year = st.selectbox("年份", [2025,2024,2023], index=1)
        with col3:
            if report_type == "月度申报":
                month = st.selectbox("月份", list(range(1,13)), index=11)
            else:
                month = None
        
        export_format = st.radio("导出格式", ["Excel (.xlsx)", "CSV (.csv)", "PDF (.pdf)"], horizontal=True)
        template_options = ["使用默认模板"] + [t['name'] for t in custom_templates] if custom_templates else ["使用默认模板"]
        selected_template_name = st.selectbox("选择模板", template_options)
        
        if st.button("🚀 生成报表", type="primary"):
            if not selected_labels:
                st.error("请至少选择一个公司")
            else:
                custom_template = None
                if selected_template_name != "使用默认模板":
                    for t in custom_templates:
                        if t['name'] == selected_template_name:
                            custom_template = t
                            break
                
                rule_df = pd.DataFrame(rules_data)
                generated_files = []
                summary_list = []
                errors = []
                
                for label in selected_labels:
                    company_name, city = company_map[label]
                    # 筛选该公司的数据
                    df_wage = df_std[df_std['公司'] == company_name].copy()
                    if df_wage.empty:
                        errors.append(f"{company_name}: 未找到员工数据")
                        continue
                    
                    # 匹配规则
                    matched = rule_df[(rule_df['city'] == city) & (rule_df['report_type'] == report_type)]
                    if matched.empty:
                        errors.append(f"{company_name}: 未找到 {city} 的规则，请先在'管理数据'中导入规则")
                        continue
                    rule = matched.iloc[0]
                    
                    unit_social = rule.get('unit_social', 0.16)
                    personal_social = rule.get('personal_social', 0.08)
                    unit_fund = rule.get('unit_fund', 0.07)
                    personal_fund = rule.get('personal_fund', 0.07)
                    social_min = rule.get('social_min', 0)
                    social_max = rule.get('social_max', float('inf'))
                    fund_min = rule.get('fund_min', 0)
                    fund_max = rule.get('fund_max', float('inf'))
                    
                    if '社保基数' not in df_wage.columns:
                        df_wage['社保基数'] = df_wage['工资基数']
                    if '公积金基数' not in df_wage.columns:
                        df_wage['公积金基数'] = df_wage['工资基数']
                    
                    df_wage['社保基数调整'] = df_wage['社保基数'].apply(lambda x: max(social_min, min(x, social_max)))
                    df_wage['公积金基数调整'] = df_wage['公积金基数'].apply(lambda x: max(fund_min, min(x, fund_max)))
                    df_wage['单位社保'] = df_wage['社保基数调整'] * unit_social
                    df_wage['个人社保'] = df_wage['社保基数调整'] * personal_social
                    df_wage['单位公积金'] = df_wage['公积金基数调整'] * unit_fund
                    df_wage['个人公积金'] = df_wage['公积金基数调整'] * personal_fund
                    df_wage['单位合计'] = df_wage['单位社保'] + df_wage['单位公积金']
                    df_wage['个人合计'] = df_wage['个人社保'] + df_wage['个人公积金']
                    df_wage['总成本'] = df_wage['单位合计'] + df_wage['个人合计']
                    
                    export_data = df_wage[['公司','城市','姓名','工资基数','社保基数调整','公积金基数调整',
                                          '单位社保','个人社保','单位公积金','个人公积金','单位合计','个人合计','总成本']].copy()
                    export_data.columns = ['公司','城市','姓名','工资基数','社保基数','公积金基数',
                                          '单位社保','个人社保','单位公积金','个人公积金','单位合计','个人合计','总成本']
                    export_data = export_data.round(2)
                    
                    if export_format == "CSV (.csv)":
                        output = BytesIO()
                        export_data.to_csv(output, index=False, encoding='utf-8-sig')
                        output.seek(0)
                        fname = f"{company_name}_{year}{month or ''}.csv"
                        generated_files.append((fname, output.getvalue(), 'text/csv'))
                    elif export_format == "PDF (.pdf)":
                        try:
                            from fpdf import FPDF
                            pdf = FPDF()
                            pdf.add_page()
                            pdf.set_font("Arial", size=12)
                            pdf.cell(200, 10, txt=f"{company_name} - 社保公积金报表", ln=True, align='C')
                            pdf.ln(10)
                            pdf.set_font("Arial", size=10)
                            pdf.cell(100, 8, f"总人数: {len(export_data)}", ln=False)
                            pdf.cell(100, 8, f"总成本: {export_data['总成本'].sum():,.2f}", ln=True)
                            pdf.ln(5)
                            headers = export_data.columns.tolist()
                            for i, h in enumerate(headers[:8]):
                                pdf.cell(20, 8, h, border=1)
                            pdf.ln()
                            for _, row in export_data.head(20).iterrows():
                                for i, val in enumerate(row[:8]):
                                    pdf.cell(20, 8, str(val), border=1)
                                pdf.ln()
                            pdf_output = pdf.output(dest='S').encode('latin1')
                            output = BytesIO(pdf_output)
                            fname = f"{company_name}_{year}{month or ''}.pdf"
                            generated_files.append((fname, output.getvalue(), 'application/pdf'))
                        except ImportError:
                            st.error("PDF导出需要安装 fpdf 库，请运行: pip install fpdf")
                    else:
                        if custom_template:
                            wb = load_workbook(BytesIO(custom_template['file_data']))
                            ws = wb.active
                            row_num = 2
                            for _, row in export_data.iterrows():
                                for col_idx, val in enumerate(row, 1):
                                    ws.cell(row=row_num, column=col_idx, value=val)
                                row_num += 1
                        else:
                            wb = Workbook()
                            ws = wb.active
                            ws.title = "员工明细"
                            ws.append(export_data.columns.tolist())
                            for _, row in export_data.iterrows():
                                ws.append(row.tolist())
                            ws_sum = wb.create_sheet("汇总")
                            ws_sum.append(["指标", "金额"])
                            ws_sum.append(["总人数", len(export_data)])
                            ws_sum.append(["总成本", round(export_data['总成本'].sum(),2)])
                            audit = wb.create_sheet("AuditTrail")
                            audit.append(["时间", "操作", "详情"])
                            audit.append([datetime.now().isoformat(), "GENERATED", f"公司:{company_name}, 规则:{rule.get('source_quote','')}"])
                            ws.insert_rows(1)
                            ws['A1'] = '⚠️ 待复核版'
                            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(export_data.columns))
                        output = BytesIO()
                        wb.save(output)
                        output.seek(0)
                        fname = f"{company_name}_{year}{month or ''}.xlsx"
                        generated_files.append((fname, output.getvalue(), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
                    
                    export_id = str(uuid.uuid4())[:8]
                    record = {
                        "id": export_id,
                        "company": company_name,
                        "city": city,
                        "report_type": report_type,
                        "year": year,
                        "month": month,
                        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "status": "待复核",
                        "source_quote": rule.get('source_quote',''),
                        "total_people": len(export_data),
                        "total_cost": round(export_data['总成本'].sum(),2),
                        "reviewer": "",
                        "reviewed_at": "",
                        "review_comment": ""
                    }
                    insert_or_update_export(record)
                    insert_audit_log({
                        "id": str(uuid.uuid4())[:8],
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "action": "生成报表",
                        "detail": f"公司:{company_name}, 人数:{len(export_data)}, 格式:{export_format}",
                        "report_id": export_id
                    })
                    summary_list.append({"公司": company_name, "人数": len(export_data), "总成本": round(export_data['总成本'].sum(),2)})
                
                if errors:
                    for err in errors:
                        st.warning(err)
                if generated_files:
                    st.success(f"✅ 成功生成 {len(generated_files)} 份报表")
                    st.dataframe(pd.DataFrame(summary_list))
                    if len(generated_files) > 1:
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w') as zf:
                            for fname, data, mime in generated_files:
                                zf.writestr(fname, data)
                        zip_buffer.seek(0)
                        st.download_button("📦 下载全部（ZIP）", data=zip_buffer, file_name=f"报表_{year}{month or ''}.zip", mime="application/zip")
                    else:
                        fname, data, mime = generated_files[0]
                        st.download_button(f"📥 下载 {fname}", data=BytesIO(data), file_name=fname, mime=mime)
    else:
        st.info("请先上传数据并完成列映射，系统将自动识别公司列表。")

# ==================== 报表历史与复核 ====================
with tab2:
    st.subheader("📋 历史记录")
    history = get_export_history()
    if not history:
        st.info("暂无历史记录")
    else:
        df_hist = pd.DataFrame(history)
        st.dataframe(df_hist[['company','city','report_type','year','month','total_people','total_cost','generated_at','status','reviewer']], use_container_width=True)
        
        st.subheader("✅ 复核报表")
        pending = [h for h in history if h.get('status') == '待复核']
        if pending:
            options = [f"{h['company']} - {h['city']} - {h['year']}{h['month'] or ''}" for h in pending]
            sel_idx = st.selectbox("选择待复核报表", range(len(options)), format_func=lambda x: options[x])
            selected = pending[sel_idx]
            st.write(f"**公司**：{selected['company']}　**城市**：{selected['city']}　**总成本**：{selected['total_cost']}")
            st.write(f"**规则来源**：{selected.get('source_quote','')}")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ 通过复核"):
                    selected['status'] = '已通过'
                    selected['reviewer'] = '复核员'
                    selected['reviewed_at'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    selected['review_comment'] = '数据无误'
                    insert_or_update_export(selected)
                    insert_audit_log({
                        "id": str(uuid.uuid4())[:8],
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "action": "复核通过",
                        "detail": f"公司:{selected['company']}",
                        "report_id": selected['id']
                    })
                    st.success("已通过复核")
                    st.rerun()
            with col2:
                if st.button("❌ 驳回"):
                    reason = st.text_input("驳回原因")
                    if st.button("确认驳回"):
                        selected['status'] = '已驳回'
                        selected['reviewer'] = '复核员'
                        selected['reviewed_at'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        selected['review_comment'] = reason or '数据有误'
                        insert_or_update_export(selected)
                        insert_audit_log({
                            "id": str(uuid.uuid4())[:8],
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "action": "复核驳回",
                            "detail": f"公司:{selected['company']}, 原因:{reason}",
                            "report_id": selected['id']
                        })
                        st.warning("已驳回")
                        st.rerun()
        else:
            st.info("所有报表已复核完成")
        
        with st.expander("📜 操作日志"):
            logs = get_audit_logs()
            if logs:
                st.dataframe(pd.DataFrame(logs))
            else:
                st.info("暂无日志")

# ==================== 数据管理 ====================
with tab3:
    st.subheader("📝 管理本地数据")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**公司管理**")
        comps = load_table('companies')
        edited_comps = st.data_editor(comps, num_rows="dynamic", key="comp_edit")
        if st.button("保存公司"):
            save_table('companies', edited_comps)
            st.success("已保存")
    with col2:
        st.markdown("**规则管理**")
        rules_data = load_table('rules')
        edited_rules = st.data_editor(rules_data, num_rows="dynamic", key="rule_edit")
        if st.button("保存规则"):
            save_table('rules', edited_rules)
            st.success("已保存")
    
    with st.expander("📤 批量导入 Excel（自动识别公司+规则）"):
        st.caption("上传Excel后，系统自动识别「公司配置表」和「基础配置表」，无需手动指定Sheet名称")
        uploaded_file = st.file_uploader("选择 Excel", type=["xlsx"], key="import_batch")
        if uploaded_file:
            try:
                xls = pd.ExcelFile(uploaded_file)
                success_count = 0
                # ===== 自动识别并导入公司 =====
                # 查找包含公司信息的Sheet
                company_sheet = None
                for sheet in xls.sheet_names:
                    if '公司' in sheet or '企业' in sheet or '配置' in sheet:
                        company_sheet = sheet
                        break
                
                if company_sheet:
                    df_comp = pd.read_excel(uploaded_file, sheet_name=company_sheet)
                    # 尝试找公司列和城市列
                    comp_col = None
                    city_col = None
                    for col in df_comp.columns:
                        if '公司' in col or '企业' in col:
                            comp_col = col
                        if '城市' in col or '市' in col:
                            city_col = col
                    if comp_col and city_col:
                        df_comp_clean = df_comp[[comp_col, city_col]].drop_duplicates()
                        df_comp_clean = df_comp_clean.rename(columns={comp_col: 'company_name', city_col: 'city'})
                        df_comp_clean['province'] = df_comp_clean['city']
                        df_comp_clean['district'] = '市区'
                        df_comp_clean['id'] = [str(uuid.uuid4())[:8] for _ in range(len(df_comp_clean))]
                        save_table('companies', df_comp_clean.to_dict('records'))
                        st.success(f"✅ 从「{company_sheet}」导入 {len(df_comp_clean)} 家公司")
                        success_count += 1
                
                # ===== 自动识别并导入规则 =====
                rule_count = import_rules_from_excel(xls)
                if rule_count > 0:
                    st.success(f"✅ 从「基础配置表」导入 {rule_count} 个城市的规则")
                    success_count += 1
                
                if success_count == 0:
                    st.warning("未识别到有效数据，请确保Excel包含「公司配置表」和「基础配置表」")
                else:
                    st.rerun()
            except Exception as e:
                st.error(f"解析失败：{e}")

# ==================== 自定义模板 ====================
with tab4:
    st.subheader("📄 自定义报表模板")
    custom_templates = load_table('custom_templates')
    col1, col2 = st.columns(2)
    with col1:
        uploaded_template = st.file_uploader("上传模板 (.xlsx)", type=["xlsx"], key="template_upload")
        if uploaded_template:
            st.success(f"已上传：{uploaded_template.name}")
            if st.button("保存模板"):
                template_data = {
                    "id": str(uuid.uuid4())[:8],
                    "name": uploaded_template.name,
                    "file_data": uploaded_template.getvalue(),
                    "field_mapping": ""
                }
                existing = [t for t in custom_templates if t['name'] == uploaded_template.name]
                if existing:
                    for t in custom_templates:
                        if t['name'] == uploaded_template.name:
                            t['file_data'] = uploaded_template.getvalue()
                            break
                    save_table('custom_templates', custom_templates)
                    st.success("模板已更新")
                else:
                    custom_templates.append(template_data)
                    save_table('custom_templates', custom_templates)
                    st.success("模板已保存")
                st.rerun()
    with col2:
        if custom_templates:
            st.write("**已保存的模板**")
            for t in custom_templates:
                st.write(f"- {t['name']}")
            if st.button("删除所有模板"):
                save_table('custom_templates', [])
                st.success("已删除")
                st.rerun()
        else:
            st.info("暂无自定义模板")

st.caption(f"📁 数据库位置：`{DB_PATH}`（请妥善备份）")
