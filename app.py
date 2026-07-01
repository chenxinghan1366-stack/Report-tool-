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

# ---------- 智能列名映射（增强版） ----------
def smart_column_mapping(df, required_cols, user_mapping=None):
    """
    智能匹配列名，返回映射字典 {标准列名: 实际列名}
    如果 user_mapping 提供，则优先使用
    """
    df_cols = list(df.columns)
    synonyms = {
        '公司': ['公司', '企业', '单位', '公司名称', '企业名称', '单位名称', 'name', 'company', '公司名'],
        '城市': ['城市', '市', 'city', '地区', '所属城市', '所在地'],
        '姓名': ['姓名', '名字', '员工', 'name', 'employee', '人员', '员工编号'],
        '工资基数': ['工资基数', '基数', '工资', '月工资', '基本工资', 'base', 'salary', '月应发工资'],
        '社保基数': ['社保基数', '养老基数', '社保', 'social_security'],
        '公积金基数': ['公积金基数', '公积金', 'fund']
    }
    
    if user_mapping:
        # 用户手动映射部分优先
        mapping = {k: v for k, v in user_mapping.items() if v}
        # 未映射的列自动匹配
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
            for df_col in df_cols:
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

# ---------- 初始化 ----------
init_db()
if not load_table('companies'):
    save_table('companies', [
        {"id": "c001", "company_name": "上海科技公司", "province": "上海", "city": "上海市", "district": "浦东新区"},
        {"id": "c002", "company_name": "深圳科技公司", "province": "广东", "city": "深圳市", "district": "南山区"}
    ])
if not load_table('rules'):
    save_table('rules', [
        {"id": "r001", "province": "上海", "city": "上海市", "report_type": "月度申报",
         "unit_social": 0.16, "personal_social": 0.08, "unit_fund": 0.07, "personal_fund": 0.07,
         "social_min": 7310, "social_max": 36549, "fund_min": 2590, "fund_max": 34188,
         "source_quote": "沪人社规〔2024〕22号"},
        {"id": "r002", "province": "广东", "city": "深圳市", "report_type": "月度申报",
         "unit_social": 0.14, "personal_social": 0.08, "unit_fund": 0.05, "personal_fund": 0.05,
         "social_min": 2360, "social_max": 29727, "fund_min": 2360, "fund_max": 29727,
         "source_quote": "深人社规〔2024〕3号"}
    ])

# ---------- Streamlit 页面 ----------
st.set_page_config(page_title="本地社保报表系统", layout="wide")
st.title("📋 本地社保公积金报表系统（智能版）")
st.markdown("🔒 **支持直接读取汇总报表，自动匹配列名**")

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

# ==================== 生成报表（升级版） ====================
with tab1:
    companies_data = load_table('companies')
    rules_data = load_table('rules')
    custom_templates = load_table('custom_templates')
    
    if not companies_data:
        st.warning("请先在「数据管理」中添加公司信息")
    else:
        st.subheader("📤 生成报表")
        # 选择公司
        company_options = [f"{c['company_name']} ({c['city']})" for c in companies_data]
        selected_indices = st.multiselect("选择公司（可多选）", range(len(company_options)), format_func=lambda x: company_options[x])
        
        # 报表参数
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
                st.info("年度汇算无月份")
        
        # 导出格式
        export_format = st.radio("导出格式", ["Excel (.xlsx)", "CSV (.csv)", "PDF (.pdf)"], horizontal=True)
        
        # 模板选择
        template_options = ["使用默认模板"] + [t['name'] for t in custom_templates] if custom_templates else ["使用默认模板"]
        selected_template_name = st.selectbox("选择模板", template_options)
        
        # ===== 上传数据（升级：支持选择Sheet和列映射） =====
        st.subheader("📤 上传数据（支持明细表或汇总表）")
        st.caption("支持 .xlsx, .xls, .csv, .txt 格式，可自动识别列名")
        uploaded = st.file_uploader("选择文件", type=["xlsx", "xls", "csv", "txt"])
        
        df_raw = None
        df_std = None
        mapping_error = None
        sheets = []
        selected_sheet = None
        
        if uploaded:
            try:
                file_ext = uploaded.name.split('.')[-1].lower()
                if file_ext in ['xlsx', 'xls']:
                    xls = pd.ExcelFile(uploaded)
                    sheets = xls.sheet_names
                    selected_sheet = st.selectbox("选择Sheet", sheets, index=0)
                    if selected_sheet:
                        df_raw = pd.read_excel(uploaded, sheet_name=selected_sheet)
                elif file_ext == 'csv':
                    df_raw = pd.read_csv(uploaded, encoding='utf-8-sig')
                elif file_ext == 'txt':
                    df_raw = pd.read_csv(uploaded, delimiter='\t', encoding='utf-8-sig')
                else:
                    st.error("不支持的文件格式")
                    df_raw = None
                
                if df_raw is not None:
                    st.info(f"📊 读取到 {len(df_raw)} 行，{len(df_raw.columns)} 列")
                    # 显示原始数据预览
                    st.dataframe(df_raw.head(5))
                    
                    # 自动检测表头行（跳过空白行和标题行）
                    # 寻找第一行包含“公司”或“城市”或“姓名”等关键字的行作为表头
                    header_row = 0
                    for i, row in df_raw.iterrows():
                        row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                        if any(key in row_text for key in ['公司', '城市', '姓名', '工资', '基数', '员工']):
                            header_row = i
                            break
                    
                    # 如果自动检测失败，让用户手动选择
                    if header_row == 0:
                        header_row = st.number_input("表头行号（从0开始）", min_value=0, max_value=len(df_raw)-1, value=0, step=1)
                    else:
                        st.info(f"自动检测到表头行：第 {header_row} 行")
                        # 让用户确认或调整
                        override = st.checkbox("手动调整表头行号")
                        if override:
                            header_row = st.number_input("表头行号（从0开始）", min_value=0, max_value=len(df_raw)-1, value=header_row, step=1)
                    
                    # 提取数据
                    if header_row < len(df_raw):
                        # 设置表头
                        new_header = df_raw.iloc[header_row]
                        df_raw = df_raw[header_row+1:]
                        df_raw.columns = new_header
                        # 重置索引
                        df_raw = df_raw.reset_index(drop=True)
                        # 删除全为空的行
                        df_raw = df_raw.dropna(how='all')
                        st.success(f"成功提取数据，共 {len(df_raw)} 行")
                        st.dataframe(df_raw.head(5))
                        
                        # 智能列名映射
                        required_cols = ['公司', '城市', '姓名', '工资基数']
                        # 检测数据中是否包含这些列（或类似列）
                        col_mapping = smart_column_mapping(df_raw, required_cols)
                        
                        # 显示映射结果，让用户确认
                        st.subheader("🔍 列名映射确认")
                        mapping_ok = True
                        for std_name, actual_col in col_mapping.items():
                            if actual_col is None or actual_col == '':
                                mapping_ok = False
                                # 让用户手动选择
                                options = [''] + list(df_raw.columns)
                                selected = st.selectbox(f"请选择 '{std_name}' 对应的列", options, key=f"map_{std_name}_{uploaded.name}")
                                if selected:
                                    col_mapping[std_name] = selected
                                else:
                                    st.warning(f"'{std_name}' 未映射，可能影响报表生成")
                        # 如果仍有未映射的，给出警告
                        missing = [k for k, v in col_mapping.items() if not v]
                        if missing:
                            st.warning(f"以下列未映射：{', '.join(missing)}，将尝试用默认值填充")
                        else:
                            st.success("列映射完成！")
                        
                        # 构建标准数据框
                        std_data = {}
                        for std_name, actual_col in col_mapping.items():
                            if actual_col and actual_col in df_raw.columns:
                                std_data[std_name] = df_raw[actual_col]
                            else:
                                std_data[std_name] = None
                        df_std = pd.DataFrame(std_data)
                        
                        # 如果缺少“姓名”列，尝试用“员工编号”或“员工”列代替
                        if '姓名' not in df_std.columns or df_std['姓名'].isnull().all():
                            for col in df_raw.columns:
                                if '编号' in col or 'ID' in col or '员工' in col:
                                    df_std['姓名'] = df_raw[col]
                                    st.info(f"使用 '{col}' 作为姓名标识")
                                    break
                        
                        # 如果仍然没有姓名，生成默认编号
                        if '姓名' not in df_std.columns or df_std['姓名'].isnull().all():
                            df_std['姓名'] = [f"员工{i+1}" for i in range(len(df_std))]
                            st.info("未找到姓名列，自动生成编号")
                        
                        # 处理工资基数：尝试从“工资基数”列读取，若无则尝试“月应发工资”
                        if '工资基数' not in df_std.columns or df_std['工资基数'].isnull().all():
                            for col in df_raw.columns:
                                if '工资' in col or '基数' in col or '月应发' in col:
                                    df_std['工资基数'] = df_raw[col]
                                    st.info(f"使用 '{col}' 作为工资基数")
                                    break
                        
                        # 如果仍然没有工资基数，报错
                        if '工资基数' not in df_std.columns or df_std['工资基数'].isnull().all():
                            st.error("未能找到工资基数列，请确保数据中包含工资基数信息")
                            df_std = None
                        else:
                            # 转换为数字
                            df_std['工资基数'] = pd.to_numeric(df_std['工资基数'], errors='coerce')
                            # 删除工资基数为空的行
                            df_std = df_std.dropna(subset=['工资基数'])
                            st.success(f"数据准备完成，共 {len(df_std)} 条有效记录")
                            st.dataframe(df_std.head(5))
            except Exception as e:
                st.error(f"❌ 读取文件失败：{str(e)}")
        
        # 生成按钮
        if st.button("🚀 生成报表", type="primary"):
            if df_std is None or df_std.empty:
                st.error("请先上传并准备数据")
            elif not selected_indices:
                st.error("请至少选择一个公司")
            else:
                # 获取自定义模板
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
                
                for idx in selected_indices:
                    company_info = companies_data[idx]
                    company_name = company_info['company_name']
                    city = company_info['city']
                    df_wage = df_std[df_std['公司'] == company_name].copy()
                    if df_wage.empty:
                        errors.append(f"{company_name}: 未找到员工数据")
                        continue
                    
                    # 匹配规则
                    matched = rule_df[(rule_df['city'] == city) & (rule_df['report_type'] == report_type)]
                    if matched.empty:
                        errors.append(f"{company_name}: 未找到 {city} 的规则")
                        continue
                    rule = matched.iloc[0]
                    
                    # 计算
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
                    
                    # 校验基数
                    df_wage['社保基数调整'] = df_wage['社保基数'].apply(lambda x: max(social_min, min(x, social_max)))
                    df_wage['公积金基数调整'] = df_wage['公积金基数'].apply(lambda x: max(fund_min, min(x, fund_max)))
                    df_wage['单位社保'] = df_wage['社保基数调整'] * unit_social
                    df_wage['个人社保'] = df_wage['社保基数调整'] * personal_social
                    df_wage['单位公积金'] = df_wage['公积金基数调整'] * unit_fund
                    df_wage['个人公积金'] = df_wage['公积金基数调整'] * personal_fund
                    df_wage['单位合计'] = df_wage['单位社保'] + df_wage['单位公积金']
                    df_wage['个人合计'] = df_wage['个人社保'] + df_wage['个人公积金']
                    df_wage['总成本'] = df_wage['单位合计'] + df_wage['个人合计']
                    
                    # 生成不同格式
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
                        # 简单PDF生成（需安装reportlab或fpdf，这里用fpdf）
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
                        # 表头（仅前8列）
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
                    else:  # Excel
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
                    
                    # 保存历史
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

# ==================== 报表历史与复核（不变） ====================
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
    
    with st.expander("📤 批量导入 Excel（公司+规则）"):
        uploaded_file = st.file_uploader("选择 Excel", type=["xlsx"], key="import_batch")
        if uploaded_file:
            try:
                xls = pd.ExcelFile(uploaded_file)
                if "公司" in xls.sheet_names and "规则" in xls.sheet_names:
                    df_comp = pd.read_excel(uploaded_file, sheet_name="公司")
                    df_rule = pd.read_excel(uploaded_file, sheet_name="规则")
                    df_comp['id'] = [str(uuid.uuid4())[:8] for _ in range(len(df_comp))]
                    df_rule['id'] = [str(uuid.uuid4())[:8] for _ in range(len(df_rule))]
                    st.write("预览公司", df_comp.head())
                    st.write("预览规则", df_rule.head())
                    if st.button("确认导入"):
                        save_table('companies', df_comp.to_dict('records'))
                        save_table('rules', df_rule.to_dict('records'))
                        st.success("导入成功！")
                        st.rerun()
                else:
                    st.warning("需要「公司」和「规则」两个 Sheet")
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
