import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
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

# ---------- 智能列名匹配 ----------
def smart_column_mapping(df, required_cols):
    """
    智能匹配列名，返回映射字典 {标准列名: 实际列名}
    如果匹配失败，返回 None 并提供候选
    """
    df_cols = list(df.columns)
    mapping = {}
    unmatched = []
    candidates = {}
    
    # 常见同义词映射
    synonyms = {
        '公司': ['公司', '企业', '单位', '公司名称', '企业名称', '单位名称', 'name', 'company'],
        '城市': ['城市', '市', 'city', '地区'],
        '姓名': ['姓名', '名字', '员工', 'name', 'employee', '人员'],
        '工资基数': ['工资基数', '基数', '工资', '月工资', '基本工资', 'base', 'salary'],
        '社保基数': ['社保基数', '养老基数', '社保', 'social_security'],
        '公积金基数': ['公积金基数', '公积金', 'fund']
    }
    
    for std_name in required_cols:
        matched = False
        for df_col in df_cols:
            # 精确匹配
            if df_col == std_name:
                mapping[std_name] = df_col
                matched = True
                break
            # 模糊匹配（包含关系）
            for syn in synonyms.get(std_name, []):
                if syn.lower() in df_col.lower() or df_col.lower() in syn.lower():
                    mapping[std_name] = df_col
                    matched = True
                    break
            if matched:
                break
        if not matched:
            unmatched.append(std_name)
            # 找最相似的候选
            for df_col in df_cols:
                if df_col not in mapping.values():
                    if std_name in candidates:
                        candidates[std_name].append(df_col)
                    else:
                        candidates[std_name] = [df_col]
    
    return mapping, unmatched, candidates

def validate_and_clean_data(df, mapping, required_cols):
    """
    验证数据并生成详细错误信息
    """
    errors = []
    warnings = []
    
    # 检查必要列是否存在
    missing_cols = [col for col in required_cols if col not in mapping]
    if missing_cols:
        return None, f"缺少必要列：{', '.join(missing_cols)}"
    
    # 构建标准化的DataFrame
    std_data = {}
    for std_name, actual_col in mapping.items():
        if actual_col in df.columns:
            std_data[std_name] = df[actual_col]
        else:
            return None, f"映射列 '{actual_col}' 在数据中不存在"
    
    try:
        std_df = pd.DataFrame(std_data)
    except Exception as e:
        return None, f"数据转换失败：{str(e)}"
    
    # 逐行验证
    row_errors = []
    for idx, row in std_df.iterrows():
        row_num = idx + 2  # Excel行号（从2开始）
        # 检查必填字段是否为空
        if pd.isna(row.get('姓名', '')) or str(row.get('姓名', '')).strip() == '':
            row_errors.append(f"第{row_num}行：姓名为空")
        if pd.isna(row.get('工资基数', '')):
            row_errors.append(f"第{row_num}行：工资基数为空")
        if pd.isna(row.get('公司', '')) or str(row.get('公司', '')).strip() == '':
            row_errors.append(f"第{row_num}行：公司为空")
        # 检查工资基数是否为数字
        if pd.notna(row.get('工资基数', '')):
            try:
                float(row.get('工资基数', 0))
            except (ValueError, TypeError):
                row_errors.append(f"第{row_num}行：工资基数不是有效数字（'{row.get('工资基数', '')}')")
    
    if row_errors:
        return std_df, "\n".join(row_errors[:5]) + (f"\n... 共{len(row_errors)}个错误" if len(row_errors) > 5 else "")
    
    return std_df, None

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
st.title("📋 本地社保公积金报表系统")
st.markdown("🔒 **所有数据仅保存在本机 `data.db` 中，绝不联网上传**")

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
        
        # ----- 自定义模板选择 -----
        st.subheader("📄 报表模板")
        template_options = ["使用默认模板"] + [t['name'] for t in custom_templates] if custom_templates else ["使用默认模板"]
        selected_template_name = st.selectbox("选择模板（或上传自定义模板）", template_options)
        
        # ----- 上传工资表（带智能列名映射） -----
        st.subheader("📤 上传员工工资表")
        st.caption("支持列名：公司、城市、姓名、工资基数、社保基数（可选）、公积金基数（可选）")
        
        uploaded = st.file_uploader("选择工资表 .xlsx", type=["xlsx"])
        df_raw = None
        std_df = None
        mapping_error = None
        
        if uploaded:
            try:
                df_raw = pd.read_excel(uploaded)
                st.info(f"📊 读取到 {len(df_raw)} 行，{len(df_raw.columns)} 列")
                
                # 智能列名匹配
                required_cols = ['公司', '城市', '姓名', '工资基数']
                mapping, unmatched, candidates = smart_column_mapping(df_raw, required_cols)
                
                if unmatched:
                    st.warning(f"⚠️ 未能自动匹配以下列：{', '.join(unmatched)}")
                    # 手动匹配
                    st.markdown("**请手动选择对应的列名：**")
                    manual_mapping = {}
                    for col in unmatched:
                        if col in candidates:
                            manual_mapping[col] = st.selectbox(
                                f"选择 '{col}' 对应的列",
                                [''] + list(df_raw.columns),
                                key=f"map_{col}_{uploaded.name}"
                            )
                        else:
                            manual_mapping[col] = st.selectbox(
                                f"选择 '{col}' 对应的列",
                                [''] + list(df_raw.columns),
                                key=f"map_{col}_{uploaded.name}"
                            )
                    
                    # 合并映射
                    for k, v in manual_mapping.items():
                        if v:
                            mapping[k] = v
                
                # 验证数据
                if mapping:
                    std_df, error = validate_and_clean_data(df_raw, mapping, required_cols)
                    if error:
                        mapping_error = error
                        st.error(f"❌ 数据验证失败：\n{error}")
                    else:
                        st.success("✅ 数据验证通过！")
                        st.dataframe(std_df.head(10))
                        
                        # ---- 报表预览 ----
                        st.subheader("📊 报表预览（生成前确认）")
                        preview_company = st.selectbox(
                            "预览公司",
                            [c['company_name'] for c in companies_data],
                            key="preview_company"
                        )
                        if st.button("🔍 预览数据", key="preview_btn"):
                            preview_df = std_df[std_df['公司'] == preview_company]
                            if preview_df.empty:
                                st.warning(f"未找到 {preview_company} 的员工数据")
                            else:
                                # 匹配规则
                                city = next((c['city'] for c in companies_data if c['company_name'] == preview_company), '')
                                rule_df = pd.DataFrame(rules_data)
                                matched = rule_df[(rule_df['city'] == city) & (rule_df['report_type'] == report_type)]
                                if matched.empty:
                                    st.warning(f"未找到 {city} 的规则")
                                else:
                                    rule = matched.iloc[0]
                                    # 计算预览数据
                                    preview_calc = preview_df.copy()
                                    unit_social = rule.get('unit_social', 0.16)
                                    personal_social = rule.get('personal_social', 0.08)
                                    unit_fund = rule.get('unit_fund', 0.07)
                                    personal_fund = rule.get('personal_fund', 0.07)
                                    if '社保基数' not in preview_calc.columns:
                                        preview_calc['社保基数'] = preview_calc['工资基数']
                                    if '公积金基数' not in preview_calc.columns:
                                        preview_calc['公积金基数'] = preview_calc['工资基数']
                                    preview_calc['单位社保'] = preview_calc['社保基数'] * unit_social
                                    preview_calc['个人社保'] = preview_calc['社保基数'] * personal_social
                                    preview_calc['单位公积金'] = preview_calc['公积金基数'] * unit_fund
                                    preview_calc['个人公积金'] = preview_calc['公积金基数'] * personal_fund
                                    preview_calc['单位合计'] = preview_calc['单位社保'] + preview_calc['单位公积金']
                                    preview_calc['个人合计'] = preview_calc['个人社保'] + preview_calc['个人公积金']
                                    preview_calc['总成本'] = preview_calc['单位合计'] + preview_calc['个人合计']
                                    
                                    st.write(f"**{preview_company} - 预览数据**")
                                    st.dataframe(preview_calc[['姓名','工资基数','单位社保','个人社保','单位公积金','个人公积金','总成本']].head(20))
                                    st.metric("总人数", len(preview_calc))
                                    st.metric("总成本", f"¥{preview_calc['总成本'].sum():,.2f}")
                                    
            except Exception as e:
                st.error(f"❌ 读取文件失败：{str(e)}")
        
        # 生成按钮
        if st.button("🚀 生成报表", type="primary"):
            if std_df is None:
                st.error("请先上传并验证工资表数据")
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
                    df_wage = std_df[std_df['公司'] == company_name].copy()
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
                    
                    # --- 生成 Excel（使用自定义模板或默认） ---
                    if custom_template:
                        # 使用自定义模板
                        wb = load_workbook(BytesIO(custom_template['file_data']))
                        ws = wb.active
                        # 从第2行开始填充数据（假设第1行是表头）
                        row_num = 2
                        for _, row in df_wage.iterrows():
                            ws.cell(row=row_num, column=1, value=row['公司'])
                            ws.cell(row=row_num, column=2, value=row['城市'])
                            ws.cell(row=row_num, column=3, value=row['姓名'])
                            ws.cell(row=row_num, column=4, value=row['工资基数'])
                            ws.cell(row=row_num, column=5, value=round(row['单位社保'],2))
                            ws.cell(row=row_num, column=6, value=round(row['个人社保'],2))
                            ws.cell(row=row_num, column=7, value=round(row['单位公积金'],2))
                            ws.cell(row=row_num, column=8, value=round(row['个人公积金'],2))
                            row_num += 1
                    else:
                        # 使用默认模板
                        wb = Workbook()
                        ws = wb.active
                        ws.title = "员工明细"
                        headers = ['公司','城市','姓名','工资基数','社保基数','公积金基数','单位社保','个人社保','单位公积金','个人公积金','单位合计','个人合计','总成本']
                        ws.append(headers)
                        for _, row in df_wage.iterrows():
                            ws.append([
                                row['公司'], row['城市'], row['姓名'], row['工资基数'],
                                row['社保基数调整'], row['公积金基数调整'],
                                round(row['单位社保'],2), round(row['个人社保'],2),
                                round(row['单位公积金'],2), round(row['个人公积金'],2),
                                round(row['单位合计'],2), round(row['个人合计'],2),
                                round(row['总成本'],2)
                            ])
                        # 汇总
                        ws_sum = wb.create_sheet("汇总")
                        ws_sum.append(["指标", "金额"])
                        ws_sum.append(["总人数", len(df_wage)])
                        ws_sum.append(["总成本", round(df_wage['总成本'].sum(),2)])
                        # 审计日志
                        audit = wb.create_sheet("AuditTrail")
                        audit.append(["时间", "操作", "详情"])
                        audit.append([datetime.now().isoformat(), "GENERATED", f"公司:{company_name}, 规则:{rule.get('source_quote','')}"])
                        # 水印
                        ws.insert_rows(1)
                        ws['A1'] = '⚠️ 待复核版'
                        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
                    
                    output = BytesIO()
                    wb.save(output)
                    output.seek(0)
                    fname = f"{company_name}_{year}{month or ''}.xlsx"
                    generated_files.append((fname, output.getvalue()))
                    
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
                        "total_people": len(df_wage),
                        "total_cost": round(df_wage['总成本'].sum(),2),
                        "reviewer": "",
                        "reviewed_at": "",
                        "review_comment": ""
                    }
                    insert_or_update_export(record)
                    insert_audit_log({
                        "id": str(uuid.uuid4())[:8],
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "action": "生成报表",
                        "detail": f"公司:{company_name}, 人数:{len(df_wage)}",
                        "report_id": export_id
                    })
                    summary_list.append({"公司": company_name, "人数": len(df_wage), "总成本": round(df_wage['总成本'].sum(),2)})
                
                if errors:
                    for err in errors:
                        st.warning(err)
                
                if generated_files:
                    st.success(f"✅ 成功生成 {len(generated_files)} 份报表")
                    st.dataframe(pd.DataFrame(summary_list))
                    if len(generated_files) > 1:
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w') as zf:
                            for fname, data in generated_files:
                                zf.writestr(fname, data)
                        zip_buffer.seek(0)
                        st.download_button("📦 下载全部（ZIP）", data=zip_buffer, file_name=f"报表_{year}{month or ''}.zip", mime="application/zip")
                    else:
                        fname, data = generated_files[0]
                        st.download_button("📥 下载报表", data=BytesIO(data), file_name=fname)

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
    st.subheader("📝 管理本地数据（永久保存）")
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
        st.caption("Excel 需包含「公司」和「规则」两个 Sheet")
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
    st.markdown("上传您自己的 Excel 模板，生成报表时将按照您的模板格式填充数据。")
    st.caption("模板第一行应为表头，数据从第二行开始填充（公司、城市、姓名、工资基数等）")
    
    custom_templates = load_table('custom_templates')
    
    col1, col2 = st.columns(2)
    with col1:
        uploaded_template = st.file_uploader("上传模板文件 (.xlsx)", type=["xlsx"], key="template_upload")
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
                    # 更新
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
                st.success("已删除所有模板")
                st.rerun()
        else:
            st.info("暂无自定义模板")

st.caption(f"📁 数据库位置：`{DB_PATH}`（请妥善备份）")
