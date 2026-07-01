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

# ===== 从“基础配置表”导入规则（自动） =====
def import_rules_from_excel(xls):
    rule_sheet = None
    for sheet in ["基础配置表", "规则", "城市规则", "配置表"]:
        if sheet in xls.sheet_names:
            rule_sheet = sheet
            break
    if rule_sheet is None:
        return 0, "未找到「基础配置表」"

    try:
        df_full = pd.read_excel(xls, sheet_name=rule_sheet, header=None)
        header_row_idx = None
        for i in range(len(df_full)):
            row_text = ' '.join([str(v) for v in df_full.iloc[i].values if pd.notna(v)])
            if '所属城市' in row_text and ('养老保险' in row_text or '公积金' in row_text):
                header_row_idx = i
                break
        
        if header_row_idx is None:
            return 0, "未找到包含列名的行"
        
        df_rules = pd.read_excel(xls, sheet_name=rule_sheet, skiprows=header_row_idx)
        df_rules.columns = [str(c).strip().replace('\n', '') for c in df_rules.columns]
        
        col_map = {}
        for col in df_rules.columns:
            col_lower = col.lower().strip()
            if '所属城市' in col_lower or '城市' in col_lower:
                col_map['城市'] = col
            elif '养老保险-单位' in col_lower or '养老单位' in col_lower:
                col_map['单位社保比例'] = col
            elif '养老保险-个人' in col_lower or '养老个人' in col_lower:
                col_map['个人社保比例'] = col
            elif '公积金-单位' in col_lower or '单位公积金' in col_lower:
                col_map['单位公积金比例'] = col
            elif '公积金-个人' in col_lower or '个人公积金' in col_lower:
                col_map['个人公积金比例'] = col
            elif '缴费基数下限' in col_lower:
                col_map['社保最低基数'] = col
            elif '缴费基数上限' in col_lower:
                col_map['社保最高基数'] = col
        
        required = ['城市', '单位社保比例', '个人社保比例', '单位公积金比例', '个人公积金比例']
        missing = [r for r in required if r not in col_map]
        if missing:
            return 0, f"缺少必要列：{', '.join(missing)}。检测到的列名：{list(df_rules.columns)}"
        
        rules_list = []
        for idx, row in df_rules.iterrows():
            city = row[col_map['城市']]
            if pd.isna(city):
                continue
            try:
                unit_social = float(row[col_map['单位社保比例']])
                personal_social = float(row[col_map['个人社保比例']])
                unit_fund = float(row[col_map['单位公积金比例']])
                personal_fund = float(row[col_map['个人公积金比例']])
                social_min = float(row[col_map['社保最低基数']]) if col_map.get('社保最低基数') is not None and not pd.isna(row[col_map['社保最低基数']]) else 0
                social_max = float(row[col_map['社保最高基数']]) if col_map.get('社保最高基数') is not None and not pd.isna(row[col_map['社保最高基数']]) else 999999
            except (ValueError, TypeError):
                continue
            
            rules_list.append({
                "id": str(uuid.uuid4())[:8],
                "province": city,
                "city": city,
                "report_type": "月度申报",
                "unit_social": unit_social,
                "personal_social": personal_social,
                "unit_fund": unit_fund,
                "personal_fund": personal_fund,
                "social_min": social_min,
                "social_max": social_max,
                "fund_min": 0,
                "fund_max": 999999,
                "source_quote": f"自动从{rule_sheet}导入"
            })
        
        if rules_list:
            save_table('rules', rules_list)
            return len(rules_list), None
        else:
            return 0, "未解析到有效数据"
    except Exception as e:
        return 0, f"解析规则时出错：{str(e)}"

# ---------- 初始化 ----------
init_db()
if not load_table('companies'):
    save_table('companies', [])
if not load_table('rules'):
    save_table('rules', [])

# ---------- Streamlit 页面 ----------
st.set_page_config(page_title="本地社保报表系统", layout="wide")
st.title("📋 社保公积金报表系统（自适配版）")
st.markdown("🔒 **上传Excel，手动映射列，自动匹配规则，一键生成报表**")

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

tab1, tab2, tab3 = st.tabs(["📊 生成报表", "📋 报表历史与复核", "✏️ 数据管理"])

# ==================== 生成报表 ====================
with tab1:
    st.subheader("📤 1. 上传Excel文件")
    uploaded = st.file_uploader("选择文件（.xlsx, .xls）", type=["xlsx", "xls"])
    
    if uploaded:
        try:
            xls = pd.ExcelFile(uploaded)
            sheets = xls.sheet_names
            selected_sheet = st.selectbox("选择数据Sheet", sheets, index=0)
            
            # 自动导入规则（从“基础配置表”）
            with st.spinner("正在导入城市规则..."):
                count, error_msg = import_rules_from_excel(xls)
                if count > 0:
                    st.success(f"✅ 成功导入 {count} 个城市的规则！")
                else:
                    st.warning(f"⚠️ 规则导入：{error_msg}，但您仍可继续生成报表（需手动输入比例）。")
            
            # 读取数据
            df_raw = pd.read_excel(uploaded, sheet_name=selected_sheet, header=None)
            # 检测表头行（查找包含“城市”或“分公司”的行）
            header_row = 0
            for i, row in df_raw.iterrows():
                row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                if any(key in row_text for key in ['城市', '公司', '参保', '缴费', '社保', '公积金']):
                    header_row = i
                    break
            if header_row > 0:
                st.info(f"自动检测到表头行：第 {header_row} 行")
                # 让用户确认或修改
                adjust = st.checkbox("手动调整表头行号")
                if adjust:
                    header_row = st.number_input("表头行号（从0开始）", min_value=0, max_value=len(df_raw)-1, value=header_row, step=1)
            
            # 提取数据
            if header_row < len(df_raw):
                new_header = df_raw.iloc[header_row]
                new_header = [str(h).strip() for h in new_header]
                df_data = df_raw.iloc[header_row+1:].copy()
                df_data.columns = new_header
                df_data = df_data.reset_index(drop=True)
                df_data = df_data.dropna(how='all')
                st.success(f"成功提取数据，共 {len(df_data)} 行")
                st.dataframe(df_data.head(5))
                
                # ===== 步骤2：手动列映射 =====
                st.subheader("📤 2. 映射列名（请为每个字段选择对应的列）")
                cols = df_data.columns.tolist()
                
                col_company = st.selectbox("公司名称（分公司）", [''] + cols, key='map_company')
                col_city = st.selectbox("所属城市", [''] + cols, key='map_city')
                col_people = st.selectbox("参保人数", [''] + cols, key='map_people')
                col_social_unit = st.selectbox("社保单位合计", [''] + cols, key='map_social_unit')
                col_social_personal = st.selectbox("社保个人合计", [''] + cols, key='map_social_personal')
                col_fund_unit = st.selectbox("公积金单位合计", [''] + cols, key='map_fund_unit')
                col_fund_personal = st.selectbox("公积金个人合计", [''] + cols, key='map_fund_personal')
                col_total = st.selectbox("总费用（可选，若无可自动计算）", [''] + cols, key='map_total')
                
                # 检查是否映射了必要字段
                required_mapped = col_company and col_city and col_social_unit and col_social_personal
                if not required_mapped:
                    st.warning("请至少映射：公司名称、城市、社保单位合计、社保个人合计。")
                else:
                    st.success("✅ 必要字段已映射，可继续生成报表。")
                    
                    # 步骤3：选择生成参数
                    st.subheader("📤 3. 选择报表参数")
                    # 提取公司列表
                    df_company = df_data[[col_company, col_city]].drop_duplicates()
                    df_company['display'] = df_company[col_company] + ' (' + df_company[col_city] + ')'
                    options = df_company['display'].tolist()
                    company_map = {row['display']: (row[col_company], row[col_city]) for _, row in df_company.iterrows()}
                    
                    selected_options = st.multiselect("选择公司（可多选）", options)
                    
                    report_type = st.selectbox("报表类型", ["月度申报", "年度汇算"])
                    year = st.selectbox("年份", [2025,2024,2023], index=0)
                    month = None
                    if report_type == "月度申报":
                        month = st.selectbox("月份", list(range(1,13)), index=11)
                    
                    export_format = st.radio("导出格式", ["Excel (.xlsx)", "CSV (.csv)"], horizontal=True)
                    
                    if st.button("🚀 生成报表", type="primary"):
                        if not selected_options:
                            st.error("请至少选择一个公司")
                        else:
                            # 获取规则
                            rule_df = pd.DataFrame(load_table('rules'))
                            generated_files = []
                            summary_list = []
                            errors = []
                            
                            for selected in selected_options:
                                company_name, city = company_map[selected]
                                # 筛选数据
                                city_data = df_data[df_data[col_city] == city]
                                if col_company:
                                    city_data = city_data[city_data[col_company] == company_name]
                                if city_data.empty:
                                    errors.append(f"{company_name}: 无数据")
                                    continue
                                
                                # 汇总
                                total_people = city_data[col_people].sum() if col_people else len(city_data)
                                total_social_unit = city_data[col_social_unit].sum()
                                total_social_personal = city_data[col_social_personal].sum()
                                total_fund_unit = city_data[col_fund_unit].sum() if col_fund_unit else 0
                                total_fund_personal = city_data[col_fund_personal].sum() if col_fund_personal else 0
                                if col_total:
                                    total_cost = city_data[col_total].sum()
                                else:
                                    total_cost = total_social_unit + total_social_personal + total_fund_unit + total_fund_personal
                                
                                # 匹配规则来源
                                matched = rule_df[(rule_df['city'] == city) & (rule_df['report_type'] == '月度申报')]
                                source_quote = matched.iloc[0]['source_quote'] if not matched.empty else '未匹配规则'
                                
                                # 生成Excel
                                wb = Workbook()
                                ws = wb.active
                                ws.title = "汇总报表"
                                ws.append(['指标', '金额'])
                                ws.append(['公司名称', company_name])
                                ws.append(['所属城市', city])
                                ws.append(['参保人数', int(total_people)])
                                ws.append(['社保单位合计', round(total_social_unit, 2)])
                                ws.append(['社保个人合计', round(total_social_personal, 2)])
                                ws.append(['公积金单位', round(total_fund_unit, 2)])
                                ws.append(['公积金个人', round(total_fund_personal, 2)])
                                ws.append(['总费用', round(total_cost, 2)])
                                
                                audit = wb.create_sheet("AuditTrail")
                                audit.append(["时间", "操作", "详情"])
                                audit.append([datetime.now().isoformat(), "GENERATED", f"公司:{company_name}, 城市:{city}, 规则来源:{source_quote}"])
                                
                                ws.insert_rows(1)
                                ws['A1'] = '⚠️ 待复核版'
                                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
                                
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
                                    "source_quote": source_quote,
                                    "total_people": int(total_people),
                                    "total_cost": round(total_cost, 2),
                                    "reviewer": "",
                                    "reviewed_at": "",
                                    "review_comment": ""
                                }
                                insert_or_update_export(record)
                                insert_audit_log({
                                    "id": str(uuid.uuid4())[:8],
                                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "action": "生成报表",
                                    "detail": f"公司:{company_name}, 城市:{city}, 人数:{total_people}, 总成本:{round(total_cost,2)}",
                                    "report_id": export_id
                                })
                                summary_list.append({"公司": company_name, "城市": city, "人数": int(total_people), "总成本": round(total_cost, 2)})
                            
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
                                    st.download_button(f"📥 下载 {fname}", data=BytesIO(data), file_name=fname)
            else:
                st.error("表头行号无效，请检查数据。")
        except Exception as e:
            st.error(f"❌ 处理文件时出错：{str(e)}")
    else:
        st.info("请上传Excel文件开始。")

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

st.caption(f"📁 数据库位置：`{DB_PATH}`（请妥善备份）")
