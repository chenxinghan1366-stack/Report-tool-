import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook, load_workbook
from io import BytesIO
import uuid
import sqlite3
import os
import zipfile

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

# ===== 专为“全地区版”定制的规则导入 =====
def import_rules_from_excel(xls):
    rule_sheet = None
    for sheet in ["基础配置表", "规则", "城市规则", "配置表"]:
        if sheet in xls.sheet_names:
            rule_sheet = sheet
            break
    if rule_sheet is None:
        return 0, "未找到「基础配置表」"

    try:
        df_rules = pd.read_excel(xls, sheet_name=rule_sheet)
        # 直接匹配您的列名
        col_map = {}
        for col in df_rules.columns:
            col_lower = str(col).lower().strip()
            if '所属城市' in col_lower or '城市' in col_lower:
                col_map['城市'] = col
            elif '养老保险-单位比例' in col_lower:
                col_map['单位社保比例'] = col
            elif '养老保险-个人比例' in col_lower:
                col_map['个人社保比例'] = col
            elif '医疗保险-单位比例' in col_lower:
                col_map['医疗单位比例'] = col
            elif '医疗保险-个人比例' in col_lower:
                col_map['医疗个人比例'] = col
            elif '失业保险-单位比例' in col_lower:
                col_map['失业单位比例'] = col
            elif '失业保险-个人比例' in col_lower:
                col_map['失业个人比例'] = col
            elif '工伤保险-单位比例' in col_lower:
                col_map['工伤单位比例'] = col
            elif '生育保险-单位比例' in col_lower:
                col_map['生育单位比例'] = col
            elif '公积金-单位比例' in col_lower:
                col_map['单位公积金比例'] = col
            elif '公积金-个人比例' in col_lower:
                col_map['个人公积金比例'] = col
            elif '缴费基数下限' in col_lower:
                col_map['社保最低基数'] = col
            elif '缴费基数上限' in col_lower:
                col_map['社保最高基数'] = col

        required = ['城市', '单位社保比例', '个人社保比例', '单位公积金比例', '个人公积金比例']
        missing = [r for r in required if r not in col_map]
        if missing:
            return 0, f"缺少必要列：{', '.join(missing)}，请检查表头行是否正确"

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
st.title("📋 本地社保公积金报表系统（适配汇总数据版）")
st.markdown("🔒 **直接读取您的汇总数据，自动匹配规则，一键生成报表**")

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
    companies_data = load_table('companies')
    rules_data = load_table('rules')

    st.subheader("📤 1. 上传数据（自动导入规则）")
    st.caption("上传Excel后，系统自动从「基础配置表」导入规则，从您选择的Sheet读取汇总数据")
    uploaded = st.file_uploader("选择文件（.xlsx, .xls）", type=["xlsx", "xls"])

    df_data = None
    selected_sheet = None
    if uploaded:
        try:
            xls = pd.ExcelFile(uploaded)
            sheets = xls.sheet_names
            selected_sheet = st.selectbox("选择Sheet", sheets, index=0)

            # 自动导入规则
            with st.spinner("正在自动识别并导入城市规则..."):
                count, error_msg = import_rules_from_excel(xls)
                if count > 0:
                    st.success(f"✅ 成功从「基础配置表」导入 {count} 个城市的规则！")
                    rules_data = load_table('rules')
                else:
                    st.warning(f"⚠️ 规则导入：{error_msg}，请检查「基础配置表」列名是否正确")

            if selected_sheet:
                df_data = pd.read_excel(uploaded, sheet_name=selected_sheet)
                if df_data is not None:
                    st.info(f"📊 读取到 {len(df_data)} 行，{len(df_data.columns)} 列")
                    st.dataframe(df_data.head(5))

                    # 检测表头行（跳过标题行）
                    header_row = 0
                    for i, row in df_data.iterrows():
                        row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                        if any(key in row_text for key in ['所属城市', '城市', '公司', '缴费基数', '参保人数']):
                            header_row = i
                            break
                    if header_row == 0:
                        header_row = st.number_input("表头行号（从0开始）", min_value=0, max_value=len(df_data)-1, value=0, step=1)
                    else:
                        st.info(f"自动检测到表头行：第 {header_row} 行")
                        override = st.checkbox("手动调整表头行号")
                        if override:
                            header_row = st.number_input("表头行号（从0开始）", min_value=0, max_value=len(df_data)-1, value=header_row, step=1)

                    if header_row < len(df_data):
                        new_header = df_data.iloc[header_row]
                        new_header = [str(h) for h in new_header]
                        df_data = df_data[header_row+1:]
                        df_data.columns = new_header
                        df_data = df_data.reset_index(drop=True)
                        df_data = df_data.dropna(how='all')
                        st.success(f"成功提取数据，共 {len(df_data)} 行")
                        st.dataframe(df_data.head(5))

        except Exception as e:
            st.error(f"❌ 读取文件失败：{str(e)}")

    # ===== 选择城市并生成报表 =====
    st.subheader("📤 2. 选择城市并生成报表")

    if df_data is not None and not df_data.empty:
        # 提取城市列
        city_col = None
        for col in df_data.columns:
            if '城市' in str(col) or '所属城市' in str(col):
                city_col = col
                break
        if city_col is None:
            st.error("未找到城市列，请确保数据包含「城市」或「所属城市」列")
        else:
            cities = df_data[city_col].unique().tolist()
            selected_cities = st.multiselect("选择城市（可多选）", cities)

            report_type = st.selectbox("报表类型", ["月度申报", "年度汇算"])
            year = st.selectbox("年份", [2025,2024,2023], index=0)
            month = None
            if report_type == "月度申报":
                month = st.selectbox("月份", list(range(1,13)), index=11)

            export_format = st.radio("导出格式", ["Excel (.xlsx)", "CSV (.csv)"], horizontal=True)

            if st.button("🚀 生成报表", type="primary"):
                if not selected_cities:
                    st.error("请至少选择一个城市")
                else:
                    generated_files = []
                    summary_list = []
                    errors = []

                    # 确定列名（根据您的实际列名）
                    # 社保单位合计、社保个人合计、公积金单位、公积金个人、总费用等
                    # 您的“月度明细数据表”中有：社保合计-单位部分, 社保合计-个人部分, 公积金-单位部分, 公积金-个人部分, 社保+公积金合计-总金额
                    # 也可能使用“月度汇总报表(12月单月)”中的列：社保单位合计, 社保个人合计, 公积金单位合计, 公积金个人合计, 社保+公积金总金额
                    # 我们智能查找
                    def find_col(possible_names):
                        for name in possible_names:
                            for col in df_data.columns:
                                if name.lower() in str(col).lower():
                                    return col
                        return None

                    col_people = find_col(['参保人数'])
                    col_social_unit = find_col(['社保单位合计', '社保合计-单位部分'])
                    col_social_personal = find_col(['社保个人合计', '社保合计-个人部分'])
                    col_fund_unit = find_col(['公积金单位合计', '公积金-单位部分'])
                    col_fund_personal = find_col(['公积金个人合计', '公积金-个人部分'])
                    col_total = find_col(['社保+公积金合计-总金额', '社保+公积金总金额'])

                    if col_social_unit is None and col_social_personal is None:
                        st.error("未找到社保金额列，请确认数据格式")
                        st.stop()

                    rule_df = pd.DataFrame(rules_data)
                    for city in selected_cities:
                        city_data = df_data[df_data[city_col] == city]
                        if city_data.empty:
                            errors.append(f"{city}: 无数据")
                            continue

                        # 汇总
                        total_people = city_data[col_people].sum() if col_people else len(city_data)
                        total_social_unit = city_data[col_social_unit].sum() if col_social_unit else 0
                        total_social_personal = city_data[col_social_personal].sum() if col_social_personal else 0
                        total_fund_unit = city_data[col_fund_unit].sum() if col_fund_unit else 0
                        total_fund_personal = city_data[col_fund_personal].sum() if col_fund_personal else 0
                        total_cost = city_data[col_total].sum() if col_total else (total_social_unit + total_social_personal + total_fund_unit + total_fund_personal)

                        # 匹配规则（用于显示来源）
                        matched = rule_df[(rule_df['city'] == city) & (rule_df['report_type'] == '月度申报')]
                        source_quote = matched.iloc[0]['source_quote'] if not matched.empty else '未匹配规则'

                        # 生成Excel
                        wb = Workbook()
                        ws = wb.active
                        ws.title = "汇总报表"
                        ws.append(['指标', '金额'])
                        ws.append(['城市', city])
                        ws.append(['参保人数', int(total_people)])
                        ws.append(['社保单位合计', round(total_social_unit, 2)])
                        ws.append(['社保个人合计', round(total_social_personal, 2)])
                        ws.append(['公积金单位', round(total_fund_unit, 2)])
                        ws.append(['公积金个人', round(total_fund_personal, 2)])
                        ws.append(['总费用', round(total_cost, 2)])

                        # 审计日志
                        audit = wb.create_sheet("AuditTrail")
                        audit.append(["时间", "操作", "详情"])
                        audit.append([datetime.now().isoformat(), "GENERATED", f"城市:{city}, 规则来源:{source_quote}"])

                        ws.insert_rows(1)
                        ws['A1'] = '⚠️ 待复核版'
                        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)

                        output = BytesIO()
                        wb.save(output)
                        output.seek(0)
                        fname = f"{city}_{year}{month or ''}.xlsx"
                        generated_files.append((fname, output.getvalue()))

                        # 保存历史
                        export_id = str(uuid.uuid4())[:8]
                        record = {
                            "id": export_id,
                            "company": city,
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
                            "detail": f"城市:{city}, 人数:{total_people}, 总成本:{round(total_cost,2)}",
                            "report_id": export_id
                        })
                        summary_list.append({"城市": city, "人数": int(total_people), "总成本": round(total_cost, 2)})

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
        st.info("请先上传数据并完成表头识别。")

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
