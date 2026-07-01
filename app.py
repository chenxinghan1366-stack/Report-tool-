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

# ===== 专为"全地区版"Excel定制的规则导入 =====
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
        # 列名映射（适配您的列名）
        col_map = {}
        for col in df_rules.columns:
            col_lower = str(col).lower().strip()
            if '所属城市' in col_lower or '城市' in col_lower:
                col_map['城市'] = col
            elif '养老保险-单位比例' in col_lower or '养老单位比例' in col_lower:
                col_map['单位社保比例'] = col
            elif '养老保险-个人比例' in col_lower or '养老个人比例' in col_lower:
                col_map['个人社保比例'] = col
            elif '医疗保险-单位比例' in col_lower or '医疗单位比例' in col_lower:
                col_map['医疗单位比例'] = col
            elif '医疗保险-个人比例' in col_lower or '医疗个人比例' in col_lower:
                col_map['医疗个人比例'] = col
            elif '失业保险-单位比例' in col_lower or '失业单位比例' in col_lower:
                col_map['失业单位比例'] = col
            elif '失业保险-个人比例' in col_lower or '失业个人比例' in col_lower:
                col_map['失业个人比例'] = col
            elif '工伤保险-单位比例' in col_lower or '工伤单位比例' in col_lower:
                col_map['工伤单位比例'] = col
            elif '生育保险-单位比例' in col_lower or '生育单位比例' in col_lower:
                col_map['生育单位比例'] = col

        required = ['城市', '单位社保比例', '个人社保比例']
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
                # 医保、失业、工伤、生育的比例，如果有则读取，否则默认
                unit_medical = float(row[col_map.get('医疗单位比例', '')]) if col_map.get('医疗单位比例') is not None and not pd.isna(row[col_map['医疗单位比例']]) else 0.08
                personal_medical = float(row[col_map.get('医疗个人比例', '')]) if col_map.get('医疗个人比例') is not None and not pd.isna(row[col_map['医疗个人比例']]) else 0.02
                unit_unemployment = float(row[col_map.get('失业单位比例', '')]) if col_map.get('失业单位比例') is not None and not pd.isna(row[col_map['失业单位比例']]) else 0.005
                personal_unemployment = float(row[col_map.get('失业个人比例', '')]) if col_map.get('失业个人比例') is not None and not pd.isna(row[col_map['失业个人比例']]) else 0.005
                unit_injury = float(row[col_map.get('工伤单位比例', '')]) if col_map.get('工伤单位比例') is not None and not pd.isna(row[col_map['工伤单位比例']]) else 0.002
                unit_maternity = float(row[col_map.get('生育单位比例', '')]) if col_map.get('生育单位比例') is not None and not pd.isna(row[col_map['生育单位比例']]) else 0.008
            except (ValueError, TypeError):
                continue

            # 社保单位合计 = 养老 + 医疗 + 失业 + 工伤 + 生育
            unit_social_total = unit_social + unit_medical + unit_unemployment + unit_injury + unit_maternity
            personal_social_total = personal_social + personal_medical + personal_unemployment

            rules_list.append({
                "id": str(uuid.uuid4())[:8],
                "province": city,
                "city": city,
                "report_type": "月度申报",
                "unit_social": unit_social_total,  # 养老+医疗+失业+工伤+生育
                "personal_social": personal_social_total,  # 养老+医疗+失业
                "unit_fund": 0.12,  # 公积金默认12%（如没有，可后续手动修改）
                "personal_fund": 0.12,
                "social_min": 0,
                "social_max": 999999,
                "fund_min": 0,
                "fund_max": 999999,
                "source_quote": f"自动从{rule_sheet}导入（合计比例）"
            })
        if rules_list:
            save_table('rules', rules_list)
            return len(rules_list), None
        else:
            return 0, "未解析到有效数据"
    except Exception as e:
        return 0, f"解析规则时出错：{str(e)}"

# ===== 从汇总数据生成报表 =====
def generate_from_summary(df_summary, rules_data, company_name, city):
    """直接从汇总数据生成报表，不依赖员工明细"""
    # 匹配规则
    rule_df = pd.DataFrame(rules_data)
    matched = rule_df[(rule_df['city'] == city) & (rule_df['report_type'] == '月度申报')]
    if matched.empty:
        return None, f"未找到 {city} 的规则"

    rule = matched.iloc[0]
    unit_social = rule.get('unit_social', 0.16)
    personal_social = rule.get('personal_social', 0.08)
    unit_fund = rule.get('unit_fund', 0.12)
    personal_fund = rule.get('personal_fund', 0.12)

    # 从汇总数据中提取关键指标
    # 假设数据包含：缴费基数、参保人数等
    # 根据您的实际列名调整
    total_people = df_summary.get('参保人数', 0)
    if isinstance(total_people, pd.Series):
        total_people = total_people.sum()
    total_base = df_summary.get('缴费基数', 0)
    if isinstance(total_base, pd.Series):
        total_base = total_base.sum()
    if total_base == 0:
        total_base = df_summary.get('社保缴费基数', 0)
        if isinstance(total_base, pd.Series):
            total_base = total_base.sum()

    # 计算
    unit_amount = total_base * unit_social
    personal_amount = total_base * personal_social
    fund_unit_amount = total_base * unit_fund
    fund_personal_amount = total_base * personal_fund
    total_cost = unit_amount + personal_amount + fund_unit_amount + fund_personal_amount

    # 生成汇总表数据
    result = {
        '公司': company_name,
        '城市': city,
        '人数': total_people,
        '缴费基数总额': total_base,
        '单位社保': unit_amount,
        '个人社保': personal_amount,
        '单位公积金': fund_unit_amount,
        '个人公积金': fund_personal_amount,
        '总费用': total_cost
    }
    return result, None

# ---------- 初始化 ----------
init_db()
if not load_table('companies'):
    save_table('companies', [])
if not load_table('rules'):
    save_table('rules', [])

# ---------- Streamlit 页面 ----------
st.set_page_config(page_title="本地社保报表系统", layout="wide")
st.title("📋 本地社保公积金报表系统")
st.markdown("🔒 **适配您的「全地区版」Excel，支持汇总数据直接生成报表**")

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

    st.subheader("📤 1. 上传数据")
    st.caption("上传Excel后，系统自动从「基础配置表」导入规则，从「月度明细数据表」读取汇总数据")
    uploaded = st.file_uploader("选择文件（.xlsx, .xls）", type=["xlsx", "xls"])

    df_summary = None
    df_std = None

    if uploaded:
        try:
            xls = pd.ExcelFile(uploaded)
            sheets = xls.sheet_names
            selected_sheet = st.selectbox("选择Sheet", sheets, index=0)

            # ===== 自动导入规则 =====
            with st.spinner("正在自动识别并导入城市规则..."):
                count, error_msg = import_rules_from_excel(xls)
                if count > 0:
                    st.success(f"✅ 成功从「基础配置表」导入 {count} 个城市的规则！")
                    rules_data = load_table('rules')
                else:
                    st.warning(f"⚠️ 规则导入：{error_msg}，请检查「基础配置表」列名是否正确")

            if selected_sheet:
                df_summary = pd.read_excel(uploaded, sheet_name=selected_sheet)

                if df_summary is not None:
                    st.info(f"📊 读取到 {len(df_summary)} 行，{len(df_summary.columns)} 列")
                    st.dataframe(df_summary.head(5))

                    # 检测表头行
                    header_row = 0
                    for i, row in df_summary.iterrows():
                        row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                        if any(key in row_text for key in ['所属城市', '城市', '公司', '缴费基数', '参保人数']):
                            header_row = i
                            break

                    if header_row == 0:
                        header_row = st.number_input("表头行号（从0开始）", min_value=0, max_value=len(df_summary)-1, value=0, step=1)
                    else:
                        st.info(f"自动检测到表头行：第 {header_row} 行")
                        override = st.checkbox("手动调整表头行号")
                        if override:
                            header_row = st.number_input("表头行号（从0开始）", min_value=0, max_value=len(df_summary)-1, value=header_row, step=1)

                    if header_row < len(df_summary):
                        new_header = df_summary.iloc[header_row]
                        new_header = [str(h) for h in new_header]
                        df_summary = df_summary[header_row+1:]
                        df_summary.columns = new_header
                        df_summary = df_summary.reset_index(drop=True)
                        df_summary = df_summary.dropna(how='all')
                        st.success(f"成功提取数据，共 {len(df_summary)} 行")
                        st.dataframe(df_summary.head(5))

                        # 提取城市列表
                        city_col = None
                        for col in df_summary.columns:
                            if '城市' in str(col) or '地区' in str(col):
                                city_col = col
                                break
                        if city_col:
                            cities = df_summary[city_col].unique().tolist()
                            st.info(f"检测到城市：{', '.join(cities)}")
                        else:
                            st.warning("未找到城市列，请确保数据包含「城市」或「所属城市」列")

                        df_std = df_summary

        except Exception as e:
            st.error(f"❌ 读取文件失败：{str(e)}")

    # ===== 选择公司和生成报表 =====
    st.subheader("📤 2. 选择公司并生成报表")

    if df_std is not None and not df_std.empty:
        # 提取城市
        city_col = None
        for col in df_std.columns:
            if '城市' in str(col) or '地区' in str(col):
                city_col = col
                break

        if city_col:
            city_options = df_std[city_col].unique().tolist()
            selected_cities = st.multiselect("选择城市（可多选）", city_options)

            report_type = st.selectbox("报表类型", ["月度申报", "年度汇算"])
            year = st.selectbox("年份", [2025, 2024, 2023], index=0)
            month = None
            if report_type == "月度申报":
                month = st.selectbox("月份", list(range(1,13)), index=11)

            export_format = st.radio("导出格式", ["Excel (.xlsx)", "CSV (.csv)"], horizontal=True)

            if st.button("🚀 生成报表", type="primary"):
                if not selected_cities:
                    st.error("请至少选择一个城市")
                else:
                    rules_data = load_table('rules')
                    generated_files = []
                    summary_list = []
                    errors = []

                    for city in selected_cities:
                        city_data = df_std[df_std[city_col] == city]

                        # 计算该城市的汇总数据
                        # 查找"缴费基数"列
                        base_col = None
                        for col in df_std.columns:
                            if '基数' in str(col) or '缴费' in str(col):
                                base_col = col
                                break
                        people_col = None
                        for col in df_std.columns:
                            if '人数' in str(col) or '参保' in str(col):
                                people_col = col
                                break

                        if base_col is None:
                            errors.append(f"{city}: 未找到缴费基数列")
                            continue

                        total_base = city_data[base_col].sum()
                        total_people = city_data[people_col].sum() if people_col else len(city_data)

                        # 从规则表匹配
                        rule_df = pd.DataFrame(rules_data)
                        matched = rule_df[(rule_df['city'] == city) & (rule_df['report_type'] == '月度申报')]
                        if matched.empty:
                            errors.append(f"{city}: 未找到规则，请先在「管理数据」中导入规则")
                            continue

                        rule = matched.iloc[0]
                        unit_social = rule.get('unit_social', 0.16)
                        personal_social = rule.get('personal_social', 0.08)
                        unit_fund = rule.get('unit_fund', 0.12)
                        personal_fund = rule.get('personal_fund', 0.12)

                        unit_amount = total_base * unit_social
                        personal_amount = total_base * personal_social
                        fund_unit_amount = total_base * unit_fund
                        fund_personal_amount = total_base * personal_fund
                        total_cost = unit_amount + personal_amount + fund_unit_amount + fund_personal_amount

                        # 生成Excel
                        wb = Workbook()
                        ws = wb.active
                        ws.title = "汇总报表"
                        ws.append(['指标', '金额'])
                        ws.append(['城市', city])
                        ws.append(['参保人数', total_people])
                        ws.append(['缴费基数总额', round(total_base, 2)])
                        ws.append(['单位社保', round(unit_amount, 2)])
                        ws.append(['个人社保', round(personal_amount, 2)])
                        ws.append(['单位公积金', round(fund_unit_amount, 2)])
                        ws.append(['个人公积金', round(fund_personal_amount, 2)])
                        ws.append(['总费用', round(total_cost, 2)])

                        # 审计日志
                        audit = wb.create_sheet("AuditTrail")
                        audit.append(["时间", "操作", "详情"])
                        audit.append([datetime.now().isoformat(), "GENERATED", f"城市:{city}, 规则:{rule.get('source_quote','')}"])

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
                            "source_quote": rule.get('source_quote', ''),
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
        st.info("请先上传数据并完成表头识别，系统将自动识别城市列表。")

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
