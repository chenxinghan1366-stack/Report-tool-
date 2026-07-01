import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from io import BytesIO
import json
import uuid
import os
import sqlite3
import zipfile

# ---------- 数据库初始化 ----------
DB_PATH = "data.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    # 公司表
    c.execute('''CREATE TABLE IF NOT EXISTS companies (
        id TEXT PRIMARY KEY,
        company_name TEXT,
        province TEXT,
        city TEXT,
        district TEXT
    )''')
    # 模板表
    c.execute('''CREATE TABLE IF NOT EXISTS templates (
        id TEXT PRIMARY KEY,
        province TEXT,
        city TEXT,
        template_name TEXT,
        report_type TEXT,
        version TEXT,
        required_fields TEXT
    )''')
    # 规则表
    c.execute('''CREATE TABLE IF NOT EXISTS rules (
        id TEXT PRIMARY KEY,
        province TEXT,
        city TEXT,
        report_type TEXT,
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
    # 报表历史表
    c.execute('''CREATE TABLE IF NOT EXISTS export_history (
        id TEXT PRIMARY KEY,
        company TEXT,
        city TEXT,
        report_type TEXT,
        year INTEGER,
        month INTEGER,
        generated_at TEXT,
        status TEXT,
        source_quote TEXT,
        total_people INTEGER,
        total_cost REAL,
        reviewer TEXT,
        reviewed_at TEXT,
        review_comment TEXT,
        file_data BLOB
    )''')
    # 操作日志表
    c.execute('''CREATE TABLE IF NOT EXISTS audit_logs (
        id TEXT PRIMARY KEY,
        timestamp TEXT,
        action TEXT,
        detail TEXT,
        report_id TEXT
    )''')
    conn.commit()
    conn.close()

# ---------- 数据库操作函数 ----------
def load_from_db(table_name):
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql(f"SELECT * FROM {table_name}", conn)
    conn.close()
    return df.to_dict('records') if not df.empty else []

def save_to_db(table_name, data_list):
    if not data_list:
        return
    conn = sqlite3.connect(DB_PATH)
    df = pd.DataFrame(data_list)
    df.to_sql(table_name, conn, if_exists='replace', index=False)
    conn.close()

def save_export_history(record):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    # 如果存在则更新，否则插入
    c.execute('''INSERT OR REPLACE INTO export_history 
        (id, company, city, report_type, year, month, generated_at, status, source_quote, total_people, total_cost, reviewer, reviewed_at, review_comment)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        (record['id'], record['company'], record['city'], record['report_type'],
         record['year'], record['month'], record['generated_at'], record['status'],
         record['source_quote'], record['total_people'], record['total_cost'],
         record.get('reviewer', ''), record.get('reviewed_at', ''), record.get('review_comment', '')))
    conn.commit()
    conn.close()

def get_export_history():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql("SELECT * FROM export_history ORDER BY generated_at DESC", conn)
    conn.close()
    return df.to_dict('records') if not df.empty else []

def save_audit_log(log):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT INTO audit_logs (id, timestamp, action, detail, report_id)
        VALUES (?, ?, ?, ?, ?)''',
        (log['id'], log['timestamp'], log['action'], log['detail'], log.get('report_id', '')))
    conn.commit()
    conn.close()

def get_audit_logs():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql("SELECT * FROM audit_logs ORDER BY timestamp DESC", conn)
    conn.close()
    return df.to_dict('records') if not df.empty else []

def init_default_data():
    # 如果数据库为空，初始化示例数据
    if not load_from_db('companies'):
        default_companies = [
            {"id": str(uuid.uuid4())[:8], "company_name": "上海科技公司", "province": "上海", "city": "上海市", "district": "浦东新区"},
            {"id": str(uuid.uuid4())[:8], "company_name": "深圳科技公司", "province": "广东", "city": "深圳市", "district": "南山区"}
        ]
        save_to_db('companies', default_companies)
    if not load_from_db('rules'):
        default_rules = [
            {"id": str(uuid.uuid4())[:8], "province": "上海", "city": "上海市", "report_type": "月度申报",
             "unit_social": 0.16, "personal_social": 0.08, "unit_fund": 0.07, "personal_fund": 0.07,
             "social_min": 7310, "social_max": 36549, "fund_min": 2590, "fund_max": 34188,
             "source_quote": "沪人社规〔2024〕22号"},
            {"id": str(uuid.uuid4())[:8], "province": "广东", "city": "深圳市", "report_type": "月度申报",
             "unit_social": 0.14, "personal_social": 0.08, "unit_fund": 0.05, "personal_fund": 0.05,
             "social_min": 2360, "social_max": 29727, "fund_min": 2360, "fund_max": 29727,
             "source_quote": "深人社规〔2024〕3号"}
        ]
        save_to_db('rules', default_rules)

# ---------- 初始化 ----------
init_db()
init_default_data()

# ---------- 页面配置 ----------
st.set_page_config(page_title="智能社保报表匹配", layout="wide")
st.title("🧠 智能社保报表匹配系统")
st.markdown("**数据持久化保存 | 支持批量生成 | 完整可追溯**")

# ---------- 仪表盘 ----------
col1, col2, col3, col4 = st.columns(4)
history = get_export_history()
companies = load_from_db('companies')
rules = load_from_db('rules')
with col1:
    st.metric("🏢 公司总数", len(companies))
with col2:
    st.metric("🏙️ 城市总数", len(set(c['city'] for c in companies)))
with col3:
    st.metric("📋 总报表数", len(history))
with col4:
    pending = len([h for h in history if h['status'] == '待复核'])
    st.metric("🟡 待复核", pending)

tab1, tab2, tab3 = st.tabs(["📊 生成报表", "📋 报表历史", "✏️ 管理数据"])

# ==================== 生成报表 ====================
with tab1:
    companies_data = load_from_db('companies')
    rules_data = load_from_db('rules')
    companies_df = pd.DataFrame(companies_data)

    if companies_df.empty:
        st.warning("请先在「管理数据」中添加公司信息。")
        st.stop()

    st.subheader("📤 批量生成报表")
    st.markdown("选择多个公司，系统将自动为每个公司生成报表，并打包下载。")

    # 选择多个公司
    company_options = [f"{c['company_name']} ({c['city']})" for c in companies_data]
    selected_indices = st.multiselect("选择公司", range(len(company_options)), format_func=lambda x: company_options[x])

    report_type = st.selectbox("报表类型", ["月度申报", "年度汇算"])
    year = st.selectbox("年份", [2025, 2024, 2023], index=0)
    month = None
    if report_type == "月度申报":
        month = st.selectbox("月份", list(range(1,13)), index=11)

    # 上传工资表（批量生成使用同一个工资表）
    st.markdown("**上传员工工资表（所有公司共用）**")
    st.markdown("Excel 格式需包含列：公司、城市、姓名、工资基数")
    uploaded_wage = st.file_uploader("选择工资表 (支持 .xlsx)", type=["xlsx"], key="batch_wage")
    df_wage_all = None
    if uploaded_wage is not None:
        try:
            df_wage_all = pd.read_excel(uploaded_wage)
            st.success(f"成功读取 {len(df_wage_all)} 条员工记录")
            st.dataframe(df_wage_all.head(5))
        except Exception as e:
            st.error(f"读取失败：{e}")

    if st.button("🚀 批量生成报表", type="primary"):
        if df_wage_all is None or df_wage_all.empty:
            st.error("请先上传员工工资表")
        elif not selected_indices:
            st.error("请至少选择一个公司")
        else:
            generated_files = []
            summary_list = []

            for idx in selected_indices:
                company_data = companies_data[idx]
                company_name = company_data['company_name']
                city = company_data['city']
                province = company_data['province']

                # 过滤该公司的员工数据
                df_wage = df_wage_all[df_wage_all['公司'] == company_name]
                if df_wage.empty:
                    st.warning(f"⚠️ 未找到 {company_name} 的员工数据，跳过")
                    continue

                # 匹配规则
                rule_df = pd.DataFrame(rules_data)
                matched_rule = rule_df[(rule_df['城市'] == city) & (rule_df['报表类型'] == report_type)]
                if matched_rule.empty:
                    st.warning(f"⚠️ 未找到 {city} 的规则，跳过")
                    continue
                rule = matched_rule.iloc[0]

                # 计算
                unit_social = rule.get('unit_social', 0.16)
                personal_social = rule.get('personal_social', 0.08)
                unit_fund = rule.get('unit_fund', 0.12)
                personal_fund = rule.get('personal_fund', 0.12)
                social_min = rule.get('social_min', 0)
                social_max = rule.get('social_max', float('inf'))

                if "社保基数" not in df_wage.columns:
                    df_wage["社保基数"] = df_wage["工资基数"]
                if "公积金基数" not in df_wage.columns:
                    df_wage["公积金基数"] = df_wage["工资基数"]

                df_wage["社保基数调整"] = df_wage["社保基数"].apply(lambda x: max(social_min, min(x, social_max)))

                df_wage["单位社保"] = df_wage["社保基数调整"] * unit_social
                df_wage["个人社保"] = df_wage["社保基数调整"] * personal_social
                df_wage["单位公积金"] = df_wage["公积金基数"] * unit_fund
                df_wage["个人公积金"] = df_wage["公积金基数"] * personal_fund
                df_wage["单位合计"] = df_wage["单位社保"] + df_wage["单位公积金"]
                df_wage["个人合计"] = df_wage["个人社保"] + df_wage["个人公积金"]
                df_wage["总成本"] = df_wage["单位合计"] + df_wage["个人合计"]

                # 生成 Excel
                wb = Workbook()
                ws_detail = wb.active
                ws_detail.title = "员工明细"
                headers = ["公司", "城市", "姓名", "工资基数", "社保基数", "单位社保", "个人社保", "单位公积金", "个人公积金", "单位合计", "个人合计", "总成本"]
                ws_detail.append(headers)
                for _, row in df_wage.iterrows():
                    ws_detail.append([
                        row["公司"], row["城市"], row["姓名"], row["工资基数"],
                        row["社保基数调整"],
                        round(row["单位社保"], 2), round(row["个人社保"], 2),
                        round(row["单位公积金"], 2), round(row["个人公积金"], 2),
                        round(row["单位合计"], 2), round(row["个人合计"], 2),
                        round(row["总成本"], 2)
                    ])

                # 汇总
                total_cost = df_wage["总成本"].sum()
                ws_summary = wb.create_sheet("汇总")
                ws_summary.append(["指标", "金额"])
                ws_summary.append(["总人数", len(df_wage)])
                ws_summary.append(["总成本", round(total_cost, 2)])

                # 审计日志
                audit = wb.create_sheet("AuditTrail")
                audit.append(["时间戳", "操作", "详情"])
                audit.append([datetime.now().isoformat(), "GENERATED", f"公司:{company_name}, 城市:{city}, 规则来源:{rule.get('source_quote','')}"])

                # 水印
                ws_detail.insert_rows(1)
                ws_detail['A1'] = '⚠️ 待复核版 (仅供核对，不可正式交付)'
                ws_detail.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))

                # 保存
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                filename = f"{company_name}_{year}{month or ''}_待复核.xlsx"
                generated_files.append((filename, output.getvalue()))

                # 保存历史
                history_entry = {
                    "id": str(uuid.uuid4())[:8],
                    "company": company_name,
                    "city": city,
                    "report_type": report_type,
                    "year": year,
                    "month": month,
                    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "status": "待复核",
                    "source_quote": rule.get('source_quote', ''),
                    "total_people": len(df_wage),
                    "total_cost": round(total_cost, 2),
                    "reviewer": "",
                    "reviewed_at": "",
                    "review_comment": ""
                }
                save_export_history(history_entry)

                # 操作日志
                save_audit_log({
                    "id": str(uuid.uuid4())[:8],
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "action": "生成报表",
                    "detail": f"公司:{company_name}, 城市:{city}, 人数:{len(df_wage)}人, 总成本:{round(total_cost,2)}",
                    "report_id": history_entry['id']
                })

                summary_list.append({"公司": company_name, "城市": city, "人数": len(df_wage), "总成本": round(total_cost, 2)})

            if generated_files:
                # 显示汇总
                st.success(f"✅ 成功生成 {len(generated_files)} 份报表")
                st.dataframe(pd.DataFrame(summary_list))

                # 下载全部（打包成ZIP）
                if len(generated_files) > 1:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w') as zf:
                        for fname, data in generated_files:
                            zf.writestr(fname, data)
                    zip_buffer.seek(0)
                    st.download_button(
                        label="📦 下载所有报表 (ZIP)",
                        data=zip_buffer,
                        file_name=f"报表_{year}{month or ''}.zip",
                        mime="application/zip"
                    )
                else:
                    # 只有一个文件直接下载
                    fname, data = generated_files[0]
                    st.download_button(
                        label=f"📥 下载 {fname}",
                        data=BytesIO(data),
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

# ==================== 报表历史 ====================
with tab2:
    st.subheader("📋 报表生成历史")
    history = get_export_history()

    if not history:
        st.info("暂无历史记录")
    else:
        df_history = pd.DataFrame(history)
        # 显示状态颜色
        def status_color(status):
            if status == "待复核":
                return "🟡"
            elif status == "已通过":
                return "🟢"
            else:
                return "🔴"
        df_history["状态标识"] = df_history["status"].apply(status_color)

        st.dataframe(df_history[["状态标识", "company", "city", "report_type", "year", "month", "total_people", "total_cost", "generated_at", "status", "reviewer"]],
                    use_container_width=True)

        # 复核操作
        st.subheader("✅ 复核报表")
        pending_reports = [h for h in history if h['status'] == '待复核']
        if pending_reports:
            report_options = [f"{h['company']} - {h['city']} - {h['year']}{h['month'] or ''}" for h in pending_reports]
            selected_idx = st.selectbox("选择要复核的报表", range(len(report_options)), format_func=lambda x: report_options[x])
            selected_report = pending_reports[selected_idx]

            st.write(f"**公司**：{selected_report['company']}")
            st.write(f"**城市**：{selected_report['city']}")
            st.write(f"**总人数**：{selected_report['total_people']}")
            st.write(f"**总成本**：{selected_report['total_cost']}")
            st.write(f"**规则来源**：{selected_report['source_quote']}")

            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ 通过复核"):
                    selected_report['status'] = '已通过'
                    selected_report['reviewer'] = '系统管理员'
                    selected_report['reviewed_at'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    selected_report['review_comment'] = '数据无误，同意交付'
                    save_export_history(selected_report)
                    save_audit_log({
                        "id": str(uuid.uuid4())[:8],
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "action": "复核通过",
                        "detail": f"报表ID:{selected_report['id']}, 公司:{selected_report['company']}",
                        "report_id": selected_report['id']
                    })
                    st.success("✅ 已通过复核！")
                    st.rerun()
            with col2:
                if st.button("❌ 驳回"):
                    reason = st.text_input("驳回原因", key="reject_reason")
                    if st.button("确认驳回"):
                        selected_report['status'] = '已驳回'
                        selected_report['reviewer'] = '系统管理员'
                        selected_report['reviewed_at'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        selected_report['review_comment'] = reason or '数据有误，需重新生成'
                        save_export_history(selected_report)
                        save_audit_log({
                            "id": str(uuid.uuid4())[:8],
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "action": "复核驳回",
                            "detail": f"报表ID:{selected_report['id']}, 原因:{reason}",
                            "report_id": selected_report['id']
                        })
                        st.warning("❌ 已驳回")
                        st.rerun()
        else:
            st.info("所有报表已复核完成 ✅")

        # 操作日志
        with st.expander("📜 操作日志"):
            logs = get_audit_logs()
            if logs:
                st.dataframe(pd.DataFrame(logs))
            else:
                st.info("暂无操作日志")

# ==================== 管理数据 ====================
with tab3:
    st.subheader("📝 添加或修改数据")

    # 公司管理
    with st.expander("🏢 公司管理"):
        companies = load_from_db('companies')
        edited = st.data_editor(companies, num_rows="dynamic", key="company_editor_db", use_container_width=True)
        if st.button("保存公司数据 (持久化)"):
            save_to_db('companies', edited)
            st.success("已保存到数据库！")

    # 规则管理
    with st.expander("⚖️ 规则管理"):
        rules = load_from_db('rules')
        st.markdown("**比例字段**：0.16 表示 16%；基数上下限为整数。")
        edited_rules = st.data_editor(rules, num_rows="dynamic", key="rule_editor_db", use_container_width=True)
        if st.button("保存规则数据 (持久化)"):
            save_to_db('rules', edited_rules)
            st.success("已保存到数据库！")

    # 批量导入Excel（简化版）
    with st.expander("📤 批量导入 Excel"):
        st.markdown("上传包含「公司」和「规则」两个 Sheet 的 Excel")
        uploaded_file = st.file_uploader("选择 Excel 文件", type=["xlsx"], key="import_db")
        if uploaded_file is not None:
            try:
                xls = pd.ExcelFile(uploaded_file)
                if "公司" in xls.sheet_names and "规则" in xls.sheet_names:
                    df_companies = pd.read_excel(uploaded_file, sheet_name="公司")
                    df_rules = pd.read_excel(uploaded_file, sheet_name="规则")
                    # 添加ID
                    df_companies['id'] = [str(uuid.uuid4())[:8] for _ in range(len(df_companies))]
                    df_rules['id'] = [str(uuid.uuid4())[:8] for _ in range(len(df_rules))]
                    # 列名映射
                    df_companies = df_companies.rename(columns={"公司名称": "company_name"})
                    df_rules = df_rules.rename(columns={
                        "单位社保比例": "unit_social", "个人社保比例": "personal_social",
                        "单位公积金比例": "unit_fund", "个人公积金比例": "personal_fund",
                        "社保最低基数": "social_min", "社保最高基数": "social_max",
                        "公积金最低基数": "fund_min", "公积金最高基数": "fund_max",
                        "来源引用": "source_quote"
                    })
                    st.write("预览公司数据", df_companies.head())
                    st.write("预览规则数据", df_rules.head())
                    if st.button("确认导入"):
                        save_to_db('companies', df_companies.to_dict('records'))
                        save_to_db('rules', df_rules.to_dict('records'))
                        st.success("导入成功！")
                        st.rerun()
                else:
                    st.warning("需要包含「公司」和「规则」两个 Sheet")
            except Exception as e:
                st.error(f"解析失败：{e}")
