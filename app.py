import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from io import BytesIO
import json

# ---------- 初始化会话状态 ----------
if "data_initialized" not in st.session_state:
    st.session_state.sources = [
        {"省份": "上海", "城市": "上海市", "机构": "上海市税务局", "来源": "https://shanghai.chinatax.gov.cn/"},
        {"省份": "广东", "城市": "广州市", "机构": "广东省税务局", "来源": "https://guangdong.chinatax.gov.cn/"},
        {"省份": "浙江", "城市": "丽水市", "机构": "浙江省税务局", "来源": "https://zhejiang.chinatax.gov.cn/"},
    ]
    st.session_state.templates = [
        {"省份": "上海", "城市": "上海市", "模板名称": "上海市社保申报表（月度）", "报表类型": "月度申报", "版本": "1.0", "必填字段": "单位名称,统一信用代码,险种,基数,单位金额,个人金额"},
        {"省份": "广东", "城市": "广州市", "模板名称": "广东省社保申报表", "报表类型": "月度申报", "版本": "1.0", "必填字段": "单位名称,统一信用代码,险种,基数,单位金额,个人金额"},
    ]
    st.session_state.rules = [
        {"省份": "上海", "城市": "上海市", "报表类型": "月度申报", "规则": "养老16%/8%，医疗9.5%/2%，失业0.5%/0.5%，工伤0.2%/0，生育1%/0，公积金7%/7%", "来源引用": "沪人社规〔2024〕22号"},
        {"省份": "广东", "城市": "广州市", "报表类型": "月度申报", "规则": "养老14%/8%，医疗5.5%/2%，失业0.32%/0.2%，工伤0.16%/0，生育0.85%/0，公积金5%-12%", "来源引用": "粤人社规〔2024〕8号"},
    ]
    st.session_state.companies = [
        {"公司名称": "上海科技公司", "省份": "上海", "城市": "上海市", "区县": "浦东新区"},
        {"公司名称": "广州科技公司", "省份": "广东", "城市": "广州市", "区县": "天河区"},
    ]
    st.session_state.data_initialized = True

# ---------- 页面 ----------
st.set_page_config(page_title="智能社保报表匹配", layout="wide")
st.title("🧠 智能社保报表匹配系统")
st.markdown("**自动选择地区 → 智能匹配官方模板 → 一键生成带审计日志的Excel**")

tab1, tab2 = st.tabs(["📊 生成报表", "✏️ 管理数据"])

# ==================== 标签页1：生成报表 ====================
with tab1:
    # 从 session_state 读取数据
    companies_df = pd.DataFrame(st.session_state.companies)
    templates_df = pd.DataFrame(st.session_state.templates)
    rules_df = pd.DataFrame(st.session_state.rules)

    if companies_df.empty:
        st.warning("请先在「管理数据」标签页中添加公司信息。")
        st.stop()

    province = st.selectbox("省份", sorted(companies_df["省份"].unique()))
    cities = sorted(companies_df[companies_df["省份"] == province]["城市"].unique())
    city = st.selectbox("城市", cities)
    districts = sorted(companies_df[(companies_df["省份"] == province) & (companies_df["城市"] == city)]["区县"].unique())
    district = st.selectbox("区县", districts)
    companies = companies_df[(companies_df["省份"] == province) & (companies_df["城市"] == city) & (companies_df["区县"] == district)]["公司名称"].tolist()
    company = st.selectbox("公司", companies)
    report_type = st.selectbox("报表类型", ["月度申报", "年度汇算"])
    year = st.selectbox("年份", [2025, 2024, 2023], index=0)
    month = None
    if report_type == "月度申报":
        month = st.selectbox("月份", list(range(1,13)), index=11)

    if st.button("智能匹配模板"):
        template = templates_df[(templates_df["城市"] == city) & (templates_df["报表类型"] == report_type)]
        rule = rules_df[(rules_df["城市"] == city) & (rules_df["报表类型"] == report_type)]
        if not template.empty and not rule.empty:
            st.success("✅ 匹配成功！")
            tpl = template.iloc[0]
            rul = rule.iloc[0]
            st.write(f"**模板**：{tpl['模板名称']} (v{tpl['版本']})")
            st.write(f"**必填字段**：{tpl['必填字段']}")
            st.write(f"**规则**：{rul['规则']}")
            st.write(f"**来源**：{rul['来源引用']}")
            
            if st.button("生成假数据Excel（待复核版）"):
                wb = Workbook()
                ws = wb.active
                ws.append(['公司', '险种', '基数', '单位金额', '个人金额'])
                ws.append([company, '养老保险', 8000, 1280, 640])
                ws.insert_rows(1)
                ws['A1'] = '⚠️ 待复核版'
                ws.merge_cells('A1:E1')
                audit = wb.create_sheet('AuditTrail')
                audit.append(['时间', '操作'])
                audit.append([datetime.now().isoformat(), '生成'])
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                st.download_button("下载Excel", data=output, file_name=f"{company}_待复核.xlsx")
        else:
            st.error("未匹配到模板，请先在「管理数据」中添加。")

# ==================== 标签页2：管理数据 ====================
with tab2:
    st.subheader("📝 添加或修改数据")
    st.info("在这里添加新的城市、公司、模板、规则。修改后会自动保存。")

    # 公司管理
    with st.expander("🏢 公司管理"):
        edited_companies = st.data_editor(
            st.session_state.companies,
            num_rows="dynamic",
            key="company_editor",
            use_container_width=True
        )
        if st.button("保存公司数据"):
            st.session_state.companies = edited_companies
            st.success("已更新！")

    # 模板管理
    with st.expander("📄 模板管理"):
        edited_templates = st.data_editor(
            st.session_state.templates,
            num_rows="dynamic",
            key="template_editor",
            use_container_width=True
        )
        if st.button("保存模板数据"):
            st.session_state.templates = edited_templates
            st.success("已更新！")

    # 规则管理
    with st.expander("⚖️ 规则管理"):
        edited_rules = st.data_editor(
            st.session_state.rules,
            num_rows="dynamic",
            key="rule_editor",
            use_container_width=True
        )
        if st.button("保存规则数据"):
            st.session_state.rules = edited_rules
            st.success("已更新！")

    # 来源管理
    with st.expander("🔗 官方来源"):
        edited_sources = st.data_editor(
            st.session_state.sources,
            num_rows="dynamic",
            key="source_editor",
            use_container_width=True
        )
        if st.button("保存来源数据"):
            st.session_state.sources = edited_sources
            st.success("已更新！")

    # ---------- 导入Excel功能 ----------
    with st.expander("📤 批量导入 Excel（自动分步解析）"):
        st.markdown("上传包含 `公司`、`模板`、`规则`、`来源` 四个 Sheet 的 Excel 文件，系统自动解析并填充。")
        uploaded_file = st.file_uploader("选择 Excel 文件", type=["xlsx"])
        if uploaded_file is not None:
            try:
                xls = pd.ExcelFile(uploaded_file)
                required_sheets = ["公司", "模板", "规则"]
                missing = [s for s in required_sheets if s not in xls.sheet_names]
                if missing:
                    st.error(f"缺少 Sheet：{', '.join(missing)}，请确保 Excel 包含这些 Sheet。")
                else:
                    df_companies = pd.read_excel(uploaded_file, sheet_name="公司")
                    df_templates = pd.read_excel(uploaded_file, sheet_name="模板")
                    df_rules = pd.read_excel(uploaded_file, sheet_name="规则")
                    df_sources = pd.read_excel(uploaded_file, sheet_name="来源") if "来源" in xls.sheet_names else pd.DataFrame()

                    st.subheader("预览数据")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write("**公司**")
                        st.dataframe(df_companies.head(5))
                    with col2:
                        st.write("**模板**")
                        st.dataframe(df_templates.head(5))
                    col3, col4 = st.columns(2)
                    with col3:
                        st.write("**规则**")
                        st.dataframe(df_rules.head(5))
                    with col4:
                        st.write("**来源**")
                        st.dataframe(df_sources.head(5))

                    if st.button("✅ 确认导入（将覆盖当前数据）"):
                        st.session_state.companies = df_companies.to_dict(orient="records")
                        st.session_state.templates = df_templates.to_dict(orient="records")
                        st.session_state.rules = df_rules.to_dict(orient="records")
                        if not df_sources.empty:
                            st.session_state.sources = df_sources.to_dict(orient="records")
                        st.success(f"成功导入 {len(df_companies)} 家公司，{len(df_templates)} 个模板，{len(df_rules)} 条规则。")
                        st.rerun()
            except Exception as e:
                st.error(f"解析失败：{e}")

    # 导出代码
    st.subheader("💾 导出当前数据为代码")
    code_snippet = f"""
# ---------- 您的自定义数据 ----------
st.session_state.sources = {json.dumps(st.session_state.sources, ensure_ascii=False, indent=2)}
st.session_state.templates = {json.dumps(st.session_state.templates, ensure_ascii=False, indent=2)}
st.session_state.rules = {json.dumps(st.session_state.rules, ensure_ascii=False, indent=2)}
st.session_state.companies = {json.dumps(st.session_state.companies, ensure_ascii=False, indent=2)}
"""
    st.code(code_snippet, language="python")
