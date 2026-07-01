import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from io import BytesIO
import json
import os

# ---------- 初始化会话状态（保存用户自定义数据） ----------
if "data_initialized" not in st.session_state:
    # 默认示例数据
    st.session_state.sources = [
        {"省份": "上海", "城市": "上海市", "机构": "上海市税务局", "来源": "https://shanghai.chinatax.gov.cn/"},
        {"省份": "广东", "城市": "广州市", "机构": "广东省税务局", "来源": "https://guangdong.chinatax.gov.cn/"},
    ]
    st.session_state.templates = [
        {"省份": "上海", "城市": "上海市", "模板名称": "上海市社保申报表（月度）", "报表类型": "月度申报", "版本": "1.0", "必填字段": "单位名称,基数"},
        {"省份": "广东", "城市": "广州市", "模板名称": "广东省社保申报表", "报表类型": "月度申报", "版本": "1.0", "必填字段": "单位名称,基数"},
    ]
    st.session_state.rules = [
        {"省份": "上海", "城市": "上海市", "报表类型": "月度申报", "规则": "养老16%/8%", "来源引用": "沪人社规〔2024〕22号"},
        {"省份": "广东", "城市": "广州市", "报表类型": "月度申报", "规则": "养老14%/8%", "来源引用": "粤人社规〔2024〕8号"},
    ]
    st.session_state.companies = [
        {"公司名称": "上海科技公司", "省份": "上海", "城市": "上海市", "区县": "浦东新区"},
        {"公司名称": "广州科技公司", "省份": "广东", "城市": "广州市", "区县": "天河区"},
    ]
    st.session_state.data_initialized = True

# ---------- 页面导航 ----------
st.set_page_config(page_title="报表匹配工具", layout="wide")
st.title("📋 社保报表匹配工具")

tab1, tab2 = st.tabs(["📊 生成报表", "✏️ 管理数据"])

# ========== 标签页1：生成报表 ==========
with tab1:
    st.markdown("---")
    # 从 session_state 读取数据
    companies_df = pd.DataFrame(st.session_state.companies)
    templates_df = pd.DataFrame(st.session_state.templates)
    rules_df = pd.DataFrame(st.session_state.rules)
    sources_df = pd.DataFrame(st.session_state.sources)

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
    year = st.selectbox("年份", [2024, 2025], index=0)
    month = None
    if report_type == "月度申报":
        month = st.selectbox("月份", list(range(1,13)), index=11)

    if st.button("匹配官方模板"):
        # 匹配模板（匹配城市+报表类型）
        template = templates_df[(templates_df["城市"] == city) & (templates_df["报表类型"] == report_type)]
        rule = rules_df[(rules_df["城市"] == city) & (rules_df["报表类型"] == report_type)]
        if not template.empty and not rule.empty:
            st.success("✅ 匹配成功！")
            tpl = template.iloc[0]
            rul = rule.iloc[0]
            st.write(f"**模板**：{tpl['模板名称']} (v{tpl['版本']})")
            st.write(f"**必填字段**：{tpl['必填字段']}")
            st.write(f"**规则**：{rul['规则']}")
            st.write(f"**来源**：[{rul['来源引用']}](https://shanghai.chinatax.gov.cn/)")
            
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
            st.error("未匹配到模板，请检查数据或添加新模板。")

# ========== 标签页2：管理数据 ==========
with tab2:
    st.subheader("📝 添加或修改数据")
    st.info("在这里添加新的城市、公司、模板、规则。修改后会自动保存（仅在当前会话有效）。如需永久保存，请导出代码并重新部署。")

    # 公司管理
    with st.expander("🏢 公司管理"):
        st.data_editor(
            st.session_state.companies,
            num_rows="dynamic",
            key="company_editor",
            use_container_width=True
        )
        if st.button("保存公司数据"):
            st.session_state.companies = st.session_state.company_editor
            st.success("已更新！")

    # 模板管理
    with st.expander("📄 模板管理"):
        st.data_editor(
            st.session_state.templates,
            num_rows="dynamic",
            key="template_editor",
            use_container_width=True
        )
        if st.button("保存模板数据"):
            st.session_state.templates = st.session_state.template_editor
            st.success("已更新！")

    # 规则管理
    with st.expander("⚖️ 规则管理"):
        st.data_editor(
            st.session_state.rules,
            num_rows="dynamic",
            key="rule_editor",
            use_container_width=True
        )
        if st.button("保存规则数据"):
            st.session_state.rules = st.session_state.rule_editor
            st.success("已更新！")

    # 来源管理（仅展示）
    with st.expander("🔗 官方来源（只读）"):
        st.dataframe(pd.DataFrame(st.session_state.sources))

    # 导出当前所有数据为代码（方便永久保存）
    st.subheader("💾 导出当前数据为代码")
    st.markdown("复制下面的代码，替换掉原来的 `app.py` 中 `st.session_state` 初始化部分，即可永久保存您的数据。")
    code_snippet = f"""
# ---------- 您的自定义数据 ----------
st.session_state.sources = {json.dumps(st.session_state.sources, ensure_ascii=False, indent=2)}
st.session_state.templates = {json.dumps(st.session_state.templates, ensure_ascii=False, indent=2)}
st.session_state.rules = {json.dumps(st.session_state.rules, ensure_ascii=False, indent=2)}
st.session_state.companies = {json.dumps(st.session_state.companies, ensure_ascii=False, indent=2)}
"""
    st.code(code_snippet, language="python")
