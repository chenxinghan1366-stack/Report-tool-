import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from io import BytesIO
import json

# ---------- 初始化会话状态 ----------
if "data_initialized" not in st.session_state:
    st.session_state.companies = []
    st.session_state.templates = []
    st.session_state.rules = []
    st.session_state.sources = []
    st.session_state.data_initialized = True

st.set_page_config(page_title="智能社保报表匹配", layout="wide")
st.title("🧠 智能社保报表匹配系统")
st.markdown("**自动选择地区 → 智能匹配官方模板 → 一键生成带审计日志的Excel**")

tab1, tab2 = st.tabs(["📊 生成报表", "✏️ 管理数据"])

# ==================== 生成报表 ====================
with tab1:
    companies_df = pd.DataFrame(st.session_state.companies)
    templates_df = pd.DataFrame(st.session_state.templates)
    rules_df = pd.DataFrame(st.session_state.rules)

    if companies_df.empty:
        st.warning("请先在「管理数据」中添加公司信息，或上传您的Excel导入。")
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

# ==================== 管理数据 ====================
with tab2:
    st.subheader("📝 添加或修改数据")

    # ---------- 智能导入（新增） ----------
    with st.expander("📤 智能导入（自动识别您的报表数据）"):
        st.markdown("上传您的 Excel 文件，系统会自动识别公司、城市、规则，并生成对应的模板。")
        uploaded_file = st.file_uploader("选择 Excel 文件（支持 .xlsx）", type=["xlsx"])
        if uploaded_file is not None:
            try:
                xls = pd.ExcelFile(uploaded_file)
                # 检测是否存在“基础数据表”和“城市规则表”
                if "基础数据表" in xls.sheet_names and "城市规则表" in xls.sheet_names:
                    df_data = pd.read_excel(uploaded_file, sheet_name="基础数据表")
                    df_city_rules = pd.read_excel(uploaded_file, sheet_name="城市规则表")
                    df_company_config = pd.read_excel(uploaded_file, sheet_name="公司配置表") if "公司配置表" in xls.sheet_names else pd.DataFrame()
                    
                    # 提取公司信息
                    companies = df_data[["公司", "城市"]].drop_duplicates().rename(columns={"公司": "公司名称"})
                    # 添加省份（这里需要映射，简单处理：用城市名推断省份，或让用户补充）
                    # 为了演示，我们假设每个城市对应一个省份，这里用城市拼音首字母或手动映射，但为了通用，我们让用户补充
                    # 更智能：根据城市规则表中的城市，生成省份（如果规则表中有省份列，则用；否则手动补）
                    # 由于规则表没有省份，我们建立一个映射字典
                    province_map = {
                        "上海": "上海", "深圳": "广东", "北京": "北京", "成都": "四川", "武汉": "湖北"
                    }
                    companies["省份"] = companies["城市"].map(province_map).fillna("未知")
                    companies["区县"] = "市区"  # 默认区县
                    
                    # 生成模板（每个城市一个模板）
                    templates = []
                    for city in companies["城市"].unique():
                        templates.append({
                            "省份": companies[companies["城市"]==city]["省份"].iloc[0],
                            "城市": city,
                            "模板名称": f"{city}社保申报表（自动生成）",
                            "报表类型": "月度申报",
                            "版本": "1.0",
                            "必填字段": "公司名称,统一信用代码,险种,基数,单位金额,个人金额,合计"
                        })
                    
                    # 生成规则（从城市规则表提取）
                    rules = []
                    for _, row in df_city_rules.iterrows():
                        city = row["城市"]
                        province = province_map.get(city, "未知")
                        rule_text = f"社保比例：单位{row['单位社保比例']:.0%}，个人{row['个人社保比例']:.0%}；公积金：单位{row['单位公积金比例']:.0%}，个人{row['个人公积金比例']:.0%}；基数范围：{row['社保最低基数']}-{row['社保最高基数']}"
                        rules.append({
                            "省份": province,
                            "城市": city,
                            "报表类型": "月度申报",
                            "规则": rule_text,
                            "来源引用": "自动从城市规则表生成"
                        })
                    
                    # 预览
                    st.subheader("预览将导入的数据")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write("**公司**")
                        st.dataframe(companies)
                    with col2:
                        st.write("**模板**")
                        st.dataframe(pd.DataFrame(templates))
                    st.write("**规则**")
                    st.dataframe(pd.DataFrame(rules))
                    
                    if st.button("✅ 确认导入（将覆盖当前数据）"):
                        st.session_state.companies = companies.to_dict(orient="records")
                        st.session_state.templates = templates
                        st.session_state.rules = rules
                        # 添加来源
                        st.session_state.sources = [{"省份": "全国", "城市": "通用", "机构": "系统", "来源": "由您的Excel自动生成"}]
                        st.success(f"成功导入 {len(companies)} 家公司，{len(templates)} 个模板，{len(rules)} 条规则。")
                        st.rerun()
                else:
                    st.warning("未检测到「基础数据表」和「城市规则表」，请确保Excel包含这两个Sheet。")
            except Exception as e:
                st.error(f"解析失败：{e}")

    # 原有的手动管理功能保留
    with st.expander("🏢 公司管理"):
        edited_companies = st.data_editor(st.session_state.companies, num_rows="dynamic", key="company_editor", use_container_width=True)
        if st.button("保存公司数据"):
            st.session_state.companies = edited_companies
            st.success("已更新！")
    with st.expander("📄 模板管理"):
        edited_templates = st.data_editor(st.session_state.templates, num_rows="dynamic", key="template_editor", use_container_width=True)
        if st.button("保存模板数据"):
            st.session_state.templates = edited_templates
            st.success("已更新！")
    with st.expander("⚖️ 规则管理"):
        edited_rules = st.data_editor(st.session_state.rules, num_rows="dynamic", key="rule_editor", use_container_width=True)
        if st.button("保存规则数据"):
            st.session_state.rules = edited_rules
            st.success("已更新！")
    with st.expander("🔗 官方来源"):
        edited_sources = st.data_editor(st.session_state.sources, num_rows="dynamic", key="source_editor", use_container_width=True)
        if st.button("保存来源数据"):
            st.session_state.sources = edited_sources
            st.success("已更新！")
