import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from io import BytesIO
import json
import re
import uuid
import os

# ---------- 初始化会话状态 ----------
if "data_initialized" not in st.session_state:
    st.session_state.companies = []
    st.session_state.templates = []
    st.session_state.rules = []
    st.session_state.sources = []
    st.session_state.export_history = []  # 报表历史记录
    st.session_state.audit_logs = []      # 操作日志
    st.session_state.data_initialized = True

st.set_page_config(page_title="智能社保报表匹配", layout="wide")
st.title("🧠 智能社保报表匹配系统")
st.markdown("**支持上传员工工资表 → 自动匹配城市规则 → 自动校验基数 → 生成合规报表 → 可复核可追溯**")

tab1, tab2, tab3 = st.tabs(["📊 生成报表", "📋 报表历史", "✏️ 管理数据"])

# ==================== 生成报表 ====================
with tab1:
    companies_df = pd.DataFrame(st.session_state.companies)
    if companies_df.empty:
        st.warning("请先在「管理数据」中添加公司信息，或上传您的Excel导入。")
        st.stop()

    col_left, col_right = st.columns(2)
    with col_left:
        province = st.selectbox("省份", sorted(companies_df["省份"].unique()))
        cities = sorted(companies_df[companies_df["省份"] == province]["城市"].unique())
        city = st.selectbox("城市", cities)
        districts = sorted(companies_df[(companies_df["省份"] == province) & (companies_df["城市"] == city)]["区县"].unique())
        district = st.selectbox("区县", districts)
        company_list = companies_df[(companies_df["省份"] == province) & (companies_df["城市"] == city) & (companies_df["区县"] == district)]["公司名称"].tolist()
        company = st.selectbox("公司", company_list)
    with col_right:
        st.write("**上传员工工资表**")
        st.markdown("Excel 格式需包含列：公司、城市、姓名、工资基数、社保基数、公积金基数（可选）")
        uploaded_wage = st.file_uploader("选择工资表 (支持 .xlsx)", type=["xlsx"], key="wage_upload")
        if uploaded_wage is not None:
            try:
                df_wage = pd.read_excel(uploaded_wage)
                st.success(f"成功读取 {len(df_wage)} 条员工记录")
                st.dataframe(df_wage.head(5))
            except Exception as e:
                st.error(f"读取失败：{e}")
                df_wage = None
        else:
            df_wage = None

    report_type = st.selectbox("报表类型", ["月度申报", "年度汇算"])
    year = st.selectbox("年份", [2025, 2024, 2023], index=0)
    month = None
    if report_type == "月度申报":
        month = st.selectbox("月份", list(range(1,13)), index=11)

    if st.button("生成正式报表"):
        if df_wage is None or df_wage.empty:
            st.error("请先上传员工工资表")
        else:
            rule_df = pd.DataFrame(st.session_state.rules)
            matched_rule = rule_df[(rule_df["城市"] == city) & (rule_df["报表类型"] == report_type)]
            if matched_rule.empty:
                st.error(f"未找到城市 {city} 的规则，请先在「管理数据」中添加规则。")
            else:
                rule = matched_rule.iloc[0]
                # 提取规则字段
                unit_social = rule.get("单位社保比例", 0.16)
                personal_social = rule.get("个人社保比例", 0.08)
                unit_fund = rule.get("单位公积金比例", 0.12)
                personal_fund = rule.get("个人公积金比例", 0.12)
                social_min = rule.get("社保最低基数", 0)
                social_max = rule.get("社保最高基数", float('inf'))
                fund_min = rule.get("公积金最低基数", 0)
                fund_max = rule.get("公积金最高基数", float('inf'))

                # 处理工资表
                required_cols = ["公司", "城市", "姓名", "工资基数"]
                for col in required_cols:
                    if col not in df_wage.columns:
                        st.error(f"工资表缺少列：{col}")
                        st.stop()

                if "社保基数" not in df_wage.columns:
                    df_wage["社保基数"] = df_wage["工资基数"]
                if "公积金基数" not in df_wage.columns:
                    df_wage["公积金基数"] = df_wage["工资基数"]

                # 基数校验
                df_wage["社保基数原始"] = df_wage["社保基数"]
                df_wage["社保基数调整"] = df_wage["社保基数"].apply(
                    lambda x: max(social_min, min(x, social_max))
                )
                df_wage["公积金基数原始"] = df_wage["公积金基数"]
                df_wage["公积金基数调整"] = df_wage["公积金基数"].apply(
                    lambda x: max(fund_min, min(x, fund_max))
                )

                # 记录调整
                adjustment_log = []
                for idx, row in df_wage.iterrows():
                    if row["社保基数调整"] != row["社保基数原始"]:
                        adjustment_log.append({"姓名": row["姓名"], "基数类型": "社保", "原始值": row["社保基数原始"], "调整后": row["社保基数调整"]})
                    if row["公积金基数调整"] != row["公积金基数原始"]:
                        adjustment_log.append({"姓名": row["姓名"], "基数类型": "公积金", "原始值": row["公积金基数原始"], "调整后": row["公积金基数调整"]})

                # 计算
                df_wage["单位社保"] = df_wage["社保基数调整"] * unit_social
                df_wage["个人社保"] = df_wage["社保基数调整"] * personal_social
                df_wage["单位公积金"] = df_wage["公积金基数调整"] * unit_fund
                df_wage["个人公积金"] = df_wage["公积金基数调整"] * personal_fund
                df_wage["单位合计"] = df_wage["单位社保"] + df_wage["单位公积金"]
                df_wage["个人合计"] = df_wage["个人社保"] + df_wage["个人公积金"]
                df_wage["总成本"] = df_wage["单位合计"] + df_wage["个人合计"]

                # 汇总
                summary = {
                    "总人数": len(df_wage),
                    "社保总基数": df_wage["社保基数调整"].sum(),
                    "公积金总基数": df_wage["公积金基数调整"].sum(),
                    "单位社保总额": df_wage["单位社保"].sum(),
                    "个人社保总额": df_wage["个人社保"].sum(),
                    "单位公积金总额": df_wage["单位公积金"].sum(),
                    "个人公积金总额": df_wage["个人公积金"].sum(),
                    "总成本": df_wage["总成本"].sum()
                }

                # 生成 Excel
                wb = Workbook()
                ws_detail = wb.active
                ws_detail.title = "员工明细"
                headers = ["公司", "城市", "姓名", "工资基数", "原始社保基数", "最终社保基数", "原始公积金基数", "最终公积金基数",
                           "单位社保", "个人社保", "单位公积金", "个人公积金", "单位合计", "个人合计", "总成本"]
                ws_detail.append(headers)
                for _, row in df_wage.iterrows():
                    ws_detail.append([
                        row["公司"], row["城市"], row["姓名"], row["工资基数"],
                        row["社保基数原始"], row["社保基数调整"],
                        row["公积金基数原始"], row["公积金基数调整"],
                        round(row["单位社保"], 2), round(row["个人社保"], 2),
                        round(row["单位公积金"], 2), round(row["个人公积金"], 2),
                        round(row["单位合计"], 2), round(row["个人合计"], 2),
                        round(row["总成本"], 2)
                    ])

                ws_summary = wb.create_sheet("汇总")
                ws_summary.append(["指标", "金额"])
                for k, v in summary.items():
                    ws_summary.append([k, round(v, 2)])

                if adjustment_log:
                    ws_adjust = wb.create_sheet("基数调整日志")
                    ws_adjust.append(["姓名", "基数类型", "原始值", "调整后"])
                    for log in adjustment_log:
                        ws_adjust.append([log["姓名"], log["基数类型"], log["原始值"], log["调整后"]])

                # 审计日志
                audit = wb.create_sheet("AuditTrail")
                audit.append(["时间戳", "操作", "详情"])
                source_info = f"规则来源:{rule.get('来源引用','')}, 模板:{rule.get('模板名称','')}"
                audit.append([datetime.now().isoformat(), "GENERATED", f"公司:{company}, 城市:{city}, 年份:{year}, 月份:{month}, {source_info}"])

                # 水印
                ws_detail.insert_rows(1)
                ws_detail['A1'] = '⚠️ 待复核版 (仅供核对，不可正式交付)'
                ws_detail.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))

                output = BytesIO()
                wb.save(output)
                output.seek(0)

                # 保存历史记录
                export_id = str(uuid.uuid4())[:8]
                history_entry = {
                    "id": export_id,
                    "公司": company,
                    "城市": city,
                    "报表类型": report_type,
                    "年份": year,
                    "月份": month,
                    "生成时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "状态": "待复核",
                    "规则来源": rule.get("来源引用", ""),
                    "模板名称": rule.get("模板名称", ""),
                    "总人数": len(df_wage),
                    "总成本": round(summary["总成本"], 2),
                    "复核人": "",
                    "复核时间": "",
                    "复核备注": ""
                }
                st.session_state.export_history.append(history_entry)

                # 记录操作日志
                st.session_state.audit_logs.append({
                    "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "操作": "生成报表",
                    "详情": f"公司:{company}, 城市:{city}, 人员:{len(df_wage)}人, 总成本:{round(summary['总成本'],2)}",
                    "报表ID": export_id
                })

                st.success("✅ 报表生成成功！")
                if adjustment_log:
                    st.info(f"ℹ️ 共调整了 {len(adjustment_log)} 个基数")
                st.download_button(
                    label="📥 下载待复核版",
                    data=output,
                    file_name=f"{company}_{year}{month or ''}_待复核.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.subheader("📊 数据汇总")
                st.dataframe(pd.DataFrame([summary]))

# ==================== 报表历史（含复核功能） ====================
with tab2:
    st.subheader("📋 报表生成历史")

    if not st.session_state.export_history:
        st.info("暂无历史记录")
    else:
        df_history = pd.DataFrame(st.session_state.export_history)
        # 显示状态颜色
        def status_color(status):
            if status == "待复核":
                return "🟡"
            elif status == "已通过":
                return "🟢"
            else:
                return "🔴"
        df_history["状态标识"] = df_history["状态"].apply(status_color)

        # 显示表格
        st.dataframe(df_history[["状态标识", "公司", "城市", "报表类型", "年份", "月份", "总人数", "总成本", "生成时间", "状态", "复核人"]])

        # 复核操作
        st.subheader("✅ 复核报表")
        pending_reports = [h for h in st.session_state.export_history if h["状态"] == "待复核"]
        if pending_reports:
            report_options = [f"{h['公司']} - {h['城市']} - {h['年份']}{h['月份'] or ''}" for h in pending_reports]
            selected_idx = st.selectbox("选择要复核的报表", range(len(report_options)), format_func=lambda x: report_options[x])
            selected_report = pending_reports[selected_idx]

            st.write(f"**公司**：{selected_report['公司']}")
            st.write(f"**城市**：{selected_report['城市']}")
            st.write(f"**总人数**：{selected_report['总人数']}")
            st.write(f"**总成本**：{selected_report['总成本']}")
            st.write(f"**规则来源**：{selected_report['规则来源']}")

            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ 通过复核"):
                    # 更新状态
                    for h in st.session_state.export_history:
                        if h["id"] == selected_report["id"]:
                            h["状态"] = "已通过"
                            h["复核人"] = "系统管理员"
                            h["复核时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            h["复核备注"] = "数据无误，同意交付"
                    st.session_state.audit_logs.append({
                        "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "操作": "复核通过",
                        "详情": f"报表ID:{selected_report['id']}, 公司:{selected_report['公司']}",
                        "报表ID": selected_report["id"]
                    })
                    st.success("✅ 已通过复核！")
                    st.rerun()
            with col2:
                if st.button("❌ 驳回"):
                    reason = st.text_input("驳回原因")
                    if st.button("确认驳回"):
                        for h in st.session_state.export_history:
                            if h["id"] == selected_report["id"]:
                                h["状态"] = "已驳回"
                                h["复核人"] = "系统管理员"
                                h["复核时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                h["复核备注"] = reason or "数据有误，需重新生成"
                        st.session_state.audit_logs.append({
                            "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "操作": "复核驳回",
                            "详情": f"报表ID:{selected_report['id']}, 原因:{reason}",
                            "报表ID": selected_report["id"]
                        })
                        st.warning("❌ 已驳回")
                        st.rerun()
        else:
            st.info("所有报表已复核完成 ✅")

        # 操作日志
        with st.expander("📜 操作日志"):
            if st.session_state.audit_logs:
                st.dataframe(pd.DataFrame(st.session_state.audit_logs))
            else:
                st.info("暂无操作日志")

# ==================== 管理数据 ====================
with tab3:
    st.subheader("📝 添加或修改数据")
    with st.expander("📤 智能导入（自动识别您的报表数据）"):
        st.markdown("上传包含「基础数据表」「城市规则表」的 Excel，自动生成公司、模板、规则。")
        uploaded_file = st.file_uploader("选择 Excel 文件", type=["xlsx"], key="import_excel")
        if uploaded_file is not None:
            try:
                xls = pd.ExcelFile(uploaded_file)
                if "基础数据表" in xls.sheet_names and "城市规则表" in xls.sheet_names:
                    df_data = pd.read_excel(uploaded_file, sheet_name="基础数据表")
                    df_city_rules = pd.read_excel(uploaded_file, sheet_name="城市规则表")
                    companies = df_data[["公司", "城市"]].drop_duplicates().rename(columns={"公司": "公司名称"})
                    province_map = {"上海": "上海", "深圳": "广东", "北京": "北京", "成都": "四川", "武汉": "湖北", "杭州": "浙江", "南京": "江苏"}
                    companies["省份"] = companies["城市"].map(province_map).fillna("未知")
                    companies["区县"] = "市区"
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
                    rules = []
                    for _, row in df_city_rules.iterrows():
                        city = row["城市"]
                        province = province_map.get(city, "未知")
                        rules.append({
                            "省份": province,
                            "城市": city,
                            "报表类型": "月度申报",
                            "单位社保比例": row["单位社保比例"],
                            "个人社保比例": row["个人社保比例"],
                            "单位公积金比例": row["单位公积金比例"],
                            "个人公积金比例": row["个人公积金比例"],
                            "社保最低基数": row["社保最低基数"],
                            "社保最高基数": row["社保最高基数"],
                            "公积金最低基数": row["公积金最低基数"],
                            "公积金最高基数": row["公积金最高基数"],
                            "模板名称": f"{city}社保申报表（自动生成）",
                            "来源引用": "自动从城市规则表生成"
                        })
                    st.subheader("预览导入数据")
                    st.write("公司", companies)
                    st.write("模板", pd.DataFrame(templates))
                    st.write("规则", pd.DataFrame(rules))
                    if st.button("确认导入"):
                        st.session_state.companies = companies.to_dict(orient="records")
                        st.session_state.templates = templates
                        st.session_state.rules = rules
                        st.session_state.sources = [{"省份": "全国", "城市": "通用", "机构": "系统", "来源": "自动生成"}]
                        st.success("导入成功！")
                        st.rerun()
                else:
                    st.warning("未检测到「基础数据表」和「城市规则表」")
            except Exception as e:
                st.error(f"解析失败：{e}")

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
        st.markdown("**规则字段说明**：比例请填写小数（如0.16表示16%），基数为整数。")
        edited_rules = st.data_editor(st.session_state.rules, num_rows="dynamic", key="rule_editor", use_container_width=True)
        if st.button("保存规则数据"):
            st.session_state.rules = edited_rules
            st.success("已更新！")
    with st.expander("🔗 官方来源"):
        edited_sources = st.data_editor(st.session_state.sources, num_rows="dynamic", key="source_editor", use_container_width=True)
        if st.button("保存来源数据"):
            st.session_state.sources = edited_sources
            st.success("已更新！")
