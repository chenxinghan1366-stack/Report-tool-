import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO
import uuid
import re

# ========== 初始化会话状态 ==========
if 'companies' not in st.session_state:
    # 初始为空，等待用户上传
    st.session_state.companies = []

if 'export_history' not in st.session_state:
    st.session_state.export_history = []

if 'cities_data' not in st.session_state:
    st.session_state.cities_data = None

# ========== 官方模板库（可扩展） ==========
TEMPLATES = [
    {
        'id': 't001',
        'province': '上海',
        'city': '上海市',
        'district': '浦东新区',
        'report_type': '增值税',
        'template_name': '上海市增值税纳税申报表（一般纳税人）',
        'template_version': 'v2024.1',
        'source_url': 'https://shanghai.chinatax.gov.cn/bsfw/bszn/2024/zzs.xlsx',
        'source_authority': '国家税务总局上海市税务局',
        'publish_date': '2024-01-15',
        'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额'
    },
    {
        'id': 't002',
        'province': '上海',
        'city': '上海市',
        'district': '浦东新区',
        'report_type': '社保',
        'template_name': '上海市社会保险费申报表（月度）',
        'template_version': 'v2024.1',
        'source_url': 'https://rsj.sh.gov.cn/sbjb/2024/sb.xlsx',
        'source_authority': '上海市人力资源和社会保障局',
        'publish_date': '2024-01-10',
        'required_fields': '单位名称,社保登记号,基数,单位金额,个人金额'
    },
    # 添加更多模板...
]

# ========== 页面配置 ==========
st.set_page_config(page_title="官方模板匹配器", layout="wide")
st.title("📋 官方模板匹配器（含自动识别与统计口径）")
st.markdown("**上传Excel → 自动提取城市/公司 → 选择模板和统计口径 → 生成待复核版Excel**")

# ===== 侧边栏：导入数据 =====
with st.sidebar:
    st.header("📤 上传数据Excel")
    st.markdown("上传包含公司/城市信息的Excel，系统自动提取所有地区")
    uploaded_file = st.file_uploader("选择Excel文件 (.xlsx)", type=["xlsx"])

    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            # 选择要读取的Sheet（通常选择汇总或明细表）
            sheet_names = xls.sheet_names
            # 优先选择包含“汇总”或“明细”的Sheet
            target_sheet = None
            for s in sheet_names:
                if '汇总' in s or '明细' in s or '月度' in s:
                    target_sheet = s
                    break
            if not target_sheet:
                target_sheet = sheet_names[0]
            
            df = pd.read_excel(uploaded_file, sheet_name=target_sheet)
            st.session_state.cities_data = df
            st.success(f"成功读取Sheet「{target_sheet}」，共 {len(df)} 行")
            st.dataframe(df.head(3))
            
            # 智能提取公司信息
            # 寻找可能的列名
            col_map = {}
            for col in df.columns:
                col_lower = str(col).lower()
                if '公司' in col_lower or '企业' in col_lower or '分公司' in col_lower:
                    col_map['company'] = col
                elif '城市' in col_lower or '市' in col_lower:
                    col_map['city'] = col
                elif '省份' in col_lower or '省' in col_lower:
                    col_map['province'] = col
                elif '区县' in col_lower or '区' in col_lower:
                    col_map['district'] = col
                elif '税号' in col_lower or '信用代码' in col_lower:
                    col_map['tax_id'] = col
            
            if 'company' in col_map and 'city' in col_map:
                # 提取唯一公司-城市-省份-区县组合
                if 'province' in col_map:
                    comp_df = df[[col_map['company'], col_map['city'], col_map['province'], col_map.get('district', '')]].drop_duplicates()
                    comp_df.columns = ['company_name', 'city', 'province', 'district'] if 'district' in col_map else ['company_name', 'city', 'province', '']
                else:
                    comp_df = df[[col_map['company'], col_map['city']]].drop_duplicates()
                    comp_df['province'] = comp_df['city']  # 默认省份同城市
                    comp_df['district'] = ''
                
                # 生成id
                companies = []
                for _, row in comp_df.iterrows():
                    companies.append({
                        'id': str(uuid.uuid4())[:8],
                        'company_name': str(row['company_name']),
                        'province': str(row['province']),
                        'city': str(row['city']),
                        'district': str(row['district']) if 'district' in comp_df.columns else '',
                        'tax_id': ''
                    })
                st.session_state.companies = companies
                st.success(f"自动提取了 {len(companies)} 家公司")
            else:
                st.warning("未自动识别到公司列和城市列，请确认数据格式")
        except Exception as e:
            st.error(f"处理文件出错：{str(e)}")

# ===== 主体：选择公司 =====
companies = st.session_state.companies
if not companies:
    st.info("请先在侧边栏上传包含公司的Excel文件，或使用下方的示例公司")
    # 可选添加示例公司（可以注释掉）
    # companies = [{'id':'c001','company_name':'上海科技','province':'上海','city':'上海市','district':'浦东新区','tax_id':''}]

if companies:
    # 动态生成省份、城市、区县列表
    provinces = sorted(set(c['province'] for c in companies if c['province']))
    col1, col2, col3 = st.columns(3)
    with col1:
        province = st.selectbox("省份", [""] + provinces)
        cities = sorted(set(c['city'] for c in companies if c['province'] == province)) if province else []
        city = st.selectbox("城市", [""] + cities)
    with col2:
        districts = sorted(set(c['district'] for c in companies if c['province'] == province and c['city'] == city)) if province and city else []
        district = st.selectbox("区县", [""] + districts)
        company_list = [c for c in companies if c['province'] == province and c['city'] == city and c['district'] == district]
        company_names = [c['company_name'] for c in company_list]
        selected_company_name = st.selectbox("公司", [""] + company_names)
    with col3:
        report_type = st.selectbox("报表类型", ["", "增值税", "社保", "公积金", "企业所得税", "个人所得税"])
        # 统计口径
        period_type = st.selectbox("统计口径", ["月度（12月单月）", "年度（1-12月累计）"])

    # 获取选中公司
    selected_company = None
    for c in company_list:
        if c['company_name'] == selected_company_name:
            selected_company = c
            break

    # ===== 匹配模板 =====
    if selected_company and report_type:
        st.markdown("---")
        st.subheader("🔍 匹配结果")

        # 匹配逻辑：区级→市级→省级
        matched = None
        for t in TEMPLATES:
            if t['province'] == province and t['city'] == city and t['district'] == district and t['report_type'] == report_type:
                matched = t
                break
        if not matched:
            for t in TEMPLATES:
                if t['province'] == province and t['city'] == city and t['report_type'] == report_type:
                    matched = t
                    break
        if not matched:
            for t in TEMPLATES:
                if t['province'] == province and t['report_type'] == report_type:
                    matched = t
                    break

        if matched:
            st.success("✅ 已匹配到官方模板")
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**📄 模板信息**")
                st.write(f"模板名称：{matched['template_name']}")
                st.write(f"版本：{matched['template_version']}")
                st.write(f"发布机构：{matched['source_authority']}")
                st.write(f"发布日期：{matched['publish_date']}")
                st.write(f"必填字段：{matched['required_fields']}")
            with col_b:
                st.markdown("**🔗 来源信息**")
                st.write(f"来源URL：[{matched['source_url']}]({matched['source_url']})")
                st.write(f"适用地区：{matched['province']} {matched['city']} {matched['district']}")
        else:
            st.warning("⚠️ 未匹配到官方模板，将使用通用模板")
            matched = {
                'template_name': f'{report_type}通用申报表',
                'template_version': 'v1.0',
                'source_authority': '系统通用',
                'publish_date': datetime.now().strftime('%Y-%m-%d'),
                'required_fields': '纳税人识别号,公司名称,申报金额',
                'source_url': '#'
            }

        # ===== 展示导入的数据（如果有） =====
        if st.session_state.cities_data is not None:
            st.subheader("📊 已导入数据预览（用于填充报表）")
            st.dataframe(st.session_state.cities_data.head(5))
            st.caption(f"共 {len(st.session_state.cities_data)} 行数据")

        # ===== 复核复选框 =====
        st.markdown("---")
        reviewed = st.checkbox("✅ 我已人工复核确认数据无误", value=False)

        if st.button("📥 生成待复核版Excel", disabled=not reviewed):
            # 生成Excel
            wb = Workbook()
            ws = wb.active
            ws.title = "申报表"

            fields = matched['required_fields'].split(',')
            ws.append(fields)

            # 尝试从导入数据中获取填充值
            if st.session_state.cities_data is not None and not st.session_state.cities_data.empty:
                # 如果有数据，尝试按公司名称筛选
                df_data = st.session_state.cities_data
                # 筛选公司（如果存在公司列）
                company_col = None
                for col in df_data.columns:
                    if '公司' in str(col):
                        company_col = col
                        break
                if company_col:
                    df_comp = df_data[df_data[company_col] == selected_company_name]
                else:
                    df_comp = df_data
                if not df_comp.empty:
                    # 使用第一行填充
                    first_row = df_comp.iloc[0]
                    row_data = []
                    for f in fields:
                        matched_col = None
                        for col in df_comp.columns:
                            if f in str(col) or str(col) in f:
                                matched_col = col
                                break
                        if matched_col:
                            row_data.append(first_row[matched_col])
                        else:
                            # 默认值
                            if '纳税人识别号' in f:
                                row_data.append(selected_company.get('tax_id', ''))
                            elif '公司名称' in f or '单位名称' in f:
                                row_data.append(selected_company_name)
                            else:
                                row_data.append('')
                    ws.append(row_data)
                else:
                    # 无匹配数据，用空值
                    ws.append([''] * len(fields))
            else:
                # 使用示例数据
                sample_data = {
                    '纳税人识别号': selected_company.get('tax_id', ''),
                    '公司名称': selected_company_name,
                    '销售额': '100,000.00',
                    '进项税额': '13,000.00',
                    '应纳税额': '0.00',
                    '单位名称': selected_company_name,
                    '社保登记号': 'SH123456',
                    '基数': '8,000.00',
                    '单位金额': '1,280.00',
                    '个人金额': '640.00'
                }
                row_data = [sample_data.get(f, '') for f in fields]
                ws.append(row_data)

            # ===== 添加水印 =====
            ws.insert_rows(1)
            ws['A1'] = '【系统生成 - 待复核版】'
            ws['A1'].font = Font(color='FF0000', bold=True, size=14)
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(fields))
            ws['A1'].alignment = Alignment(horizontal='center')

            # ===== 模板版本信息 =====
            ws.insert_rows(2)
            ws['A2'] = f'模板名称：{matched["template_name"]}  版本：{matched["template_version"]}'
            ws['A2'].font = Font(color='666666', size=10)
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(fields))

            # ===== 来源信息 =====
            ws.insert_rows(3)
            ws['A3'] = f'来源：{matched["source_authority"]}  发布日期：{matched["publish_date"]}'
            ws['A3'].font = Font(color='666666', size=10)
            ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(fields))

            # ===== 统计口径信息 =====
            ws.insert_rows(4)
            ws['A4'] = f'统计口径：{period_type}'
            ws['A4'].font = Font(color='666666', size=10)
            ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=len(fields))

            # ===== 审计日志 =====
            audit = wb.create_sheet("审计日志")
            audit.append(['操作时间', '操作类型', '操作人', '详情'])
            audit.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 'GENERATED', '系统', f'公司:{selected_company_name}, 模板:{matched["template_name"]}, 统计口径:{period_type}'])

            output = BytesIO()
            wb.save(output)
            output.seek(0)

            # 记录历史
            export_record = {
                'id': str(uuid.uuid4())[:8],
                'company': selected_company_name,
                'template': matched['template_name'],
                'period': period_type,
                'generated_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'status': '待复核'
            }
            st.session_state.export_history.append(export_record)

            st.success("✅ Excel已生成！")
            st.download_button(
                label="📥 下载待复核版",
                data=output,
                file_name=f"{selected_company_name}_{report_type}_{period_type.replace('（','_').replace('）','')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.info("📌 该文件为【待复核版】，正式使用前请完成人工复核流程。")

    else:
        if not selected_company:
            st.info("👆 请先选择公司")
        elif not report_type:
            st.info("👆 请选择报表类型")

# ===== 历史记录 =====
with st.expander("📋 导出历史记录"):
    if st.session_state.export_history:
        st.dataframe(pd.DataFrame(st.session_state.export_history))
    else:
        st.info("暂无导出记录")

# ===== 查看知识库 =====
with st.expander("📚 官方模板知识库"):
    st.dataframe(pd.DataFrame(TEMPLATES))

# ===== 显示当前公司列表（便于调试） =====
with st.expander("🏢 当前公司列表（自动提取）"):
    if st.session_state.companies:
        st.dataframe(pd.DataFrame(st.session_state.companies))
    else:
        st.info("暂无公司数据，请上传Excel")
