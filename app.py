import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO
import uuid

# ========== 初始化会话状态 ==========
if 'companies' not in st.session_state:
    st.session_state.companies = [
        {'id': 'c001', 'company_name': '上海科技发展有限公司', 'province': '上海', 'city': '上海市', 'district': '浦东新区', 'tax_id': '91310115MA1KXXXXX'},
        {'id': 'c002', 'company_name': '上海商贸有限公司', 'province': '上海', 'city': '上海市', 'district': '黄浦区', 'tax_id': '91310101MA1HXXXXX'},
        {'id': 'c003', 'company_name': '广州创新科技有限公司', 'province': '广东', 'city': '广州市', 'district': '天河区', 'tax_id': '91440106MA1YXXXXX'},
        {'id': 'c004', 'company_name': '北京智汇科技有限公司', 'province': '北京', 'city': '北京市', 'district': '海淀区', 'tax_id': '91110108MA01XXXXX'},
    ]

if 'export_history' not in st.session_state:
    st.session_state.export_history = []

if 'imported_data' not in st.session_state:
    st.session_state.imported_data = None

# ========== 官方模板库 ==========
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
        'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额',
        'period_type': '月度'  # 月度/年度/累计
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
        'required_fields': '单位名称,社保登记号,基数,单位金额,个人金额',
        'period_type': '月度'
    },
    {
        'id': 't003',
        'province': '上海',
        'city': '上海市',
        'district': '黄浦区',
        'report_type': '增值税',
        'template_name': '上海市增值税纳税申报表（小规模纳税人）',
        'template_version': 'v2024.1',
        'source_url': 'https://shanghai.chinatax.gov.cn/bsfw/bszn/2024/zzs_small.xlsx',
        'source_authority': '国家税务总局上海市税务局',
        'publish_date': '2024-01-15',
        'required_fields': '纳税人识别号,公司名称,销售额,应纳税额',
        'period_type': '月度'
    },
    {
        'id': 't004',
        'province': '广东',
        'city': '广州市',
        'district': '天河区',
        'report_type': '增值税',
        'template_name': '广东省增值税纳税申报表',
        'template_version': 'v2024.1',
        'source_url': 'https://guangdong.chinatax.gov.cn/bsfw/2024/zzs.xlsx',
        'source_authority': '国家税务总局广东省税务局',
        'publish_date': '2024-01-20',
        'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额',
        'period_type': '月度'
    },
    {
        'id': 't005',
        'province': '北京',
        'city': '北京市',
        'district': '海淀区',
        'report_type': '增值税',
        'template_name': '北京市增值税纳税申报表（一般纳税人）',
        'template_version': 'v2024.2',
        'source_url': 'https://beijing.chinatax.gov.cn/bsfw/2024/zzs.xlsx',
        'source_authority': '国家税务总局北京市税务局',
        'publish_date': '2024-02-01',
        'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额',
        'period_type': '月度'
    },
    # 示例年度报表（如社保年度汇算）
    {
        'id': 't006',
        'province': '上海',
        'city': '上海市',
        'district': '浦东新区',
        'report_type': '社保',
        'template_name': '上海市社会保险费年度汇算清缴表',
        'template_version': 'v2024.1',
        'source_url': 'https://rsj.sh.gov.cn/sbjb/2024/annual.xlsx',
        'source_authority': '上海市人力资源和社会保障局',
        'publish_date': '2024-01-10',
        'required_fields': '单位名称,社保登记号,全年基数合计,全年单位合计,全年个人合计',
        'period_type': '年度'
    },
]

# ========== 页面配置 ==========
st.set_page_config(page_title="官方模板匹配器", layout="wide")
st.title("📋 官方模板匹配器（含统计口径）")
st.markdown("**上传公司 → 上传数据 → 选择模板 → 选择统计口径 → 生成待复核版Excel**")

# ===== 侧边栏：导入公司数据 =====
with st.sidebar:
    st.header("📤 导入公司数据")
    uploaded_company = st.file_uploader("选择公司Excel (.xlsx)", type=["xlsx"], key="company_upload")
    if uploaded_company:
        try:
            df_comp = pd.read_excel(uploaded_company)
            col_map = {}
            for col in df_comp.columns:
                col_lower = str(col).lower()
                if '公司' in col_lower or '企业' in col_lower:
                    col_map['company'] = col
                elif '省份' in col_lower or '省' in col_lower:
                    col_map['province'] = col
                elif '城市' in col_lower or '市' in col_lower:
                    col_map['city'] = col
                elif '区县' in col_lower or '区' in col_lower:
                    col_map['district'] = col
                elif '税号' in col_lower or '信用代码' in col_lower:
                    col_map['tax_id'] = col
            if 'company' not in col_map or 'province' not in col_map or 'city' not in col_map:
                st.error("未识别到必要列（公司名称、省份、城市）")
            else:
                new_companies = []
                for _, row in df_comp.iterrows():
                    new_companies.append({
                        'id': str(uuid.uuid4())[:8],
                        'company_name': str(row[col_map['company']]),
                        'province': str(row[col_map['province']]),
                        'city': str(row[col_map['city']]),
                        'district': str(row[col_map['district']]) if col_map.get('district') else '',
                        'tax_id': str(row[col_map['tax_id']]) if col_map.get('tax_id') else ''
                    })
                existing_names = {c['company_name'] for c in st.session_state.companies}
                for c in new_companies:
                    if c['company_name'] not in existing_names:
                        st.session_state.companies.append(c)
                        existing_names.add(c['company_name'])
                st.success(f"成功导入 {len(new_companies)} 家公司")
                st.dataframe(pd.DataFrame(new_companies))
        except Exception as e:
            st.error(f"导入失败：{str(e)}")

    st.header("📊 导入数据Excel")
    uploaded_data = st.sidebar.file_uploader("选择数据Excel", type=["xlsx"], key="data_upload")
    if uploaded_data:
        try:
            df_data = pd.read_excel(uploaded_data)
            st.sidebar.success(f"读取到 {len(df_data)} 行数据")
            st.session_state.imported_data = df_data
        except Exception as e:
            st.sidebar.error(f"读取失败：{str(e)}")

# ===== 主体 =====
companies = st.session_state.companies
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
    period_options = ["月度", "年度", "累计（1-12月）", "单月（本期）"]
    selected_period = st.selectbox("统计口径", period_options)

# 获取选中公司
selected_company = None
for c in company_list:
    if c['company_name'] == selected_company_name:
        selected_company = c
        break

# ===== 匹配模板 =====
if selected_company and report_type and selected_period:
    st.markdown("---")
    st.subheader("🔍 匹配结果")

    # 匹配逻辑：地区 + 报表类型 + 统计口径
    period_map = {
        "月度": "月度",
        "年度": "年度",
        "累计（1-12月）": "年度",
        "单月（本期）": "月度"
    }
    matched_period = period_map.get(selected_period, "月度")

    matched = None
    for t in TEMPLATES:
        if t['province'] == province and t['city'] == city and t['district'] == district and t['report_type'] == report_type and t.get('period_type') == matched_period:
            matched = t
            break
    if not matched:
        for t in TEMPLATES:
            if t['province'] == province and t['city'] == city and t['report_type'] == report_type and t.get('period_type') == matched_period:
                matched = t
                break
    if not matched:
        for t in TEMPLATES:
            if t['province'] == province and t['report_type'] == report_type and t.get('period_type') == matched_period:
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
            st.write(f"统计口径：{matched.get('period_type', '未指定')}")
        with col_b:
            st.markdown("**🔗 来源信息**")
            st.write(f"来源URL：[{matched['source_url']}]({matched['source_url']})")
            st.write(f"适用地区：{matched['province']} {matched['city']} {matched['district']}")
    else:
        st.warning("⚠️ 未匹配到官方模板，将使用通用模板")
        matched = {
            'template_name': f'{report_type}通用申报表（{selected_period}）',
            'template_version': 'v1.0',
            'source_authority': '系统通用',
            'publish_date': datetime.now().strftime('%Y-%m-%d'),
            'required_fields': '纳税人识别号,公司名称,申报金额',
            'source_url': '#',
            'period_type': matched_period
        }

    # ===== 导入数据预览 =====
    if st.session_state.imported_data is not None:
        st.subheader("📊 已导入数据预览")
        st.dataframe(st.session_state.imported_data.head(5))
        st.caption(f"共 {len(st.session_state.imported_data)} 行数据")

    # ===== 复核 =====
    st.markdown("---")
    reviewed = st.checkbox("✅ 我已人工复核确认数据无误", value=False)

    if st.button("📥 生成待复核版Excel", disabled=not reviewed):
        wb = Workbook()
        ws = wb.active
        ws.title = "申报表"

        fields = matched['required_fields'].split(',')
        ws.append(fields)

        # 填充数据
        if st.session_state.imported_data is not None and not st.session_state.imported_data.empty:
            # 根据统计口径聚合数据
            df = st.session_state.imported_data
            # 尝试识别日期列
            date_col = None
            for col in df.columns:
                if '月' in str(col) or '日期' in str(col) or '期间' in str(col):
                    date_col = col
                    break
            if date_col and selected_period in ["累计（1-12月）", "年度"]:
                # 全年累计：sum所有数据
                agg_data = df.sum(numeric_only=True).to_dict()
            elif date_col and selected_period in ["单月（本期）", "月度"]:
                # 单月：取最新月份或用户选择
                # 简化：取第一行
                agg_data = df.iloc[0].to_dict()
            else:
                # 默认取第一行
                agg_data = df.iloc[0].to_dict()

            # 生成行数据
            row_data = []
            for f in fields:
                matched_col = None
                for col in df.columns:
                    if f in str(col) or str(col) in f:
                        matched_col = col
                        break
                if matched_col:
                    row_data.append(agg_data.get(matched_col, ''))
                else:
                    # 默认值
                    if '纳税人识别号' in f:
                        row_data.append(selected_company['tax_id'])
                    elif '公司名称' in f or '单位名称' in f:
                        row_data.append(selected_company['company_name'])
                    else:
                        row_data.append('')
            ws.append(row_data)
        else:
            # 示例数据
            sample_data = {
                '纳税人识别号': selected_company['tax_id'],
                '公司名称': selected_company['company_name'],
                '销售额': '100,000.00',
                '进项税额': '13,000.00',
                '应纳税额': '0.00',
                '单位名称': selected_company['company_name'],
                '社保登记号': 'SH123456',
                '基数': '8,000.00',
                '单位金额': '1,280.00',
                '个人金额': '640.00',
                '全年基数合计': '96,000.00',
                '全年单位合计': '15,360.00',
                '全年个人合计': '7,680.00'
            }
            row_data = [sample_data.get(f, '') for f in fields]
            ws.append(row_data)

        # 水印
        ws.insert_rows(1)
        ws['A1'] = '【系统生成 - 待复核版】'
        ws['A1'].font = Font(color='FF0000', bold=True, size=14)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(fields))
        ws['A1'].alignment = Alignment(horizontal='center')

        # 版本
        ws.insert_rows(2)
        ws['A2'] = f'模板名称：{matched["template_name"]}  版本：{matched["template_version"]}  统计口径：{selected_period}'
        ws['A2'].font = Font(color='666666', size=10)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(fields))

        # 来源
        ws.insert_rows(3)
        ws['A3'] = f'来源：{matched["source_authority"]}  发布日期：{matched["publish_date"]}'
        ws['A3'].font = Font(color='666666', size=10)
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(fields))

        # 审计
        audit = wb.create_sheet("审计日志")
        audit.append(['操作时间', '操作类型', '操作人', '详情'])
        audit.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 'GENERATED', '系统', f'公司:{selected_company["company_name"]}, 模板:{matched["template_name"]}, 口径:{selected_period}'])

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        export_record = {
            'id': str(uuid.uuid4())[:8],
            'company': selected_company['company_name'],
            'template': matched['template_name'],
            'period': selected_period,
            'generated_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'status': '待复核'
        }
        st.session_state.export_history.append(export_record)

        st.success("✅ Excel已生成！")
        st.download_button(
            label="📥 下载待复核版",
            data=output,
            file_name=f"{selected_company['company_name']}_{report_type}_{selected_period}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.info("📌 该文件为【待复核版】，正式使用前请完成人工复核流程。")

else:
    if not selected_company:
        st.info("👆 请先选择公司")
    elif not report_type:
        st.info("👆 请选择报表类型")
    elif not selected_period:
        st.info("👆 请选择统计口径")

# ===== 历史 =====
with st.expander("📋 导出历史记录"):
    if st.session_state.export_history:
        st.dataframe(pd.DataFrame(st.session_state.export_history))
    else:
        st.info("暂无导出记录")

# ===== 知识库 =====
with st.expander("📚 官方模板知识库"):
    st.dataframe(pd.DataFrame(TEMPLATES))

# ===== 公司列表 =====
with st.expander("🏢 当前公司列表"):
    st.dataframe(pd.DataFrame(st.session_state.companies))
