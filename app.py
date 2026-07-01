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
    st.session_state.companies = []  # 将从Excel自动填充
if 'export_history' not in st.session_state:
    st.session_state.export_history = []
if 'imported_data' not in st.session_state:
    st.session_state.imported_data = None
if 'uploaded_sheets' not in st.session_state:
    st.session_state.uploaded_sheets = {}
if 'all_cities' not in st.session_state:
    st.session_state.all_cities = set()

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
    {
        'id': 't003',
        'province': '广东',
        'city': '广州市',
        'district': '天河区',
        'report_type': '增值税',
        'template_name': '广东省增值税纳税申报表',
        'template_version': 'v2024.1',
        'source_url': 'https://guangdong.chinatax.gov.cn/bsfw/2024/zzs.xlsx',
        'source_authority': '国家税务总局广东省税务局',
        'publish_date': '2024-01-20',
        'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额'
    },
    {
        'id': 't004',
        'province': '北京',
        'city': '北京市',
        'district': '海淀区',
        'report_type': '增值税',
        'template_name': '北京市增值税纳税申报表（一般纳税人）',
        'template_version': 'v2024.2',
        'source_url': 'https://beijing.chinatax.gov.cn/bsfw/2024/zzs.xlsx',
        'source_authority': '国家税务总局北京市税务局',
        'publish_date': '2024-02-01',
        'required_fields': '纳税人识别号,公司名称,销售额,进项税额,应纳税额'
    },
]

# ========== 解析上传的Excel，自动提取公司/城市 ==========
def parse_uploaded_excel(file):
    """解析上传的Excel，提取所有公司和城市信息"""
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    all_companies = []
    all_cities = set()
    
    # 遍历所有Sheet，尝试提取数据
    for sheet in sheets:
        try:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            # 尝试识别表头行
            header_row = None
            for i, row in df.iterrows():
                row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                if '所属城市' in row_text or '城市' in row_text or '公司' in row_text:
                    header_row = i
                    break
            if header_row is not None:
                df = pd.read_excel(file, sheet_name=sheet, skiprows=header_row)
                df.columns = [str(c).strip() for c in df.columns]
                
                # 查找关键列
                city_col = None
                company_col = None
                district_col = None
                for col in df.columns:
                    col_lower = col.lower()
                    if '所属城市' in col_lower or '城市' in col_lower:
                        city_col = col
                    elif '分公司' in col_lower or '公司' in col_lower:
                        company_col = col
                    elif '区县' in col_lower or '区' in col_lower:
                        district_col = col
                
                if city_col and company_col:
                    # 提取唯一的公司-城市-区县组合
                    for _, row in df.iterrows():
                        city = str(row[city_col]) if pd.notna(row[city_col]) else ''
                        company = str(row[company_col]) if pd.notna(row[company_col]) else ''
                        district = str(row[district_col]) if district_col and pd.notna(row[district_col]) else ''
                        if city and company:
                            all_cities.add(city)
                            all_companies.append({
                                'company_name': company,
                                'province': city,  # 省份与城市名相同（适用于直辖市/省城）
                                'city': city,
                                'district': district,
                                'tax_id': ''
                            })
        except:
            continue
    
    # 去重
    unique_companies = []
    seen = set()
    for c in all_companies:
        key = (c['company_name'], c['city'])
        if key not in seen:
            seen.add(key)
            unique_companies.append(c)
    
    return unique_companies, list(all_cities)

# ========== 页面 ==========
st.set_page_config(page_title="官方模板匹配器", layout="wide")
st.title("📋 官方模板匹配器（含自动识别与统计口径）")
st.markdown("**上传Excel → 自动提取城市/公司 → 选择模板和统计口径 → 生成待复核版Excel**")

# ===== 侧边栏：上传Excel =====
with st.sidebar:
    st.header("📤 上传数据Excel")
    st.markdown("上传包含公司/城市信息的Excel，系统自动提取所有地区")
    uploaded_file = st.file_uploader("选择Excel文件（.xlsx）", type=["xlsx"], key="main_upload")
    
    if uploaded_file:
        # 解析并提取公司/城市
        with st.spinner("正在解析Excel并提取地区信息..."):
            companies, cities = parse_uploaded_excel(uploaded_file)
            if companies:
                # 更新session_state
                st.session_state.companies = companies
                st.session_state.all_cities = set(cities)
                st.success(f"成功提取 {len(companies)} 家公司，{len(cities)} 个城市")
                # 同时存储数据用于填充模板
                # 读取数据（用于报表填充）
                try:
                    # 尝试读取第一个数据Sheet
                    xls = pd.ExcelFile(uploaded_file)
                    # 优先选择“月度明细数据表”
                    data_sheet = None
                    for s in xls.sheet_names:
                        if '明细' in s or '月度' in s or '数据' in s:
                            data_sheet = s
                            break
                    if data_sheet:
                        df_data = pd.read_excel(uploaded_file, sheet_name=data_sheet)
                        st.session_state.imported_data = df_data
                        st.sidebar.success(f"成功读取Sheet「{data_sheet}」，共{len(df_data)}行")
                    else:
                        st.sidebar.warning("未找到数据Sheet，请手动选择")
                except Exception as e:
                    st.sidebar.error(f"读取数据失败：{str(e)}")
            else:
                st.sidebar.warning("未自动识别到公司列和城市列，请确认数据格式")
    
    # 显示当前公司列表
    with st.sidebar.expander("🏢 当前公司列表（自动提取）"):
        if st.session_state.companies:
            st.dataframe(pd.DataFrame(st.session_state.companies))
            st.caption(f"共 {len(st.session_state.companies)} 家公司")
        else:
            st.info("暂无公司数据，请上传Excel")

# ===== 主体：筛选 =====
companies = st.session_state.companies
if not companies:
    st.info("👈 请先在侧边栏上传包含公司/城市数据的Excel")
    st.stop()

# 获取所有省份/城市/区县
provinces = sorted(set(c['province'] for c in companies if c['province']))
all_cities = sorted(set(c['city'] for c in companies if c['city']))

col1, col2, col3 = st.columns(3)
with col1:
    province = st.selectbox("省份", [""] + provinces)
    # 根据省份过滤城市
    cities = sorted(set(c['city'] for c in companies if c['province'] == province)) if province else all_cities
    city = st.selectbox("城市", [""] + cities)
with col2:
    districts = sorted(set(c['district'] for c in companies if c['province'] == province and c['city'] == city)) if province and city else []
    district = st.selectbox("区县", [""] + districts)
    company_list = [c for c in companies if c['province'] == province and c['city'] == city and (not district or c['district'] == district)]
    company_names = [c['company_name'] for c in company_list]
    selected_company_name = st.selectbox("公司", [""] + company_names)
with col3:
    report_type = st.selectbox("报表类型", ["", "增值税", "社保", "公积金", "企业所得税", "个人所得税"])
    # 统计口径
    period_type = st.selectbox("统计口径", ["月度（12月单月）", "累计（1-12月）"])

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

    # ===== 展示导入的数据 =====
    if st.session_state.imported_data is not None:
        st.subheader("📊 已导入数据预览")
        st.dataframe(st.session_state.imported_data.head(5))
        st.caption(f"共 {len(st.session_state.imported_data)} 行数据，统计口径：{period_type}")

    # ===== 复核复选框 =====
    reviewed = st.checkbox("✅ 我已人工复核确认数据无误", value=False)

    if st.button("📥 生成待复核版Excel", disabled=not reviewed):
        # 生成Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "申报表"

        # 填充数据
        fields = matched['required_fields'].split(',')
        ws.append(fields)

        if st.session_state.imported_data is not None and not st.session_state.imported_data.empty:
            # 使用导入数据的第一行填充
            first_row = st.session_state.imported_data.iloc[0]
            row_data = []
            for f in fields:
                matched_col = None
                for col in st.session_state.imported_data.columns:
                    if f in str(col) or str(col) in f:
                        matched_col = col
                        break
                if matched_col:
                    row_data.append(first_row[matched_col])
                else:
                    if '纳税人识别号' in f:
                        row_data.append(selected_company.get('tax_id', ''))
                    elif '公司名称' in f or '单位名称' in f:
                        row_data.append(selected_company['company_name'])
                    else:
                        row_data.append('')
            ws.append(row_data)
        else:
            # 示例数据
            sample_data = {
                '纳税人识别号': selected_company.get('tax_id', ''),
                '公司名称': selected_company['company_name'],
                '销售额': '100,000.00',
                '进项税额': '13,000.00',
                '应纳税额': '0.00',
                '单位名称': selected_company['company_name'],
                '社保登记号': 'SH123456',
                '基数': '8,000.00',
                '单位金额': '1,280.00',
                '个人金额': '640.00'
            }
            row_data = [sample_data.get(f, '') for f in fields]
            ws.append(row_data)

        # 水印
        ws.insert_rows(1)
        ws['A1'] = f'【系统生成 - 待复核版】统计口径：{period_type}'
        ws['A1'].font = Font(color='FF0000', bold=True, size=14)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(fields))
        ws['A1'].alignment = Alignment(horizontal='center')

        # 模板版本
        ws.insert_rows(2)
        ws['A2'] = f'模板名称：{matched["template_name"]}  版本：{matched["template_version"]}  统计口径：{period_type}'
        ws['A2'].font = Font(color='666666', size=10)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(fields))

        # 来源信息
        ws.insert_rows(3)
        ws['A3'] = f'来源：{matched["source_authority"]}  发布日期：{matched["publish_date"]}'
        ws['A3'].font = Font(color='666666', size=10)
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(fields))

        # 审计日志
        audit = wb.create_sheet("审计日志")
        audit.append(['操作时间', '操作类型', '操作人', '详情'])
        audit.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 'GENERATED', '系统', f'公司:{selected_company["company_name"]}, 模板:{matched["template_name"]}, 口径:{period_type}'])

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # 记录历史
        export_record = {
            'id': str(uuid.uuid4())[:8],
            'company': selected_company['company_name'],
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
            file_name=f"{selected_company['company_name']}_{report_type}_{period_type.replace('（','_').replace('）','')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.info("📌 该文件为【待复核版】，正式使用前请完成人工复核流程。")

else:
    if not selected_company:
        st.info("👆 请先选择公司")
    elif not report_type:
        st.info("👆 请选择报表类型")

# ===== 导出历史 =====
with st.expander("📋 导出历史记录"):
    if st.session_state.export_history:
        st.dataframe(pd.DataFrame(st.session_state.export_history))
    else:
        st.info("暂无导出记录")

# ===== 查看知识库 =====
with st.expander("📚 官方模板知识库"):
    st.dataframe(pd.DataFrame(TEMPLATES))
